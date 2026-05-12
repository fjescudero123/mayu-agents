param(
  [ValidateSet("poll", "single_file", "test")]
  [string]$Mode = "poll",
  [string]$FileName = "",
  [int]$ProcessLimit = 10,
  [bool]$SendEmail = $true
)

$ErrorActionPreference = "Stop"
$VerbosePreference = "SilentlyContinue"

function Get-RunbookVariable {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Name,
    [bool]$Required = $true,
    [string]$Default = ""
  )

  $value = [Environment]::GetEnvironmentVariable($Name)
  if ([string]::IsNullOrWhiteSpace([string]$value)) {
    $value = $Default
  }
  if ($Required -and [string]::IsNullOrWhiteSpace([string]$value)) {
    throw "Falta variable requerida: $Name"
  }
  return [string]$value
}

function ConvertTo-DrivePath {
  param([string]$Path)
  if ([string]::IsNullOrWhiteSpace($Path)) {
    return ""
  }
  (($Path.Trim("/") -split "/") | Where-Object { $_ } | ForEach-Object { [Uri]::EscapeDataString($_) }) -join "/"
}

function HtmlEscape {
  param([object]$Value)
  [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Invoke-GraphGet {
  param([string]$Token, [string]$Uri)
  Invoke-RestMethod -Method Get -Uri $Uri -Headers @{ Authorization = "Bearer $Token" }
}

function Invoke-GraphJson {
  param([string]$Token, [string]$Method, [string]$Uri, [object]$Body)
  $json = $Body | ConvertTo-Json -Depth 50
  Invoke-RestMethod -Method $Method -Uri $Uri -Headers @{ Authorization = "Bearer $Token" } -ContentType "application/json" -Body $json
}

function Invoke-GraphPutBytes {
  param([string]$Token, [string]$Uri, [byte[]]$Bytes, [string]$ContentType)
  Invoke-RestMethod -Method Put -Uri $Uri -Headers @{ Authorization = "Bearer $Token" } -ContentType $ContentType -Body $Bytes
}

function Get-GraphToken {
  $tenantId = Get-RunbookVariable "MayuTenantId"
  $clientId = Get-RunbookVariable "MayuClientId"
  $clientSecret = Get-RunbookVariable "MayuClientSecret"

  $token = Invoke-RestMethod `
    -Method Post `
    -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" `
    -ContentType "application/x-www-form-urlencoded" `
    -Body @{
      client_id = $clientId
      client_secret = $clientSecret
      scope = "https://graph.microsoft.com/.default"
      grant_type = "client_credentials"
    }

  return $token.access_token
}

function Get-SiteId {
  param([string]$Token, [string]$HostName)
  (Invoke-GraphGet -Token $Token -Uri "https://graph.microsoft.com/v1.0/sites/$HostName").id
}

function Get-DriveChildren {
  param([string]$Token, [string]$SiteId, [string]$FolderPath)
  $items = @()
  $encoded = ConvertTo-DrivePath $FolderPath
  $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/root:/$encoded" + ":/children?`$top=200"

  while ($uri) {
    $data = Invoke-GraphGet -Token $Token -Uri $uri
    if ($null -ne $data.value) {
      $items += @($data.value)
    }
    $next = $null
    if ($null -ne $data.PSObject.Properties["@odata.nextLink"]) {
      $next = [string]$data.PSObject.Properties["@odata.nextLink"].Value
    }
    if ([string]::IsNullOrWhiteSpace($next)) {
      $uri = $null
    } else {
      $uri = $next
    }
  }

  foreach ($item in @($items)) {
    if ($null -eq $item) { continue }
    $item | Add-Member -NotePropertyName "_mayuRelativePath" -NotePropertyValue "$FolderPath/$($item.name)" -Force
  }
  return $items
}

function Ensure-GraphFolder {
  param([string]$Token, [string]$SiteId, [string]$FolderPath)
  if ([string]::IsNullOrWhiteSpace($FolderPath)) {
    return
  }

  $parts = @($FolderPath.Trim("/") -split "/" | Where-Object { $_ })
  $current = ""
  foreach ($part in $parts) {
    $candidate = if ($current) { "$current/$part" } else { $part }
    $encoded = ConvertTo-DrivePath $candidate
    $exists = $true
    try {
      Invoke-GraphGet -Token $Token -Uri "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/root:/$encoded" | Out-Null
    } catch {
      $exists = $false
    }

    if (-not $exists) {
      $parentEncoded = ConvertTo-DrivePath $current
      if ($parentEncoded) {
        $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/root:/$parentEncoded" + ":/children"
      } else {
        $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/root/children"
      }
      Invoke-GraphJson -Token $Token -Method Post -Uri $uri -Body @{
        name = $part
        folder = @{}
        "@microsoft.graph.conflictBehavior" = "fail"
      } | Out-Null
    }
    $current = $candidate
  }
}

function Get-FileBytes {
  param([string]$Token, [string]$SiteId, [string]$FilePath)
  $encoded = ConvertTo-DrivePath $FilePath
  $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/root:/$encoded" + ":/content"
  $tempFile = Join-Path $env:TEMP ("mayu_minuta_" + [Guid]::NewGuid().ToString("N") + ".bin")
  try {
    Invoke-WebRequest -Method Get -Uri $uri -Headers @{ Authorization = "Bearer $Token" } -OutFile $tempFile -UseBasicParsing
    return [System.IO.File]::ReadAllBytes($tempFile)
  } finally {
    if (Test-Path $tempFile) {
      Remove-Item $tempFile -Force
    }
  }
}

function Write-TextFileToGraph {
  param([string]$Token, [string]$SiteId, [string]$FilePath, [string]$Text, [string]$ContentType)
  $encoded = ConvertTo-DrivePath $FilePath
  $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/root:/$encoded" + ":/content"
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
  Invoke-GraphPutBytes -Token $Token -Uri $uri -Bytes $bytes -ContentType $ContentType | Out-Null
}

function Send-GraphMail {
  param(
    [string]$Token,
    [string]$Sender,
    [string[]]$To,
    [string]$Subject,
    [string]$HtmlBody,
    [string]$AttachmentName = "",
    [byte[]]$AttachmentBytes = $null
  )

  $attachments = @()
  if ($AttachmentName -and $AttachmentBytes) {
    $attachments += @{
      "@odata.type" = "#microsoft.graph.fileAttachment"
      name = $AttachmentName
      contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      contentBytes = [Convert]::ToBase64String($AttachmentBytes)
    }
  }

  $message = @{
    subject = $Subject
    body = @{ contentType = "HTML"; content = $HtmlBody }
    toRecipients = @($To | Where-Object { $_ } | ForEach-Object { @{ emailAddress = @{ address = $_ } } })
    attachments = $attachments
  }

  Invoke-GraphJson -Token $Token -Method Post -Uri "https://graph.microsoft.com/v1.0/users/$Sender/sendMail" -Body @{
    message = $message
    saveToSentItems = $true
  } | Out-Null
}

function Move-GraphItemToFolder {
  param(
    [string]$Token,
    [string]$SiteId,
    [string]$ItemId,
    [string]$DestinationFolderPath,
    [string]$NewName
  )

  $destEncoded = ConvertTo-DrivePath $DestinationFolderPath
  $destFolder = Invoke-GraphGet -Token $Token -Uri "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/root:/$destEncoded"
  $body = @{
    parentReference = @{
      id = [string]$destFolder.id
    }
  }
  if (-not [string]::IsNullOrWhiteSpace($NewName)) {
    $body.name = $NewName
  }

  Invoke-GraphJson -Token $Token -Method Patch -Uri "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/items/$ItemId" -Body $body | Out-Null
}

function Convert-DocxBytesToText {
  param([byte[]]$Bytes)
  Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
  $tempRoot = Join-Path $env:TEMP ("mayu_docx_" + [Guid]::NewGuid().ToString("N"))
  $tempDocx = Join-Path $env:TEMP ("mayu_docx_" + [Guid]::NewGuid().ToString("N") + ".docx")
  try {
    [System.IO.File]::WriteAllBytes($tempDocx, $Bytes)
    [System.IO.Compression.ZipFile]::ExtractToDirectory($tempDocx, $tempRoot)
    $documentPath = Join-Path $tempRoot "word/document.xml"
    if (-not (Test-Path $documentPath)) {
      return ""
    }

    $xmlText = Get-Content -Raw -Path $documentPath
    $xmlText = $xmlText -replace "</w:tr>", "`n"
    $xmlText = $xmlText -replace "</w:p>", "`n"
    $xmlText = $xmlText -replace "</w:tc>", "`t"
    $xmlText = $xmlText -replace "<w:tab[^>]*/>", "`t"
    $xmlText = [regex]::Replace($xmlText, "<[^>]+>", "")
    $xmlText = [System.Net.WebUtility]::HtmlDecode($xmlText)
    $lines = @()
    foreach ($line in ($xmlText -split "(`r`n|`n|`r)")) {
      $clean = [regex]::Replace($line, "[ \t]+", " ").Trim()
      if ($clean) {
        $lines += $clean
      }
    }
    return ($lines -join "`n").Trim()
  } catch {
    Write-Warning "No se pudo leer texto DOCX. $($_.Exception.Message)"
    return ""
  } finally {
    if (Test-Path $tempDocx) {
      Remove-Item $tempDocx -Force
    }
    if (Test-Path $tempRoot) {
      Remove-Item $tempRoot -Recurse -Force
    }
  }
}

function Invoke-OpenAiText {
  param([string]$Prompt)

  $apiKey = Get-RunbookVariable "MayuOpenAiApiKey" $false ""
  if ([string]::IsNullOrWhiteSpace($apiKey)) {
    return ""
  }
  $model = Get-RunbookVariable "MayuOpenAiModel" $false "gpt-5.4-mini"

  try {
    $body = @{
      model = $model
      input = $Prompt
      max_output_tokens = 2500
    } | ConvertTo-Json -Depth 10

    $response = Invoke-RestMethod `
      -Method Post `
      -Uri "https://api.openai.com/v1/responses" `
      -Headers @{ Authorization = "Bearer $apiKey" } `
      -ContentType "application/json" `
      -Body $body `
      -TimeoutSec 60

    if ($response.output_text) {
      return [string]$response.output_text
    }

    $texts = @()
    foreach ($output in @($response.output)) {
      foreach ($content in @($output.content)) {
        if ($content.text) {
          $texts += [string]$content.text
        }
      }
    }
    return ($texts -join "`n").Trim()
  } catch {
    Write-Warning "OpenAI no pudo generar la minuta HTML. Se usara fallback. $($_.Exception.Message)"
    return ""
  }
}

function ConvertFrom-ModelJson {
  param([string]$Text)
  if ([string]::IsNullOrWhiteSpace($Text)) {
    return $null
  }
  $candidate = $Text.Trim()
  $candidate = [regex]::Replace($candidate, "^```(?:json)?\s*", "")
  $candidate = [regex]::Replace($candidate, "\s*```$", "")
  $start = $candidate.IndexOf("{")
  $end = $candidate.LastIndexOf("}")
  if ($start -ge 0 -and $end -gt $start) {
    $candidate = $candidate.Substring($start, $end - $start + 1)
  }
  try {
    return ($candidate | ConvertFrom-Json)
  } catch {
    Write-Warning "OpenAI no devolvio JSON parseable para la minuta."
    return $null
  }
}

function Normalize-ModelHtml {
  param([string]$Text)
  if ([string]::IsNullOrWhiteSpace($Text)) {
    return ""
  }
  $candidate = $Text.Trim()
  $candidate = [regex]::Replace($candidate, "^```(?:html)?\s*", "")
  $candidate = [regex]::Replace($candidate, "\s*```$", "")
  return $candidate.Trim()
}

function Get-FallbackHtml {
  param(
    [string]$MeetingTitle,
    [string]$Date,
    [string]$SourceFileName,
    [string]$ExtractedText
  )

  $safeTitle = HtmlEscape $MeetingTitle
  $safeDate = HtmlEscape $Date
  $safeFile = HtmlEscape $SourceFileName
  $safeText = HtmlEscape $ExtractedText
  @"
<h1>$safeTitle</h1>
<p><strong>Fecha:</strong> $safeDate</p>
<p><strong>Archivo fuente:</strong> $safeFile</p>
<p>No se pudo estructurar la minuta con IA. Se envia el contenido extraido del AGENDA para no frenar la operacion.</p>
<pre style="white-space:pre-wrap;font-family:Consolas,monospace;background:#f6f6f6;padding:12px;border-radius:6px;">$safeText</pre>
"@
}

function Get-MeetingConfigMap {
  param([object]$Config)
  $map = @{}
  foreach ($entry in $Config.meetings.PSObject.Properties) {
    $type = [string]$entry.Name
    $meeting = $entry.Value
    $meeting | Add-Member -NotePropertyName tipo -NotePropertyValue $type -Force
    $map[$type.ToLowerInvariant()] = $meeting
    foreach ($label in @($meeting.labels)) {
      if (-not [string]::IsNullOrWhiteSpace([string]$label)) {
        $map[[string]$label.ToLowerInvariant()] = $meeting
      }
    }
  }
  return $map
}

function Parse-MinutaFileName {
  param([string]$Name, [hashtable]$MeetingMap)
  $m = [regex]::Match($Name, "^AGENDA_(.+)_(\d{2})_(\d{2})_(\d{4})\.docx$", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
  if (-not $m.Success) {
    return $null
  }
  $label = [string]$m.Groups[1].Value
  $normalized = $label.ToLowerInvariant()
  if (-not $MeetingMap.ContainsKey($normalized)) {
    return $null
  }
  $meeting = $MeetingMap[$normalized]
  $dateIso = "{0}-{1}-{2}" -f $m.Groups[4].Value, $m.Groups[3].Value, $m.Groups[2].Value
  [pscustomobject]@{
    FileName = $Name
    Label = $label
    Meeting = $meeting
    Tipo = [string]$meeting.tipo
    DateIso = $dateIso
  }
}

function Find-PendingAgendaItems {
  param([string]$Token, [string]$SiteId, [object]$Config, [string]$RequestedFileName = "", [int]$Limit = 10)
  $meetingMap = Get-MeetingConfigMap -Config $Config
  $inputFolder = [string]$Config.sharepoint.input_folder
  $archiveFolder = [string]$Config.sharepoint.archive_folder
  $inputItems = @(Get-DriveChildren -Token $Token -SiteId $SiteId -FolderPath $inputFolder)
  $archiveItems = @(Get-DriveChildren -Token $Token -SiteId $SiteId -FolderPath $archiveFolder)
  $archiveNames = @{}
  foreach ($item in $archiveItems) {
    $archiveNames[[string]$item.name.ToLowerInvariant()] = $true
  }

  $pending = @()
  foreach ($item in $inputItems) {
    if ($null -eq $item -or $item.folder) { continue }
    $name = [string]$item.name
    if (-not $name.EndsWith(".docx", [System.StringComparison]::OrdinalIgnoreCase)) { continue }
    if ($RequestedFileName -and $name -ne $RequestedFileName) { continue }

    $parsed = Parse-MinutaFileName -Name $name -MeetingMap $meetingMap
    if ($null -eq $parsed) {
      Write-Warning "Se ignora archivo sin patron soportado: $name"
      continue
    }

    $archiveDocx = $name.ToLowerInvariant()
    $archiveHtml = ("minuta_{0}_{1}.html" -f $parsed.Tipo, $parsed.DateIso).ToLowerInvariant()
    if ($archiveNames.ContainsKey($archiveDocx) -and $archiveNames.ContainsKey($archiveHtml)) {
      Write-Output "Ya archivado y generado HTML: $name"
      continue
    }

    $pending += [pscustomobject]@{
      Item = $item
      Parsed = $parsed
      ArchiveHtmlName = "minuta_{0}_{1}.html" -f $parsed.Tipo, $parsed.DateIso
      SourceFolder = $inputFolder
      ArchiveDocxExists = $archiveNames.ContainsKey($archiveDocx)
      ArchiveHtmlExists = $archiveNames.ContainsKey($archiveHtml)
    }
  }

  if ($RequestedFileName -and -not $pending.Count) {
    foreach ($item in $archiveItems) {
      if ($null -eq $item -or $item.folder) { continue }
      $name = [string]$item.name
      if ($name -ne $RequestedFileName) { continue }
      $parsed = Parse-MinutaFileName -Name $name -MeetingMap $meetingMap
      if ($null -eq $parsed) { continue }
      $pending += [pscustomobject]@{
        Item = $item
        Parsed = $parsed
        ArchiveHtmlName = "minuta_{0}_{1}.html" -f $parsed.Tipo, $parsed.DateIso
        SourceFolder = $archiveFolder
        ArchiveDocxExists = $true
        ArchiveHtmlExists = $archiveNames.ContainsKey(("minuta_{0}_{1}.html" -f $parsed.Tipo, $parsed.DateIso).ToLowerInvariant())
      }
      break
    }
  }

  $sorted = @($pending | Sort-Object { [datetime]($_.Item.createdDateTime) }, { [string]$_.Item.name })
  if ($Limit -gt 0) {
    return @($sorted | Select-Object -First $Limit)
  }
  return $sorted
}

function Build-MinutaPrompt {
  param(
    [string]$MeetingType,
    [string]$MeetingTitle,
    [string]$DateIso,
    [string]$SourceFileName,
    [string]$ExtractedText
  )

@"
Eres el asistente de minutas de MAYU, empresa chilena de construccion industrializada.

Convierte el contenido extraido desde un AGENDA.docx rellenado en una minuta HTML lista para enviar por correo.

Reglas:
- No inventes datos.
- Si una seccion no aparece, omitela.
- Escribe en espanol claro y ejecutivo.
- Devuelve SOLO HTML, sin JSON, sin markdown, sin backticks.

Requisitos del HTML:
- Sin etiquetas <html> ni <body>.
- Debe seguir este estilo visual, no un HTML plano:
  - contenedor principal: `<div style="font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; color: #333;">`
  - header con borde inferior azul `#2E5A8B`
  - titulo principal en azul `#2E5A8B`
  - subtitulo gris con fecha/hora/lugar
  - secciones con `h2` azul y borde izquierdo azul
  - bloque de objetivo con fondo `#f5f7fa`, padding y borde redondeado
  - tablas con encabezados coloreados y filas alternadas
  - footer final con reglas transversales o cierre ejecutivo
- Usa un tono ejecutivo y concreto.
- Incluye, si existe evidencia en el texto: asistentes, objetivo, estado/avances/hitos, bloqueos, decisiones, compromisos y proxima reunion.
- Si hay listas de proyectos o compromisos, conviertelas en tablas HTML con encabezados y estilos inline.
- Evita responder solo con parrafos simples. Debe verse como minuta formal MAYU del estilo historico.

Plantilla visual de referencia a imitar:
`<div style="font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; color: #333;">`
`<div style="border-bottom: 3px solid #2E5A8B; padding-bottom: 15px; margin-bottom: 20px;">`
`<h1 style="color: #2E5A8B; margin: 0; font-size: 24px;">...`
`<p style="color: #666; margin: 5px 0 0 0; font-size: 14px;">...`
`</div>`
`<div style="margin-bottom: 20px; background-color: #f5f7fa; padding: 12px 15px; border-radius: 5px;">...`
`<table style="width: 100%; border-collapse: collapse; font-size: 13px;">...`
`</div>`

Contexto:
- tipo interno: $MeetingType
- titulo esperado: $MeetingTitle
- fecha: $DateIso
- archivo fuente: $SourceFileName

Texto extraido del documento:
$ExtractedText
"@
}

function Process-MinutaItem {
  param([string]$Token, [string]$SiteId, [object]$Config, [object]$PendingItem, [bool]$DoSendEmail)

  $item = $PendingItem.Item
  $parsed = $PendingItem.Parsed
  $sender = [string]$Config.mail.sender
  $inputPath = [string]$item._mayuRelativePath
  $archiveFolder = [string]$Config.sharepoint.archive_folder
  $sourceFolder = [string]$PendingItem.SourceFolder

  Write-Output "Procesando $($item.name) -> tipo=$($parsed.Tipo) fecha=$($parsed.DateIso)"
  $docxBytes = Get-FileBytes -Token $Token -SiteId $SiteId -FilePath $inputPath
  $text = Convert-DocxBytesToText -Bytes $docxBytes
  if ([string]::IsNullOrWhiteSpace($text)) {
    throw "No se pudo extraer texto util desde $($item.name)"
  }

  $meetingTitle = [string]$parsed.Meeting.asunto_prefix
  $prompt = Build-MinutaPrompt `
    -MeetingType $parsed.Tipo `
    -MeetingTitle $meetingTitle `
    -DateIso $parsed.DateIso `
    -SourceFileName ([string]$item.name) `
    -ExtractedText $text

  $modelText = Invoke-OpenAiText -Prompt $prompt
  $htmlBody = ""
  $candidateHtml = Normalize-ModelHtml -Text $modelText
  if ($candidateHtml -and $candidateHtml -match "<(div|table|h1|h2|p|ul|ol)\b") {
    $htmlBody = $candidateHtml
  } else {
    $modelJson = ConvertFrom-ModelJson -Text $modelText
    if ($modelJson -and $modelJson.html_body) {
      $htmlBody = [string]$modelJson.html_body
    }
  }
  if ([string]::IsNullOrWhiteSpace($htmlBody)) {
    $htmlBody = Get-FallbackHtml `
      -MeetingTitle $meetingTitle `
      -Date $parsed.DateIso `
      -SourceFileName ([string]$item.name) `
      -ExtractedText $text
  }

  $subject = "{0} - {1}" -f [string]$parsed.Meeting.asunto_prefix, $parsed.DateIso
  if ($DoSendEmail) {
    Send-GraphMail `
      -Token $Token `
      -Sender $sender `
      -To @($parsed.Meeting.to) `
      -Subject $subject `
      -HtmlBody $htmlBody `
      -AttachmentName ([string]$item.name) `
      -AttachmentBytes $docxBytes
  }

  Write-TextFileToGraph `
    -Token $Token `
    -SiteId $SiteId `
    -FilePath "$archiveFolder/$($PendingItem.ArchiveHtmlName)" `
    -Text $htmlBody `
    -ContentType "text/html; charset=utf-8"

  if ($sourceFolder -eq [string]$Config.sharepoint.input_folder) {
    Move-GraphItemToFolder `
      -Token $Token `
      -SiteId $SiteId `
      -ItemId ([string]$item.id) `
      -DestinationFolderPath $archiveFolder `
      -NewName ([string]$item.name)
  }
}

$configJson = Get-RunbookVariable "MayuMinutasConfigJson"
$config = $configJson | ConvertFrom-Json
$token = Get-GraphToken
$siteId = Get-SiteId -Token $token -HostName ([string]$config.sharepoint.host)

Ensure-GraphFolder -Token $token -SiteId $siteId -FolderPath ([string]$config.sharepoint.input_folder)
Ensure-GraphFolder -Token $token -SiteId $siteId -FolderPath ([string]$config.sharepoint.archive_folder)

if ($Mode -eq "test") {
  $subject = "Prueba pipeline minutas MAYU"
  $body = "<p>Pipeline de minutas en GitHub Actions operativo.</p>"
  if ($SendEmail) {
    Send-GraphMail -Token $token -Sender ([string]$config.mail.sender) -To @("fjescudero@imayu.cl") -Subject $subject -HtmlBody $body
  }
  Write-Output "Modo test completado."
  exit 0
}

$requestedFileName = ""
if ($Mode -eq "single_file") {
  if ([string]::IsNullOrWhiteSpace($FileName)) {
    throw "single_file requiere -FileName"
  }
  $requestedFileName = $FileName
}

$pendingItems = @(Find-PendingAgendaItems -Token $token -SiteId $siteId -Config $config -RequestedFileName $requestedFileName -Limit $ProcessLimit)
if (-not $pendingItems.Count) {
  Write-Output "No hay archivos pendientes para procesar."
  exit 0
}

foreach ($pending in $pendingItems) {
  Process-MinutaItem -Token $token -SiteId $siteId -Config $config -PendingItem $pending -DoSendEmail $SendEmail
}

Write-Output "Procesamiento de minutas completado. Archivos procesados: $($pendingItems.Count)"
