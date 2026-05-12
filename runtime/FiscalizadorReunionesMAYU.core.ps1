param(
  [ValidateSet("pre_reunion", "post_reunion", "manual_due_sweep", "weekly_report", "monthly_report", "test")]
  [string]$Mode = "post_reunion",
  [string]$Tipo = "",
  [string]$Date = "",
  [int]$LookbackDays = 7
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

function Invoke-GraphGet {
  param([string]$Token, [string]$Uri)
  Invoke-RestMethod -Method Get -Uri $Uri -Headers @{ Authorization = "Bearer $Token" }
}

function Invoke-GraphJson {
  param([string]$Token, [string]$Method, [string]$Uri, [object]$Body)
  $json = $Body | ConvertTo-Json -Depth 30
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

function Get-Mayutime {
  param([string]$TimeZoneName)
  $candidateIds = @($TimeZoneName, "Pacific SA Standard Time", "Chile Standard Time")
  foreach ($id in $candidateIds) {
    if ([string]::IsNullOrWhiteSpace($id)) {
      continue
    }
    try {
      $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById($id)
      return [System.TimeZoneInfo]::ConvertTime([DateTimeOffset]::UtcNow, $tz).DateTime
    } catch {
      continue
    }
  }
  return [DateTime]::UtcNow.AddHours(-4)
}

function Get-DayCode {
  param([datetime]$TargetDate)
  $codes = @("SU", "MO", "TU", "WE", "TH", "FR", "SA")
  $codes[[int]$TargetDate.DayOfWeek]
}

function Test-MonthlyDueOnDate {
  param([object]$Meeting, [datetime]$TargetDate)
  $iso = $TargetDate.ToString("yyyy-MM-dd")
  if ($Meeting.exceptions -and @($Meeting.exceptions) -contains $iso) {
    return $true
  }
  if ($Meeting.not_before -and $iso -lt [string]$Meeting.not_before) {
    return $false
  }
  if ($Meeting.rule -eq "first_tuesday") {
    return $TargetDate.Day -le 7 -and $TargetDate.DayOfWeek -eq [DayOfWeek]::Tuesday
  }
  if ($Meeting.rule -eq "third_thursday") {
    return $TargetDate.DayOfWeek -eq [DayOfWeek]::Thursday -and ([Math]::Floor(($TargetDate.Day - 1) / 7) + 1) -eq 3
  }
  if ($Meeting.rule -eq "third_friday") {
    return $TargetDate.DayOfWeek -eq [DayOfWeek]::Friday -and ([Math]::Floor(($TargetDate.Day - 1) / 7) + 1) -eq 3
  }
  return $false
}

function Get-DateAtTime {
  param([datetime]$TargetDate, [string]$Time)
  $parts = $Time.Split(":")
  $TargetDate.Date.AddHours([int]$parts[0]).AddMinutes([int]$parts[1])
}

function Get-MeetingFileLabels {
  param([string]$MeetingType)
  if ($MeetingType -eq "id") {
    return @("I+D", "id")
  }
  return @($MeetingType)
}

function Get-ExpectedFileName {
  param([string]$MeetingType, [datetime]$TargetDate)
  $labels = @(Get-MeetingFileLabels -MeetingType $MeetingType)
  $label = $labels[0]
  "AGENDA_{0}_{1}_{2}_{3}.docx" -f $label, $TargetDate.ToString("dd"), $TargetDate.ToString("MM"), $TargetDate.ToString("yyyy")
}

function Get-ExpectedEvidenceNames {
  param([string]$MeetingType, [datetime]$TargetDate)
  $names = @()
  foreach ($label in (Get-MeetingFileLabels -MeetingType $MeetingType)) {
    $names += ("AGENDA_{0}_{1}_{2}_{3}.docx" -f $label, $TargetDate.ToString("dd"), $TargetDate.ToString("MM"), $TargetDate.ToString("yyyy")).ToLowerInvariant()
    $names += ("minuta_{0}_{1}.html" -f $label, $TargetDate.ToString("yyyy-MM-dd")).ToLowerInvariant()
  }
  return @($names | Select-Object -Unique)
}

function Get-ExpectedOccurrences {
  param([object]$Config, [datetime]$StartDate, [datetime]$EndDate)

  $items = @()
  $dateCursor = $StartDate.Date
  while ($dateCursor -le $EndDate.Date) {
    $iso = $dateCursor.ToString("yyyy-MM-dd")
    if (-not ($Config.holidays -and @($Config.holidays) -contains $iso)) {
      foreach ($meeting in @($Config.meetings)) {
        if (@($meeting.days) -contains (Get-DayCode $dateCursor)) {
          $items += [pscustomobject]@{
            Tipo = [string]$meeting.tipo
            Nombre = [string]$meeting.nombre
            Fecha = $dateCursor
            Meeting = $meeting
            DueAt = (Get-DateAtTime -TargetDate $dateCursor -Time ([string]$meeting.end)).AddHours(2)
          }
        }
      }

      foreach ($meeting in @($Config.monthly_meetings)) {
        if (Test-MonthlyDueOnDate -Meeting $meeting -TargetDate $dateCursor) {
          $items += [pscustomobject]@{
            Tipo = [string]$meeting.tipo
            Nombre = [string]$meeting.nombre
            Fecha = $dateCursor
            Meeting = $meeting
            DueAt = (Get-DateAtTime -TargetDate $dateCursor -Time ([string]$meeting.end)).AddHours(2)
          }
        }
      }
    }
    $dateCursor = $dateCursor.AddDays(1)
  }
  return $items
}

function Get-DriveChildren {
  param([string]$Token, [string]$SiteId, [string]$FolderPath)
  $items = @()
  $encoded = ConvertTo-DrivePath $FolderPath
  if ([string]::IsNullOrWhiteSpace($encoded)) {
    $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/root/children?`$top=200"
  } else {
    $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/root:/$encoded" + ":/children?`$top=200"
  }

  while ($uri) {
    $data = Invoke-GraphGet -Token $Token -Uri $uri
    if ($null -eq $data) {
      break
    }
    if ($null -ne $data.value) {
      $items += @($data.value)
    }
    $next = $null
    if ($null -ne $data.PSObject -and $null -ne $data.PSObject.Properties) {
      $next = $data.PSObject.Properties["@odata.nextLink"]
    }
    if ($next -and $next.Value) {
      $uri = [string]$next.Value
    } else {
      $uri = $null
    }
  }
  return $items
}

function Get-FolderItemsRecursive {
  param(
    [string]$Token,
    [string]$SiteId,
    [string]$FolderPath,
    [int]$Depth = 0,
    [int]$MaxDepth = 4
  )

  $all = @()
  try {
    $children = Get-DriveChildren -Token $Token -SiteId $SiteId -FolderPath $FolderPath
  } catch {
    return @()
  }

  foreach ($item in @($children)) {
    if ($null -eq $item -or [string]::IsNullOrWhiteSpace([string]$item.name)) {
      continue
    }
    $relative = if ([string]::IsNullOrWhiteSpace($FolderPath)) { [string]$item.name } else { "$FolderPath/$($item.name)" }
    $item | Add-Member -NotePropertyName "_mayuRelativePath" -NotePropertyValue $relative -Force
    $all += $item
    if ($item.folder -and $Depth -lt $MaxDepth) {
      $all += Get-FolderItemsRecursive -Token $Token -SiteId $SiteId -FolderPath $relative -Depth ($Depth + 1) -MaxDepth $MaxDepth
    }
  }
  return $all
}

function Get-FolderItemsFlatSafe {
  param([string]$Token, [string]$SiteId, [string]$FolderPath)

  try {
    $items = Get-DriveChildren -Token $Token -SiteId $SiteId -FolderPath $FolderPath
    $safe = @()
    foreach ($item in @($items)) {
      if ($null -eq $item -or [string]::IsNullOrWhiteSpace([string]$item.name)) {
        continue
      }
      $relative = if ([string]::IsNullOrWhiteSpace($FolderPath)) { [string]$item.name } else { "$FolderPath/$($item.name)" }
      try {
        $item | Add-Member -NotePropertyName "_mayuRelativePath" -NotePropertyValue $relative -Force
      } catch {
        # Si el objeto no admite Add-Member, igual sirve para nombre simple.
      }
      $safe += $item
    }
    return $safe
  } catch {
    Write-Warning "No se pudo leer carpeta SharePoint '$FolderPath'. Se continua sin esa evidencia. $($_.Exception.Message)"
    return @()
  }
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
  $tempFile = Join-Path $env:TEMP ("mayu_attach_" + [Guid]::NewGuid().ToString("N") + ".bin")
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
    [string[]]$Cc,
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
    ccRecipients = @($Cc | Where-Object { $_ } | ForEach-Object { @{ emailAddress = @{ address = $_ } } })
    attachments = $attachments
  }

  Invoke-GraphJson -Token $Token -Method Post -Uri "https://graph.microsoft.com/v1.0/users/$Sender/sendMail" -Body @{
    message = $message
    saveToSentItems = $true
  } | Out-Null
}

function Test-EvidencePresent {
  param([object[]]$Items, [string]$MeetingType, [datetime]$TargetDate)
  $expectedNames = @(Get-ExpectedEvidenceNames -MeetingType $MeetingType -TargetDate $TargetDate)
  $processedMd = ("{0}.md" -f $TargetDate.ToString("yyyy-MM-dd")).ToLowerInvariant()
  $pathLabels = @(Get-MeetingFileLabels -MeetingType $MeetingType | ForEach-Object { $_.ToLowerInvariant() })

  foreach ($item in @($Items)) {
    if ($null -eq $item) {
      continue
    }
    $name = ([string]$item.name).ToLowerInvariant()
    $path = ([string]$item._mayuRelativePath).ToLowerInvariant()
    if ($expectedNames -contains $name) {
      return $true
    }
    foreach ($label in $pathLabels) {
      if ($path -like "*$label*" -and $name -eq $processedMd) {
        return $true
      }
    }
  }
  return $false
}

function Find-EvidenceItem {
  param([object[]]$Items, [string]$MeetingType, [datetime]$TargetDate)
  $expectedNames = @(Get-ExpectedEvidenceNames -MeetingType $MeetingType -TargetDate $TargetDate)
  $processedMd = ("{0}.md" -f $TargetDate.ToString("yyyy-MM-dd")).ToLowerInvariant()
  $pathLabels = @(Get-MeetingFileLabels -MeetingType $MeetingType | ForEach-Object { $_.ToLowerInvariant() })

  foreach ($item in @($Items)) {
    if ($null -eq $item -or $item.folder) {
      continue
    }
    $name = ([string]$item.name).ToLowerInvariant()
    $path = ([string]$item._mayuRelativePath).ToLowerInvariant()
    if ($expectedNames -contains $name) {
      return $item
    }
    foreach ($label in $pathLabels) {
      if ($path -like "*$label*" -and $name -eq $processedMd) {
        return $item
      }
    }
  }
  return $null
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
    [xml]$xml = Get-Content -Raw -Path $documentPath
    $texts = @()
    foreach ($node in $xml.GetElementsByTagName("w:t")) {
      if ($node.InnerText) {
        $texts += $node.InnerText
      }
    }
    return ($texts -join " ").Trim()
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

function Convert-BytesToMeetingText {
  param([byte[]]$Bytes, [string]$FileName)
  $extension = [System.IO.Path]::GetExtension($FileName).ToLowerInvariant()
  if ($extension -eq ".docx") {
    return Convert-DocxBytesToText -Bytes $Bytes
  }
  $text = [System.Text.Encoding]::UTF8.GetString($Bytes)
  if ($extension -eq ".html" -or $extension -eq ".htm") {
    $text = [regex]::Replace($text, "<script[\s\S]*?</script>", " ")
    $text = [regex]::Replace($text, "<style[\s\S]*?</style>", " ")
    $text = [regex]::Replace($text, "<[^>]+>", " ")
    $text = [System.Net.WebUtility]::HtmlDecode($text)
  }
  return ([regex]::Replace($text, "\s+", " ")).Trim()
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
    Write-Warning "OpenAI no devolvio JSON parseable para seguimiento. Se guardara respuesta como texto."
    return $null
  }
}

function Get-IntelligenceFileName {
  param([string]$MeetingType, [datetime]$TargetDate)
  "seguimiento_{0}_{1}.json" -f $MeetingType, $TargetDate.ToString("yyyy-MM-dd")
}

function Get-IntelligenceFolder {
  param([object]$Config)
  $folder = [string]$Config.sharepoint.seguimiento_folder
  if ([string]::IsNullOrWhiteSpace($folder)) {
    return "minutas_archivadas/seguimiento_compromisos"
  }
  return $folder
}

function Find-PreviousIntelligenceItem {
  param([object[]]$Items, [string]$MeetingType, [datetime]$TargetDate)
  $pattern = "^seguimiento_$([regex]::Escape($MeetingType))_(\d{4}-\d{2}-\d{2})\.json$"
  $candidates = @()
  foreach ($item in @($Items)) {
    if ($null -eq $item -or [string]::IsNullOrWhiteSpace([string]$item.name)) {
      continue
    }
    $match = [regex]::Match([string]$item.name, $pattern)
    if ($match.Success) {
      $date = [datetime]::Parse($match.Groups[1].Value)
      if ($date.Date -lt $TargetDate.Date) {
        $candidates += [pscustomobject]@{ Date = $date; Item = $item }
      }
    }
  }
  if (@($candidates).Count -eq 0) {
    return $null
  }
  return (@($candidates | Sort-Object Date -Descending)[0]).Item
}

function Read-TextDriveItem {
  param([string]$Token, [string]$SiteId, [object]$Item)
  if ($null -eq $Item -or [string]::IsNullOrWhiteSpace([string]$Item._mayuRelativePath)) {
    return ""
  }
  $bytes = Get-FileBytes -Token $Token -SiteId $SiteId -FilePath ([string]$Item._mayuRelativePath)
  return [System.Text.Encoding]::UTF8.GetString($bytes)
}

function New-MeetingIntelligence {
  param(
    [object]$Occurrence,
    [string]$EvidenceText,
    [object]$PreviousIntelligence
  )

  $previousJson = if ($PreviousIntelligence) { $PreviousIntelligence | ConvertTo-Json -Depth 20 } else { "{}" }
  if ($EvidenceText.Length -gt 14000) {
    $EvidenceText = $EvidenceText.Substring(0, 14000)
  }

  $prompt = @"
Eres el Agente de Continuidad Operacional de MAYU.
Analiza la minuta de una reunion y comparala con el seguimiento anterior si existe.
No inventes compromisos, responsables ni fechas. Si algo no esta en la minuta, usa null o [].
Devuelve SOLO JSON valido con esta estructura:
{
  "tipo": "$($Occurrence.Tipo)",
  "fecha": "$($Occurrence.Fecha.ToString("yyyy-MM-dd"))",
  "reunion": "$($Occurrence.Nombre)",
  "responsable_reunion": "$([string]$Occurrence.Meeting.responsable)",
  "resumen": "",
  "asistencia": { "realizada": true, "asistentes_detectados": [] },
  "compromisos": [
    { "descripcion": "", "responsable": "", "fecha_objetivo": null, "estado_inferido": "nuevo|cumplido|pendiente|vencido|bloqueado|parcial", "evidencia": "" }
  ],
  "compromisos_cumplidos_desde_anterior": [],
  "pendientes_arrastrados": [],
  "bloqueos": [
    { "descripcion": "", "responsable": "", "impacto": "", "accion_recomendada": "" }
  ],
  "decisiones": [],
  "riesgos": [],
  "seguimiento_recomendado": "",
  "confianza": 0.0
}

Seguimiento anterior:
$previousJson

Minuta actual:
$EvidenceText
"@

  $ai = Invoke-OpenAiText -Prompt $prompt
  $parsed = ConvertFrom-ModelJson -Text $ai
  if ($parsed) {
    return $parsed
  }
  return [pscustomobject]@{
    tipo = [string]$Occurrence.Tipo
    fecha = $Occurrence.Fecha.ToString("yyyy-MM-dd")
    reunion = [string]$Occurrence.Nombre
    responsable_reunion = [string]$Occurrence.Meeting.responsable
    resumen = "No se pudo parsear el seguimiento IA."
    asistencia = @{ realizada = $true; asistentes_detectados = @() }
    compromisos = @()
    compromisos_cumplidos_desde_anterior = @()
    pendientes_arrastrados = @()
    bloqueos = @()
    decisiones = @()
    riesgos = @()
    seguimiento_recomendado = $ai
    confianza = 0
  }
}

function Get-MeetingIntelligenceRows {
  param(
    [object]$Config,
    [string]$Token,
    [string]$SiteId,
    [object[]]$Occurrences,
    [object[]]$EvidenceItems,
    [object[]]$ExistingIntelligenceItems
  )

  $folder = Get-IntelligenceFolder -Config $Config
  Ensure-GraphFolder -Token $Token -SiteId $SiteId -FolderPath $folder

  $rows = @()
  foreach ($occ in @($Occurrences)) {
    $evidenceItem = Find-EvidenceItem -Items $EvidenceItems -MeetingType $occ.Tipo -TargetDate $occ.Fecha
    if ($null -eq $evidenceItem) {
      continue
    }
    try {
      $bytes = Get-FileBytes -Token $Token -SiteId $SiteId -FilePath ([string]$evidenceItem._mayuRelativePath)
      $text = Convert-BytesToMeetingText -Bytes $bytes -FileName ([string]$evidenceItem.name)
      if ([string]::IsNullOrWhiteSpace($text)) {
        continue
      }

      $previousItem = Find-PreviousIntelligenceItem -Items $ExistingIntelligenceItems -MeetingType $occ.Tipo -TargetDate $occ.Fecha
      $previous = $null
      if ($previousItem) {
        try {
          $previousText = Read-TextDriveItem -Token $Token -SiteId $SiteId -Item $previousItem
          if (-not [string]::IsNullOrWhiteSpace($previousText)) {
            $previous = $previousText | ConvertFrom-Json
          }
        } catch {
          Write-Warning "No se pudo leer seguimiento previo para $($occ.Tipo) $($occ.Fecha.ToString("yyyy-MM-dd")). $($_.Exception.Message)"
        }
      }

      $intel = New-MeetingIntelligence -Occurrence $occ -EvidenceText $text -PreviousIntelligence $previous
      $intel | Add-Member -NotePropertyName "_mayuTo" -NotePropertyValue @($occ.Meeting.to) -Force
      $intel | Add-Member -NotePropertyName "_mayuCc" -NotePropertyValue @($occ.Meeting.cc) -Force
      $intel | Add-Member -NotePropertyName "_mayuResponsable" -NotePropertyValue ([string]$occ.Meeting.responsable) -Force
      $json = $intel | ConvertTo-Json -Depth 30
      $fileName = Get-IntelligenceFileName -MeetingType $occ.Tipo -TargetDate $occ.Fecha
      Write-TextFileToGraph -Token $Token -SiteId $SiteId -FilePath "$folder/$fileName" -Text $json -ContentType "application/json; charset=utf-8"
      $rows += $intel
    } catch {
      Write-Warning "No se pudo generar seguimiento para $($occ.Tipo) $($occ.Fecha.ToString("yyyy-MM-dd")). $($_.Exception.Message)"
    }
  }
  return $rows
}

function Convert-IntelligenceRowsToHtml {
  param([object[]]$Rows)
  if (@($Rows).Count -eq 0) {
    return "<p>No hay minutas procesadas para seguimiento de compromisos en este periodo.</p>"
  }
  $html = ""
  foreach ($row in @($Rows)) {
    $commitments = @($row.compromisos).Count
    $done = @($row.compromisos | Where-Object { $_.estado_inferido -eq "cumplido" }).Count
    $blocked = @($row.bloqueos).Count
    $carried = @($row.pendientes_arrastrados).Count
    $html += "<div style='border:1px solid #ddd;padding:10px;margin:8px 0;'><strong>$($row.fecha) - $($row.reunion)</strong><br/>"
    $html += "<span>Compromisos: $commitments. Cumplidos: $done. Arrastrados: $carried. Bloqueos: $blocked.</span><br/>"
    if ($row.resumen) {
      $html += "<span>$([System.Security.SecurityElement]::Escape([string]$row.resumen))</span><br/>"
    }
    if ($row.seguimiento_recomendado) {
      $html += "<em>$([System.Security.SecurityElement]::Escape([string]$row.seguimiento_recomendado))</em>"
    }
    $html += "</div>"
  }
  return $html
}

function Send-ResponsibleIntelligenceReports {
  param(
    [object]$Config,
    [string]$Token,
    [string]$Title,
    [object[]]$Rows
  )

  $groups = @{}
  foreach ($row in @($Rows)) {
    $to = @($row._mayuTo | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    if (@($to).Count -eq 0) {
      continue
    }
    $key = ($to | Sort-Object) -join ";"
    if (-not $groups.ContainsKey($key)) {
      $groups[$key] = @()
    }
    $groups[$key] += $row
  }

  foreach ($key in $groups.Keys) {
    $recipientRows = @($groups[$key])
    $to = @($key -split ";" | Where-Object { $_ })
    $cc = @($Config.mail.felix)
    $meetingNames = @($recipientRows | ForEach-Object { [string]$_.reunion } | Where-Object { $_ } | Select-Object -Unique)
    $meetingsLabel = if (@($meetingNames).Count -gt 0) { ($meetingNames -join ", ") } else { "sin reunion identificada" }
    $subject = "$Title [$meetingsLabel]"
    $html = @"
<html>
<body style="font-family: Arial, sans-serif; font-size: 14px; color: #222; max-width: 850px;">
  <h2>$Title</h2>
  <p><strong>Reuniones:</strong> $meetingsLabel</p>
  <p>Resumen de compromisos, bloqueos y continuidad de las reuniones bajo tu responsabilidad.</p>
  $(Convert-IntelligenceRowsToHtml -Rows $recipientRows)
  <p style="font-size:12px;color:#777;margin-top:24px;">Generado por el Agente de Continuidad Operacional MAYU.</p>
</body>
</html>
"@
    try {
      Send-GraphMail `
        -Token $Token `
        -Sender $Config.mail.sender `
        -To $to `
        -Cc $cc `
        -Subject $subject `
        -HtmlBody $html
    } catch {
      Write-Warning "No se pudo enviar seguimiento a responsables $key. $($_.Exception.Message)"
    }
  }
}

function Get-ScoreBand {
  param([int]$Score)
  if ($Score -ge 90) {
    return "excelente"
  }
  if ($Score -ge 75) {
    return "bien"
  }
  if ($Score -ge 60) {
    return "requiere seguimiento"
  }
  return "critico"
}

function Get-NormalizedState {
  param([object]$Value)
  return ([string]$Value).Trim().ToLowerInvariant()
}

function Get-ResponsibleScorecards {
  param([object[]]$StatusRows, [object[]]$IntelligenceRows)

  $groups = @{}
  foreach ($row in @($StatusRows)) {
    $to = @($row.to | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    if (@($to).Count -eq 0) {
      continue
    }
    $key = ($to | Sort-Object) -join ";"
    if (-not $groups.ContainsKey($key)) {
      $groups[$key] = [pscustomobject]@{ to = $to; status = @(); intelligence = @() }
    }
    $groups[$key].status += $row
  }

  foreach ($row in @($IntelligenceRows)) {
    $to = @($row._mayuTo | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    if (@($to).Count -eq 0) {
      continue
    }
    $key = ($to | Sort-Object) -join ";"
    if (-not $groups.ContainsKey($key)) {
      $groups[$key] = [pscustomobject]@{ to = $to; status = @(); intelligence = @() }
    }
    $groups[$key].intelligence += $row
  }

  $cards = @()
  foreach ($key in $groups.Keys) {
    $group = $groups[$key]
    $statusRows = @($group.status)
    $intelRows = @($group.intelligence)
    $meetings = @($statusRows).Count
    $fulfilled = @($statusRows | Where-Object { $_.estado -eq "CUMPLIDA" }).Count
    $alerts = @($statusRows | Where-Object { $_.alerta -eq "SI" }).Count

    $evidenceScore = if ($meetings -gt 0) { [Math]::Round(($fulfilled / $meetings) * 100, 0) } else { 75 }
    $commitments = @()
    foreach ($intel in $intelRows) {
      $commitments += @($intel.compromisos)
    }
    $commitmentCount = @($commitments).Count
    $doneCommitments = @($commitments | Where-Object { (Get-NormalizedState $_.estado_inferido) -eq "cumplido" }).Count
    $partialCommitments = @($commitments | Where-Object { (Get-NormalizedState $_.estado_inferido) -eq "parcial" }).Count
    $badCommitments = @($commitments | Where-Object { @("pendiente", "vencido", "bloqueado") -contains (Get-NormalizedState $_.estado_inferido) }).Count
    $taskScore = if ($commitmentCount -gt 0) { [Math]::Round((($doneCommitments + ($partialCommitments * 0.5)) / $commitmentCount) * 100, 0) } else { 75 }

    $blockers = 0
    $carried = 0
    $risks = 0
    foreach ($intel in $intelRows) {
      $blockers += @($intel.bloqueos).Count
      $carried += @($intel.pendientes_arrastrados).Count
      $risks += @($intel.riesgos).Count
    }
    $problemScore = [Math]::Max(0, 100 - ($blockers * 15) - ($carried * 10))
    $cleanAdvanceScore = [Math]::Max(0, 100 - ($blockers * 15) - ($carried * 10) - ($risks * 8) - ($alerts * 8))
    $score = [Math]::Round(($evidenceScore * 0.25) + ($taskScore * 0.35) + ($problemScore * 0.20) + ($cleanAdvanceScore * 0.20), 0)

    $cards += [pscustomobject]@{
      to = @($group.to)
      responsable = (@($statusRows.responsable) + @($intelRows._mayuResponsable) | Where-Object { $_ } | Select-Object -First 1)
      score = [int]$score
      banda = Get-ScoreBand -Score ([int]$score)
      evidencia_asistencia = [int]$evidenceScore
      cumplimiento_tareas = [int]$taskScore
      resolucion_problemas = [int]$problemScore
      avances_limpios = [int]$cleanAdvanceScore
      reuniones_esperadas = $meetings
      reuniones_con_evidencia = $fulfilled
      alertas = $alerts
      compromisos_total = $commitmentCount
      compromisos_cumplidos = $doneCommitments
      compromisos_parciales = $partialCommitments
      compromisos_pendientes = $badCommitments
      bloqueos = $blockers
      pendientes_arrastrados = $carried
      riesgos = $risks
      reuniones = @($intelRows | ForEach-Object { [pscustomobject]@{ fecha = $_.fecha; reunion = $_.reunion; resumen = $_.resumen; seguimiento_recomendado = $_.seguimiento_recomendado } })
    }
  }
  return $cards
}

function Convert-ScorecardToHtml {
  param([object]$Card, [string]$PeriodLabel)
  $color = if ($Card.score -ge 90) { "#166534" } elseif ($Card.score -ge 75) { "#2563eb" } elseif ($Card.score -ge 60) { "#a16207" } else { "#991b1b" }
  $meetingsHtml = ""
  foreach ($meeting in @($Card.reuniones)) {
    $meetingsHtml += "<li><strong>$($meeting.fecha) - $($meeting.reunion):</strong> $([System.Security.SecurityElement]::Escape([string]$meeting.resumen))</li>"
  }
  if (-not $meetingsHtml) {
    $meetingsHtml = "<li>No hubo seguimiento inteligente disponible para reuniones con evidencia.</li>"
  }
  @"
<html>
<body style="font-family: Arial, sans-serif; font-size: 14px; color: #222; max-width: 850px;">
  <h2>Calificacion mensual MAYU - $PeriodLabel</h2>
  <p><strong>Responsable:</strong> $($Card.responsable)</p>
  <p style="font-size:24px;color:$color;"><strong>$($Card.score)/100</strong> - $($Card.banda)</p>
  <table style="border-collapse:collapse;width:100%;font-size:13px;">
    <tr><th style="text-align:left;border:1px solid #ddd;padding:6px;">Criterio</th><th style="text-align:left;border:1px solid #ddd;padding:6px;">Puntaje</th></tr>
    <tr><td style="border:1px solid #ddd;padding:6px;">Asistencia / evidencia</td><td style="border:1px solid #ddd;padding:6px;">$($Card.evidencia_asistencia)</td></tr>
    <tr><td style="border:1px solid #ddd;padding:6px;">Cumplimiento de tareas</td><td style="border:1px solid #ddd;padding:6px;">$($Card.cumplimiento_tareas)</td></tr>
    <tr><td style="border:1px solid #ddd;padding:6px;">Resolucion de problemas</td><td style="border:1px solid #ddd;padding:6px;">$($Card.resolucion_problemas)</td></tr>
    <tr><td style="border:1px solid #ddd;padding:6px;">Avances limpios</td><td style="border:1px solid #ddd;padding:6px;">$($Card.avances_limpios)</td></tr>
  </table>
  <p><strong>Reuniones:</strong> $($Card.reuniones_con_evidencia) con evidencia de $($Card.reuniones_esperadas) esperadas. <strong>Alertas:</strong> $($Card.alertas).</p>
  <p><strong>Compromisos:</strong> $($Card.compromisos_cumplidos) cumplidos, $($Card.compromisos_parciales) parciales, $($Card.compromisos_pendientes) pendientes/vencidos/bloqueados, de $($Card.compromisos_total) detectados.</p>
  <p><strong>Bloqueos:</strong> $($Card.bloqueos). <strong>Pendientes arrastrados:</strong> $($Card.pendientes_arrastrados). <strong>Riesgos:</strong> $($Card.riesgos).</p>
  <h3>Lectura por reunion</h3>
  <ul>$meetingsHtml</ul>
  <p style="font-size:12px;color:#777;margin-top:24px;">Generado por el Agente de Continuidad Operacional MAYU. Puntaje automatico 0-100 basado en evidencia, compromisos, bloqueos y continuidad.</p>
</body>
</html>
"@
}

function Send-MonthlyScorecards {
  param(
    [object]$Config,
    [string]$Token,
    [string]$SiteId,
    [object[]]$StatusRows,
    [object[]]$IntelligenceRows,
    [datetime]$StartDate
  )
  $folder = "$($Config.sharepoint.reportes_folder)/scorecards_$($StartDate.ToString("yyyy-MM"))"
  Ensure-GraphFolder -Token $Token -SiteId $SiteId -FolderPath $folder
  $cards = Get-ResponsibleScorecards -StatusRows $StatusRows -IntelligenceRows $IntelligenceRows
  foreach ($card in @($cards)) {
    $html = Convert-ScorecardToHtml -Card $card -PeriodLabel $StartDate.ToString("yyyy-MM")
    $safeName = (($card.to -join "_") -replace "[^a-zA-Z0-9_.-]", "_")
    Write-TextFileToGraph -Token $Token -SiteId $SiteId -FilePath "$folder/scorecard_$safeName.html" -Text $html -ContentType "text/html; charset=utf-8"
    $meetingNames = @($card.reuniones | ForEach-Object { [string]$_.reunion } | Where-Object { $_ } | Select-Object -Unique)
    $meetingsLabel = if (@($meetingNames).Count -gt 0) { ($meetingNames -join ", ") } else { "sin reunion identificada" }
    try {
      Send-GraphMail `
        -Token $Token `
        -Sender $Config.mail.sender `
        -To @($card.to) `
        -Cc @($Config.mail.felix) `
        -Subject "Calificacion mensual MAYU [$meetingsLabel] - $($StartDate.ToString("yyyy-MM")) - $($card.score)/100" `
        -HtmlBody $html
    } catch {
      Write-Warning "No se pudo enviar scorecard mensual a $($card.to -join ', '). $($_.Exception.Message)"
    }
  }
  return $cards
}

function Get-SentAlertKeys {
  param([object[]]$Items)
  $keys = New-Object System.Collections.Generic.HashSet[string]
  foreach ($item in @($Items)) {
    if ($null -eq $item) {
      continue
    }
    $name = [string]$item.name
    $match = [regex]::Match($name, "^alerta_(.+)_(\d{4}-\d{2}-\d{2})\.json$")
    if ($match.Success) {
      [void]$keys.Add("$($match.Groups[1].Value)|$($match.Groups[2].Value)")
    }
  }
  return $keys
}

function Get-SentReminderKeys {
  param([object[]]$Items)
  $keys = New-Object System.Collections.Generic.HashSet[string]
  foreach ($item in @($Items)) {
    if ($null -eq $item) {
      continue
    }
    $name = [string]$item.name
    $match = [regex]::Match($name, "^recordatorio_(.+)_(\d{4}-\d{2}-\d{2})\.json$")
    if ($match.Success) {
      [void]$keys.Add("$($match.Groups[1].Value)|$($match.Groups[2].Value)")
    }
  }
  return $keys
}

function New-AlertBody {
  param([object]$Config, [object]$Occurrence, [string]$FileName)
  $uploadLink = $Config.sharepoint.minutas_entrada_link
  $tipo = $Occurrence.Tipo
  $fecha = $Occurrence.Fecha.ToString("yyyy-MM-dd")
  $responsable = $Occurrence.Meeting.responsable
  @"
<html>
<body style="font-family: Arial, sans-serif; font-size: 14px; color: #222; max-width: 760px;">
  <h2 style="margin-bottom: 4px;">Alerta MAYU - falta minuta/reporte de reunion</h2>
  <p style="margin-top: 0; color: #666;">Control automatico del sistema de reuniones.</p>
  <p>Hola $responsable,</p>
  <p>El fiscalizador detecto que falta evidencia de la reunion <strong>$tipo</strong> del <strong>$fecha</strong>.</p>
  <table style="border-collapse: collapse; width: 100%; margin: 18px 0;">
    <tr><td style="font-weight:bold; padding:6px; border:1px solid #ddd;">Estado</td><td style="padding:6px; border:1px solid #ddd;">ROJO</td></tr>
    <tr><td style="font-weight:bold; padding:6px; border:1px solid #ddd;">Falta</td><td style="padding:6px; border:1px solid #ddd;">No aparece una minuta/reporte valido subido en SharePoint</td></tr>
    <tr><td style="font-weight:bold; padding:6px; border:1px solid #ddd;">Accion requerida</td><td style="padding:6px; border:1px solid #ddd;">Rellenar el AGENDA.docx adjunto y subirlo a SharePoint en minutas_entrada</td></tr>
    <tr><td style="font-weight:bold; padding:6px; border:1px solid #ddd;">Fecha limite</td><td style="padding:6px; border:1px solid #ddd;">Hoy</td></tr>
  </table>
  <p><strong>Carpeta de subida:</strong><br><a href="$uploadLink">$uploadLink</a></p>
  <p><strong>Nombre obligatorio del archivo:</strong><br><code>$FileName</code></p>
  <p style="color:#9a3412;"><strong>Importante:</strong> si el archivo se sube con otro nombre, el flujo automatico puede no procesarlo correctamente.</p>
  <p style="font-size: 12px; color: #777; margin-top: 28px;">Generado por el fiscalizador de reuniones MAYU.</p>
</body>
</html>
"@
}

function New-ReminderBody {
  param([object]$Config, [object]$Occurrence, [string]$FileName)
  $uploadLink = $Config.sharepoint.minutas_entrada_link
  $tipo = $Occurrence.Tipo
  $fecha = $Occurrence.Fecha.ToString("yyyy-MM-dd")
  $hora = [string]$Occurrence.Meeting.start
  $responsable = $Occurrence.Meeting.responsable
@"
<html>
<body style="font-family: Arial, sans-serif; font-size: 14px; color: #222; max-width: 760px;">
  <h2 style="margin-bottom: 4px;">Recordatorio MAYU - reunion de hoy</h2>
  <p style="margin-top: 0; color: #666;">Control automatico del sistema de reuniones.</p>
  <p>Hola $responsable,</p>
  <p>Te recordamos la reunion <strong>$tipo</strong> de hoy <strong>$fecha</strong> a las <strong>$hora</strong>.</p>
  <table style="border-collapse: collapse; width: 100%; margin: 18px 0;">
    <tr><td style="font-weight:bold; padding:6px; border:1px solid #ddd;">Accion requerida</td><td style="padding:6px; border:1px solid #ddd;">Usar el AGENDA.docx adjunto durante la reunion y subirlo luego a SharePoint en minutas_entrada</td></tr>
    <tr><td style="font-weight:bold; padding:6px; border:1px solid #ddd;">Fecha</td><td style="padding:6px; border:1px solid #ddd;">$fecha</td></tr>
    <tr><td style="font-weight:bold; padding:6px; border:1px solid #ddd;">Hora</td><td style="padding:6px; border:1px solid #ddd;">$hora</td></tr>
  </table>
  <p><strong>Carpeta de subida:</strong><br><a href="$uploadLink">$uploadLink</a></p>
  <p><strong>Nombre obligatorio del archivo:</strong><br><code>$FileName</code></p>
  <p style="color:#9a3412;"><strong>Importante:</strong> si el archivo se sube con otro nombre, el flujo automatico puede no procesarlo correctamente.</p>
  <p style="font-size: 12px; color: #777; margin-top: 28px;">Generado por el fiscalizador de reuniones MAYU.</p>
</body>
</html>
"@
}

function Send-MissingAlert {
  param(
    [object]$Config,
    [string]$Token,
    [string]$SiteId,
    [object]$Occurrence
  )

  $tipo = $Occurrence.Tipo
  $fileName = Get-ExpectedFileName -MeetingType $tipo -TargetDate $Occurrence.Fecha
  $attachment = $null
  try {
    $attachment = Get-FileBytes -Token $Token -SiteId $SiteId -FilePath "$($Config.sharepoint.templates_folder)/AGENDA_$tipo.docx"
  } catch {
    Write-Warning "No se pudo adjuntar template AGENDA_$tipo.docx. Se enviara el aviso sin adjunto."
  }

  Send-GraphMail `
    -Token $Token `
    -Sender $Config.mail.sender `
    -To @($Occurrence.Meeting.to) `
    -Cc @($Occurrence.Meeting.cc) `
    -Subject "Alerta MAYU - falta minuta/reporte - $tipo $($Occurrence.Fecha.ToString("yyyy-MM-dd"))" `
    -HtmlBody (New-AlertBody -Config $Config -Occurrence $Occurrence -FileName $fileName) `
    -AttachmentName $(if ($attachment) { "AGENDA_$tipo.docx" } else { "" }) `
    -AttachmentBytes $attachment

  $alertDoc = @{
    key = "$tipo|$($Occurrence.Fecha.ToString("yyyy-MM-dd"))"
    tipo = $tipo
    fecha = $Occurrence.Fecha.ToString("yyyy-MM-dd")
    sent_at = (Get-Date).ToUniversalTime().ToString("o")
    to = @($Occurrence.Meeting.to)
    cc = @($Occurrence.Meeting.cc)
  } | ConvertTo-Json -Depth 10

  try {
    Write-TextFileToGraph `
      -Token $Token `
      -SiteId $SiteId `
      -FilePath "$($Config.sharepoint.alertas_enviadas_folder)/alerta_${tipo}_$($Occurrence.Fecha.ToString("yyyy-MM-dd")).json" `
      -Text $alertDoc `
      -ContentType "application/json; charset=utf-8"
  } catch {
    Write-Warning "La alerta fue enviada, pero no se pudo registrar en SharePoint. $($_.Exception.Message)"
  }
}

function Send-PreMeetingReminder {
  param(
    [object]$Config,
    [string]$Token,
    [string]$SiteId,
    [object]$Occurrence
  )

  $tipo = $Occurrence.Tipo
  $fileName = Get-ExpectedFileName -MeetingType $tipo -TargetDate $Occurrence.Fecha
  $attachment = $null
  try {
    $attachment = Get-FileBytes -Token $Token -SiteId $SiteId -FilePath "$($Config.sharepoint.templates_folder)/AGENDA_$tipo.docx"
  } catch {
    Write-Warning "No se pudo adjuntar template AGENDA_$tipo.docx. Se enviara el recordatorio sin adjunto."
  }

  Send-GraphMail `
    -Token $Token `
    -Sender $Config.mail.sender `
    -To @($Occurrence.Meeting.to) `
    -Cc @($Occurrence.Meeting.cc) `
    -Subject "Recordatorio MAYU - reunion hoy - $tipo $($Occurrence.Fecha.ToString("yyyy-MM-dd"))" `
    -HtmlBody (New-ReminderBody -Config $Config -Occurrence $Occurrence -FileName $fileName) `
    -AttachmentName $(if ($attachment) { "AGENDA_$tipo.docx" } else { "" }) `
    -AttachmentBytes $attachment

  $recordDoc = @{
    key = "$tipo|$($Occurrence.Fecha.ToString("yyyy-MM-dd"))"
    tipo = $tipo
    fecha = $Occurrence.Fecha.ToString("yyyy-MM-dd")
    sent_at = (Get-Date).ToUniversalTime().ToString("o")
    to = @($Occurrence.Meeting.to)
    cc = @($Occurrence.Meeting.cc)
  } | ConvertTo-Json -Depth 10

  try {
    Write-TextFileToGraph `
      -Token $Token `
      -SiteId $SiteId `
      -FilePath "$($Config.sharepoint.recordatorios_enviados_folder)/recordatorio_${tipo}_$($Occurrence.Fecha.ToString("yyyy-MM-dd")).json" `
      -Text $recordDoc `
      -ContentType "application/json; charset=utf-8"
  } catch {
    Write-Warning "El recordatorio fue enviado, pero no se pudo registrar en SharePoint. $($_.Exception.Message)"
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
      max_output_tokens = 1800
    } | ConvertTo-Json -Depth 10

    $response = Invoke-RestMethod `
      -Method Post `
      -Uri "https://api.openai.com/v1/responses" `
      -Headers @{ Authorization = "Bearer $apiKey" } `
      -ContentType "application/json" `
      -Body $body `
      -TimeoutSec 45

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
    Write-Warning "OpenAI no pudo generar el resumen inteligente. Se usara resumen deterministicamente calculado. $($_.Exception.Message)"
    return ""
  }
}

function Get-StatusRows {
  param(
    [object]$Config,
    [object[]]$EvidenceItems,
    [object[]]$AlertItems,
    [datetime]$StartDate,
    [datetime]$EndDate
  )
  $alertKeys = Get-SentAlertKeys -Items $AlertItems
  $rows = @()
  foreach ($occ in (Get-ExpectedOccurrences -Config $Config -StartDate $StartDate -EndDate $EndDate)) {
    $hasEvidence = Test-EvidencePresent -Items $EvidenceItems -MeetingType $occ.Tipo -TargetDate $occ.Fecha
    $key = "$($occ.Tipo)|$($occ.Fecha.ToString("yyyy-MM-dd"))"
    $hasAlert = $false
    if ($null -ne $alertKeys) {
      $hasAlert = $alertKeys.Contains($key)
    }
    $rows += [pscustomobject]@{
      fecha = $occ.Fecha.ToString("yyyy-MM-dd")
      tipo = $occ.Tipo
      reunion = $occ.Nombre
      responsable = [string]$occ.Meeting.responsable
      to = @($occ.Meeting.to)
      cc = @($occ.Meeting.cc)
      estado = $(if ($hasEvidence) { "CUMPLIDA" } else { "PENDIENTE" })
      alerta = $(if ($hasAlert) { "SI" } else { "NO" })
    }
  }
  return $rows
}

function Convert-RowsToHtmlTable {
  param([object[]]$Rows)
  $html = "<table style='border-collapse:collapse;width:100%;font-size:13px;'><tr><th style='text-align:left;border:1px solid #ddd;padding:6px;'>Fecha</th><th style='text-align:left;border:1px solid #ddd;padding:6px;'>Reunion</th><th style='text-align:left;border:1px solid #ddd;padding:6px;'>Responsable</th><th style='text-align:left;border:1px solid #ddd;padding:6px;'>Estado</th><th style='text-align:left;border:1px solid #ddd;padding:6px;'>Alerta</th></tr>"
  foreach ($row in @($Rows)) {
    $color = if ($row.estado -eq "CUMPLIDA") { "#166534" } else { "#991b1b" }
    $html += "<tr><td style='border:1px solid #ddd;padding:6px;'>$($row.fecha)</td><td style='border:1px solid #ddd;padding:6px;'>$($row.reunion)</td><td style='border:1px solid #ddd;padding:6px;'>$($row.responsable)</td><td style='border:1px solid #ddd;padding:6px;color:$color;font-weight:bold;'>$($row.estado)</td><td style='border:1px solid #ddd;padding:6px;'>$($row.alerta)</td></tr>"
  }
  $html += "</table>"
  return $html
}

function New-ReportHtml {
  param([string]$Title, [object[]]$Rows, [string]$AiSummary, [object[]]$IntelligenceRows = @())
  $total = @($Rows).Count
  $done = @($Rows | Where-Object { $_.estado -eq "CUMPLIDA" }).Count
  $pending = $total - $done
  $rate = if ($total -gt 0) { [Math]::Round(($done / $total) * 100, 0) } else { 0 }
  $summary = if ($AiSummary) { "<pre style='white-space:pre-wrap;font-family:Arial,sans-serif;background:#f7f7f7;padding:12px;border:1px solid #ddd;'>$AiSummary</pre>" } else { "<p>No se genero resumen IA; se informa solo el control operativo.</p>" }
  $intelligenceHtml = Convert-IntelligenceRowsToHtml -Rows $IntelligenceRows
  @"
<html>
<body style="font-family: Arial, sans-serif; font-size: 14px; color: #222; max-width: 900px;">
  <h2>$Title</h2>
  <p><strong>Cumplimiento:</strong> $done de $total ($rate%). <strong>Pendientes:</strong> $pending.</p>
  <h3>Lectura inteligente</h3>
  $summary
  <h3>Compromisos, bloqueos y continuidad</h3>
  $intelligenceHtml
  <h3>Detalle fiscalizado</h3>
  $(Convert-RowsToHtmlTable -Rows $Rows)
  <p style="font-size:12px;color:#777;margin-top:24px;">Generado por el fiscalizador de reuniones MAYU.</p>
</body>
</html>
"@
}

function Send-ControlReport {
  param(
    [object]$Config,
    [string]$Token,
    [string]$SiteId,
    [string]$Title,
    [string]$FileName,
    [object[]]$Rows,
    [object[]]$IntelligenceRows = @()
  )

  $jsonRows = $Rows | ConvertTo-Json -Depth 10
  $jsonIntelligenceRows = @($IntelligenceRows) | ConvertTo-Json -Depth 20
  Write-Output "Reporte: filas fiscalizadas $(@($Rows).Count)."
  Write-Output "Reporte: seguimientos inteligentes $(@($IntelligenceRows).Count)."
  $prompt = @"
Eres el Agente de Continuidad Operacional de MAYU. Tu tarea es escribir un resumen ejecutivo breve para Felix Escudero, GG.
No inventes reuniones ni personas. Lee estos datos y resume:
- que se cumplio,
- que no se cumplio,
- donde hubo alertas,
- que compromisos se cumplieron o quedaron pendientes,
- que bloqueos requieren intervencion,
- que seguimiento conviene hacer esta semana.

Control documental:
$jsonRows

Seguimiento de compromisos y bloqueos:
$jsonIntelligenceRows
"@
  Write-Output "Reporte: generando lectura inteligente."
  $ai = Invoke-OpenAiText -Prompt $prompt
  $html = New-ReportHtml -Title $Title -Rows $Rows -AiSummary $ai -IntelligenceRows $IntelligenceRows

  Write-Output "Reporte: enviando correo a $($Config.mail.felix)."
  Send-GraphMail `
    -Token $Token `
    -Sender $Config.mail.sender `
    -To @($Config.mail.felix) `
    -Cc @() `
    -Subject $Title `
    -HtmlBody $html

  try {
    Write-Output "Reporte: guardando copia en SharePoint."
    Ensure-GraphFolder -Token $Token -SiteId $SiteId -FolderPath $Config.sharepoint.reportes_folder
    Write-TextFileToGraph -Token $Token -SiteId $SiteId -FilePath "$($Config.sharepoint.reportes_folder)/$FileName" -Text $html -ContentType "text/html; charset=utf-8"
  } catch {
    Write-Warning "El reporte fue enviado por correo, pero no se pudo guardar copia en SharePoint. $($_.Exception.Message)"
  }
}

function Invoke-PostReunion {
  param(
    [object]$Config,
    [string]$Token,
    [string]$SiteId,
    [datetime]$Now,
    [string]$TargetTipo,
    [string]$TargetDate
  )

  $checkDate = if ($TargetDate) { [datetime]::Parse($TargetDate) } else { $Now.Date }
  $occurrences = Get-ExpectedOccurrences -Config $Config -StartDate $checkDate -EndDate $checkDate
  $occurrences = @($occurrences | Where-Object { $Now -ge $_.DueAt })
  if ($TargetTipo) {
    $occurrences = @($occurrences | Where-Object { $_.Tipo -eq $TargetTipo })
  }

  $evidence = @()
  $evidence += Get-FolderItemsRecursive -Token $Token -SiteId $SiteId -FolderPath $Config.sharepoint.minutas_entrada
  $evidence += Get-FolderItemsRecursive -Token $Token -SiteId $SiteId -FolderPath $Config.sharepoint.minutas_archivadas

  Ensure-GraphFolder -Token $Token -SiteId $SiteId -FolderPath $Config.sharepoint.alertas_enviadas_folder
  $alertItems = Get-FolderItemsRecursive -Token $Token -SiteId $SiteId -FolderPath $Config.sharepoint.alertas_enviadas_folder
  $sentKeys = Get-SentAlertKeys -Items $alertItems
  if ($null -eq $sentKeys) {
    $sentKeys = New-Object System.Collections.Generic.HashSet[string]
  }

  $sentCount = 0
  foreach ($occ in @($occurrences)) {
    $key = "$($occ.Tipo)|$($occ.Fecha.ToString("yyyy-MM-dd"))"
    if ($sentKeys.Contains($key)) {
      Write-Output "Ya existia alerta: $key"
      continue
    }
    if (Test-EvidencePresent -Items $evidence -MeetingType $occ.Tipo -TargetDate $occ.Fecha) {
      Write-Output "OK evidencia encontrada: $key"
      continue
    }
    Send-MissingAlert -Config $Config -Token $Token -SiteId $SiteId -Occurrence $occ
    Write-Output "Alerta enviada: $key"
    $sentCount += 1
  }
  Write-Output "Revision post reunion completa. Alertas enviadas: $sentCount"
}

function Invoke-PreReunion {
  param(
    [object]$Config,
    [string]$Token,
    [string]$SiteId,
    [datetime]$Now,
    [string]$TargetTipo,
    [string]$TargetDate
  )

  $checkDate = if ($TargetDate) { [datetime]::Parse($TargetDate) } else { $Now.Date }
  $occurrences = Get-ExpectedOccurrences -Config $Config -StartDate $checkDate -EndDate $checkDate
  if ($TargetTipo) {
    $occurrences = @($occurrences | Where-Object { $_.Tipo -eq $TargetTipo })
  }

  Ensure-GraphFolder -Token $Token -SiteId $SiteId -FolderPath $Config.sharepoint.recordatorios_enviados_folder
  $reminderItems = Get-FolderItemsRecursive -Token $Token -SiteId $SiteId -FolderPath $Config.sharepoint.recordatorios_enviados_folder
  $sentKeys = Get-SentReminderKeys -Items $reminderItems
  if ($null -eq $sentKeys) {
    $sentKeys = New-Object System.Collections.Generic.HashSet[string]
  }

  $sentCount = 0
  foreach ($occ in @($occurrences)) {
    $key = "$($occ.Tipo)|$($occ.Fecha.ToString("yyyy-MM-dd"))"
    if ($sentKeys.Contains($key)) {
      Write-Output "Ya existia recordatorio: $key"
      continue
    }
    Send-PreMeetingReminder -Config $Config -Token $Token -SiteId $SiteId -Occurrence $occ
    Write-Output "Recordatorio enviado: $key"
    $sentCount += 1
  }
  Write-Output "Revision pre reunion completa. Recordatorios enviados: $sentCount"
}

function Invoke-ManualDueSweep {
  param([object]$Config, [string]$Token, [string]$SiteId, [datetime]$Now, [int]$Days)

  $start = $Now.Date.AddDays(-1 * [Math]::Max(1, $Days))
  $end = $Now.Date
  Write-Output "Barrido manual: periodo $($start.ToString("yyyy-MM-dd")) a $($end.ToString("yyyy-MM-dd"))."
  Write-Output "Barrido manual: leyendo evidencia SharePoint."
  $evidence = @()
  $evidence += Get-FolderItemsFlatSafe -Token $Token -SiteId $SiteId -FolderPath $Config.sharepoint.minutas_entrada
  $evidence += Get-FolderItemsFlatSafe -Token $Token -SiteId $SiteId -FolderPath $Config.sharepoint.minutas_archivadas
  $evidence += Get-FolderItemsFlatSafe -Token $Token -SiteId $SiteId -FolderPath $Config.sharepoint.reportes_folder

  Ensure-GraphFolder -Token $Token -SiteId $SiteId -FolderPath $Config.sharepoint.alertas_enviadas_folder
  $alertItems = Get-FolderItemsFlatSafe -Token $Token -SiteId $SiteId -FolderPath $Config.sharepoint.alertas_enviadas_folder
  $sentKeys = Get-SentAlertKeys -Items $alertItems
  if ($null -eq $sentKeys) {
    $sentKeys = New-Object System.Collections.Generic.HashSet[string]
  }

  $sentCount = 0
  foreach ($occ in (Get-ExpectedOccurrences -Config $Config -StartDate $start -EndDate $end)) {
    if ($Now -lt $occ.DueAt) {
      continue
    }
    $key = "$($occ.Tipo)|$($occ.Fecha.ToString("yyyy-MM-dd"))"
    if ($sentKeys.Contains($key)) {
      continue
    }
    if (Test-EvidencePresent -Items $evidence -MeetingType $occ.Tipo -TargetDate $occ.Fecha) {
      continue
    }
    Send-MissingAlert -Config $Config -Token $Token -SiteId $SiteId -Occurrence $occ
    Write-Output "Alerta manual enviada: $key"
    $sentCount += 1
  }
  Write-Output "Barrido manual completo. Alertas enviadas: $sentCount"
}

function Invoke-WeeklyReport {
  param([object]$Config, [string]$Token, [string]$SiteId, [datetime]$Now)
  $anchor = if ($Date) { [datetime]::Parse($Date) } else { $Now }
  $daysSinceMonday = (([int]$anchor.DayOfWeek + 6) % 7)
  $thisMonday = $anchor.Date.AddDays(-1 * $daysSinceMonday)
  $start = $thisMonday.AddDays(-7)
  $end = $thisMonday.AddDays(-1)

  Write-Output "Reporte semanal: periodo $($start.ToString("yyyy-MM-dd")) a $($end.ToString("yyyy-MM-dd"))."
  Write-Output "Reporte semanal: leyendo evidencia SharePoint."
  $evidence = @()
  $evidence += Get-FolderItemsFlatSafe -Token $Token -SiteId $SiteId -FolderPath $Config.sharepoint.minutas_entrada
  $evidence += Get-FolderItemsFlatSafe -Token $Token -SiteId $SiteId -FolderPath $Config.sharepoint.minutas_archivadas
  $evidence += Get-FolderItemsFlatSafe -Token $Token -SiteId $SiteId -FolderPath $Config.sharepoint.reportes_folder
  $alertItems = Get-FolderItemsFlatSafe -Token $Token -SiteId $SiteId -FolderPath $Config.sharepoint.alertas_enviadas_folder
  $intelligenceItems = Get-FolderItemsFlatSafe -Token $Token -SiteId $SiteId -FolderPath (Get-IntelligenceFolder -Config $Config)
  $occurrences = Get-ExpectedOccurrences -Config $Config -StartDate $start -EndDate $end
  $rows = Get-StatusRows -Config $Config -EvidenceItems $evidence -AlertItems $alertItems -StartDate $start -EndDate $end
  $intelligenceRows = Get-MeetingIntelligenceRows -Config $Config -Token $Token -SiteId $SiteId -Occurrences $occurrences -EvidenceItems $evidence -ExistingIntelligenceItems $intelligenceItems
  Write-Output "Reporte semanal: filas calculadas $(@($rows).Count)."
  Send-ControlReport -Config $Config -Token $Token -SiteId $SiteId -Title "Reporte semanal reuniones MAYU - $($start.ToString("yyyy-MM-dd")) a $($end.ToString("yyyy-MM-dd"))" -FileName "reporte_semanal_$($end.ToString("yyyy-MM-dd")).html" -Rows $rows -IntelligenceRows $intelligenceRows
  Send-ResponsibleIntelligenceReports -Config $Config -Token $Token -Title "Seguimiento semanal compromisos MAYU - $($start.ToString("yyyy-MM-dd")) a $($end.ToString("yyyy-MM-dd"))" -Rows $intelligenceRows
  Write-Output "Reporte semanal enviado."
}

function Invoke-MonthlyReport {
  param([object]$Config, [string]$Token, [string]$SiteId, [datetime]$Now)
  $anchor = if ($Date) { [datetime]::Parse($Date) } else { $Now }
  if (-not $Date -and $anchor.Date -ne $anchor.Date.AddDays(1).AddDays(-1 * $anchor.Date.AddDays(1).Day).Date) {
    Write-Output "Hoy no es cierre de mes. No se envia reporte mensual."
    return
  }
  $start = Get-Date -Year $anchor.Year -Month $anchor.Month -Day 1 -Hour 0 -Minute 0 -Second 0
  $end = $start.AddMonths(1).AddDays(-1)

  $evidence = @()
  $evidence += Get-FolderItemsRecursive -Token $Token -SiteId $SiteId -FolderPath $Config.sharepoint.minutas_entrada
  $evidence += Get-FolderItemsRecursive -Token $Token -SiteId $SiteId -FolderPath $Config.sharepoint.minutas_archivadas
  $alertItems = Get-FolderItemsRecursive -Token $Token -SiteId $SiteId -FolderPath $Config.sharepoint.alertas_enviadas_folder
  $intelligenceItems = Get-FolderItemsRecursive -Token $Token -SiteId $SiteId -FolderPath (Get-IntelligenceFolder -Config $Config)
  $occurrences = Get-ExpectedOccurrences -Config $Config -StartDate $start -EndDate $end
  $rows = Get-StatusRows -Config $Config -EvidenceItems $evidence -AlertItems $alertItems -StartDate $start -EndDate $end
  $intelligenceRows = Get-MeetingIntelligenceRows -Config $Config -Token $Token -SiteId $SiteId -Occurrences $occurrences -EvidenceItems $evidence -ExistingIntelligenceItems $intelligenceItems
  Send-ControlReport -Config $Config -Token $Token -SiteId $SiteId -Title "Cierre mensual reuniones MAYU - $($start.ToString("yyyy-MM"))" -FileName "reporte_mensual_$($start.ToString("yyyy-MM")).html" -Rows $rows -IntelligenceRows $intelligenceRows
  $scorecards = Send-MonthlyScorecards -Config $Config -Token $Token -SiteId $SiteId -StatusRows $rows -IntelligenceRows $intelligenceRows -StartDate $start
  Write-Output "Scorecards mensuales enviados: $(@($scorecards).Count)."
  Write-Output "Reporte mensual enviado."
}

$configJson = Get-RunbookVariable "MayuFiscalizadorConfigJson"
$config = $configJson | ConvertFrom-Json
$now = Get-Mayutime -TimeZoneName $config.timezone
$token = Get-GraphToken
$siteId = Get-SiteId -Token $token -HostName $config.sharepoint.host

Write-Output "Fiscalizador MAYU iniciado. Modo=$Mode Tipo=$Tipo Fecha=$Date HoraMAYU=$($now.ToString("s"))"

Ensure-GraphFolder -Token $token -SiteId $siteId -FolderPath $config.sharepoint.alertas_enviadas_folder
Ensure-GraphFolder -Token $token -SiteId $siteId -FolderPath $config.sharepoint.recordatorios_enviados_folder
Ensure-GraphFolder -Token $token -SiteId $siteId -FolderPath $config.sharepoint.evaluaciones_folder
Ensure-GraphFolder -Token $token -SiteId $siteId -FolderPath (Get-IntelligenceFolder -Config $config)
Ensure-GraphFolder -Token $token -SiteId $siteId -FolderPath $config.sharepoint.reportes_folder

if ($Mode -eq "pre_reunion") {
  Invoke-PreReunion -Config $config -Token $token -SiteId $siteId -Now $now -TargetTipo $Tipo -TargetDate $Date
} elseif ($Mode -eq "post_reunion") {
  Invoke-PostReunion -Config $config -Token $token -SiteId $siteId -Now $now -TargetTipo $Tipo -TargetDate $Date
} elseif ($Mode -eq "manual_due_sweep") {
  try {
    Invoke-ManualDueSweep -Config $config -Token $token -SiteId $siteId -Now $now -Days $LookbackDays
  } catch {
    $errorText = $_.Exception.Message
    $positionText = ""
    if ($_.InvocationInfo -and $_.InvocationInfo.PositionMessage) {
      $positionText = [string]$_.InvocationInfo.PositionMessage
    }
    $stackText = [string]$_.ScriptStackTrace
    Write-Error "Fallo manual_due_sweep: $errorText`n$positionText`n$stackText"
    throw
  }
} elseif ($Mode -eq "weekly_report") {
  try {
    Invoke-WeeklyReport -Config $config -Token $token -SiteId $siteId -Now $now
  } catch {
    $errorText = $_.Exception.Message
    $positionText = ""
    if ($_.InvocationInfo -and $_.InvocationInfo.PositionMessage) {
      $positionText = [string]$_.InvocationInfo.PositionMessage
    }
    $stackText = [string]$_.ScriptStackTrace
    Write-Error "Fallo weekly_report: $errorText`n$positionText`n$stackText"
    try {
      Send-GraphMail `
        -Token $token `
        -Sender $config.mail.sender `
        -To @($config.mail.felix) `
        -Cc @() `
        -Subject "Error reporte semanal fiscalizador MAYU" `
        -HtmlBody "<html><body><p>El fiscalizador no pudo generar el reporte semanal.</p><p><strong>Error:</strong> $errorText</p><pre>$positionText`n$stackText</pre><p>Hora MAYU: $($now.ToString("yyyy-MM-dd HH:mm"))</p></body></html>"
    } catch {
      Write-Warning "Tampoco se pudo enviar correo de diagnostico. $($_.Exception.Message)"
    }
    throw
  }
} elseif ($Mode -eq "monthly_report") {
  Invoke-MonthlyReport -Config $config -Token $token -SiteId $siteId -Now $now
} elseif ($Mode -eq "test") {
  Send-GraphMail `
    -Token $token `
    -Sender $config.mail.sender `
    -To @($config.mail.felix) `
    -Cc @() `
    -Subject "Prueba fiscalizador reuniones MAYU" `
    -HtmlBody "<html><body><p>Fiscalizador MAYU operativo en GitHub Actions.</p><p>Hora MAYU: $($now.ToString("yyyy-MM-dd HH:mm"))</p></body></html>"
  Write-Output "Prueba enviada a $($config.mail.felix)."
}
