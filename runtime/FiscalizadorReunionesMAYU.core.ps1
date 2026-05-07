param(
  [ValidateSet("post_reunion", "manual_due_sweep", "weekly_report", "monthly_report", "test")]
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

  $value = $null
  $cmd = Get-Command -Name Get-AutomationVariable -ErrorAction SilentlyContinue
  if ($cmd) {
    try {
      $value = Get-AutomationVariable -Name $Name
    } catch {
      $value = $null
    }
  }

  if ([string]::IsNullOrWhiteSpace([string]$value)) {
    $value = [Environment]::GetEnvironmentVariable($Name)
  }
  if ([string]::IsNullOrWhiteSpace([string]$value)) {
    $value = $Default
  }
  if ($Required -and [string]::IsNullOrWhiteSpace([string]$value)) {
    throw "Falta variable de Automation: $Name"
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

function Get-ExpectedFileName {
  param([string]$MeetingType, [datetime]$TargetDate)
  "AGENDA_{0}_{1}_{2}_{3}.docx" -f $MeetingType, $TargetDate.ToString("dd"), $TargetDate.ToString("MM"), $TargetDate.ToString("yyyy")
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
  $expected = (Get-ExpectedFileName -MeetingType $MeetingType -TargetDate $TargetDate).ToLowerInvariant()
  $processed = ("minuta_{0}_{1}.html" -f $MeetingType, $TargetDate.ToString("yyyy-MM-dd")).ToLowerInvariant()
  $processedMd = ("{0}.md" -f $TargetDate.ToString("yyyy-MM-dd")).ToLowerInvariant()

  foreach ($item in @($Items)) {
    if ($null -eq $item) {
      continue
    }
    $name = ([string]$item.name).ToLowerInvariant()
    $path = ([string]$item._mayuRelativePath).ToLowerInvariant()
    if ($name -eq $expected -or $name -eq $processed) {
      return $true
    }
    if ($path -like "*$MeetingType*" -and $name -eq $processedMd) {
      return $true
    }
  }
  return $false
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
      max_output_tokens = 900
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
  param([string]$Title, [object[]]$Rows, [string]$AiSummary)
  $total = @($Rows).Count
  $done = @($Rows | Where-Object { $_.estado -eq "CUMPLIDA" }).Count
  $pending = $total - $done
  $rate = if ($total -gt 0) { [Math]::Round(($done / $total) * 100, 0) } else { 0 }
  $summary = if ($AiSummary) { "<pre style='white-space:pre-wrap;font-family:Arial,sans-serif;background:#f7f7f7;padding:12px;border:1px solid #ddd;'>$AiSummary</pre>" } else { "<p>No se genero resumen IA; se informa solo el control operativo.</p>" }
  @"
<html>
<body style="font-family: Arial, sans-serif; font-size: 14px; color: #222; max-width: 900px;">
  <h2>$Title</h2>
  <p><strong>Cumplimiento:</strong> $done de $total ($rate%). <strong>Pendientes:</strong> $pending.</p>
  <h3>Lectura inteligente</h3>
  $summary
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
    [object[]]$Rows
  )

  $jsonRows = $Rows | ConvertTo-Json -Depth 10
  Write-Output "Reporte: filas fiscalizadas $(@($Rows).Count)."
  $prompt = @"
Eres el fiscalizador de reuniones de MAYU. Tu tarea es escribir un resumen ejecutivo breve para Felix Escudero, GG.
No inventes reuniones ni personas. Lee estos datos y resume:
- que se cumplio,
- que no se cumplio,
- donde hubo alertas,
- que seguimiento conviene hacer esta semana.

Datos:
$jsonRows
"@
  Write-Output "Reporte: generando lectura inteligente."
  $ai = Invoke-OpenAiText -Prompt $prompt
  $html = New-ReportHtml -Title $Title -Rows $Rows -AiSummary $ai

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
  $rows = Get-StatusRows -Config $Config -EvidenceItems $evidence -AlertItems $alertItems -StartDate $start -EndDate $end
  Write-Output "Reporte semanal: filas calculadas $(@($rows).Count)."
  Send-ControlReport -Config $Config -Token $Token -SiteId $SiteId -Title "Reporte semanal reuniones MAYU - $($start.ToString("yyyy-MM-dd")) a $($end.ToString("yyyy-MM-dd"))" -FileName "reporte_semanal_$($end.ToString("yyyy-MM-dd")).html" -Rows $rows
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
  $rows = Get-StatusRows -Config $Config -EvidenceItems $evidence -AlertItems $alertItems -StartDate $start -EndDate $end
  Send-ControlReport -Config $Config -Token $Token -SiteId $SiteId -Title "Cierre mensual reuniones MAYU - $($start.ToString("yyyy-MM"))" -FileName "reporte_mensual_$($start.ToString("yyyy-MM")).html" -Rows $rows
  Write-Output "Reporte mensual enviado."
}

$configJson = Get-RunbookVariable "MayuFiscalizadorConfigJson"
$config = $configJson | ConvertFrom-Json
$now = Get-Mayutime -TimeZoneName $config.timezone
$token = Get-GraphToken
$siteId = Get-SiteId -Token $token -HostName $config.sharepoint.host

Write-Output "Fiscalizador MAYU iniciado. Modo=$Mode Tipo=$Tipo Fecha=$Date HoraMAYU=$($now.ToString("s"))"

Ensure-GraphFolder -Token $token -SiteId $siteId -FolderPath $config.sharepoint.alertas_enviadas_folder
Ensure-GraphFolder -Token $token -SiteId $siteId -FolderPath $config.sharepoint.evaluaciones_folder
Ensure-GraphFolder -Token $token -SiteId $siteId -FolderPath $config.sharepoint.reportes_folder

if ($Mode -eq "post_reunion") {
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
    -HtmlBody "<html><body><p>Fiscalizador MAYU operativo en Azure Automation.</p><p>Hora MAYU: $($now.ToString("yyyy-MM-dd HH:mm"))</p></body></html>"
  Write-Output "Prueba enviada a $($config.mail.felix)."
}
