param(
  [ValidateSet("morning_reports", "daily_pulse", "bodega_materiales", "bodega_materiales_respuestas", "finanzas", "finanzas_admin", "finanzas_dte_inbox", "finanzas_respuestas", "bice_cartola_mail", "test")]
  [string]$Mode = "daily_pulse",
  [string]$Date = "",
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
  [string]$value
}

function ConvertTo-DrivePath {
  param([string]$Path)
  if ([string]::IsNullOrWhiteSpace($Path)) { return "" }
  (($Path.Trim("/") -split "/") | Where-Object { $_ } | ForEach-Object { [Uri]::EscapeDataString($_) }) -join "/"
}

function Invoke-GraphGet {
  param([string]$Token, [string]$Uri)
  Invoke-RestMethod -Method Get -Uri $Uri -Headers @{ Authorization = "Bearer $Token" }
}

function Invoke-GraphJson {
  param([string]$Token, [string]$Method, [string]$Uri, [object]$Body)
  $json = $Body | ConvertTo-Json -Depth 40
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
  $token.access_token
}

function Get-SiteId {
  param([string]$Token, [string]$HostName)
  (Invoke-GraphGet -Token $Token -Uri "https://graph.microsoft.com/v1.0/sites/$HostName").id
}

function Ensure-GraphFolder {
  param([string]$Token, [string]$SiteId, [string]$FolderPath)
  if ([string]::IsNullOrWhiteSpace($FolderPath)) { return }
  $parts = @($FolderPath.Trim("/") -split "/" | Where-Object { $_ })
  $current = ""
  foreach ($part in $parts) {
    $candidate = if ($current) { "$current/$part" } else { $part }
    $encoded = ConvertTo-DrivePath $candidate
    $exists = $true
    try { Invoke-GraphGet -Token $Token -Uri "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/root:/$encoded" | Out-Null } catch { $exists = $false }
    if (-not $exists) {
      $parentEncoded = ConvertTo-DrivePath $current
      $uri = if ($parentEncoded) {
        "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/root:/$parentEncoded" + ":/children"
      } else {
        "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/root/children"
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

function Write-TextFileToGraph {
  param([string]$Token, [string]$SiteId, [string]$FilePath, [string]$Text, [string]$ContentType)
  $encoded = ConvertTo-DrivePath $FilePath
  $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/root:/$encoded" + ":/content"
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
  Invoke-GraphPutBytes -Token $Token -Uri $uri -Bytes $bytes -ContentType $ContentType | Out-Null
}

function Write-BytesFileToGraph {
  param([string]$Token, [string]$SiteId, [string]$FilePath, [byte[]]$Bytes, [string]$ContentType)
  if ([string]::IsNullOrWhiteSpace($ContentType)) { $ContentType = "application/octet-stream" }
  $encoded = ConvertTo-DrivePath $FilePath
  $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/root:/$encoded" + ":/content"
  Invoke-GraphPutBytes -Token $Token -Uri $uri -Bytes $Bytes -ContentType $ContentType | Out-Null
}

function Read-TextFileFromGraph {
  param([string]$Token, [string]$SiteId, [string]$FilePath)
  $encoded = ConvertTo-DrivePath $FilePath
  $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/root:/$encoded" + ":/content"
  try {
    $result = Invoke-GraphGet -Token $Token -Uri $uri
    if ($result -is [string]) { return $result }
    return ($result | ConvertTo-Json -Depth 20)
  } catch {
    return ""
  }
}

function Send-GraphMail {
  param(
    [string]$Token,
    [string]$Sender,
    [string[]]$To,
    [string[]]$Cc,
    [string]$Subject,
    [string]$HtmlBody
  )
  $message = @{
    subject = $Subject
    body = @{ contentType = "HTML"; content = $HtmlBody }
    toRecipients = @($To | Where-Object { $_ } | ForEach-Object { @{ emailAddress = @{ address = $_ } } })
    ccRecipients = @($Cc | Where-Object { $_ } | ForEach-Object { @{ emailAddress = @{ address = $_ } } })
  }
  Invoke-GraphJson -Token $Token -Method Post -Uri "https://graph.microsoft.com/v1.0/users/$Sender/sendMail" -Body @{
    message = $message
    saveToSentItems = $true
  } | Out-Null
}

function Get-UniqueEmails {
  param(
    [string[]]$Emails,
    [string[]]$Exclude = @()
  )
  $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($email in @($Exclude | Where-Object { $_ })) {
    [void]$seen.Add(([string]$email).Trim())
  }
  $result = [System.Collections.ArrayList]::new()
  foreach ($email in @($Emails | Where-Object { $_ })) {
    $clean = ([string]$email).Trim()
    if ($clean -and -not $seen.Contains($clean)) {
      [void]$seen.Add($clean)
      [void]$result.Add($clean)
    }
  }
  return @($result)
}

function Get-BodegaMaterialesMailAudience {
  param([object]$Config)
  return Get-UniqueEmails -Emails @(
    $Config.mail.felix,
    $Config.mail.valentina,
    $Config.mail.carlos,
    $Config.mail.mauricio
  )
}

function Reply-GraphMail {
  param(
    [string]$Token,
    [string]$Mailbox,
    [string]$MessageId,
    [string]$HtmlBody
  )
  Invoke-GraphJson -Token $Token -Method Post -Uri "https://graph.microsoft.com/v1.0/users/$Mailbox/messages/$MessageId/reply" -Body @{
    comment = $HtmlBody
  } | Out-Null
}

function Set-GraphMailRead {
  param([string]$Token, [string]$Mailbox, [string]$MessageId)
  Invoke-GraphJson -Token $Token -Method Patch -Uri "https://graph.microsoft.com/v1.0/users/$Mailbox/messages/$MessageId" -Body @{
    isRead = $true
  } | Out-Null
}

function Disable-GraphMailboxAutoReplies {
  param([string]$Token, [string]$Mailbox)
  try {
    Invoke-GraphJson -Token $Token -Method Patch -Uri "https://graph.microsoft.com/v1.0/users/$Mailbox/mailboxSettings" -Body @{
      automaticRepliesSetting = @{
        status = "disabled"
      }
    } | Out-Null
    Write-Output "Mailbox ${Mailbox}: respuestas automaticas desactivadas."
  } catch {
    Write-Output "Mailbox ${Mailbox}: no se pudieron desactivar respuestas automaticas ($($_.Exception.Message))."
  }
}

function Get-Mayutime {
  param([string]$TimeZoneName)
  $candidateIds = @($TimeZoneName, "Pacific SA Standard Time", "Chile Standard Time")
  foreach ($id in $candidateIds) {
    if ([string]::IsNullOrWhiteSpace($id)) { continue }
    try {
      $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById($id)
      return [System.TimeZoneInfo]::ConvertTime([DateTimeOffset]::UtcNow, $tz).DateTime
    } catch { continue }
  }
  [DateTime]::UtcNow.AddHours(-4)
}

function HtmlEscape {
  param([object]$Value)
  [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Get-MayuEmailTone {
  param([string]$Tone)
  switch (([string]$Tone).ToLowerInvariant()) {
    "rojo" { return [pscustomobject]@{ label = "Rojo"; accent = "#d92d20"; bg = "#fff5f5"; soft = "#fef3f2" } }
    "critico" { return [pscustomobject]@{ label = "Critico"; accent = "#d92d20"; bg = "#fff5f5"; soft = "#fef3f2" } }
    "amarillo" { return [pscustomobject]@{ label = "Amarillo"; accent = "#d97706"; bg = "#fffbeb"; soft = "#fef7e0" } }
    "alto" { return [pscustomobject]@{ label = "Alto"; accent = "#d97706"; bg = "#fffbeb"; soft = "#fef7e0" } }
    "medio" { return [pscustomobject]@{ label = "Medio"; accent = "#ca8a04"; bg = "#fffbeb"; soft = "#fef7e0" } }
    "bajo" { return [pscustomobject]@{ label = "Bajo"; accent = "#2563eb"; bg = "#f4f7fb"; soft = "#eef4ff" } }
    "verde" { return [pscustomobject]@{ label = "Verde"; accent = "#168a50"; bg = "#f0fdf4"; soft = "#ecfdf3" } }
    default { return [pscustomobject]@{ label = "Info"; accent = "#2563eb"; bg = "#f4f7fb"; soft = "#eef4ff" } }
  }
}

function New-MayuEmailMetric {
  param([string]$Label, [object]$Value, [string]$Tone = "info")
  $toneInfo = Get-MayuEmailTone -Tone $Tone
  "<td style='padding:0 8px 8px 0;vertical-align:top;'><div style='border:1px solid #e5e7eb;border-left:4px solid $($toneInfo.accent);background:#ffffff;padding:10px 12px;min-width:110px;'><div style='font-size:11px;line-height:1.25;color:#6b7280;text-transform:uppercase;letter-spacing:.3px;'>$(HtmlEscape $Label)</div><div style='font-size:22px;line-height:1.15;color:#202124;font-weight:700;margin-top:3px;'>$(HtmlEscape $Value)</div></div></td>"
}

function New-MayuEmailSection {
  param([string]$Title, [string]$Html)
  @"
<div style="margin-top:22px;">
  <h3 style="font-size:16px;line-height:1.3;color:#202124;margin:0 0 10px 0;">$(HtmlEscape $Title)</h3>
  $Html
</div>
"@
}

function New-MayuEmptyState {
  param([string]$Text)
  "<div style='border:1px solid #d9eadf;border-left:4px solid #168a50;background:#f7fbf8;padding:12px 14px;color:#234234;'>$(HtmlEscape $Text)</div>"
}

function New-MayuEmailLayout {
  param([string]$Title, [string]$Subtitle, [string]$ContentHtml, [string]$Footer = "Generado por MAYU Agents.")
  @"
<html>
<body style="margin:0;padding:0;background:#f6f8fb;font-family:Arial,sans-serif;color:#202124;font-size:14px;line-height:1.45;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f6f8fb;padding:20px 0;">
    <tr>
      <td align="center">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:980px;background:#ffffff;border:1px solid #e5e7eb;">
          <tr>
            <td style="border-left:4px solid #0078d4;background:#f4f7fb;padding:18px 22px;">
              <h2 style="font-size:22px;line-height:1.25;color:#202124;margin:0 0 4px 0;">$(HtmlEscape $Title)</h2>
              <p style="margin:0;color:#5f6368;">$(HtmlEscape $Subtitle)</p>
            </td>
          </tr>
          <tr>
            <td style="padding:20px 22px;">
              $ContentHtml
              <p style="font-size:12px;line-height:1.4;color:#777;margin:28px 0 0 0;">$(HtmlEscape $Footer)</p>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>
"@
}

function Get-Number {
  param([object]$Value, [double]$Default = 0)
  if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $Default }
  try { return [double]$Value } catch { return $Default }
}

function Format-Clp {
  param([double]$Value)
  "$([Math]::Round($Value, 0)) CLP"
}

function Get-ConfigApiKey {
  param([object]$Config)
  $primary = [string]$Config.firebase.api_key_env
  $fallback = [string]$Config.firebase.fallback_api_key_env
  $apiKey = ""
  if ($primary) { $apiKey = [Environment]::GetEnvironmentVariable($primary) }
  if ([string]::IsNullOrWhiteSpace($apiKey) -and $fallback) { $apiKey = [Environment]::GetEnvironmentVariable($fallback) }
  if ([string]::IsNullOrWhiteSpace($apiKey)) {
    throw "Falta variable Firebase API key ($primary). Configurala como GitHub Variable, no como secreto de codigo."
  }
  $apiKey
}

function Get-FirebaseIdToken {
  param([string]$ApiKey)
  $body = @{ returnSecureToken = $true } | ConvertTo-Json
  $resp = Invoke-RestMethod -Method Post -Uri "https://identitytoolkit.googleapis.com/v1/accounts:signUp?key=$ApiKey" -ContentType "application/json" -Body $body -TimeoutSec 30
  [string]$resp.idToken
}

function ConvertFrom-FirestoreValue {
  param([object]$Value)
  if ($null -eq $Value) { return $null }
  $props = $Value.PSObject.Properties.Name
  if ($props -contains "stringValue") { return [string]$Value.stringValue }
  if ($props -contains "integerValue") { return [int64]$Value.integerValue }
  if ($props -contains "doubleValue") { return [double]$Value.doubleValue }
  if ($props -contains "booleanValue") { return [bool]$Value.booleanValue }
  if ($props -contains "timestampValue") { return [string]$Value.timestampValue }
  if ($props -contains "nullValue") { return $null }
  if ($props -contains "arrayValue") {
    $values = @()
    if ($Value.arrayValue.values) {
      foreach ($v in @($Value.arrayValue.values)) { $values += ConvertFrom-FirestoreValue $v }
    }
    return $values
  }
  if ($props -contains "mapValue") {
    $hash = [ordered]@{}
    if ($Value.mapValue.fields) {
      foreach ($p in $Value.mapValue.fields.PSObject.Properties) {
        $hash[$p.Name] = ConvertFrom-FirestoreValue $p.Value
      }
    }
    return [pscustomobject]$hash
  }
  $null
}

function ConvertFrom-FirestoreDocument {
  param([object]$Doc)
  $hash = [ordered]@{}
  if ($Doc.fields) {
    foreach ($p in $Doc.fields.PSObject.Properties) {
      $hash[$p.Name] = ConvertFrom-FirestoreValue $p.Value
    }
  }
  $id = ([string]$Doc.name -split "/")[-1]
  $hash["id"] = $id
  [pscustomobject]$hash
}

function Get-FirestoreCollection {
  param(
    [object]$Config,
    [string]$Token,
    [string]$CollectionName,
    [int]$PageSize = 300
  )
  $projectId = [string]$Config.firebase.project_id
  $database = [string]$Config.firebase.database
  if ([string]::IsNullOrWhiteSpace($database)) { $database = "(default)" }
  $encodedDb = [Uri]::EscapeDataString($database)
  $base = "https://firestore.googleapis.com/v1/projects/$projectId/databases/$encodedDb/documents/$CollectionName"
  $items = @()
  $pageToken = ""
  do {
    $uri = "$base`?pageSize=$PageSize"
    if ($pageToken) { $uri += "&pageToken=$([Uri]::EscapeDataString($pageToken))" }
    $resp = Invoke-RestMethod -Method Get -Uri $uri -Headers @{ Authorization = "Bearer $Token" } -TimeoutSec 45
    foreach ($doc in @($resp.documents)) {
      if ($null -ne $doc) { $items += ConvertFrom-FirestoreDocument $doc }
    }
    $pageToken = [string]$resp.nextPageToken
  } while (-not [string]::IsNullOrWhiteSpace($pageToken))
  $items
}

function Get-FirestoreData {
  param([object]$Config)
  $apiKey = Get-ConfigApiKey -Config $Config
  $idToken = Get-FirebaseIdToken -ApiKey $apiKey
  $data = [ordered]@{}
  foreach ($p in $Config.collections.PSObject.Properties) {
    $logical = $p.Name
    $collectionName = [string]$p.Value
    try {
      $data[$logical] = @(Get-FirestoreCollection -Config $Config -Token $idToken -CollectionName $collectionName)
      Write-Host "Firestore: $collectionName -> $(@($data[$logical]).Count) docs."
    } catch {
      Write-Warning "No se pudo leer Firestore '$collectionName'. $($_.Exception.Message)"
      $data[$logical] = @()
    }
  }
  [pscustomobject]$data
}

function ConvertTo-FirestoreValue {
  param([object]$Value)
  if ($null -eq $Value) { return @{ nullValue = $null } }
  if ($Value -is [bool]) { return @{ booleanValue = [bool]$Value } }
  if ($Value -is [int] -or $Value -is [int64] -or $Value -is [long]) { return @{ integerValue = [string]$Value } }
  if ($Value -is [double] -or $Value -is [decimal] -or $Value -is [float]) {
    $rounded = [Math]::Round([double]$Value, 0)
    if ([Math]::Abs(([double]$Value) - $rounded) -lt 0.0001) { return @{ integerValue = [string][int64]$rounded } }
    return @{ doubleValue = [double]$Value }
  }
  if ($Value -is [datetime]) { return @{ timestampValue = ([datetime]$Value).ToUniversalTime().ToString("o") } }
  if ($Value -is [System.Collections.IDictionary]) {
    $fields = [ordered]@{}
    foreach ($key in $Value.Keys) {
      $fields[[string]$key] = ConvertTo-FirestoreValue $Value[$key]
    }
    return @{ mapValue = @{ fields = $fields } }
  }
  if ($Value -is [System.Array]) {
    return @{ arrayValue = @{ values = @($Value | ForEach-Object { ConvertTo-FirestoreValue $_ }) } }
  }
  $props = $Value.PSObject.Properties
  if ($props -and $props.Count -gt 0 -and -not ($Value -is [string])) {
    $fields = [ordered]@{}
    foreach ($p in $props) {
      $fields[$p.Name] = ConvertTo-FirestoreValue $p.Value
    }
    return @{ mapValue = @{ fields = $fields } }
  }
  @{ stringValue = [string]$Value }
}

function ConvertTo-FirestoreFields {
  param([object]$Object)
  $fields = [ordered]@{}
  foreach ($p in $Object.PSObject.Properties) {
    if ($p.Name -eq "id") { continue }
    if ($null -eq $p.Value) { continue }
    $fields[$p.Name] = ConvertTo-FirestoreValue $p.Value
  }
  $fields
}

function Set-FirestoreDocument {
  param(
    [object]$Config,
    [string]$Token,
    [string]$CollectionName,
    [string]$DocumentId,
    [object]$Data
  )
  $projectId = [string]$Config.firebase.project_id
  $database = [string]$Config.firebase.database
  if ([string]::IsNullOrWhiteSpace($database)) { $database = "(default)" }
  $encodedDb = [Uri]::EscapeDataString($database)
  $encodedCollection = [Uri]::EscapeDataString($CollectionName)
  $encodedDoc = [Uri]::EscapeDataString($DocumentId)
  $uri = "https://firestore.googleapis.com/v1/projects/$projectId/databases/$encodedDb/documents/$encodedCollection/$encodedDoc"
  $body = @{ fields = (ConvertTo-FirestoreFields -Object $Data) } | ConvertTo-Json -Depth 80
  Invoke-RestMethod -Method Patch -Uri $uri -Headers @{ Authorization = "Bearer $Token" } -ContentType "application/json" -Body $body -TimeoutSec 45 | Out-Null
}

function Add-FirestoreDocument {
  param(
    [object]$Config,
    [string]$Token,
    [string]$CollectionName,
    [object]$Data
  )
  $projectId = [string]$Config.firebase.project_id
  $database = [string]$Config.firebase.database
  if ([string]::IsNullOrWhiteSpace($database)) { $database = "(default)" }
  $encodedDb = [Uri]::EscapeDataString($database)
  $encodedCollection = [Uri]::EscapeDataString($CollectionName)
  $uri = "https://firestore.googleapis.com/v1/projects/$projectId/databases/$encodedDb/documents/$encodedCollection"
  $body = @{ fields = (ConvertTo-FirestoreFields -Object $Data) } | ConvertTo-Json -Depth 80
  $resp = Invoke-RestMethod -Method Post -Uri $uri -Headers @{ Authorization = "Bearer $Token" } -ContentType "application/json" -Body $body -TimeoutSec 45
  ([string]$resp.name -split "/")[-1]
}

function Get-FirestoreWriteToken {
  param([object]$Config)
  $apiKey = Get-ConfigApiKey -Config $Config
  Get-FirebaseIdToken -ApiKey $apiKey
}

function Get-DocList {
  param([object]$Project, [string]$Area)
  $areaObj = $null
  if ($Project.areas -and $Project.areas.PSObject.Properties[$Area]) { $areaObj = $Project.areas.$Area }
  if ($areaObj -and $areaObj.docs) { return @($areaObj.docs) }
  @()
}

function Find-ControlDoc {
  param([object]$Project, [object]$Requirement)
  $docs = Get-DocList -Project $Project -Area ([string]$Requirement.area)
  $ids = @($Requirement.doc_ids)
  @($docs | Where-Object { $ids -contains $_.id })
}

function Test-DocApproved {
  param([object]$Doc)
  if ($null -eq $Doc) { return $false }
  ([string]$Doc.status) -in @("Aprobado", "Aprobado con observaciones")
}

function Get-RequirementStatus {
  param([object]$Project, [string]$Key, [object]$Requirement)
  $docs = @(Find-ControlDoc -Project $Project -Requirement $Requirement)
  if (@($docs).Count -eq 0) {
    return [pscustomobject]@{ key = $Key; estado = "rojo"; detalle = "Documento no existe en Control"; owner = [string]$Requirement.owner; version = ""; file = "" }
  }
  $missing = @($docs | Where-Object { -not $_.fileUrl })
  $notApproved = @($docs | Where-Object { -not (Test-DocApproved $_) })
  $versions = @($docs | ForEach-Object { [string]$_.version } | Where-Object { $_ -and $_ -ne "-" })
  $files = @($docs | ForEach-Object { [string]$_.originalFileName } | Where-Object { $_ })
  if (@($missing).Count -gt 0) {
    return [pscustomobject]@{ key = $Key; estado = "rojo"; detalle = "$(@($missing).Count) doc(s) sin archivo"; owner = [string]$Requirement.owner; version = ($versions -join ", "); file = ($files -join ", ") }
  }
  if (@($notApproved).Count -gt 0) {
    return [pscustomobject]@{ key = $Key; estado = "amarillo"; detalle = "$(@($notApproved).Count) doc(s) sin aprobacion final"; owner = [string]$Requirement.owner; version = ($versions -join ", "); file = ($files -join ", ") }
  }
  [pscustomobject]@{ key = $Key; estado = "verde"; detalle = "Aprobado"; owner = [string]$Requirement.owner; version = ($versions -join ", "); file = ($files -join ", ") }
}

function New-Issue {
  param(
    [ValidateSet("rojo", "amarillo", "info")]
    [string]$Severity,
    [string]$Area,
    [string]$Title,
    [string]$Detail,
    [string]$Owner = "",
    [string]$Action = "",
    [string]$Ref = ""
  )
  [pscustomobject]@{
    severity = $Severity
    area = $Area
    title = $Title
    detail = $Detail
    owner = $Owner
    action = $Action
    ref = $Ref
  }
}

function New-FinanceIssue {
  param(
    [string]$Code,
    [ValidateSet("rojo", "amarillo", "info")]
    [string]$Severity,
    [string]$Area,
    [string]$Title,
    [string]$Detail,
    [string]$Owner = "",
    [string]$Action = "",
    [string]$Ref = ""
  )
  [pscustomobject]@{
    code = $Code
    severity = $Severity
    area = $Area
    title = $Title
    detail = $Detail
    owner = $Owner
    action = $Action
    ref = $Ref
  }
}

function Add-Issue {
  param([System.Collections.ArrayList]$List, [object]$Issue)
  [void]$List.Add($Issue)
}

function Get-IssueKey {
  param([object]$Issue)
  "$($Issue.severity)|$($Issue.area)|$($Issue.title)|$($Issue.ref)"
}

function Select-UniqueIssues {
  param([object[]]$Issues)
  $seen = New-Object System.Collections.Generic.HashSet[string]
  $out = @()
  foreach ($issue in @($Issues)) {
    if ($null -eq $issue) { continue }
    $key = Get-IssueKey -Issue $issue
    if ($seen.Add($key)) { $out += $issue }
  }
  $out
}

function Test-OperationalMatProject {
  param([object]$MatProject, [object[]]$Units, [object[]]$Packs)
  if ($null -eq $MatProject) { return $false }
  if (@($Units | Where-Object { $_.matProjectId -eq $MatProject.id }).Count -gt 0) { return $true }
  if (@($Packs | Where-Object { $_.matProjectId -eq $MatProject.id }).Count -gt 0) { return $true }
  if ((Get-Number $MatProject.totalPods) -gt 0) { return $true }
  if ($MatProject.podsPorTip) {
    foreach ($prop in $MatProject.podsPorTip.PSObject.Properties) {
      if ((Get-Number $prop.Value) -gt 0) { return $true }
    }
  }
  $false
}

function Get-BibliaProjectRows {
  param([object]$Config, [object[]]$ChkProjects)
  $rows = @()
  foreach ($p in @($ChkProjects)) {
    $reqRows = @()
    foreach ($req in $Config.doc_requirements.PSObject.Properties) {
      $reqRows += Get-RequirementStatus -Project $p -Key $req.Name -Requirement $req.Value
    }
    $rojos = @($reqRows | Where-Object { $_.estado -eq "rojo" }).Count
    $amarillos = @($reqRows | Where-Object { $_.estado -eq "amarillo" }).Count
    $estado = if ($rojos -gt 0) { "rojo" } elseif ($amarillos -gt 0) { "amarillo" } else { "verde" }
    $rows += [pscustomobject]@{
      id = $p.id
      name = $p.name
      client = $p.client
      crmId = $p.crmId
      status = $p.status
      estado = $estado
      rojos = $rojos
      amarillos = $amarillos
      requirements = $reqRows
    }
  }
  $rows
}

function Get-LatestApprovedBomVersion {
  param([object]$ChkProject)
  $bom = @(Find-ControlDoc -Project $ChkProject -Requirement ([pscustomobject]@{ area = "ingenieria"; doc_ids = @("i1") })) | Select-Object -First 1
  if ($bom -and (Test-DocApproved $bom)) { return [string]$bom.version }
  ""
}

function Get-OperationalIssues {
  param([object]$Config, [object]$Data, [object[]]$BibliaRows)
  $issues = [System.Collections.ArrayList]::new()
  $matProjects = @($Data.mat_projects)
  $chkProjects = @($Data.chk_projects)
  $crmProjects = @($Data.crm_projects)
  $recetas = @($Data.fab_proyecto_recetas)
  $units = @($Data.fab_units)
  $packs = @($Data.fab_packs)

  foreach ($mp in $matProjects) {
    $isOperational = Test-OperationalMatProject -MatProject $mp -Units $units -Packs $packs
    if (-not $isOperational) {
      continue
    }
    if (-not $mp.crmId) {
      Add-Issue $issues (New-Issue -Severity "rojo" -Area "Traspaso Control-Operacion" -Title "Proyecto operativo sin vinculo comercial" -Detail "$($mp.name) esta en operacion, pero no esta vinculado al negocio comercial. Finanzas no puede cruzar bien los costos por proyecto." -Owner "Carlos / Valentina" -Action "Vincular el proyecto operativo con su negocio comercial." -Ref $mp.id)
    }
    if (-not $mp.chkId) {
      Add-Issue $issues (New-Issue -Severity "rojo" -Area "Traspaso Control-Operacion" -Title "Proyecto operativo sin biblia vinculada" -Detail "$($mp.name) esta en operacion, pero no esta conectado con Control Documental. No puede sincronizar la biblia del proyecto." -Owner "Carlos / Martin" -Action "Vincularlo al proyecto correcto en Control Documental." -Ref $mp.id)
      continue
    }
    $chk = @($chkProjects | Where-Object { $_.id -eq $mp.chkId } | Select-Object -First 1)
    if (-not $chk) {
      Add-Issue $issues (New-Issue -Severity "rojo" -Area "Traspaso Control-Operacion" -Title "Vinculo documental roto" -Detail "$($mp.name) esta conectado a un proyecto documental que ya no existe o no se encontro." -Owner "Carlos / Martin" -Action "Corregir el vinculo con Control Documental." -Ref $mp.id)
      continue
    }
    $bomVigente = Get-LatestApprovedBomVersion -ChkProject $chk
    if ($bomVigente -and $bomVigente -ne [string]$mp.bomVersionUsada) {
      Add-Issue $issues (New-Issue -Severity "rojo" -Area "Control/Materiales" -Title "BOM aprobado pendiente de traspaso" -Detail "$($mp.name): Control Documental tiene BOM $bomVigente y Materiales todavia usa '$($mp.bomVersionUsada)'." -Owner "Carlos" -Action "Sincronizar la BOM aprobada desde Control Documental." -Ref $mp.id)
    }
    $rec = @($recetas | Where-Object { $_.matProjectId -eq $mp.id -or $_.id -eq $mp.id } | Select-Object -First 1)
    if ($mp.bomVersionUsada -and $rec -and ([string]$rec.bomVersionSincronizada) -ne ([string]$mp.bomVersionUsada)) {
      Add-Issue $issues (New-Issue -Severity "rojo" -Area "Materiales/Fabricacion" -Title "Fabricacion no alineada al BOM vigente" -Detail "$($mp.name): Materiales usa BOM $($mp.bomVersionUsada), pero Fabricacion quedo con '$($rec.bomVersionSincronizada)'." -Owner "Carlos / Martin" -Action "Alinear Fabricacion con la BOM vigente y regenerar packs si corresponde." -Ref $mp.id)
    }
    $unitsMp = @($units | Where-Object { $_.matProjectId -eq $mp.id })
    if ([string]$chk.status -eq "Aprobado para ejecucion" -or [string]$chk.status -eq "Aprobado para ejecución") {
      if (@($unitsMp).Count -eq 0) {
        Add-Issue $issues (New-Issue -Severity "amarillo" -Area "Control/Fabricacion" -Title "Proyecto aprobado sin unidades creadas en fabrica" -Detail "$($mp.name) esta aprobado en Control, pero aun no tiene unidades creadas en Fabricacion." -Owner "Martin / Felipe" -Action "Crear las unidades de Fabricacion si el proyecto ya debe entrar a planta." -Ref $mp.id)
      }
    }
    $activeUnits = @($unitsMp | Where-Object { @("M00","M01","M02","M03","M04","M05","RF_FABRICA") -contains [string]$_.status })
    if (@($activeUnits).Count -gt 0) {
      $bib = @($BibliaRows | Where-Object { $_.id -eq $mp.chkId } | Select-Object -First 1)
      if ($bib -and $bib.estado -ne "verde") {
        Add-Issue $issues (New-Issue -Severity "rojo" -Area "Control/Fabricacion" -Title "Produccion activa con documentacion incompleta" -Detail "$($mp.name) tiene $(@($activeUnits).Count) unidad(es) activas y la biblia del proyecto sigue en $($bib.estado)." -Owner "Martin / Carlos" -Action "Cerrar documentos criticos o definir formalmente como sigue la produccion." -Ref $mp.id)
      }
    }
    $packsMp = @($packs | Where-Object { $_.matProjectId -eq $mp.id })
    if (@($unitsMp).Count -gt 0 -and @($packsMp).Count -eq 0) {
      Add-Issue $issues (New-Issue -Severity "amarillo" -Area "Fabricacion/Packs" -Title "Produccion sin packs generados" -Detail "$($mp.name) tiene unidades creadas, pero aun no tiene packs de materiales." -Owner "Felipe / Mauricio" -Action "Generar packs por modulo desde Fabricacion." -Ref $mp.id)
    }
  }
  @($issues)
}

function Get-PackIssues {
  param([object]$Config, [object]$Data, [datetime]$Now)
  $issues = [System.Collections.ArrayList]::new()
  $catalog = @{}
  foreach ($sku in @($Data.inv_catalogo)) { if ($sku.code) { $catalog[[string]$sku.code] = $sku } elseif ($sku.id) { $catalog[[string]$sku.id] = $sku } }
  $activePacks = @($Data.fab_packs | Where-Object { @("planificado","armado","entregado","recibido") -contains [string]$_.estado })
  $packsSinCodigo = [ordered]@{}
  foreach ($pack in $activePacks) {
    $label = "$($pack.projectName) / $($pack.modulo) / $($pack.fecha)"
    $itemsSinSku = @($pack.items | Where-Object { -not $_.skuCode })
    if ($itemsSinSku.Count -gt 0 -or $pack.recetaIncompleta) {
      $projectKey = if ($pack.matProjectId) { [string]$pack.matProjectId } elseif ($pack.projectName) { [string]$pack.projectName } else { [string]$pack.id }
      if (-not $packsSinCodigo.Contains($projectKey)) {
        $packsSinCodigo[$projectKey] = [pscustomobject]@{
          projectName = if ($pack.projectName) { [string]$pack.projectName } else { $projectKey }
          packCount = 0
          itemCount = 0
          examples = [System.Collections.ArrayList]::new()
        }
      }
      $group = $packsSinCodigo[$projectKey]
      $group.packCount += 1
      $group.itemCount += $itemsSinSku.Count
      if ($group.examples.Count -lt 3) {
        [void]$group.examples.Add("$($pack.modulo) / $($pack.fecha)")
      }
    }
    if ([string]$pack.estado -eq "planificado" -and -not $pack.compromisoEntregaBodega) {
      Add-Issue $issues (New-Issue -Severity "amarillo" -Area "Packs" -Title "Pack sin compromiso de Bodega" -Detail "$label esta planificado sin fecha comprometida por Bodega." -Owner "Mauricio" -Action "Asignar compromiso de entrega en Bodega." -Ref $pack.id)
    }
    if ([string]$pack.estado -eq "armado") {
      foreach ($it in @($pack.items | Where-Object { $_.skuCode })) {
        $sku = $catalog[[string]$it.skuCode]
        $qty = Get-Number $it.qtyArmada
        if ($qty -eq 0) { $qty = Get-Number $it.qtyPlanificada }
        $stock = if ($sku) { Get-Number $sku.stock } else { 0 }
        if (-not $sku -or $stock -lt $qty) {
          Add-Issue $issues (New-Issue -Severity "rojo" -Area "Packs" -Title "Pack armado con stock insuficiente" -Detail "$label tiene productos con stock insuficiente para entregar." -Owner "Mauricio" -Action "Ingresar stock o recepcionar compra antes de entregar." -Ref $pack.id)
          break
        }
      }
    }
    if ([string]$pack.estado -eq "entregado") {
      Add-Issue $issues (New-Issue -Severity "amarillo" -Area "Packs/Fabricacion" -Title "Pack entregado no recibido" -Detail "$label fue entregado a fabrica y falta recepcion conforme." -Owner "Felipe" -Action "Recibir pack en Fabricacion." -Ref $pack.id)
    }
    if ([string]$pack.estado -eq "recibido") {
      Add-Issue $issues (New-Issue -Severity "amarillo" -Area "Packs/Fabricacion" -Title "Pack recibido no cerrado" -Detail "$label esta recibido y falta cierre." -Owner "Felipe" -Action "Cerrar pack o registrar desvio/devolucion." -Ref $pack.id)
    }
  }
  foreach ($projectKey in $packsSinCodigo.Keys) {
    $group = $packsSinCodigo[$projectKey]
    $examples = @($group.examples)
    $exampleText = if ($examples.Count -gt 0) { " Ejemplos: $($examples -join '; ')." } else { "" }
    $itemText = if ($group.itemCount -gt 0) { "$($group.itemCount) producto(s) sin codigo de bodega" } else { "receta incompleta" }
    Add-Issue $issues (New-Issue -Severity "rojo" -Area "Packs" -Title "Packs no armables por productos sin definir" -Detail "$($group.projectName) tiene $($group.packCount) pack(s) no armable(s) por $itemText.$exampleText" -Owner "Carlos" -Action "Completar codificacion y cotizacion antes de pedir armado a Bodega." -Ref $projectKey)
  }
  @($issues)
}

function Get-AbastecimientoIssues {
  param([object]$Config, [object]$Data)
  $issues = [System.Collections.ArrayList]::new()
  $skuDemand = @{}
  $skuDemandPacks = @{}
  foreach ($pack in @($Data.fab_packs | Where-Object { @("planificado","armado") -contains [string]$_.estado })) {
    foreach ($it in @($pack.items | Where-Object { $_.skuCode })) {
      $code = [string]$it.skuCode
      $qty = Get-Number $it.qtyArmada
      if ($qty -eq 0) { $qty = Get-Number $it.qtyPlanificada }
      if (-not $skuDemand.ContainsKey($code)) {
        $skuDemand[$code] = 0.0
        $skuDemandPacks[$code] = [System.Collections.ArrayList]::new()
      }
      $skuDemand[$code] += $qty
      [void]$skuDemandPacks[$code].Add($pack)
    }
  }
  foreach ($mp in @($Data.mat_projects)) {
    if (-not (Test-OperationalMatProject -MatProject $mp -Units @($Data.fab_units) -Packs @($Data.fab_packs))) {
      continue
    }
    $sinSku = @($mp.items | Where-Object { -not $_.skuCode -and ([string]$_.status) -ne "inactivo" })
    if ($sinSku.Count -gt 0) {
      Add-Issue $issues (New-Issue -Severity "rojo" -Area "Abastecimiento" -Title "BOM con productos sin codificar" -Detail "$($mp.name) tiene $($sinSku.Count) producto(s) de la BOM sin codigo de bodega." -Owner "Carlos" -Action "Completar codificacion y cotizacion para poder comprar y armar packs." -Ref $mp.id)
    }
  }
  $stockGaps = @()
  foreach ($sku in @($Data.inv_catalogo | Where-Object { $_.activo -ne $false -and -not $_.aliasDe })) {
    $stock = Get-Number $sku.stock
    $cost = Get-Number $sku.costoPromedio
    $code = if ($sku.code) { $sku.code } else { $sku.id }
    if ($stock -le [double]$Config.thresholds.stock_critical_qty -and $skuDemand.ContainsKey($code)) {
      $stockGaps += [pscustomobject]@{
        code = $code
        demand = $skuDemand[$code]
        packs = @($skuDemandPacks[$code])
      }
    }
  }
  if ($stockGaps.Count -gt 0) {
    $packRefs = New-Object System.Collections.Generic.HashSet[string]
    $projectRefs = New-Object System.Collections.Generic.HashSet[string]
    foreach ($gap in $stockGaps) {
      foreach ($pack in @($gap.packs)) {
        if ($pack.id) { [void]$packRefs.Add([string]$pack.id) }
        if ($pack.projectName) { [void]$projectRefs.Add([string]$pack.projectName) }
      }
    }
    $projectList = @($projectRefs.GetEnumerator() | Select-Object -First 3)
    $projectText = if ($projectList.Count -gt 0) { " Proyectos impactados: $($projectList -join '; ')." } else { "" }
    Add-Issue $issues (New-Issue -Severity "amarillo" -Area "Bodega" -Title "Productos requeridos sin stock suficiente" -Detail "Hay $($stockGaps.Count) producto(s) requeridos por packs proximos con stock bajo o cero. Impacta $($packRefs.Count) pack(s).$projectText" -Owner "Mauricio / Carlos" -Action "Definir si se recepciona compra, se reemplaza material o se reprograma entrega." -Ref "stock_packs")
  }
  $sinValor = @($Data.inv_catalogo | Where-Object {
    $_.activo -ne $false -and
    -not $_.aliasDe -and
    (Get-Number $_.stock) -gt 0 -and
    (Get-Number $_.costoPromedio) -le 0
  })
  if ($sinValor.Count -gt 0) {
    Add-Issue $issues (New-Issue -Severity "rojo" -Area "Bodega/Finanzas" -Title "Inventario con costos incompletos" -Detail "$($sinValor.Count) producto(s) con stock tienen costo registrado en cero. Esto distorsiona costos por proyecto y reportes de directorio." -Owner "Mauricio / Valentina" -Action "Valorizar inventario desde Kardex." -Ref "inv_catalogo")
  }
  foreach ($oc in @($Data.mat_ordenes)) {
    if (-not $oc.precioUnit -or [double]$oc.precioUnit -le 0) {
      Add-Issue $issues (New-Issue -Severity "amarillo" -Area "Compras" -Title "Orden de compra sin precio" -Detail "$($oc.id) / $($oc.itemDesc) no tiene precio unitario valido." -Owner "Carlos / Valentina" -Action "Completar precio para medir costo comprometido." -Ref $oc.id)
    }
  }
  @($issues)
}

function Get-FirstText {
  param([object]$Object, [string[]]$Names)
  if ($null -eq $Object) { return "" }
  foreach ($name in $Names) {
    $prop = $Object.PSObject.Properties[$name]
    if ($prop -and -not [string]::IsNullOrWhiteSpace([string]$prop.Value)) {
      return [string]$prop.Value
    }
  }
  ""
}

function Normalize-MayuText {
  param([object]$Value)
  $text = ([string]$Value).ToLowerInvariant()
  $text = $text -replace "[^a-z0-9áéíóúñü ]", " "
  ($text -replace "\s+", " ").Trim()
}

function Get-TextSimilarity {
  param([object]$A, [object]$B)
  $aText = Normalize-MayuText $A
  $bText = Normalize-MayuText $B
  if ([string]::IsNullOrWhiteSpace($aText) -or [string]::IsNullOrWhiteSpace($bText)) { return 0.0 }
  if ($aText -eq $bText) { return 1.0 }
  $aWords = @($aText -split " " | Where-Object { $_.Length -gt 2 } | Select-Object -Unique)
  $bWords = @($bText -split " " | Where-Object { $_.Length -gt 2 } | Select-Object -Unique)
  if ($aWords.Count -eq 0 -or $bWords.Count -eq 0) { return 0.0 }
  $set = New-Object System.Collections.Generic.HashSet[string]
  foreach ($w in $aWords) { [void]$set.Add($w) }
  $intersection = 0
  foreach ($w in $bWords) { if ($set.Contains($w)) { $intersection++ } }
  $union = @($aWords + $bWords | Select-Object -Unique).Count
  if ($union -eq 0) { return 0.0 }
  [double]$intersection / [double]$union
}

function Get-NumberTokens {
  param([object]$Value)
  $text = ([string]$Value).ToLowerInvariant()
  @([regex]::Matches($text, "\d+(?:[.,]\d+)?") | ForEach-Object { $_.Value -replace ",", "." } | Select-Object -Unique)
}

function Test-MeaningfullyDifferentSpecs {
  param([object]$A, [object]$B)
  $aNums = @(Get-NumberTokens $A)
  $bNums = @(Get-NumberTokens $B)
  if ($aNums.Count -eq 0 -or $bNums.Count -eq 0) { return $false }
  $a = ($aNums | Sort-Object) -join "|"
  $b = ($bNums | Sort-Object) -join "|"
  $a -ne $b
}

function New-BodegaIssue {
  param(
    [string]$CheckId,
    [ValidateSet("CRITICO", "ALTO", "MEDIO", "BAJO")]
    [string]$Level,
    [string]$Block,
    [string]$Title,
    [string]$Detail,
    [string]$Owner,
    [string]$Action,
    [string]$Ref
  )
  $severity = if ($Level -eq "CRITICO") { "rojo" } elseif ($Level -in @("ALTO", "MEDIO")) { "amarillo" } else { "info" }
  [pscustomobject]@{
    severity = $severity
    level = $Level
    checkId = $CheckId
    block = $Block
    area = "Bodega + Materiales"
    title = "$CheckId - $Title"
    detail = $Detail
    owner = $Owner
    action = $Action
    ref = $Ref
  }
}

function Add-BodegaIssue {
  param([System.Collections.ArrayList]$List, [object]$Issue)
  [void]$List.Add($Issue)
}

function Get-BodegaBomItems {
  param([object[]]$MatProjects)
  $items = @()
  foreach ($mp in @($MatProjects)) {
    foreach ($it in @($mp.items)) {
      $code = Get-FirstText $it @("bomItemCode", "itemCode", "code", "matchCode", "id")
      $desc = Get-FirstText $it @("descripcion", "description", "desc", "itemDesc", "nombre", "name")
      $skuCode = Get-FirstText $it @("skuCode", "catalogCode", "codigoBodega")
      if ([string]::IsNullOrWhiteSpace($code) -and [string]::IsNullOrWhiteSpace($desc)) { continue }
      $items += [pscustomobject]@{
        matProjectId = $mp.id
        projectName = $mp.name
        code = $code
        skuCode = $skuCode
        desc = $desc
        raw = $it
      }
    }
  }
  $items
}

function Get-BodegaCatalogItems {
  param([object[]]$Catalog)
  $items = @()
  foreach ($sku in @($Catalog)) {
    $code = Get-FirstText $sku @("code", "skuCode", "id")
    $legacy = Get-FirstText $sku @("codigoLegacy", "legacyCode", "bomItemCode")
    $desc = Get-FirstText $sku @("descripcion", "description", "desc", "nombre", "name")
    $items += [pscustomobject]@{ code = $code; legacy = $legacy; desc = $desc; raw = $sku }
  }
  $items
}

function New-BodegaBomSkuIndex {
  param([object[]]$BomItems)
  $byProjectItem = @{}
  $byItem = @{}
  foreach ($b in @($BomItems)) {
    if ([string]::IsNullOrWhiteSpace([string]$b.code)) { continue }
    if ($b.matProjectId) { $byProjectItem["$($b.matProjectId)|$($b.code)"] = $b }
    if (-not $byItem.ContainsKey([string]$b.code)) { $byItem[[string]$b.code] = $b }
  }
  [pscustomobject]@{ byProjectItem = $byProjectItem; byItem = $byItem }
}

function Resolve-BodegaBomSku {
  param([object]$Index, [string]$ProjectId, [string]$ItemCode)
  if ([string]::IsNullOrWhiteSpace($ItemCode)) { return "" }
  $bomItem = $null
  if (-not [string]::IsNullOrWhiteSpace($ProjectId)) {
    $key = "$ProjectId|$ItemCode"
    if ($Index.byProjectItem.ContainsKey($key)) { $bomItem = $Index.byProjectItem[$key] }
  }
  if ($null -eq $bomItem -and $Index.byItem.ContainsKey($ItemCode)) { $bomItem = $Index.byItem[$ItemCode] }
  if ($null -eq $bomItem) { return "" }
  Get-FirstText $bomItem @("skuCode", "catalogCode", "codigoBodega")
}

function Build-BodegaMaterialesReport {
  param([object]$Config, [object]$Data, [datetime]$Now)
  $issues = [System.Collections.ArrayList]::new()
  $bomItems = @(Get-BodegaBomItems -MatProjects @($Data.mat_projects))
  $bomSkuIndex = New-BodegaBomSkuIndex -BomItems $bomItems
  $catalogItems = @(Get-BodegaCatalogItems -Catalog @($Data.inv_catalogo))
  $bomCodes = New-Object System.Collections.Generic.HashSet[string]
  foreach ($b in $bomItems) { if ($b.code) { [void]$bomCodes.Add([string]$b.code) } }
  $catalogCodes = New-Object System.Collections.Generic.HashSet[string]
  $legacyCodes = New-Object System.Collections.Generic.HashSet[string]
  foreach ($c in $catalogItems) {
    if ($c.code) { [void]$catalogCodes.Add([string]$c.code) }
    if ($c.legacy) { [void]$legacyCodes.Add([string]$c.legacy) }
  }
  $ocBomCodes = New-Object System.Collections.Generic.HashSet[string]
  foreach ($oc in @($Data.mat_ordenes)) {
    $ocItemCode = Get-FirstText $oc @("bomItemCode", "itemCode", "matchCode", "code")
    if ($ocItemCode) { [void]$ocBomCodes.Add([string]$ocItemCode) }
  }
  $similarThreshold = Get-Number $Config.thresholds.bodega_materiales_similarity_threshold 0.7
  $dupThreshold = Get-Number $Config.thresholds.bodega_materiales_duplicate_threshold 0.85
  $deliveryDays = [int](Get-Number $Config.thresholds.bodega_materiales_delivery_days 7)
  $staleOcDays = [int](Get-Number $Config.thresholds.bodega_materiales_oc_stale_days 30)
  $priceDiffPct = Get-Number $Config.thresholds.bodega_materiales_price_diff_pct 0.05
  $today = $Now.Date

  foreach ($oc in @($Data.mat_ordenes)) {
    $ocId = Get-FirstText $oc @("id", "folio", "ocId")
    $projectId = Get-FirstText $oc @("matProjectId", "projectId", "proyectoId")
    $itemCode = Get-FirstText $oc @("bomItemCode", "itemCode", "matchCode", "code")
    $proveedor = Get-FirstText $oc @("proveedor", "supplierName", "vendor", "razonSocial")
    $ocSkuCode = Get-FirstText $oc @("skuCode", "catalogCode", "codigoBodega")
    $bomSkuCode = Resolve-BodegaBomSku -Index $bomSkuIndex -ProjectId $projectId -ItemCode $itemCode
    $skuCode = if ($ocSkuCode) { $ocSkuCode } else { $bomSkuCode }
    if ([string]::IsNullOrWhiteSpace($projectId) -or [string]::IsNullOrWhiteSpace($itemCode)) {
      Add-BodegaIssue $issues (New-BodegaIssue "B-0001" "CRITICO" "Trazabilidad" "OC sin identidad BOM completa" "OC $ocId no tiene matProjectId y bomItemCode trazables." "Carlos" "Regularizar el vinculo OC -> proyecto -> item BOM antes de nuevas recepciones." $ocId)
    } elseif (-not $bomCodes.Contains($itemCode)) {
      Add-BodegaIssue $issues (New-BodegaIssue "B-0002" "CRITICO" "Trazabilidad" "OC con item fuera del BOM activo" "OC $ocId usa item $itemCode, pero ese codigo no aparece en los BOM activos leidos." "Carlos" "Corregir itemCode o registrar excepcion autorizada por Carlos." $ocId)
    }
    if ([string]::IsNullOrWhiteSpace($skuCode)) {
      $itemDesc = Get-FirstText $oc @("itemDesc", "descripcion", "description", "nombre", "name")
      Add-BodegaIssue $issues (New-BodegaIssue "B-0003" "CRITICO" "Trazabilidad" "OC sin SKU candidato de bodega" "OC $ocId no tiene SKU de bodega asociado antes de recepcionar. Proyecto=$projectId ItemBOM=$itemCode Desc='$itemDesc' Proveedor='$proveedor'." "Carlos / Mauricio" "Linkear a SKU candidato antes de recepcion o dejar excepcion autorizada." $ocId)
    } elseif (-not $catalogCodes.Contains($skuCode)) {
      $itemDesc = Get-FirstText $oc @("itemDesc", "descripcion", "description", "nombre", "name")
      $source = if ($ocSkuCode) { "OC" } else { "BOM" }
      Add-BodegaIssue $issues (New-BodegaIssue "B-0003" "CRITICO" "Trazabilidad" "OC con SKU que no existe en bodega" "OC $ocId resuelve SKU $skuCode desde $source, pero no existe en inv_catalogo. Proyecto=$projectId ItemBOM=$itemCode Desc='$itemDesc'." "Carlos / Mauricio" "Corregir link a SKU existente o crear el codigo real antes de recepcionar." $ocId)
    }
    if (-not (Get-FirstText $oc @("cotizId", "cotizacionId", "quoteId"))) {
      Add-BodegaIssue $issues (New-BodegaIssue "B-D04" "ALTO" "Ordenes de compra" "OC sin cotizacion asociada" "OC $ocId no tiene cotizacion vinculada." "Carlos / Valentina" "Vincular cotizacion formal o justificar compra directa." $ocId)
    }
    $proveedor = Get-FirstText $oc @("proveedor", "supplierName", "vendor", "razonSocial")
    if ($proveedor -match "(?i)cotizaci[oó]n|mayu tx|proveedor|pendiente") {
      Add-BodegaIssue $issues (New-BodegaIssue "B-D06" "MEDIO" "Ordenes de compra" "OC con proveedor generico" "OC $ocId tiene proveedor '$proveedor'." "Carlos / Valentina" "Identificar proveedor real para trazabilidad y reclamos." $ocId)
    }
    $fechaOc = Get-FirstText $oc @("fecha", "createdAt", "fechaOC")
    $status = [string](Get-FirstText $oc @("status", "estado"))
    if ($fechaOc) {
      try {
        $age = ($today - ([datetime]::Parse($fechaOc)).Date).Days
        if ($age -gt $staleOcDays -and $status -in @("adjuntada", "emitida", "abierta")) {
          Add-BodegaIssue $issues (New-BodegaIssue "B-D05" "MEDIO" "Ordenes de compra" "OC antigua sin recepcion" "OC $ocId lleva $age dias en estado $status." "Carlos" "Confirmar entrega proveedor o cerrar/reprogramar OC." $ocId)
        }
      } catch {}
    }
  }

  foreach ($rec in @($Data.mat_recepciones)) {
    $recId = Get-FirstText $rec @("id", "recepcionId")
    $ocId = Get-FirstText $rec @("ocId", "ordenId", "ordenCompraId")
    $skuCode = Get-FirstText $rec @("skuCode", "catalogCode", "code", "codigoBodega")
    $itemCode = Get-FirstText $rec @("bomItemCode", "itemCode", "matchCode")
    if ([string]::IsNullOrWhiteSpace($skuCode)) {
      Add-BodegaIssue $issues (New-BodegaIssue "B-0003" "CRITICO" "Trazabilidad" "Recepcion sin SKU real de bodega" "Recepcion $recId no tiene SKU real; no entra trazablemente al inventario." "Mauricio" "Linkear recepcion a SKU existente o crear SKU con justificacion y codigoLegacy." $recId)
    }
    if ([string]::IsNullOrWhiteSpace($itemCode) -and [string]::IsNullOrWhiteSpace($ocId)) {
      Add-BodegaIssue $issues (New-BodegaIssue "B-0005" "CRITICO" "Trazabilidad" "Recepcion sin OC ni item BOM" "Recepcion $recId no conserva OC ni item BOM." "Mauricio / Carlos" "Completar ocId y bomItemCode para que el costo llegue al proyecto." $recId)
    }
    $precioFactura = Get-Number (Get-FirstText $rec @("precioUnitFactura", "costoUnit", "precioFactura", "precioUnit"))
    $precioOc = Get-Number (Get-FirstText $rec @("precioUnitOC", "precioUnitOc", "ocPrecioUnit"))
    if ($precioFactura -gt 0 -and $precioOc -gt 0) {
      $diff = [Math]::Abs($precioFactura - $precioOc) / [Math]::Max($precioOc, 1)
      if ($diff -gt $priceDiffPct -and -not $rec.alertaConciliacion) {
        Add-BodegaIssue $issues (New-BodegaIssue "B-E08" "ALTO" "Recepciones" "Precio factura distinto a OC sin conciliacion" "Recepcion $recId tiene precio factura $precioFactura vs OC $precioOc." "Felix / Valentina / Carlos / Mauricio" "Abrir conciliacion: reclamo, nota de credito, ajuste o aceptar diferencia." $recId)
      }
    }
  }

  $receivedByOc = @{}
  foreach ($rec in @($Data.mat_recepciones)) {
    $ocId = Get-FirstText $rec @("ocId", "ordenId", "ordenCompraId")
    if (-not $ocId) { continue }
    if (-not $receivedByOc.ContainsKey($ocId)) { $receivedByOc[$ocId] = 0.0 }
    $receivedByOc[$ocId] += Get-Number (Get-FirstText $rec @("qtyRecibida", "qty", "cantidad"))
  }
  foreach ($oc in @($Data.mat_ordenes)) {
    $ocId = Get-FirstText $oc @("id", "folio", "ocId")
    $qty = Get-Number (Get-FirstText $oc @("qty", "cantidad", "cantidadPedida"))
    $status = [string](Get-FirstText $oc @("status", "estado"))
    $received = if ($receivedByOc.ContainsKey($ocId)) { $receivedByOc[$ocId] } else { 0.0 }
    if ($qty -gt 0 -and $status -in @("recibida_completa", "completa", "total", "recibida_total") -and $received -lt ($qty * 0.99)) {
      Add-BodegaIssue $issues (New-BodegaIssue "B-D01" "CRITICO" "Ordenes de compra" "OC completa con recepcion parcial" "OC $ocId esta cerrada como completa, pero recepciones suman $received de $qty." "Carlos / Mauricio" "Cambiar a recepcion parcial y reabrir saldo pendiente." $ocId)
    }
  }

  foreach ($mov in @($Data.inv_movimientos)) {
    if ($mov.anulado -eq $true -or $mov.ignorarAuditoria -eq $true) { continue }
    if ($mov.stockGeneral -eq $true -or $mov.auditoriaBodegaMaterialesExcepcion -eq $true) { continue }
    $movId = Get-FirstText $mov @("id", "movementId")
    $tipo = Get-FirstText $mov @("tipo", "type")
    $skuCode = Get-FirstText $mov @("skuCode", "code", "itemCode")
    $projectId = Get-FirstText $mov @("matProjectId", "projectId", "proyectoRef", "proyectoId")
    $ocId = Get-FirstText $mov @("ocId", "ordenId")
    $recId = Get-FirstText $mov @("recepcionId", "receiptId")
    $itemCode = Get-FirstText $mov @("bomItemCode", "itemCode", "matchCode")
    if ($mov.refDoc) {
      if (-not $ocId) { $ocId = Get-FirstText $mov.refDoc @("ocId", "ordenId") }
      if (-not $ocId -and (Get-FirstText $mov.refDoc @("tipo")) -eq "mat_orden") { $ocId = Get-FirstText $mov.refDoc @("id") }
      if (-not $recId) { $recId = Get-FirstText $mov.refDoc @("recepcionId", "receiptId") }
      if (-not $itemCode) { $itemCode = Get-FirstText $mov.refDoc @("bomItemCode", "itemCode") }
    }
    if ($tipo -in @("recepcion_oc", "ingreso_directo") -and [string]::IsNullOrWhiteSpace($skuCode)) {
      Add-BodegaIssue $issues (New-BodegaIssue "B-0005" "CRITICO" "Trazabilidad" "Movimiento sin identidad completa" "Movimiento $movId tipo $tipo no conserva proyecto/SKU suficientes." "Mauricio" "Completar trazabilidad en movimiento o corregir recepcion origen." $movId)
    } elseif ($tipo -eq "recepcion_oc" -and [string]::IsNullOrWhiteSpace($projectId)) {
      Add-BodegaIssue $issues (New-BodegaIssue "B-0005" "CRITICO" "Trazabilidad" "Movimiento sin proyecto" "Movimiento $movId tipo recepcion_oc no conserva proyecto." "Mauricio" "Completar proyecto desde OC/recepcion origen." $movId)
    } elseif ($tipo -eq "ingreso_directo" -and [string]::IsNullOrWhiteSpace($projectId)) {
      Add-BodegaIssue $issues (New-BodegaIssue "B-0005" "ALTO" "Trazabilidad" "Ingreso directo sin proyecto" "Movimiento $movId tipo ingreso_directo tiene SKU, pero no proyecto." "Mauricio / Carlos" "Confirmar si es stock general; si corresponde a proyecto, vincularlo a OC/BOM." $movId)
    }
    if ($tipo -eq "recepcion_oc" -and ([string]::IsNullOrWhiteSpace($ocId) -or [string]::IsNullOrWhiteSpace($itemCode))) {
      Add-BodegaIssue $issues (New-BodegaIssue "B-0005" "CRITICO" "Trazabilidad" "Kardex sin OC/recepcion/BOM" "Movimiento $movId no hereda ocId, recepcionId y bomItemCode." "Mauricio / Felix" "Corregir app para heredar campos y regularizar movimiento." $movId)
    } elseif ($tipo -eq "recepcion_oc" -and [string]::IsNullOrWhiteSpace($recId)) {
      Add-BodegaIssue $issues (New-BodegaIssue "B-0005" "ALTO" "Trazabilidad" "Kardex historico sin recepcionId" "Movimiento $movId conserva OC y BOM, pero no tiene recepcionId historico." "Mauricio / Felix" "Mantener como pendiente historico o asociar recepcion si existe evidencia." $movId)
    }
    if ($tipo -eq "ingreso_directo") {
      $desc = Get-FirstText $mov @("descSku", "description", "descripcion", "itemDesc")
      $similarBom = @($bomItems | Where-Object { (Get-TextSimilarity $desc $_.desc) -ge $similarThreshold } | Select-Object -First 1)
      if ($similarBom.Count -gt 0) {
        Add-BodegaIssue $issues (New-BodegaIssue "B-E02" "CRITICO" "Recepciones" "Ingreso directo parece item BOM activo" "Movimiento $movId ingreso_directo '$desc' se parece a BOM $($similarBom[0].code) / $($similarBom[0].projectName)." "Mauricio / Carlos" "Revisar si debio recepcionarse contra OC/BOM en vez de crear identidad paralela." $movId)
      }
    }
  }

  foreach ($c in $catalogItems) {
    if (-not $c.legacy -and $c.desc) {
      $similarBom = @($bomItems | Where-Object { (Get-TextSimilarity $c.desc $_.desc) -ge $similarThreshold } | Select-Object -First 1)
      if ($similarBom.Count -gt 0) {
        Add-BodegaIssue $issues (New-BodegaIssue "B-0004" "CRITICO" "Trazabilidad" "SKU nuevo similar a BOM sin link" "$($c.code) '$($c.desc)' se parece a BOM $($similarBom[0].code) / $($similarBom[0].projectName), pero no tiene codigoLegacy." "Mauricio / Carlos" "Linkear a BOM existente o justificar SKU nuevo." $c.code)
      }
    }
    $stock = Get-Number $c.raw.stock
    if ($stock -gt 0 -and -not $c.legacy -and -not $c.raw.refDoc -and -not $c.raw.origen) {
      Add-BodegaIssue $issues (New-BodegaIssue "B-0006" "ALTO" "Trazabilidad" "SKU con stock sin origen trazable" "$($c.code) tiene stock $stock sin codigoLegacy/origen claro." "Mauricio / Valentina" "Regularizar origen del stock para valorizar y atribuir proyecto." $c.code)
    }
    $std = Get-Number $c.raw.costoEstandar
    $avg = Get-Number $c.raw.costoPromedio
    if ($std -gt 0 -and $avg -gt 0) {
      $diff = [Math]::Abs($std - $avg) / [Math]::Max($avg, 1)
      if ($diff -gt 0.5) {
        Add-BodegaIssue $issues (New-BodegaIssue "B-A03" "ALTO" "Catalogo" "Costo estandar muy distinto al promedio" "$($c.code): costoEstandar $std vs costoPromedio $avg." "Mauricio / Valentina" "Revisar valorizacion desde factura/guia y kardex." $c.code)
      }
    }
    if ($avg -le 1 -and $stock -gt 0) {
      Add-BodegaIssue $issues (New-BodegaIssue "B-A05" "MEDIO" "Catalogo" "SKU con costo placeholder y stock" "$($c.code) tiene stock $stock y costoPromedio $avg." "Mauricio / Valentina" "Valorizar desde documento fisico recepcionado." $c.code)
    }
  }

  for ($i = 0; $i -lt $catalogItems.Count; $i++) {
    for ($j = $i + 1; $j -lt $catalogItems.Count; $j++) {
      if ([string]::IsNullOrWhiteSpace($catalogItems[$i].desc) -or [string]::IsNullOrWhiteSpace($catalogItems[$j].desc)) { continue }
      if (Test-MeaningfullyDifferentSpecs $catalogItems[$i].desc $catalogItems[$j].desc) { continue }
      if ((Get-TextSimilarity $catalogItems[$i].desc $catalogItems[$j].desc) -ge $dupThreshold) {
        Add-BodegaIssue $issues (New-BodegaIssue "B-A04" "ALTO" "Catalogo" "SKUs posiblemente duplicados" "$($catalogItems[$i].code) y $($catalogItems[$j].code) tienen descripciones casi identicas." "Mauricio / Carlos" "Unificar identidad o documentar diferencia fisica." "$($catalogItems[$i].code)|$($catalogItems[$j].code)")
        if (@($issues | Where-Object { $_.checkId -eq "B-A04" }).Count -ge 10) { break }
      }
    }
  }

  foreach ($b in $bomItems) {
    $hasStrict = ($b.code -and ($legacyCodes.Contains($b.code) -or $catalogCodes.Contains($b.code))) -or ($b.skuCode -and $catalogCodes.Contains($b.skuCode))
    if (-not $hasStrict) {
      $usedInOc = $b.code -and $ocBomCodes.Contains([string]$b.code)
      if ($usedInOc) {
        Add-BodegaIssue $issues (New-BodegaIssue "B-B01" "CRITICO" "BOM-Catalogo" "Item BOM sin match estricto en catalogo" "$($b.projectName): $($b.code) '$($b.desc)' no tiene codigoLegacy/code/SKU estricto en catalogo." "Carlos / Mauricio" "Crear o linkear SKU de bodega antes de comprar/recepcionar." "$($b.matProjectId)|$($b.code)")
      }
    }
  }

  foreach ($cot in @($Data.mat_cotizaciones)) {
    $cotId = Get-FirstText $cot @("id", "cotizId")
    $file = Get-FirstText $cot @("fileName", "filename", "pdfUrl", "archivo", "documentUrl")
    if ($file -and $file -match "(?i)\.xlsx?$") {
      Add-BodegaIssue $issues (New-BodegaIssue "B-C01" "CRITICO" "Cotizaciones" "Cotizacion solo en Excel interno" "Cotizacion $cotId parece Excel, no PDF formal del proveedor." "Carlos / Valentina" "Adjuntar PDF formal con proveedor, RUT, fecha y validez." $cotId)
    }
    foreach ($it in @($cot.items)) {
      $matchCode = Get-FirstText $it @("matchCode", "bomItemCode", "itemCode")
      $matchScore = Get-Number (Get-FirstText $it @("matchScore", "score"))
      if ([string]::IsNullOrWhiteSpace($matchCode)) {
        Add-BodegaIssue $issues (New-BodegaIssue "B-C05" "MEDIO" "Cotizaciones" "Item cotizado sin match BOM" "Cotizacion $cotId tiene item sin matchCode." "Carlos" "Asignar item BOM o descartar linea de cotizacion." $cotId)
      } elseif ($matchScore -gt 0 -and $matchScore -lt 90) {
        Add-BodegaIssue $issues (New-BodegaIssue "B-C04" "MEDIO" "Cotizaciones" "Match de cotizacion dudoso" "Cotizacion $cotId item $matchCode tiene matchScore $matchScore." "Carlos" "Validar manualmente match de item cotizado contra BOM." "$cotId|$matchCode")
      }
    }
  }

  foreach ($mp in @($Data.mat_projects)) {
    $skuCancelados = @()
    foreach ($x in @($mp.skuCancelados)) { if ($x) { $skuCancelados += [string]$x } }
    if ($mp.recetaFabInternaPorTipologia -and $mp.recetaFabInternaPorTipologia.skuCancelados) {
      foreach ($x in @($mp.recetaFabInternaPorTipologia.skuCancelados)) { if ($x) { $skuCancelados += [string]$x } }
    }
    foreach ($sku in @($skuCancelados | Select-Object -Unique)) {
      $hasOc = @($Data.mat_ordenes | Where-Object { (Get-FirstText $_ @("bomItemCode", "itemCode", "matchCode", "code")) -eq $sku }).Count -gt 0
      if ($hasOc) {
        Add-BodegaIssue $issues (New-BodegaIssue "B-G03" "ALTO" "Fabricacion interna" "Item fab interna con OC externa" "$($mp.name): $sku esta en skuCancelados pero tiene OC." "Martin / Carlos" "Decidir si se fabrica o se compra; no ambos." "$($mp.id)|$sku")
      }
    }
  }

  foreach ($ent in @($Data.mat_entregas)) {
    $entId = Get-FirstText $ent @("id", "entregaId")
    $skuCode = Get-FirstText $ent @("skuCode", "itemCode", "code")
    $projectId = Get-FirstText $ent @("matProjectId", "projectId", "proyectoId")
    $podNum = Get-FirstText $ent @("podNum", "pod", "unitCode")
    $tipologia = Get-FirstText $ent @("tipologia", "tipo")
    if ([string]::IsNullOrWhiteSpace($skuCode) -or [string]::IsNullOrWhiteSpace($projectId)) {
      Add-BodegaIssue $issues (New-BodegaIssue "B-0007" "CRITICO" "Entregas a fabrica" "Entrega sin trazabilidad a SKU/proyecto" "Entrega $entId no conserva SKU o proyecto." "Carlos / Mauricio" "Regularizar entrega para atribuir consumo a proyecto/POD." $entId)
    }
    if ([string]::IsNullOrWhiteSpace($podNum) -or [string]::IsNullOrWhiteSpace($tipologia)) {
      Add-BodegaIssue $issues (New-BodegaIssue "B-F03" "MEDIO" "Entregas a fabrica" "Entrega sin POD o tipologia" "Entrega $entId no tiene POD/tipologia completos." "Mauricio" "Completar POD y tipologia para consumo real por unidad." $entId)
    }
    if ($ent.esMerma -eq $true -and -not (Get-FirstText $ent @("motivoMerma", "nota", "observacion"))) {
      Add-BodegaIssue $issues (New-BodegaIssue "B-F02" "ALTO" "Entregas a fabrica" "Merma sin motivo" "Entrega $entId marcada como merma no tiene motivo documentado." "Mauricio / Carlos" "Documentar causa de merma y responsable." $entId)
    }
  }

  $deliveryKeys = New-Object System.Collections.Generic.HashSet[string]
  foreach ($ent in @($Data.mat_entregas)) {
    $sku = Get-FirstText $ent @("skuCode", "itemCode", "code")
    $proj = Get-FirstText $ent @("matProjectId", "projectId", "proyectoId")
    if ($sku -and $proj) { [void]$deliveryKeys.Add("$proj|$sku") }
  }
  foreach ($rec in @($Data.mat_recepciones)) {
    $recId = Get-FirstText $rec @("id", "recepcionId")
    $sku = Get-FirstText $rec @("skuCode", "catalogCode", "code", "codigoBodega")
    $proj = Get-FirstText $rec @("matProjectId", "projectId", "proyectoId")
    $dateText = Get-FirstText $rec @("fecha", "createdAt", "fechaRecepcion")
    if (-not $sku -or -not $proj -or -not $dateText) { continue }
    try {
      $age = ($today - ([datetime]::Parse($dateText)).Date).Days
      if ($age -gt $deliveryDays -and -not $deliveryKeys.Contains("$proj|$sku")) {
        Add-BodegaIssue $issues (New-BodegaIssue "B-F01" "CRITICO" "Entregas a fabrica" "Recepcion sin entrega formal" "Recepcion $recId / SKU $sku lleva $age dias sin entrega formal a fabrica." "Carlos / Mauricio" "Registrar entrega formal o justificar stock retenido." $recId)
      }
    } catch {}
  }

  $selected = @(Select-UniqueIssues -Issues @($issues))
  [pscustomobject]@{
    generatedAt = $Now.ToString("o")
    date = $Now.ToString("yyyy-MM-dd")
    summary = [pscustomobject]@{
      criticas = @($selected | Where-Object { $_.level -eq "CRITICO" }).Count
      altas = @($selected | Where-Object { $_.level -eq "ALTO" }).Count
      medias = @($selected | Where-Object { $_.level -eq "MEDIO" }).Count
      bajas = @($selected | Where-Object { $_.level -eq "BAJO" }).Count
      total = @($selected).Count
    }
    blocks = [pscustomobject]@{
      trazabilidad = @($selected | Where-Object { $_.block -eq "Trazabilidad" })
      catalogo = @($selected | Where-Object { $_.block -eq "Catalogo" })
      bomCatalogo = @($selected | Where-Object { $_.block -eq "BOM-Catalogo" })
      cotizaciones = @($selected | Where-Object { $_.block -eq "Cotizaciones" })
      ordenesCompra = @($selected | Where-Object { $_.block -eq "Ordenes de compra" })
      recepciones = @($selected | Where-Object { $_.block -eq "Recepciones" })
      entregasFabrica = @($selected | Where-Object { $_.block -eq "Entregas a fabrica" })
      fabricacionInterna = @($selected | Where-Object { $_.block -eq "Fabricacion interna" })
    }
    issues = $selected
  }
}

function Write-BodegaPprPvcCandidates {
  param([object]$Data)
  $catalogItems = @(Get-BodegaCatalogItems -Catalog @($Data.inv_catalogo))
  $bomItems = @(Get-BodegaBomItems -MatProjects @($Data.mat_projects))
  $bomSkuIndex = New-BodegaBomSkuIndex -BomItems $bomItems
  $targets = @()
  foreach ($oc in @($Data.mat_ordenes)) {
    $projectId = Get-FirstText $oc @("matProjectId", "projectId", "proyectoId")
    $itemCode = Get-FirstText $oc @("bomItemCode", "itemCode", "matchCode", "code")
    $skuCode = Get-FirstText $oc @("skuCode", "catalogCode", "codigoBodega")
    if (-not $skuCode) { $skuCode = Resolve-BodegaBomSku -Index $bomSkuIndex -ProjectId $projectId -ItemCode $itemCode }
    if ($skuCode) { continue }
    $desc = Get-FirstText $oc @("itemDesc", "descripcion", "description", "nombre", "name")
    if ($desc -notmatch "(?i)\b(PPR|PVC|ABRAZADERA|TUBER|CODO|COPLA|TEE|TERMINAL|LLAVE DE PASO)\b") { continue }
    $targets += [pscustomobject]@{
      oc = Get-FirstText $oc @("id", "folio", "ocId")
      item = $itemCode
      desc = $desc
      proveedor = Get-FirstText $oc @("proveedor", "supplierName", "vendor", "razonSocial")
      project = $projectId
    }
  }
  Write-Output "Bodega+Materiales PPR/PVC: targets=$($targets.Count)"
  foreach ($target in @($targets | Select-Object -First 80)) {
    $candidates = @(
      $catalogItems | ForEach-Object {
        $score = Get-TextSimilarity $target.desc $_.desc
        $bonus = 0.0
        if ($_.legacy -and $target.item -and $_.legacy -eq $target.item) { $bonus += 1.0 }
        if ($_.code -and $target.item -and $_.code -eq $target.item) { $bonus += 1.0 }
        if ($target.desc -match "(?i)PPR" -and $_.desc -match "(?i)PPR") { $bonus += 0.2 }
        if ($target.desc -match "(?i)PVC" -and $_.desc -match "(?i)PVC") { $bonus += 0.2 }
        [pscustomobject]@{
          score = $score + $bonus
          code = $_.code
          legacy = $_.legacy
          desc = $_.desc
          stock = Get-Number $_.raw.stock
          costo = Get-Number $_.raw.costoPromedio
        }
      } | Sort-Object score -Descending | Select-Object -First 5
    )
    $candidateText = (@($candidates) | ForEach-Object {
      "$($_.code)|legacy=$($_.legacy)|stock=$($_.stock)|score=$([Math]::Round($_.score,2))|$($_.desc)"
    }) -join " || "
    Write-Output "Bodega+Materiales PPR/PVC candidate: oc=$($target.oc) item=$($target.item) desc=$($target.desc) proveedor=$($target.proveedor) :: $candidateText"
  }
}

function Render-BodegaMaterialesHtml {
  param([object]$Report)
  return Render-BodegaMaterialesHtmlV2 -Report $Report
  $s = $Report.summary
  $rows = Render-IssueList -Items @($Report.issues | Select-Object -First 80)
  @"
<html>
<body style="font-family:Arial,sans-serif;color:#222;max-width:980px;font-size:14px;">
  <h2 style="margin-bottom:4px;">Agente Bodega + Materiales - $($Report.date)</h2>
  <p style="margin-top:0;color:#555;">Criticas: <strong style="color:#991b1b;">$($s.criticas)</strong> · Altas: <strong style="color:#9a3412;">$($s.altas)</strong> · Medias: <strong>$($s.medias)</strong> · Bajas: <strong>$($s.bajas)</strong></p>
  <p style="background:#f7fafc;border-left:4px solid #2563eb;padding:10px 12px;">
    Puedes responder este correo con preguntas como <strong>"como resuelvo B-D06"</strong>, <strong>"explicame B-A05"</strong> o copiando una alerta. El agente respondera solo con explicacion operativa y pasos manuales en las apps.
  </p>
  <h3>Alertas</h3>
  $rows
  <p style="font-size:12px;color:#777;">Destinatarios productivos esperados: Felix, Valentina, Carlos y Mauricio. Este HTML es evidencia del agente; el envio queda controlado por SendEmail.</p>
</body>
</html>
"@
}

function Invoke-BodegaMateriales {
  param([object]$Config, [string]$GraphToken, [string]$SiteId, [datetime]$Now, [bool]$DoSendEmail)
  Write-Output "Bodega+Materiales: leyendo Firestore."
  $data = Get-FirestoreData -Config $Config
  $report = Build-BodegaMaterialesReport -Config $Config -Data $data -Now $Now
  Write-BodegaPprPvcCandidates -Data $data
  Write-Output "Bodega+Materiales: resumen criticas=$($report.summary.criticas) altas=$($report.summary.altas) medias=$($report.summary.medias) bajas=$($report.summary.bajas) total=$($report.summary.total)."
  foreach ($group in @($report.issues | Group-Object checkId | Sort-Object Count -Descending | Select-Object -First 20)) {
    Write-Output "Bodega+Materiales breakdown: $($group.Name)=$($group.Count)"
  }
  foreach ($group in @($report.issues | Group-Object checkId | Sort-Object Count -Descending | Select-Object -First 12)) {
    foreach ($issue in @($group.Group | Select-Object -First 5)) {
      Write-Output "Bodega+Materiales ejemplo $($group.Name): [$($issue.level)] ref=$($issue.ref) detalle=$($issue.detail) accion=$($issue.action)"
    }
  }
  foreach ($issue in @($report.issues | Where-Object { $_.checkId -eq "B-0003" } | Select-Object -First 200)) {
    Write-Output "Bodega+Materiales B-0003 full: ref=$($issue.ref) detalle=$($issue.detail)"
  }
  foreach ($issue in @($report.issues | Select-Object -First 25)) {
    Write-Output "Bodega+Materiales alerta: [$($issue.level)] $($issue.checkId) $($issue.title) | $($issue.ref) | $($issue.action)"
  }
  $html = Render-BodegaMaterialesHtmlV2 -Report $report
  $dateKey = $report.date
  Ensure-GraphFolder -Token $GraphToken -SiteId $SiteId -FolderPath $Config.sharepoint.bodega_materiales_folder
  Write-TextFileToGraph -Token $GraphToken -SiteId $SiteId -FilePath "$($Config.sharepoint.bodega_materiales_folder)/$dateKey.json" -Text ($report | ConvertTo-Json -Depth 80) -ContentType "application/json; charset=utf-8"
  Write-TextFileToGraph -Token $GraphToken -SiteId $SiteId -FilePath "$($Config.sharepoint.bodega_materiales_folder)/$dateKey.html" -Text $html -ContentType "text/html; charset=utf-8"
  if ($DoSendEmail) {
    $to = Get-BodegaMaterialesMailAudience -Config $Config
    $subject = "[Bodega-Materiales] Alertas $dateKey - $($report.summary.criticas) CRITICAS, $($report.summary.altas) ALTAS"
    Send-GraphMail -Token $GraphToken -Sender $Config.mail.sender -To $to -Cc @() -Subject $subject -HtmlBody $html
    Write-Output "Bodega+Materiales: correo enviado."
  } else {
    Write-Output "Bodega+Materiales: SendEmail=false, no se envia correo."
  }
  Write-Output "Bodega+Materiales: outputs guardados en SharePoint."
}

function Get-BodegaHelpDefinition {
  param([string]$CheckId)
  switch ($CheckId) {
    "B-D06" {
      [pscustomobject]@{
        title = "B-D06 - Proveedor mal nombrado"
        problem = "La OC tiene como proveedor un texto tipo 'Cotizacion ...' en vez del proveedor real. Esto dificulta reclamos, trazabilidad y analisis de compras."
        steps = @(
          "Entrar a la app Materiales.",
          "Abrir Ordenes de compra y buscar la OC indicada.",
          "Editar el proveedor para dejar el nombre real: por ejemplo Hoffens, Gobantes, CHC, APSA, StrongTie, Steelfix o Sodimac.",
          "Guardar la OC. No cambiar cantidades, precios ni recepciones."
        )
      }
    }
    "B-A04" {
      [pscustomobject]@{
        title = "B-A04 - Posibles SKUs duplicados"
        problem = "Hay dos codigos de bodega con descripciones muy parecidas. El riesgo es comprar o descontar stock desde identidades distintas para el mismo producto."
        steps = @(
          "Entrar a la app Bodega.",
          "Buscar ambos codigos SKU indicados en la alerta.",
          "Comparar medida, material, marca y uso real en bodega.",
          "Si son el mismo producto, dejar un SKU activo y mover/regularizar stock al codigo correcto.",
          "Si son distintos, ajustar la descripcion para que la diferencia quede clara."
        )
      }
    }
    "B-A05" {
      [pscustomobject]@{
        title = "B-A05 - SKU con stock y costo cero"
        problem = "El producto tiene stock, pero costo promedio cero. Eso distorsiona el costo de proyecto y el inventario valorizado."
        steps = @(
          "Entrar a la app Bodega.",
          "Buscar el SKU indicado.",
          "Revisar la recepcion, factura o guia asociada.",
          "Actualizar costo promedio o costo unitario segun el documento real.",
          "Si no existe documento, dejar pendiente para Valentina/Mauricio antes de usarlo en costos."
        )
      }
    }
    "B-C01" {
      [pscustomobject]@{
        title = "B-C01 - Cotizacion Excel sin PDF formal"
        problem = "La compra se apoya solo en Excel interno. MAYU necesita cotizacion formal del proveedor para respaldo y aprobacion."
        steps = @(
          "Entrar a la app Materiales.",
          "Abrir Cotizaciones y buscar el codigo COT indicado.",
          "Adjuntar PDF formal del proveedor con razon social, fecha, precios y validez.",
          "No marcar excepcion: la regla exige cotizacion formal."
        )
      }
    }
    "B-C04" {
      [pscustomobject]@{
        title = "B-C04 - Match dudoso entre cotizacion y BOM"
        problem = "La linea cotizada fue asociada a un item BOM con baja confianza. Puede estar comprandose algo parecido, pero no exactamente lo pedido."
        steps = @(
          "Entrar a la app Materiales.",
          "Abrir la cotizacion indicada.",
          "Comparar descripcion cotizada contra descripcion del item BOM.",
          "Si corresponde, confirmar o corregir el item BOM asociado.",
          "Si no corresponde, reasignar la linea al BOM correcto o descartarla."
        )
      }
    }
    "B-C05" {
      [pscustomobject]@{
        title = "B-C05 - Item de cotizacion sin match BOM"
        problem = "Hay una linea cotizada que no esta asociada a ningun item del BOM. Si se compra asi, despues no se puede cargar bien al proyecto."
        steps = @(
          "Entrar a la app Materiales.",
          "Abrir la cotizacion indicada.",
          "Buscar la linea sin match.",
          "Asignarla al item BOM correcto.",
          "Si la linea no corresponde al proyecto, eliminarla o dejarla fuera de la cotizacion."
        )
      }
    }
    "B-0004" {
      [pscustomobject]@{
        title = "B-0004 - SKU nuevo parecido a BOM sin link"
        problem = "Existe un SKU de bodega parecido a un item BOM, pero no esta linkeado. Puede crear identidad paralela entre compra, recepcion y bodega."
        steps = @(
          "Entrar a la app Bodega.",
          "Buscar el SKU indicado.",
          "Compararlo con el item BOM sugerido.",
          "Si es el mismo producto, completar codigoLegacy o link al BOM.",
          "Si es distinto, mejorar la descripcion o justificar el SKU nuevo."
        )
      }
    }
    "B-0003" {
      [pscustomobject]@{
        title = "B-0003 - OC sin SKU real de bodega"
        problem = "La OC apunta a un item BOM, pero no tiene SKU real de bodega. Antes de recepcionar, debe existir el codigo de bodega correcto."
        steps = @(
          "Entrar a la app Materiales.",
          "Abrir la OC indicada.",
          "Revisar item BOM y descripcion del producto.",
          "Buscar si ya existe SKU en Bodega.",
          "Si existe, linkearlo a la OC/BOM. Si no existe, crear el SKU en Bodega y luego linkearlo."
        )
      }
    }
    "B-B01" {
      [pscustomobject]@{
        title = "B-B01 - BOM usado en OC sin SKU"
        problem = "Un item del BOM ya fue usado en una OC, pero todavia no tiene match estricto con SKU de bodega."
        steps = @(
          "Entrar a la app Materiales.",
          "Buscar el item BOM indicado.",
          "Buscar o crear el SKU real en Bodega.",
          "Linkear el item BOM al SKU antes de nuevas recepciones.",
          "Verificar que la OC quede apuntando al mismo SKU."
        )
      }
    }
    default {
      [pscustomobject]@{
        title = "Agente Bodega + Materiales"
        problem = "Puedo explicar solo alertas de Bodega + Materiales y pasos manuales en las apps."
        steps = @(
          "Responde indicando el codigo de alerta: B-D06, B-A04, B-A05, B-C01, B-C04, B-C05, B-0004, B-0003 o B-B01.",
          "Tambien puedes copiar una linea completa de la alerta.",
          "No ejecuto cambios desde el correo; solo explico el problema y como resolverlo manualmente."
        )
      }
    }
  }
}

function Resolve-BodegaHelpCheckId {
  param([string]$Text)
  $m = [regex]::Match($Text, "(?i)\bB-(D06|A04|A05|C01|C04|C05|0004|0003|B01)\b")
  if ($m.Success) { return $m.Value.ToUpperInvariant() }
  if ($Text -match "(?i)proveedor|cotizacion hoffens|gobantes|sodimac|strongtie|steelfix") { return "B-D06" }
  if ($Text -match "(?i)duplicad|dos sku|unificar") { return "B-A04" }
  if ($Text -match "(?i)costo cero|costo 0|valorizar|stock.*costo") { return "B-A05" }
  if ($Text -match "(?i)pdf formal|excel|cotizacion formal") { return "B-C01" }
  if ($Text -match "(?i)match dudoso|score|confianza") { return "B-C04" }
  if ($Text -match "(?i)sin match|matchcode") { return "B-C05" }
  if ($Text -match "(?i)sku nuevo|similar a bom") { return "B-0004" }
  if ($Text -match "(?i)oc sin sku|sku real") { return "B-0003" }
  if ($Text -match "(?i)bom.*sin sku|item bom") { return "B-B01" }
  ""
}

function Render-BodegaHelpReplyHtml {
  param([object]$Help, [string]$CheckId)
  $steps = (@($Help.steps) | ForEach-Object { "<li>$(HtmlEscape $_)</li>" }) -join ""
  $codeText = if ($CheckId) { "<p><strong>Alerta:</strong> $(HtmlEscape $CheckId)</p>" } else { "" }
  @"
<div style="font-family:Arial,sans-serif;color:#222;font-size:14px;line-height:1.45;">
  <p>Hola. Respondo solo sobre como resolver manualmente alertas de Bodega + Materiales.</p>
  $codeText
  <h3 style="margin-bottom:4px;">$(HtmlEscape $Help.title)</h3>
  <p><strong>Cual es el problema:</strong> $(HtmlEscape $Help.problem)</p>
  <p><strong>Como resolverlo en la app:</strong></p>
  <ol>$steps</ol>
  <p style="font-size:12px;color:#666;">Si necesitas otra alerta, responde con su codigo o copia la linea exacta del correo del agente.</p>
</div>
"@
}

function Invoke-BodegaMaterialesResponder {
  param([object]$Config, [string]$GraphToken, [string]$SiteId, [datetime]$Now)
  $mailbox = [string]$Config.mail.sender
  Disable-GraphMailboxAutoReplies -Token $GraphToken -Mailbox $mailbox
  $stateFile = "$($Config.sharepoint.bodega_materiales_folder)/responder_procesados.json"
  Ensure-GraphFolder -Token $GraphToken -SiteId $SiteId -FolderPath $Config.sharepoint.bodega_materiales_folder
  $stateText = Read-TextFileFromGraph -Token $GraphToken -SiteId $SiteId -FilePath $stateFile
  $processedIds = @()
  if (-not [string]::IsNullOrWhiteSpace($stateText)) {
    try { $processedIds = @($stateText | ConvertFrom-Json) } catch { $processedIds = @() }
  }
  $processedSet = New-Object System.Collections.Generic.HashSet[string]
  foreach ($id in @($processedIds)) { if ($id) { [void]$processedSet.Add([string]$id) } }

  $uri = "https://graph.microsoft.com/v1.0/users/$mailbox/mailFolders/inbox/messages?`$top=25&`$orderby=receivedDateTime desc&`$select=id,subject,from,bodyPreview,body,isRead,receivedDateTime"
  $messages = @((Invoke-GraphGet -Token $GraphToken -Uri $uri).value)
  $processed = 0
  foreach ($msg in @($messages | Where-Object { $_.isRead -eq $false })) {
    $messageId = [string]$msg.id
    if ($processedSet.Contains($messageId)) { continue }
    $from = [string]$msg.from.emailAddress.address
    if ($from -eq $mailbox) { continue }
    $subject = [string]$msg.subject
    $bodyText = "$subject`n$($msg.bodyPreview)`n$($msg.body.content)"
    if ($bodyText -notmatch "(?i)Bodega.?Materiales|B-[A-Z0-9]{2,4}|bodega|materiales") { continue }
    if ($bodyText -notmatch "(?i)como|c[oó]mo|resolver|resuelvo|explica|explicame|problema|que significa|qu[eé] significa|ayuda") { continue }

    $checkId = Resolve-BodegaHelpCheckId -Text $bodyText
    $help = Get-BodegaHelpDefinition -CheckId $checkId
    $reply = Render-HelpReplyCard -Intro "Hola. Respondo solo sobre como resolver manualmente alertas de Bodega + Materiales." -Help $help -Code $checkId
    $audience = Get-UniqueEmails -Emails @((Get-BodegaMaterialesMailAudience -Config $Config) + $from) -Exclude @($mailbox)
    $replySubject = if ($subject -match "^(?i)re:") { $subject } else { "RE: $subject" }
    Send-GraphMail -Token $GraphToken -Sender $mailbox -To $audience -Cc @() -Subject $replySubject -HtmlBody $reply
    [void]$processedSet.Add($messageId)
    $processed++
    Write-Output "Bodega+Materiales responder: respuesta enviada a audiencia oficial ($($audience -join ', ')) para $checkId."
  }
  $nextState = @($processedSet.GetEnumerator() | Select-Object -Last 500)
  Write-TextFileToGraph -Token $GraphToken -SiteId $SiteId -FilePath $stateFile -Text ($nextState | ConvertTo-Json -Depth 5) -ContentType "application/json; charset=utf-8"
  Write-Output "Bodega+Materiales responder: mensajes procesados=$processed."
}

function Get-FinanceIssues {
  param([object]$Config, [object]$Data, [datetime]$Now)
  $issues = [System.Collections.ArrayList]::new()
  $today = $Now.ToString("yyyy-MM-dd")
  function Get-MaxIsoDateFinanzas([object[]]$Rows, [string]$FieldName) {
    $dates = @($Rows | ForEach-Object {
      $prop = $_.PSObject.Properties[$FieldName]
      if ($prop -and -not [string]::IsNullOrWhiteSpace([string]$prop.Value)) { [string]$prop.Value }
    } | Sort-Object)
    if ($dates.Count -eq 0) { return "" }
    [string]$dates[-1]
  }
  function Get-LatestImportFinanzas([object[]]$Rows, [string]$Tipo) {
    @($Rows | Where-Object { [string]$_.tipo -eq $Tipo } | Sort-Object { -1 * (Get-Number $_.createdAt) } | Select-Object -First 1)
  }
  function Add-FuenteFinancieraIssue([string]$Nombre, [string]$TipoImportacion, [string]$FechaMaxima, [int]$DiasRojo, [int]$DiasAmarillo) {
    $latest = @(Get-LatestImportFinanzas -Rows @($Data.fin_importaciones) -Tipo $TipoImportacion | Select-Object -First 1)
    $archivo = if ($latest.Count -gt 0) { [string]$latest[0].archivo } else { "sin registro de importacion" }
    $importado = "sin fecha"
    $diasImportacion = $null
    if ($latest.Count -gt 0 -and $latest[0].createdAt) {
      $importDate = ([DateTimeOffset]::FromUnixTimeMilliseconds([int64](Get-Number $latest[0].createdAt))).Date
      $importado = $importDate.ToString("yyyy-MM-dd")
      $diasImportacion = [int](New-TimeSpan -Start $importDate -End $Now.Date).TotalDays
      if ($diasImportacion -lt $DiasAmarillo) { return }
    }
    if ([string]::IsNullOrWhiteSpace($FechaMaxima)) {
      Add-Issue $issues (New-FinanceIssue -Code "F-D01" -Severity "rojo" -Area "Datos Finanzas" -Title "Fuente sin datos: $Nombre" -Detail "No hay datos cargados para $Nombre. Ultima importacion registrada: $archivo, importada $importado." -Owner "Valentina" -Action "Revisar la automatizacion o importar archivo actualizado en Finanzas > Importar." -Ref "fin_importaciones")
      return
    }
    $dias = [int](New-TimeSpan -Start ([datetime]::Parse($FechaMaxima)) -End $Now.Date).TotalDays
    if ($dias -lt $DiasAmarillo -and $null -eq $diasImportacion) { return }
    $diasBase = if ($null -ne $diasImportacion) { $diasImportacion } else { $dias }
    $sev = if ($diasBase -ge $DiasRojo) { "rojo" } else { "amarillo" }
    Add-Issue $issues (New-FinanceIssue -Code "F-D02" -Severity $sev -Area "Datos Finanzas" -Title "Fuente desactualizada: $Nombre" -Detail "$Nombre tuvo su ultima importacion el $importado. Ultimo documento en datos: $FechaMaxima ($dias dia(s) sin documentos nuevos). Archivo/fuente: $archivo." -Owner "Valentina" -Action "Revisar que la automatizacion de la fuente este corriendo; si corrio sin novedades, no corresponde cierre contable sino solo seguimiento de fuente." -Ref "fin_importaciones")
  }
  function Test-NotaCreditoFinanzas([object]$Factura) {
    $tipo = [int](Get-Number $Factura.tipoDte)
    return ($tipo -eq 61 -or $tipo -eq 112)
  }
  function Test-ControlEspecialFinanzas([object]$Factura) {
    return (([string]$Factura.tratamientoFinanciero) -in @("CONCESION", "CONSIGNACION") -or ([string]$Factura.auditoriaFinanzasEstado) -eq "CONTROL_ESPECIAL")
  }
  $apSinProyecto = @()
  $apSinClasificar = @()
  $apVencidas = @()
  $arSinProyecto = @()
  $arVencidas = @()
  $movCargosSinConciliar = @()
  $movAbonosSinConciliar = @()
  foreach ($f in @($Data.fin_facturas_ap)) {
    if ((Test-NotaCreditoFinanzas $f) -or (Test-ControlEspecialFinanzas $f)) { continue }
    $estado = [string]$f.estado
    if (@("PAGADA","ANULADA","RECHAZADA") -contains $estado) { continue }
    if ((-not $f.proyectoId) -and (-not $f.asignaciones) -and ([string]$f.lineaNegocio) -ne "OPEX_CORP" -and ([string]$f.categoriaContable) -notin @("OPEX_FIJO","OPEX_VARIABLE")) {
      $apSinProyecto += $f
    }
    if ([string]$f.lineaNegocio -eq "SIN_CLASIFICAR") {
      $apSinClasificar += $f
    }
    if ($f.fechaVencimiento -and [string]$f.fechaVencimiento -lt $today) {
      $apVencidas += $f
    }
  }
  foreach ($f in @($Data.fin_facturas_ar)) {
    if (Test-NotaCreditoFinanzas $f) { continue }
    $estado = [string]$f.estado
    if (@("COBRADA","ANULADA") -contains $estado) { continue }
    if ((-not $f.crmProjectId) -and (-not $f.proyectoId) -and (-not $f.asignaciones)) {
      $arSinProyecto += $f
    }
    if ($f.fechaVencimiento -and [string]$f.fechaVencimiento -lt $today) {
      $arVencidas += $f
    }
  }
  foreach ($m in @($Data.fin_mov_bancarios)) {
    if ($m.conciliado -or $m.linkNomina -or $m.linkOperacional) { continue }
    $cargo = Get-Number $m.cargo
    $abono = Get-Number $m.abono
    if ($cargo -gt 0) {
      $movCargosSinConciliar += $m
    } elseif ($abono -gt 0) {
      $movAbonosSinConciliar += $m
    }
  }
  $diasRojo = [int](Get-Number $Config.thresholds.finance_source_stale_days_red 7)
  $diasAmarillo = [int](Get-Number $Config.thresholds.finance_source_stale_days_yellow 3)
  Add-FuenteFinancieraIssue -Nombre "Cartola BICE" -TipoImportacion "CARTOLA_BICE" -FechaMaxima (Get-MaxIsoDateFinanzas -Rows @($Data.fin_mov_bancarios) -FieldName "fecha") -DiasRojo $diasRojo -DiasAmarillo $diasAmarillo
  Add-FuenteFinancieraIssue -Nombre "RCV Compras" -TipoImportacion "RCV_COMPRAS" -FechaMaxima (Get-MaxIsoDateFinanzas -Rows @($Data.fin_facturas_ap) -FieldName "fechaEmision") -DiasRojo $diasRojo -DiasAmarillo $diasAmarillo
  Add-FuenteFinancieraIssue -Nombre "RCV Ventas" -TipoImportacion "RCV_VENTAS" -FechaMaxima (Get-MaxIsoDateFinanzas -Rows @($Data.fin_facturas_ar) -FieldName "fechaEmision") -DiasRojo $diasRojo -DiasAmarillo $diasAmarillo
  $maxOverdue = [int]$Config.thresholds.max_finance_overdue_items
  if ($maxOverdue -le 0) { $maxOverdue = 8 }
  if ($movCargosSinConciliar.Count -gt 0) {
    $monto = ($movCargosSinConciliar | ForEach-Object { Get-Number $_.cargo } | Measure-Object -Sum).Sum
    Add-Issue $issues (New-FinanceIssue -Code "F-B01" -Severity "rojo" -Area "Caja" -Title "Cargos bancarios sin conciliar" -Detail "$($movCargosSinConciliar.Count) salida(s) de banco sin trazabilidad por aprox. $(Format-Clp $monto). No se muestra como total banco para evitar una alerta demasiado bruta." -Owner "Valentina / Felix" -Action "Separar pagos a proveedores, remuneraciones, impuestos, prestamos, comisiones y otros movimientos operacionales." -Ref "fin_mov_bancarios_cargos")
  }
  if ($movAbonosSinConciliar.Count -gt 0) {
    $monto = ($movAbonosSinConciliar | ForEach-Object { Get-Number $_.abono } | Measure-Object -Sum).Sum
    Add-Issue $issues (New-FinanceIssue -Code "F-B02" -Severity "rojo" -Area "Caja" -Title "Abonos bancarios sin conciliar" -Detail "$($movAbonosSinConciliar.Count) entrada(s) de banco sin trazabilidad por aprox. $(Format-Clp $monto). No se muestra como total banco para evitar una alerta demasiado bruta." -Owner "Valentina / Felix" -Action "Separar cobros de clientes, aportes, vale vista, rechazos y transferencias internas." -Ref "fin_mov_bancarios_abonos")
  }
  foreach ($f in @($apVencidas | Sort-Object { -1 * (Get-Number $_.montoTotal) } | Select-Object -First $maxOverdue)) {
    Add-Issue $issues (New-FinanceIssue -Code "F-CXP01" -Severity "rojo" -Area "Caja" -Title "CxP vencida" -Detail "$($f.razonSocialContraparte) folio $($f.folio) vencio $($f.fechaVencimiento), monto $(Format-Clp (Get-Number $f.montoTotal))." -Owner "Valentina / Felix" -Action "Decidir pago o renegociacion." -Ref $f.id)
  }
  if ($apVencidas.Count -gt $maxOverdue) {
    $rest = @($apVencidas | Sort-Object { -1 * (Get-Number $_.montoTotal) } | Select-Object -Skip $maxOverdue)
    $restMonto = ($rest | ForEach-Object { Get-Number $_.montoTotal } | Measure-Object -Sum).Sum
    Add-Issue $issues (New-FinanceIssue -Code "F-CXP02" -Severity "rojo" -Area "Caja" -Title "CxP vencida adicional agrupada" -Detail "$($rest.Count) factura(s) CxP vencidas adicionales por aprox. $(Format-Clp $restMonto)." -Owner "Valentina / Felix" -Action "Revisar lista completa en Finanzas." -Ref "fin_facturas_ap")
  }
  foreach ($f in @($arVencidas | Sort-Object { -1 * (Get-Number $_.montoTotal) } | Select-Object -First $maxOverdue)) {
    Add-Issue $issues (New-FinanceIssue -Code "F-CXC01" -Severity "rojo" -Area "Caja" -Title "CxC vencida" -Detail "$($f.razonSocialContraparte) folio $($f.folio) vencio $($f.fechaVencimiento), monto $(Format-Clp (Get-Number $f.montoTotal))." -Owner "Valentina / Comercial" -Action "Gestionar cobranza." -Ref $f.id)
  }
  if ($arVencidas.Count -gt $maxOverdue) {
    $rest = @($arVencidas | Sort-Object { -1 * (Get-Number $_.montoTotal) } | Select-Object -Skip $maxOverdue)
    $restMonto = ($rest | ForEach-Object { Get-Number $_.montoTotal } | Measure-Object -Sum).Sum
    Add-Issue $issues (New-FinanceIssue -Code "F-CXC02" -Severity "rojo" -Area "Caja" -Title "CxC vencida adicional agrupada" -Detail "$($rest.Count) factura(s) CxC vencidas adicionales por aprox. $(Format-Clp $restMonto)." -Owner "Valentina / Comercial" -Action "Revisar lista completa en Finanzas." -Ref "fin_facturas_ar")
  }
  if ($apSinProyecto.Count -gt 0) {
    $monto = ($apSinProyecto | ForEach-Object { Get-Number $_.montoTotal } | Measure-Object -Sum).Sum
    Add-Issue $issues (New-FinanceIssue -Code "F-CXP03" -Severity "amarillo" -Area "Finanzas" -Title "Facturas CxP sin proyecto agrupadas" -Detail "$($apSinProyecto.Count) factura(s) CxP sin proyecto por aprox. $(Format-Clp $monto). Pueden superponerse con sin clasificar." -Owner "Valentina" -Action "Clasificar por lote en Finanzas para destrabar costos por proyecto." -Ref "fin_facturas_ap")
  }
  if ($apSinClasificar.Count -gt 0) {
    $monto = ($apSinClasificar | ForEach-Object { Get-Number $_.montoTotal } | Measure-Object -Sum).Sum
    Add-Issue $issues (New-FinanceIssue -Code "F-CXP04" -Severity "amarillo" -Area "Finanzas" -Title "Facturas CxP sin clasificar agrupadas" -Detail "$($apSinClasificar.Count) factura(s) CxP en SIN_CLASIFICAR por aprox. $(Format-Clp $monto)." -Owner "Valentina" -Action "Clasificar linea/cuenta/proyecto por lote." -Ref "fin_facturas_ap")
  }
  if ($arSinProyecto.Count -gt 0) {
    $monto = ($arSinProyecto | ForEach-Object { Get-Number $_.montoTotal } | Measure-Object -Sum).Sum
    Add-Issue $issues (New-FinanceIssue -Code "F-CXC03" -Severity "amarillo" -Area "Finanzas" -Title "Facturas CxC sin proyecto agrupadas" -Detail "$($arSinProyecto.Count) factura(s) CxC sin proyecto por aprox. $(Format-Clp $monto)." -Owner "Valentina" -Action "Vincular a CRM/proyecto para lectura comercial y directorio." -Ref "fin_facturas_ar")
  }
  $ocPressure = 0.0
  foreach ($oc in @($Data.mat_ordenes | Where-Object { ([string]$_.status) -notin @("total","recibida_total","cerrada","anulada") })) {
    $ocPressure += ((Get-Number $oc.qty) * (Get-Number $oc.precioUnit))
  }
  if ($ocPressure -gt 0) {
    Add-Issue $issues (New-FinanceIssue -Code "F-OC01" -Severity "info" -Area "Caja" -Title "Presion de OCs sobre caja" -Detail "OCs abiertas comprometen aprox. $([Math]::Round($ocPressure,0)) CLP." -Owner "Felix / Valentina" -Action "Contrastar contra flujo 13 semanas." -Ref "mat_ordenes")
  }
  @($issues)
}

function Build-FinanzasReport {
  param([object]$Config, [object]$Data, [datetime]$Now)
  $issues = @(Get-FinanceIssues -Config $Config -Data $Data -Now $Now)
  [pscustomobject]@{
    generatedAt = $Now.ToString("o")
    date = $Now.ToString("yyyy-MM-dd")
    summary = [pscustomobject]@{
      rojas = @($issues | Where-Object { $_.severity -eq "rojo" }).Count
      amarillas = @($issues | Where-Object { $_.severity -eq "amarillo" }).Count
      informativas = @($issues | Where-Object { $_.severity -eq "info" }).Count
      total = @($issues).Count
    }
    issues = $issues
  }
}

function Get-FinanzasAdminSkillDefinition {
  param([string]$Code)
  switch -Wildcard ($Code) {
    "F-B01" {
      return [pscustomobject]@{
        key = "FIN-BAN-CARGOS"
        name = "Destino recurrente de cargos bancarios"
        condition = "Cargos bancarios repetidos con contraparte/glosa/monto dentro de rango aprobado."
        action = "Proponer clasificacion y link operacional; ejecutar solo cuando exista regla aprobada por Valentina."
      }
    }
    "F-B02" {
      return [pscustomobject]@{
        key = "FIN-BAN-ABONOS"
        name = "Destino recurrente de abonos bancarios"
        condition = "Abonos repetidos con cliente, aporte, reembolso o transferencia identificable."
        action = "Proponer cobranza/aporte/reembolso; ejecutar solo con regla aprobada y evidencia suficiente."
      }
    }
    "F-B03" {
      return [pscustomobject]@{
        key = "FIN-BAN-RESPALDOS"
        name = "Respaldo pendiente por patron operacional"
        condition = "Movimientos ya explicados operacionalmente, pero con respaldo documental pendiente."
        action = "Agrupar y pedir una decision por tipo de respaldo; no cerrar contablemente hasta resolver evidencia."
      }
    }
    "F-CXP03" {
      return [pscustomobject]@{
        key = "FIN-CXP-DESTINO"
        name = "Destino CxP por proveedor/proyecto"
        condition = "Facturas CxP sin destino que calzan con OC, proyecto, OPEX o proveedor recurrente."
        action = "Proponer proyecto/OPEX/categoria; activar regla solo para casos identicos futuros."
      }
    }
    "F-CXP04" {
      return [pscustomobject]@{
        key = "FIN-CXP-CLASIFICACION"
        name = "Clasificacion CxP recurrente"
        condition = "Facturas CxP SIN_CLASIFICAR con proveedor, glosa y cuenta repetible."
        action = "Proponer linea/cuenta/proyecto; ejecutar automaticamente solo despues de aprobacion de skill."
      }
    }
    "F-CXC03" {
      return [pscustomobject]@{
        key = "FIN-CXC-PROYECTO"
        name = "Vinculo CxC con proyecto/CRM"
        condition = "Facturas CxC sin proyecto donde cliente/folio/negocio identifica un proyecto probable."
        action = "Proponer vinculo a proyecto; dejar excepcion si hay ambiguedad comercial."
      }
    }
    "F-BHE*" {
      return [pscustomobject]@{
        key = "FIN-BHE-CONCILIACION"
        name = "Conciliacion de boletas de honorarios"
        condition = "BHE con folio, prestador y monto liquido que calzan contra banco."
        action = "Proponer match seguro; evitar duplicar como AP normal."
      }
    }
    "F-DTE*" {
      return [pscustomobject]@{
        key = "FIN-DTE-RESPALDO"
        name = "Respaldo tributario DTE/RCV"
        condition = "Documento operativo cargado, pero pendiente de cruce DTE/RCV."
        action = "Mantener como gestion hasta validar respaldo tributario; no usar como cierre."
      }
    }
    default { return $null }
  }
}

function New-FinanzasAdminTask {
  param([object]$Issue, [int]$Index)
  $code = [string]$Issue.code
  $skill = Get-FinanzasAdminSkillDefinition -Code $code
  $needsValentina = ([string]$Issue.owner -match "Valentina") -or ($code -like "F-CXP*") -or ($code -like "F-CXC*") -or ($code -like "F-B*") -or ($code -like "F-DTE*")
  $sensitive = ($code -like "F-CXP01" -or $code -like "F-CXP02" -or $code -like "F-CXC01" -or $code -like "F-CXC02" -or $code -like "F-AUTO*" -or $code -like "F-CONF*")
  $mode = if ($sensitive) { "preguntar_aprobacion" } elseif ($null -ne $skill) { "proponer_skill" } else { "preparar_tarea" }
  $question = if ($needsValentina) {
    "Valentina, confirmame criterio para '$($Issue.title)'. Propongo: $($Issue.action)"
  } else {
    "Preparar resolucion operativa y dejar evidencia antes de pedir aprobacion."
  }
  [pscustomobject][ordered]@{
    id = ("FIN-ADMIN-{0:000}" -f $Index)
    code = $code
    severity = [string]$Issue.severity
    area = [string]$Issue.area
    title = [string]$Issue.title
    detail = [string]$Issue.detail
    action = [string]$Issue.action
    owner = if ($needsValentina) { "Valentina" } else { [string]$Issue.owner }
    ref = [string]$Issue.ref
    mode = $mode
    needsValentina = [bool]$needsValentina
    canAutoExecuteNow = $false
    safetyReason = "Piloto Nivel 1: no modifica pagos, cierres, impuestos, deuda, caja ni datos financieros sin regla aprobada."
    question = $question
    skillCandidate = [bool]($null -ne $skill)
    skillKey = if ($skill) { [string]$skill.key } else { "" }
    skillName = if ($skill) { [string]$skill.name } else { "" }
  }
}

function Get-FinanzasAdminText {
  param([object]$Item, [string[]]$Names, [string]$Default = "")
  if ($null -eq $Item) { return $Default }
  foreach ($name in @($Names)) {
    $prop = $Item.PSObject.Properties[$name]
    if ($prop -and -not [string]::IsNullOrWhiteSpace([string]$prop.Value)) {
      return [string]$prop.Value
    }
  }
  $Default
}

function Test-FinanzasAdminBankExplained {
  param([object]$Mov)
  if ($null -eq $Mov) { return $true }
  if ($Mov.conciliado -or $Mov.pagoId -or $Mov.cobranzaId -or $Mov.linkNomina -or $Mov.linkOperacional) { return $true }
  if ($Mov.pagoIds -and @($Mov.pagoIds).Count -gt 0) { return $true }
  if ($Mov.cobranzaIds -and @($Mov.cobranzaIds).Count -gt 0) { return $true }
  $false
}

function Get-FinanzasAdminBankSuggestion {
  param([object]$Mov, [string]$Direction)
  $desc = (Get-FinanzasAdminText -Item $Mov -Names @("descripcion", "detalle", "glosa", "documento") -Default "").ToLowerInvariant()
  $amount = if ($Direction -eq "cargo") { Get-Number $Mov.cargo } else { Get-Number $Mov.abono }
  $kind = if ($Direction -eq "cargo") { "PAGO_PROVEEDOR_O_GASTO" } else { "COBRANZA_CLIENTE_O_APORTE" }
  $confidence = "baja"
  $reason = "No hay link a factura, cobranza, nomina ni concepto operacional."
  if ($Direction -eq "cargo" -and $desc -match "previred|afp|isapre|fonasa|cotiz") {
    $kind = "NOMINA_PREVIRED"; $confidence = "media"; $reason = "La glosa parece pago previsional/remuneraciones."
  } elseif ($desc -match "tesoreria|tgr|sii|impuesto|iva") {
    $kind = "IMPUESTO_SII"; $confidence = "media"; $reason = "La glosa parece impuesto o pago al fisco."
  } elseif ($Direction -eq "cargo" -and $desc -match "comision|comisión|mantencion|mantención|interes|interés|cargo banco") {
    $kind = "COMISIONES_O_INTERESES"; $confidence = "media"; $reason = "La glosa parece cargo bancario recurrente."
  } elseif ($desc -match "traspaso|transferencia interna|cuenta propia") {
    $kind = "TRANSFERENCIA_INTERNA"; $confidence = "media"; $reason = "La glosa sugiere movimiento entre cuentas."
  } elseif ($Direction -eq "abono" -and $desc -match "aporte|capital|socio|prestamo|préstamo") {
    $kind = "APORTE_O_PRESTAMO_RECIBIDO"; $confidence = "media"; $reason = "La glosa parece aporte o prestamo recibido."
  } elseif ($Direction -eq "abono" -and $desc -match "vale vista|rechazo|devolucion|devolución|reembolso") {
    $kind = "VALE_VISTA_RECHAZO_O_REEMBOLSO"; $confidence = "media"; $reason = "La glosa parece regularizacion bancaria o devolucion."
  } elseif ($Direction -eq "abono") {
    $kind = "COBRANZA_CLIENTE"; $reason = "Es una entrada sin cobranza asociada; primer supuesto a revisar."
  }
  [pscustomobject][ordered]@{
    kind = $kind
    confidence = $confidence
    reason = $reason
    proposedAction = if ($Direction -eq "cargo") {
      "Revisar si corresponde a CxP, nomina, impuesto, comision, prestamo o transferencia interna antes de usarlo para caja/EERR."
    } else {
      "Revisar si corresponde a cobranza cliente, aporte/prestamo, reembolso, vale vista/rechazo o transferencia interna."
    }
    amount = $amount
  }
}

function New-FinanzasAdminChoice {
  param(
    [string]$Key,
    [string]$Label,
    [string]$Value,
    [string]$Effect,
    [string]$Outcome = "CRITERIO_REGISTRADO",
    [bool]$Recommended = $false
  )
  [pscustomobject][ordered]@{
    key = $Key
    label = $Label
    value = $Value
    effect = $Effect
    outcome = $Outcome
    recommended = $Recommended
  }
}

function Get-FinanzasAdminChoiceSet {
  param([string]$Code, [string]$SuggestionKind)
  $choices = @()
  switch -Wildcard ($Code) {
    "F-B01" {
      $defs = @(
        @("A", "Pago a proveedor / CxP", "PAGO_PROVEEDOR_O_GASTO", "Buscar factura CxP, OC o respaldo operacional y dejar match propuesto."),
        @("B", "Nomina / Previred", "NOMINA_PREVIRED", "Tratar como remuneraciones o pago previsional; pedir respaldo de nomina si falta."),
        @("C", "Impuesto SII / TGR", "IMPUESTO_SII", "Tratar como impuesto; pedir respaldo tributario antes de cierre."),
        @("D", "Comision o interes bancario", "COMISIONES_O_INTERESES", "Clasificar como gasto financiero si la glosa lo confirma."),
        @("E", "Transferencia interna", "TRANSFERENCIA_INTERNA", "Buscar contraparte bancaria interna y evitar duplicar gasto."),
        @("F", "Bloquear: no calza ninguna", "BLOQUEAR_REVISION", "Mantener pendiente y pedir revision manual.", "BLOQUEADO")
      )
      foreach ($d in $defs) {
        $choices += New-FinanzasAdminChoice -Key $d[0] -Label $d[1] -Value $d[2] -Effect $d[3] -Outcome $(if ($d.Count -ge 5) { $d[4] } else { "CRITERIO_REGISTRADO" }) -Recommended ([string]$SuggestionKind -eq [string]$d[2])
      }
    }
    "F-B02" {
      $defs = @(
        @("A", "Cobranza de cliente / CxC", "COBRANZA_CLIENTE", "Buscar factura CxC o cobranza asociada y dejar match propuesto."),
        @("B", "Aporte o prestamo recibido", "APORTE_O_PRESTAMO_RECIBIDO", "Tratar como financiamiento/aporte; pedir respaldo si falta."),
        @("C", "Reembolso o devolucion", "VALE_VISTA_RECHAZO_O_REEMBOLSO", "Tratar como regularizacion bancaria o devolucion."),
        @("D", "Transferencia interna", "TRANSFERENCIA_INTERNA", "Buscar contraparte bancaria interna y evitar duplicar ingreso."),
        @("E", "Otro abono operativo", "OTRO_ABONO", "Mantener como abono explicado con criterio manual."),
        @("F", "Bloquear: no calza ninguna", "BLOQUEAR_REVISION", "Mantener pendiente y pedir revision manual.", "BLOQUEADO")
      )
      foreach ($d in $defs) {
        $choices += New-FinanzasAdminChoice -Key $d[0] -Label $d[1] -Value $d[2] -Effect $d[3] -Outcome $(if ($d.Count -ge 5) { $d[4] } else { "CRITERIO_REGISTRADO" }) -Recommended ([string]$SuggestionKind -eq [string]$d[2])
      }
    }
    "F-CXP03" {
      $choices = @(
        New-FinanzasAdminChoice -Key "A" -Label "Proyecto / costo directo" -Value "PROYECTO_COSTO_DIRECTO" -Effect "Asignar a proyecto o pedir proyecto si falta; categoria costo directo."
        New-FinanzasAdminChoice -Key "B" -Label "OPEX fijo corporativo" -Value "OPEX_CORP_OPEX_FIJO" -Effect "Clasificar como OPEX_CORP + OPEX_FIJO."
        New-FinanzasAdminChoice -Key "C" -Label "OPEX variable" -Value "OPEX_CORP_OPEX_VARIABLE" -Effect "Clasificar como OPEX_CORP + OPEX_VARIABLE."
        New-FinanzasAdminChoice -Key "D" -Label "Stock / control especial" -Value "STOCK_CONTROL_ESPECIAL" -Effect "No llevar directo a proyecto; pedir respaldo de bodega/control."
        New-FinanzasAdminChoice -Key "E" -Label "Inversion, impuesto, financiero u otro" -Value "NO_OPERACIONAL_ESPECIAL" -Effect "Clasificar fuera de costo directo/OPEX normal segun respaldo."
        New-FinanzasAdminChoice -Key "F" -Label "Bloquear: falta respaldo" -Value "BLOQUEAR_REVISION" -Effect "Mantener pendiente hasta que exista respaldo suficiente." -Outcome "BLOQUEADO"
      )
    }
    "F-CXP04" {
      $choices = @(
        New-FinanzasAdminChoice -Key "A" -Label "Pods / costo directo" -Value "POD_COSTO_DIRECTO" -Effect "Linea Pods + costo directo de obra."
        New-FinanzasAdminChoice -Key "B" -Label "Momentum / costo directo" -Value "MODULAR_COSTO_DIRECTO" -Effect "Linea Momentum + costo directo de obra."
        New-FinanzasAdminChoice -Key "C" -Label "Soluciones transitorias / costo directo" -Value "EMERGENCIA_COSTO_DIRECTO" -Effect "Linea Soluciones transitorias + costo directo de obra."
        New-FinanzasAdminChoice -Key "D" -Label "Galpones / industrial / costo directo" -Value "INDUSTRIAL_COSTO_DIRECTO" -Effect "Linea Galpones/industrial + costo directo de obra."
        New-FinanzasAdminChoice -Key "E" -Label "OPEX corporativo" -Value "OPEX_CORP" -Effect "Gasto admin/comercial/I+D; luego precisar OPEX fijo o variable."
        New-FinanzasAdminChoice -Key "F" -Label "Impuesto, financiero, inversion u otro" -Value "NO_OPERACIONAL_ESPECIAL" -Effect "Clasificar fuera de costo directo normal."
        New-FinanzasAdminChoice -Key "G" -Label "Bloquear: falta criterio" -Value "BLOQUEAR_REVISION" -Effect "Mantener sin clasificar hasta revision manual." -Outcome "BLOQUEADO"
      )
    }
    "F-CXC03" {
      $choices = @(
        New-FinanzasAdminChoice -Key "A" -Label "Vincular a proyecto CRM probable" -Value "CXC_PROYECTO_CRM" -Effect "Buscar proyecto/negocio CRM por cliente, folio o glosa antes de cierre."
        New-FinanzasAdminChoice -Key "B" -Label "Ingreso Pods" -Value "POD_INGRESO" -Effect "Linea Pods + ingreso; pedir proyecto si corresponde."
        New-FinanzasAdminChoice -Key "C" -Label "Ingreso Momentum" -Value "MODULAR_INGRESO" -Effect "Linea Momentum + ingreso; pedir proyecto si corresponde."
        New-FinanzasAdminChoice -Key "D" -Label "Ingreso Soluciones transitorias" -Value "EMERGENCIA_INGRESO" -Effect "Linea Soluciones transitorias + ingreso; pedir proyecto si corresponde."
        New-FinanzasAdminChoice -Key "E" -Label "Ingreso Galpones / industrial" -Value "INDUSTRIAL_INGRESO" -Effect "Linea Galpones/industrial + ingreso; pedir proyecto si corresponde."
        New-FinanzasAdminChoice -Key "F" -Label "Excepcion sin proyecto" -Value "CXC_EXCEPCION_SIN_PROYECTO" -Effect "Dejar documentado que no corresponde vinculo a proyecto."
        New-FinanzasAdminChoice -Key "G" -Label "Bloquear: revisar con Comercial" -Value "BLOQUEAR_REVISION" -Effect "Escalar antes de clasificar." -Outcome "BLOQUEADO"
      )
    }
    default {
      $choices = @(
        New-FinanzasAdminChoice -Key "A" -Label "Aprobar propuesta del agente" -Value "APROBAR_PROPUESTA" -Effect "Registrar criterio como aprobado para este caso." -Outcome "APROBADO"
        New-FinanzasAdminChoice -Key "B" -Label "Corregir criterio" -Value "CORREGIR_CRITERIO" -Effect "Responder con el criterio correcto para que el agente aprenda." -Outcome "CORREGIDO"
        New-FinanzasAdminChoice -Key "C" -Label "Bloquear" -Value "BLOQUEAR_REVISION" -Effect "No usar este caso hasta revision manual." -Outcome "BLOQUEADO"
        New-FinanzasAdminChoice -Key "D" -Label "Rechazar" -Value "RECHAZAR_PROPUESTA" -Effect "Descartar la propuesta para este caso." -Outcome "RECHAZADO"
      )
    }
  }
  @($choices)
}

function New-FinanzasAdminCase {
  param(
    [string]$Id,
    [string]$Code,
    [string]$Severity,
    [string]$Domain,
    [string]$Title,
    [string]$Detail,
    [double]$Amount,
    [string]$SourceCollection,
    [string]$SourceId,
    [string]$SourceDate,
    [string]$SourceLabel,
    [string]$SuggestionKind,
    [string]$Confidence,
    [string]$ProposedAction,
    [string]$Question,
    [object[]]$Choices = @(),
    [string]$SkillKey,
    [string]$SkillName,
    [datetime]$Now,
    [object]$Existing = $null
  )
  $status = if ($Existing -and $Existing.status) { [string]$Existing.status } else { "PENDIENTE_VALENTINA" }
  $decision = if ($Existing -and $Existing.decision) { [string]$Existing.decision } else { "" }
  $decidedBy = if ($Existing -and $Existing.decidedBy) { [string]$Existing.decidedBy } else { "" }
  $decidedAt = if ($Existing -and $Existing.decidedAt) { [string]$Existing.decidedAt } else { "" }
  $mailCode = if ($Existing -and $Existing.mailCode) { [string]$Existing.mailCode } else { "ADM-" + (Get-StableShortCode -Value $Id) }
  [pscustomobject][ordered]@{
    id = $Id
    mailCode = $mailCode
    generatedAt = $Now.ToString("o")
    date = $Now.ToString("yyyy-MM-dd")
    status = $status
    decision = $decision
    decidedBy = $decidedBy
    decidedAt = $decidedAt
    code = $Code
    severity = $Severity
    domain = $Domain
    title = $Title
    detail = $Detail
    amount = [Math]::Round($Amount, 0)
    sourceCollection = $SourceCollection
    sourceId = $SourceId
    sourceDate = $SourceDate
    sourceLabel = $SourceLabel
    suggestionKind = $SuggestionKind
    confidence = $Confidence
    proposedAction = $ProposedAction
    question = $Question
    choices = @($Choices)
    replyFormat = "${mailCode}: A"
    skillKey = $SkillKey
    skillName = $SkillName
    mode = "shadow_proposal"
    canAutoExecuteNow = $false
    safetyReason = "Nivel 2 sin Valentina: propuesta individual en modo sombra; no modifica datos financieros."
    nextStep = "Valentina debe aprobar, corregir o rechazar. La respuesta alimenta skills candidatas."
  }
}

function Get-StableShortCode {
  param([string]$Value)
  $bytes = [System.Text.Encoding]::UTF8.GetBytes([string]$Value)
  $sha = [System.Security.Cryptography.SHA1]::Create()
  try {
    $hash = $sha.ComputeHash($bytes)
    (($hash | Select-Object -First 4 | ForEach-Object { $_.ToString("x2") }) -join "").ToUpperInvariant()
  } finally {
    $sha.Dispose()
  }
}

function Get-FinanzasAdminExistingMap {
  param([object[]]$Cases)
  $map = @{}
  foreach ($case in @($Cases)) {
    if ($case.id) { $map[[string]$case.id] = $case }
  }
  $map
}

function Get-FinanzasAdminCases {
  param([object]$Config, [object]$Data, [datetime]$Now)
  $existing = Get-FinanzasAdminExistingMap -Cases @($Data.fin_admin_cases)
  $cases = @()
  foreach ($mov in @($Data.fin_mov_bancarios | Sort-Object fecha -Descending)) {
    if (Test-FinanzasAdminBankExplained -Mov $mov) { continue }
    $cargo = Get-Number $mov.cargo
    $abono = Get-Number $mov.abono
    if ($cargo -le 0 -and $abono -le 0) { continue }
    $direction = if ($cargo -gt 0) { "cargo" } else { "abono" }
    $code = if ($direction -eq "cargo") { "F-B01" } else { "F-B02" }
    $skill = Get-FinanzasAdminSkillDefinition -Code $code
    $suggestion = Get-FinanzasAdminBankSuggestion -Mov $mov -Direction $direction
    $sourceId = Get-FinanzasAdminText -Item $mov -Names @("id", "externalId", "clayId") -Default ([guid]::NewGuid().ToString("N"))
    $id = "bank-$direction-$sourceId" -replace '[^A-Za-z0-9_-]', '_'
    $label = "$(Get-FinanzasAdminText -Item $mov -Names @("fecha")) / $(Get-FinanzasAdminText -Item $mov -Names @("descripcion", "documento"))"
    $title = if ($direction -eq "cargo") { "Cargo bancario sin destino" } else { "Abono bancario sin destino" }
    $amount = [double]$suggestion.amount
    $question = "Valentina, este $direction bancario por $(Format-Clp $amount) figura sin destino. Propongo revisarlo como $($suggestion.kind). ¿Lo apruebo para este caso, lo corrijo o lo dejo bloqueado?"
    $choices = @(Get-FinanzasAdminChoiceSet -Code $code -SuggestionKind ([string]$suggestion.kind))
    $cases += New-FinanzasAdminCase -Id $id -Code $code -Severity "rojo" -Domain "Banco" -Title $title -Detail $suggestion.reason -Amount $amount -SourceCollection ([string]$Config.collections.fin_mov_bancarios) -SourceId $sourceId -SourceDate ([string]$mov.fecha) -SourceLabel $label -SuggestionKind ([string]$suggestion.kind) -Confidence ([string]$suggestion.confidence) -ProposedAction ([string]$suggestion.proposedAction) -Question $question -Choices $choices -SkillKey ([string]$skill.key) -SkillName ([string]$skill.name) -Now $Now -Existing $existing[$id]
  }

  foreach ($f in @($Data.fin_facturas_ap | Sort-Object fechaEmision -Descending)) {
    $estado = [string]$f.estado
    if (@("PAGADA", "ANULADA", "RECHAZADA") -contains $estado) { continue }
    $isNotaCredito = (([int](Get-Number $f.tipoDte)) -eq 61 -or ([int](Get-Number $f.tipoDte)) -eq 112)
    if ($isNotaCredito) { continue }
    $sourceId = Get-FinanzasAdminText -Item $f -Names @("id") -Default ([guid]::NewGuid().ToString("N"))
    $label = "$($f.razonSocialContraparte) folio $($f.folio)"
    $sinDestino = ((-not $f.proyectoId) -and (-not $f.asignaciones) -and ([string]$f.lineaNegocio) -ne "OPEX_CORP" -and ([string]$f.categoriaContable) -notin @("OPEX_FIJO", "OPEX_VARIABLE"))
    if ($sinDestino) {
      $code = "F-CXP03"
      $skill = Get-FinanzasAdminSkillDefinition -Code $code
      $id = "ap-destino-$sourceId" -replace '[^A-Za-z0-9_-]', '_'
      $question = "Valentina, esta CxP de $($f.razonSocialContraparte) por $(Format-Clp (Get-Number $f.montoTotal)) no tiene destino. Propongo clasificarla como proyecto, OPEX o control especial segun respaldo. ¿Cual criterio uso?"
      $choices = @(Get-FinanzasAdminChoiceSet -Code $code -SuggestionKind "CLASIFICAR_DESTINO_CXP")
      $cases += New-FinanzasAdminCase -Id $id -Code $code -Severity "amarillo" -Domain "CxP" -Title "Factura CxP sin destino" -Detail "No tiene proyecto/asignacion ni clasificacion OPEX suficiente." -Amount (Get-Number $f.montoTotal) -SourceCollection ([string]$Config.collections.fin_facturas_ap) -SourceId $sourceId -SourceDate ([string]$f.fechaEmision) -SourceLabel $label -SuggestionKind "CLASIFICAR_DESTINO_CXP" -Confidence "media" -ProposedAction "Asignar proyecto, OPEX, stock/control especial o excepcion documentada." -Question $question -Choices $choices -SkillKey ([string]$skill.key) -SkillName ([string]$skill.name) -Now $Now -Existing $existing[$id]
    }
    if ([string]$f.lineaNegocio -eq "SIN_CLASIFICAR") {
      $code = "F-CXP04"
      $skill = Get-FinanzasAdminSkillDefinition -Code $code
      $id = "ap-clasificacion-$sourceId" -replace '[^A-Za-z0-9_-]', '_'
      $question = "Valentina, esta CxP de $($f.razonSocialContraparte) esta SIN_CLASIFICAR. Propongo definir linea/categoria/cuenta y guardar el criterio como candidato de skill si se repite."
      $choices = @(Get-FinanzasAdminChoiceSet -Code $code -SuggestionKind "CLASIFICAR_CXP")
      $cases += New-FinanzasAdminCase -Id $id -Code $code -Severity "amarillo" -Domain "CxP" -Title "Factura CxP sin clasificacion" -Detail "Linea de negocio SIN_CLASIFICAR; EERR de gestion incompleto." -Amount (Get-Number $f.montoTotal) -SourceCollection ([string]$Config.collections.fin_facturas_ap) -SourceId $sourceId -SourceDate ([string]$f.fechaEmision) -SourceLabel $label -SuggestionKind "CLASIFICAR_CXP" -Confidence "media" -ProposedAction "Completar linea, categoria, cuenta y proyecto/OPEX; proponer regla si proveedor/glosa se repite." -Question $question -Choices $choices -SkillKey ([string]$skill.key) -SkillName ([string]$skill.name) -Now $Now -Existing $existing[$id]
    }
  }

  foreach ($f in @($Data.fin_facturas_ar | Sort-Object fechaEmision -Descending)) {
    $estado = [string]$f.estado
    if (@("COBRADA", "ANULADA") -contains $estado) { continue }
    $isNotaCredito = (([int](Get-Number $f.tipoDte)) -eq 61 -or ([int](Get-Number $f.tipoDte)) -eq 112)
    if ($isNotaCredito) { continue }
    $sinProyecto = ((-not $f.crmProjectId) -and (-not $f.proyectoId) -and (-not $f.asignaciones))
    if (-not $sinProyecto) { continue }
    $code = "F-CXC03"
    $skill = Get-FinanzasAdminSkillDefinition -Code $code
    $sourceId = Get-FinanzasAdminText -Item $f -Names @("id") -Default ([guid]::NewGuid().ToString("N"))
    $id = "ar-proyecto-$sourceId" -replace '[^A-Za-z0-9_-]', '_'
    $label = "$($f.razonSocialContraparte) folio $($f.folio)"
    $question = "Valentina, esta CxC de $($f.razonSocialContraparte) por $(Format-Clp (Get-Number $f.montoTotal)) no tiene proyecto. Propongo vincularla a CRM/proyecto antes de usarla para directorio."
    $choices = @(Get-FinanzasAdminChoiceSet -Code $code -SuggestionKind "VINCULAR_CXC_PROYECTO")
    $cases += New-FinanzasAdminCase -Id $id -Code $code -Severity "amarillo" -Domain "CxC" -Title "Factura CxC sin proyecto" -Detail "No tiene CRM/projectId/asignaciones; afecta lectura comercial y directorio." -Amount (Get-Number $f.montoTotal) -SourceCollection ([string]$Config.collections.fin_facturas_ar) -SourceId $sourceId -SourceDate ([string]$f.fechaEmision) -SourceLabel $label -SuggestionKind "VINCULAR_CXC_PROYECTO" -Confidence "media" -ProposedAction "Vincular a proyecto/CRM o dejar excepcion comercial documentada." -Question $question -Choices $choices -SkillKey ([string]$skill.key) -SkillName ([string]$skill.name) -Now $Now -Existing $existing[$id]
  }

  @($cases | Sort-Object @{ Expression = { if ($_.severity -eq "rojo") { 0 } else { 1 } } }, @{ Expression = "amount"; Descending = $true })
}

function Build-FinanzasAdminReport {
  param([object]$Config, [object]$Data, [object]$FinanceReport, [datetime]$Now)
  $queue = @()
  $idx = 0
  foreach ($issue in @($FinanceReport.issues | Where-Object { $_.severity -in @("rojo", "amarillo") })) {
    $idx++
    $queue += New-FinanzasAdminTask -Issue $issue -Index $idx
  }

  $cases = @(Get-FinanzasAdminCases -Config $Config -Data $Data -Now $Now)
  $questions = @($cases | Where-Object { $_.status -eq "PENDIENTE_VALENTINA" } | Select-Object -First 12)
  $skillCandidates = @()
  foreach ($group in @($cases | Where-Object { $_.skillKey } | Group-Object skillKey)) {
    $first = @($group.Group | Select-Object -First 1)[0]
    $def = Get-FinanzasAdminSkillDefinition -Code ([string]$first.code)
    if ($null -eq $def) { continue }
    $skillCandidates += [pscustomobject][ordered]@{
      id = [string]$def.key
      name = [string]$def.name
      cases = @($group.Group).Count
      condition = [string]$def.condition
      proposedAction = [string]$def.action
      activation = "Requiere aprobacion explicita de Valentina: Activar skill desde ahora."
      scope = "Solo casos futuros; backlog historico requiere aprobacion por lote."
      examples = @($group.Group | Select-Object -First 3 | ForEach-Object { "$($_.code) - $($_.sourceLabel)" })
    }
  }

  [pscustomobject][ordered]@{
    generatedAt = $Now.ToString("o")
    date = $Now.ToString("yyyy-MM-dd")
    stage = "piloto_productivo_nivel_2"
    mandate = "Administrar operativamente Finanzas: ordenar pendientes, bajar alertas a casos individuales, proponer resoluciones en modo sombra y convertir respuestas repetidas en skills aprobables."
    safety = "Si no es 100% seguro, no ejecuta; pregunta. En Nivel 2 propone caso a caso y persiste la bandeja, sin autoejecutar cambios financieros."
    summary = [pscustomobject][ordered]@{
      pendientes = @($queue).Count
      preguntasValentina = @($questions).Count
      skillsCandidatas = @($skillCandidates).Count
      autoejecutadas = 0
      casosIndividuales = @($cases).Count
      casosModoSombra = @($cases | Where-Object { $_.mode -eq "shadow_proposal" }).Count
      rojas = @($queue | Where-Object { $_.severity -eq "rojo" }).Count
      amarillas = @($queue | Where-Object { $_.severity -eq "amarillo" }).Count
    }
    queue = $queue
    questions = $questions
    cases = $cases
    skillCandidates = $skillCandidates
    fiscalizerSummary = $FinanceReport.summary
  }
}

function Render-FinanzasAdminQuestionCards {
  param([object[]]$Cases)
  if (@($Cases).Count -eq 0) { return (New-MayuEmptyState -Text "Sin preguntas individuales pendientes.") }
  $html = ""
  foreach ($case in @($Cases)) {
    $tone = Get-MayuEmailTone -Tone ([string]$case.severity)
    $choices = @($case.choices)
    $choiceHtml = ""
    if ($choices.Count -gt 0) {
      foreach ($choice in $choices) {
        $recommended = if ($choice.recommended) { " <span style='color:#166534;font-weight:700;'>(propuesta del agente)</span>" } else { "" }
        $choiceHtml += "<li style='margin:0 0 6px 0;'><strong>$(HtmlEscape $choice.key)) $(HtmlEscape $choice.label)</strong>$recommended<br><span style='color:#5f6368;'>$(HtmlEscape $choice.effect)</span></li>"
      }
    } else {
      $choiceHtml = "<li><strong>A) Aprobar</strong></li><li><strong>B) Corregir criterio</strong></li><li><strong>C) Bloquear</strong></li><li><strong>D) Rechazar</strong></li>"
    }
    $replyExample = "$(HtmlEscape $case.mailCode): A"
    $html += @"
<div style="border:1px solid #e5e7eb;border-left:4px solid $($tone.accent);background:#ffffff;padding:12px 14px;margin:0 0 12px 0;">
  <div style="margin-bottom:6px;">
    <span style="display:inline-block;background:$($tone.soft);border:1px solid $($tone.accent);color:#202124;font-size:11px;line-height:1;text-transform:uppercase;letter-spacing:.3px;padding:5px 7px;">$(HtmlEscape $case.severity)</span>
    <span style="color:#6b7280;font-size:12px;margin-left:6px;">$(HtmlEscape $case.domain)</span>
  </div>
  <div style="font-weight:700;color:#202124;margin-bottom:4px;">$(HtmlEscape $case.mailCode) / $(HtmlEscape $case.code) - $(HtmlEscape $case.title)</div>
  <div style="color:#4b5563;margin-bottom:8px;">$(HtmlEscape $case.sourceLabel) - $(HtmlEscape $case.detail) - Monto: $(HtmlEscape (Format-Clp (Get-Number $case.amount)))</div>
  <div style="background:#f8fafc;border:1px solid #edf2f7;padding:8px 10px;color:#30343b;margin-bottom:10px;"><strong>Pregunta:</strong> $(HtmlEscape $case.question)</div>
  <div style="background:#fffef7;border:1px solid #f3e8b8;padding:8px 10px;color:#30343b;margin-bottom:10px;">
    <strong>Alternativas para responder:</strong>
    <ul style="margin:8px 0 0 18px;padding:0;">$choiceHtml</ul>
  </div>
  <div style="font-size:13px;background:#f4f7fb;border:1px solid #dbe7f3;padding:8px 10px;color:#30343b;">
    Responde una linea asi: <strong>$replyExample</strong>. Si ninguna alternativa calza: <strong>$(HtmlEscape $case.mailCode): corregir: [criterio]</strong>
  </div>
</div>
"@
  }
  $html
}

function Render-FinanzasAdminHtml {
  param([object]$Report)
  $s = $Report.summary
  $metrics = @(
    New-MayuEmailMetric -Label "Pendientes" -Value $s.pendientes -Tone "amarillo"
    New-MayuEmailMetric -Label "Casos detalle" -Value $s.casosIndividuales -Tone "info"
    New-MayuEmailMetric -Label "Preguntas" -Value $s.preguntasValentina -Tone "rojo"
    New-MayuEmailMetric -Label "Skills candidatas" -Value $s.skillsCandidatas -Tone "info"
    New-MayuEmailMetric -Label "Autoejecutadas" -Value $s.autoejecutadas -Tone "verde"
  ) -join ""

  $skillCards = @($Report.skillCandidates | ForEach-Object {
    [pscustomobject]@{
      severity = "info"
      area = "Skill candidata"
      code = $_.id
      title = $_.name
      detail = "Casos detectados: $($_.cases). Condicion: $($_.condition)"
      action = "$($_.proposedAction) Activacion: $($_.activation)"
      owner = "Valentina"
      ref = $_.scope
    }
  })
  $caseCards = @($Report.cases | Select-Object -First 40 | ForEach-Object {
    [pscustomobject]@{
      severity = $_.severity
      area = "$($_.domain) / $($_.confidence)"
      code = if ($_.mailCode) { "$($_.mailCode) / $($_.code)" } else { $_.code }
      title = $_.title
      detail = "$($_.sourceLabel) - $($_.detail) - Monto: $(Format-Clp (Get-Number $_.amount))"
      action = "$($_.proposedAction) Estado: $($_.status)."
      owner = "Administrador IA"
      ref = $_.id
    }
  })
  $skillsHtml = if ($skillCards.Count -eq 0) { New-MayuEmptyState -Text "Sin skills candidatas nuevas en esta corrida." } else { Render-IssueListCards -Items $skillCards }
  $questionsHtml = Render-FinanzasAdminQuestionCards -Cases @($Report.questions)
  $queueHtml = Render-IssueListCards -Items @($Report.queue | Select-Object -First 40)
  $casesHtml = if ($caseCards.Count -eq 0) { New-MayuEmptyState -Text "Sin casos individuales en modo sombra." } else { Render-IssueListCards -Items $caseCards }

  $content = @"
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;margin:0 0 12px 0;"><tr>$metrics</tr></table>
<div style="background:#f4f7fb;border-left:4px solid #0078d4;padding:12px 14px;color:#30343b;margin:12px 0 18px 0;">
  El Administrador Finanzas esta activo como piloto Nivel 2. Baja alertas agrupadas a casos individuales y trabaja en modo sombra: propone, persiste bandeja y aprende patrones candidatos, pero no modifica datos financieros sin aprobacion.
  <br><br>
  Para responder por correo, Valentina solo debe elegir una alternativa por caso, por ejemplo: <strong>ADM-XXXXXX: A</strong>. Si ninguna alternativa calza, puede responder: <strong>ADM-XXXXXX: corregir: [criterio]</strong>.
</div>
$(New-MayuEmailSection -Title "Preguntas para Valentina" -Html $questionsHtml)
$(New-MayuEmailSection -Title "Casos individuales modo sombra" -Html $casesHtml)
$(New-MayuEmailSection -Title "Skills candidatas" -Html $skillsHtml)
$(New-MayuEmailSection -Title "Cola operativa" -Html $queueHtml)
"@
  New-MayuEmailLayout -Title "Administrador Finanzas MAYU - $($Report.date)" -Subtitle "Piloto productivo Nivel 2: casos individuales, modo sombra y skills candidatas." -ContentHtml $content -Footer "El fiscalizador sigue separado: el administrador ordena y propone; el fiscalizador valida confiabilidad."
}

function Save-FinanzasAdminCases {
  param([object]$Config, [object]$Report)
  $collection = if ($Config.collections.fin_admin_cases) { [string]$Config.collections.fin_admin_cases } else { "fin_admin_cases" }
  $token = Get-FirestoreWriteToken -Config $Config
  $saved = 0
  foreach ($case in @($Report.cases)) {
    if ([string]::IsNullOrWhiteSpace([string]$case.id)) { continue }
    Set-FirestoreDocument -Config $Config -Token $token -CollectionName $collection -DocumentId ([string]$case.id) -Data $case
    $saved++
  }
  $saved
}

function Invoke-FinanzasAdmin {
  param([object]$Config, [string]$GraphToken, [string]$SiteId, [datetime]$Now, [bool]$DoSendEmail)
  if ($Config.agents.finanzas_admin -and $Config.agents.finanzas_admin.enabled -eq $false) {
    Write-Output "Administrador Finanzas: deshabilitado por config."
    return
  }
  Write-Output "Administrador Finanzas: leyendo Firestore."
  $data = Get-FirestoreData -Config $Config
  $financeReport = Build-FinanzasReport -Config $Config -Data $data -Now $Now
  $report = Build-FinanzasAdminReport -Config $Config -Data $data -FinanceReport $financeReport -Now $Now
  $folder = if ($Config.sharepoint.finanzas_admin_folder) { [string]$Config.sharepoint.finanzas_admin_folder } else { "agentes_mayu/finanzas/administrador" }
  $dateKey = $report.date
  $html = Render-FinanzasAdminHtml -Report $report
  Ensure-GraphFolder -Token $GraphToken -SiteId $SiteId -FolderPath $folder
  Write-TextFileToGraph -Token $GraphToken -SiteId $SiteId -FilePath "$folder/$dateKey.json" -Text ($report | ConvertTo-Json -Depth 80) -ContentType "application/json; charset=utf-8"
  Write-TextFileToGraph -Token $GraphToken -SiteId $SiteId -FilePath "$folder/$dateKey.html" -Text $html -ContentType "text/html; charset=utf-8"
  $savedCases = Save-FinanzasAdminCases -Config $Config -Report $report
  Write-Output "Administrador Finanzas: pendientes=$($report.summary.pendientes) casos=$($report.summary.casosIndividuales) guardados=$savedCases preguntas=$($report.summary.preguntasValentina) skills=$($report.summary.skillsCandidatas)."
  if ($DoSendEmail) {
    $to = Get-UniqueEmails -Emails @($Config.mail.valentina)
    $cc = Get-UniqueEmails -Emails @($Config.mail.felix)
    $subject = "[Finanzas Admin] Piloto $dateKey - $($report.summary.pendientes) pendientes, $($report.summary.skillsCandidatas) skills"
    Send-GraphMail -Token $GraphToken -Sender $Config.mail.sender -To $to -Cc $cc -Subject $subject -HtmlBody $html
    Write-Output "Administrador Finanzas: correo enviado a $($to -join ', ')."
  } else {
    Write-Output "Administrador Finanzas: SendEmail=false, no se envia correo."
  }
  Write-Output "Administrador Finanzas: outputs guardados en SharePoint."
}

function Render-FinanzasHtml {
  param([object]$Report)
  return Render-FinanzasHtmlV2 -Report $Report
  $s = $Report.summary
  @"
<html>
<body style="font-family:Arial,sans-serif;color:#222;max-width:980px;font-size:14px;">
  <h2 style="margin-bottom:4px;">Agente Finanzas MAYU - $($Report.date)</h2>
  <p style="margin-top:0;color:#555;">Rojas: <strong style="color:#991b1b;">$($s.rojas)</strong> &middot; Amarillas: <strong style="color:#9a3412;">$($s.amarillas)</strong> &middot; Informativas: <strong style="color:#2563eb;">$($s.informativas)</strong></p>

  <h3>Alertas accionables</h3>
  $(Render-IssueList -Items @($Report.issues | Where-Object { $_.severity -in @("rojo", "amarillo") }))

  <h3>Contexto gerencial</h3>
  $(Render-IssueList -Items @($Report.issues | Where-Object { $_.severity -eq "info" }))

  <h3>Como pedir ayuda</h3>
  <p>Responde este correo con el codigo de la alerta, por ejemplo <strong>F-B01</strong>, o copia la linea del problema. El respondedor de Finanzas te dira que hacer en la app para resolverlo.</p>
  <p style="font-size:12px;color:#777;">Este agente no corrige datos por correo. Solo reporta inconsistencias que afectan caja, CxP, CxC, clasificacion, proyectos o fuentes desactualizadas.</p>
</body>
</html>
"@
}

function Get-FinanzasHelpDefinition {
  param([string]$Code)
  switch -Regex ($Code) {
    "^F-D0" {
      [pscustomobject]@{
        title = "$Code - Fuente financiera desactualizada o sin datos"
        problem = "La app Finanzas no tiene una fuente al dia, como cartola bancaria o RCV. Esto afecta lectura gerencial de caja, CxP, CxC y EERR."
        steps = @(
          "Entrar a Finanzas > Importar.",
          "Subir la fuente actualizada que corresponda: cartola BICE, RCV Compras o RCV Ventas.",
          "Revisar el resumen de importacion y confirmar que la fecha maxima avance.",
          "Volver a correr el agente Finanzas.",
          "Si el archivo nuevo igual queda atrasado, revisar si el archivo descargado del banco/SII venia incompleto."
        )
      }
    }
    "^F-B01" {
      [pscustomobject]@{
        title = "F-B01 - Cargos bancarios sin conciliar"
        problem = "Hay salidas de banco sin trazabilidad. Gerencialmente esto impide saber que proveedor, remuneracion, impuesto, prestamo o gasto ya fue pagado."
        steps = @(
          "Entrar a Finanzas > Banco / Conciliacion.",
          "Filtrar cargos sin conciliar.",
          "Para cada cargo, vincular factura CxP, remuneracion, impuesto, prestamo, transferencia interna o gasto operacional.",
          "Si no existe documento, clasificarlo con nota operacional y responsable.",
          "Si afecta proyecto o inventario, dejar proyecto/linea/cuenta antes de cerrar la conciliacion."
        )
      }
    }
    "^F-B02" {
      [pscustomobject]@{
        title = "F-B02 - Abonos bancarios sin conciliar"
        problem = "Hay entradas de banco sin trazabilidad. Esto puede mezclar cobros de clientes, aportes de capital, prestamos de socios, reembolsos o reversas."
        steps = @(
          "Entrar a Finanzas > Banco / Conciliacion.",
          "Filtrar abonos sin conciliar.",
          "Separar cobros de clientes, aportes de capital, prestamos de socios, reembolsos, vale vista rechazados y reversas.",
          "Vincular CxC cuando sea cobranza real.",
          "Cuando sea aporte/prestamo/reversa, dejar categoria y nota para que no se lea como venta."
        )
      }
    }
    "^F-CXP" {
      [pscustomobject]@{
        title = "$Code - CxP por corregir"
        problem = "La cuenta por pagar requiere accion porque puede estar vencida, sin proyecto o sin clasificacion. Esto afecta caja, deuda real, costo por proyecto y OPEX."
        steps = @(
          "Entrar a Finanzas > Cuentas por pagar.",
          "Abrir la factura o el grupo indicado.",
          "Si ya esta pagada, vincular el movimiento bancario correspondiente.",
          "Si no esta pagada, decidir pago, renegociacion o excepcion documentada.",
          "Completar proyecto, linea de negocio y categoria contable. Usar OPEX solo cuando no sea costo de proyecto, inventario, herramientas, garantia, prestamo o remuneracion."
        )
      }
    }
    "^F-CXC" {
      [pscustomobject]@{
        title = "$Code - CxC por corregir"
        problem = "La cuenta por cobrar requiere accion porque puede estar vencida o sin proyecto. Esto afecta cobranza, pipeline y lectura comercial."
        steps = @(
          "Entrar a Finanzas > Cuentas por cobrar.",
          "Abrir la factura indicada.",
          "Si ya fue cobrada, vincular el abono bancario.",
          "Si sigue pendiente, coordinar cobranza con Comercial.",
          "Completar CRM/proyecto para que el ingreso quede conectado al negocio correcto."
        )
      }
    }
    "^F-OC01" {
      [pscustomobject]@{
        title = "F-OC01 - Presion de OCs sobre caja"
        problem = "Las OCs abiertas no son deuda bancaria ni CxP, pero si representan compromisos potenciales de caja. Sirven para mirar costo real versus comprometido."
        steps = @(
          "Revisar OCs abiertas en Materiales.",
          "Separar OCs emitidas, recibidas parcial, cerradas o anuladas.",
          "Contrastar contra flujo de caja de 13 semanas.",
          "Cuando una OC ya tenga factura, revisar que Finanzas la tome como CxP y no como compromiso duplicado.",
          "Usar esta alerta como contexto gerencial, no como deuda real."
        )
      }
    }
    default {
      [pscustomobject]@{
        title = "Agente Finanzas"
        problem = "Puedo explicar alertas financieras y pasos manuales en la app Finanzas."
        steps = @(
          "Responde indicando el codigo de alerta: F-D01, F-D02, F-B01, F-B02, F-CXP01, F-CXP03, F-CXP04, F-CXC01, F-CXC03 o F-OC01.",
          "Tambien puedes copiar una linea completa del informe.",
          "No modifico datos por correo; solo explico que hacer para resolver la inconsistencia."
        )
      }
    }
  }
}

function Resolve-FinanzasHelpCode {
  param([string]$Text)
  $m = [regex]::Match($Text, "(?i)\bF-(D0[12]|B0[12]|CXP0[1-4]|CXC0[1-3]|OC01)\b")
  if ($m.Success) { return $m.Value.ToUpperInvariant() }
  if ($Text -match "(?i)cartola|rcv|libro|fuente|desactualizad|sin datos") { return "F-D02" }
  if ($Text -match "(?i)cargo|salida|banco|concili") { return "F-B01" }
  if ($Text -match "(?i)abono|entrada|aporte|prestamo|reembolso|vale vista|reversa") { return "F-B02" }
  if ($Text -match "(?i)cxp|proveedor|pagar|vencida|sin proyecto|sin clasificar|opex") { return "F-CXP01" }
  if ($Text -match "(?i)cxc|cobrar|cobranza|cliente") { return "F-CXC01" }
  if ($Text -match "(?i)oc|orden de compra|comprometido") { return "F-OC01" }
  ""
}

function Render-FinanzasHelpReplyHtml {
  param([object]$Help, [string]$Code)
  $steps = (@($Help.steps) | ForEach-Object { "<li>$(HtmlEscape $_)</li>" }) -join ""
  $codeText = if ($Code) { "<p><strong>Alerta:</strong> $(HtmlEscape $Code)</p>" } else { "" }
  @"
<div style="font-family:Arial,sans-serif;color:#222;font-size:14px;line-height:1.45;">
  <p>Hola. Respondo sobre como resolver manualmente alertas de Finanzas.</p>
  $codeText
  <h3 style="margin-bottom:4px;">$(HtmlEscape $Help.title)</h3>
  <p><strong>Cual es el problema:</strong> $(HtmlEscape $Help.problem)</p>
  <p><strong>Como resolverlo en la app:</strong></p>
  <ol>$steps</ol>
  <p style="font-size:12px;color:#666;">Si necesitas otra alerta, responde con su codigo o copia la linea exacta del correo del agente.</p>
</div>
"@
}

function ConvertTo-SafeFileName {
  param([string]$Name)
  $clean = ([string]$Name).Trim()
  if ([string]::IsNullOrWhiteSpace($clean)) { $clean = "archivo" }
  $clean = $clean -replace '[\\/:*?"<>|#%&{}$!@+`=]', "_"
  $clean = $clean -replace "\s+", " "
  if ($clean.Length -gt 120) { $clean = $clean.Substring(0, 120) }
  $clean
}

function ConvertFrom-ClpText {
  param([string]$Text)
  $raw = ([string]$Text).Trim()
  if ([string]::IsNullOrWhiteSpace($raw)) { return 0.0 }
  $negative = $raw -match "^\s*-|\(\s*[\$0-9]"
  $clean = $raw -replace "[^\d,.\-]", ""
  $clean = $clean -replace "^-", ""
  if ([string]::IsNullOrWhiteSpace($clean)) { return 0.0 }
  $lastComma = $clean.LastIndexOf(",")
  $lastDot = $clean.LastIndexOf(".")
  if ($lastComma -ge 0 -and $lastDot -ge 0) {
    if ($lastComma -gt $lastDot) {
      $clean = ($clean -replace "\.", "")
      $clean = ($clean -replace ",\d{1,2}$", "")
    } else {
      $clean = ($clean -replace ",", "")
      $clean = ($clean -replace "\.\d{1,2}$", "")
    }
  } elseif ($lastComma -ge 0) {
    $clean = ($clean -replace ",\d{1,2}$", "")
    $clean = ($clean -replace ",", "")
  } elseif ($lastDot -ge 0) {
    $groups = @($clean -split "\.")
    if ($groups.Count -gt 1 -and $groups[-1].Length -eq 3) {
      $clean = $clean -replace "\.", ""
    } else {
      $clean = ($clean -replace "\.\d{1,2}$", "")
      $clean = $clean -replace "\.", ""
    }
  }
  if ([string]::IsNullOrWhiteSpace($clean)) { return 0.0 }
  try {
    $value = [double]$clean
    if ($negative) { $value = -1 * $value }
    return $value
  } catch {
    return 0.0
  }
}

function ConvertTo-BiceIsoDate {
  param([string]$Text)
  $value = [string]$Text
  $m = [regex]::Match($value, "(?<!\d)(?<y>20\d{2})[/-](?<m>\d{1,2})[/-](?<d>\d{1,2})(?!\d)")
  if ($m.Success) {
    return "{0:0000}-{1:00}-{2:00}" -f [int]$m.Groups["y"].Value, [int]$m.Groups["m"].Value, [int]$m.Groups["d"].Value
  }
  $m = [regex]::Match($value, "(?<!\d)(?<d>\d{1,2})[/-](?<m>\d{1,2})[/-](?<y>\d{2,4})(?!\d)")
  if ($m.Success) {
    $year = [int]$m.Groups["y"].Value
    if ($year -lt 100) { $year += 2000 }
    return "{0:0000}-{1:00}-{2:00}" -f $year, [int]$m.Groups["m"].Value, [int]$m.Groups["d"].Value
  }
  ""
}

function ConvertFrom-XlsxBytesToText {
  param([byte[]]$Bytes)
  try {
    Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
    $stream = [System.IO.MemoryStream]::new($Bytes)
    $zip = [System.IO.Compression.ZipArchive]::new($stream, [System.IO.Compression.ZipArchiveMode]::Read)
    $shared = @()
    $sharedEntry = $zip.GetEntry("xl/sharedStrings.xml")
    if ($sharedEntry) {
      $reader = [System.IO.StreamReader]::new($sharedEntry.Open())
      $sharedXml = $reader.ReadToEnd()
      $reader.Dispose()
      foreach ($si in [regex]::Matches($sharedXml, "<si\b[^>]*>(.*?)</si>", [System.Text.RegularExpressions.RegexOptions]::Singleline)) {
        $parts = @()
        foreach ($t in [regex]::Matches($si.Groups[1].Value, "<t\b[^>]*>(.*?)</t>", [System.Text.RegularExpressions.RegexOptions]::Singleline)) {
          $parts += [System.Net.WebUtility]::HtmlDecode(($t.Groups[1].Value -replace "<[^>]+>", ""))
        }
        $shared += ($parts -join "")
      }
    }
    $lines = @()
    foreach ($entry in @($zip.Entries | Where-Object { $_.FullName -match "^xl/worksheets/sheet\d+\.xml$" } | Sort-Object FullName)) {
      $reader = [System.IO.StreamReader]::new($entry.Open())
      $sheetXml = $reader.ReadToEnd()
      $reader.Dispose()
      foreach ($row in [regex]::Matches($sheetXml, "<row\b[^>]*>(.*?)</row>", [System.Text.RegularExpressions.RegexOptions]::Singleline)) {
        $cells = @()
        foreach ($cell in [regex]::Matches($row.Groups[1].Value, "<c\b([^>]*)>(.*?)</c>", [System.Text.RegularExpressions.RegexOptions]::Singleline)) {
          $attrs = $cell.Groups[1].Value
          $body = $cell.Groups[2].Value
          $value = ""
          $vm = [regex]::Match($body, "<v>(.*?)</v>", [System.Text.RegularExpressions.RegexOptions]::Singleline)
          if ($vm.Success) { $value = [System.Net.WebUtility]::HtmlDecode($vm.Groups[1].Value) }
          if ($attrs -match 't="s"' -and $value -match "^\d+$") {
            $idx = [int]$value
            if ($idx -ge 0 -and $idx -lt $shared.Count) { $value = $shared[$idx] }
          } elseif ($attrs -match 't="inlineStr"') {
            $tm = [regex]::Match($body, "<t\b[^>]*>(.*?)</t>", [System.Text.RegularExpressions.RegexOptions]::Singleline)
            if ($tm.Success) { $value = [System.Net.WebUtility]::HtmlDecode($tm.Groups[1].Value) }
          }
          $cells += $value
        }
        $line = ($cells -join "`t").Trim()
        if ($line) { $lines += $line }
      }
    }
    $zip.Dispose()
    $stream.Dispose()
    return ($lines -join "`n")
  } catch {
    return ""
  }
}

function ConvertFrom-BiceAttachmentBytesToText {
  param([byte[]]$Bytes, [string]$FileName, [string]$ContentType)
  $name = ([string]$FileName).ToLowerInvariant()
  if ($name.EndsWith(".xlsx") -or ([string]$ContentType).ToLowerInvariant().Contains("spreadsheetml")) {
    return ConvertFrom-XlsxBytesToText -Bytes $Bytes
  }
  if ($name.EndsWith(".xls") -or ([string]$ContentType).ToLowerInvariant().Contains("vnd.ms-excel")) {
    $legacyText = ConvertFrom-LegacyXlsBytesToText -Bytes $Bytes
    if (-not [string]::IsNullOrWhiteSpace($legacyText)) { return $legacyText }
  }
  $text = [System.Text.Encoding]::UTF8.GetString($Bytes)
  if ($text -match "\x00") {
    try { $text = [System.Text.Encoding]::Unicode.GetString($Bytes) } catch { }
  }
  if ($name.EndsWith(".html") -or $name.EndsWith(".htm") -or ([string]$ContentType).ToLowerInvariant().Contains("html")) {
    $text = $text -replace "(?i)</t[dh]>", "`t"
    $text = $text -replace "(?i)</tr>", "`n"
    $text = $text -replace "(?is)<script.*?</script>", " "
    $text = $text -replace "(?is)<style.*?</style>", " "
    $text = $text -replace "(?s)<[^>]+>", " "
    $text = [System.Net.WebUtility]::HtmlDecode($text)
  }
  $text
}

function ConvertFrom-LegacyXlsBytesToText {
  param([byte[]]$Bytes)
  $tmpDir = if (-not [string]::IsNullOrWhiteSpace($env:RUNNER_TEMP)) { $env:RUNNER_TEMP } elseif (-not [string]::IsNullOrWhiteSpace($env:TEMP)) { $env:TEMP } else { [System.IO.Path]::GetTempPath() }
  $inputPath = Join-Path $tmpDir ("bice-cartola-" + [guid]::NewGuid().ToString("N") + ".xls")
  $scriptPath = Join-Path $tmpDir ("bice-xls-parser-" + [guid]::NewGuid().ToString("N") + ".py")
  try {
    [System.IO.File]::WriteAllBytes($inputPath, $Bytes)
    $script = @'
import sys
from datetime import datetime

import xlrd

path = sys.argv[1]
book = xlrd.open_workbook(path)
for sheet in book.sheets():
    for r in range(sheet.nrows):
        values = []
        for c in range(sheet.ncols):
            cell = sheet.cell(r, c)
            value = cell.value
            if cell.ctype == xlrd.XL_CELL_DATE:
                try:
                    value = xlrd.xldate_as_datetime(value, book.datemode).strftime("%Y-%m-%d")
                except Exception:
                    value = str(value)
            elif isinstance(value, float):
                if value.is_integer():
                    value = str(int(value))
                else:
                    value = str(value)
            else:
                value = str(value)
            values.append(value.replace("\r", " ").replace("\n", " ").strip())
        line = "\t".join(values).strip()
        if line:
            print(line)
'@
    [System.IO.File]::WriteAllText($scriptPath, $script, [System.Text.Encoding]::UTF8)
    $output = & python $scriptPath $inputPath 2>$null
    if ($LASTEXITCODE -eq 0 -and $output) { return ($output -join "`n") }
    return ""
  } catch {
    return ""
  } finally {
    Remove-Item -LiteralPath $inputPath -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $scriptPath -Force -ErrorAction SilentlyContinue
  }
}

function ConvertFrom-BiceCartolaText {
  param([string]$Text, [string]$SourceFile)
  $rows = @()
  $lines = @(([string]$Text -split "`r?`n") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
  $lineNo = 0
  $header = @{}
  foreach ($line in $lines) {
    $lineNo++
    $raw = ([string]$line).Trim()
    $cols = @($raw -split "`t|;")
    if ($cols.Count -ge 3 -and $raw -match "(?i)fecha" -and $raw -match "(?i)cargo|abono|haber|debe") {
      $header = @{}
      for ($i = 0; $i -lt $cols.Count; $i++) {
        $label = ([string]$cols[$i]).ToLowerInvariant()
        if ($label -match "fecha" -and -not $header.ContainsKey("fecha")) { $header["fecha"] = $i }
        if ($label -match "descrip|detalle|glosa|movimiento" -and -not $header.ContainsKey("descripcion")) { $header["descripcion"] = $i }
        if ($label -match "cargo|debe" -and -not $header.ContainsKey("cargo")) { $header["cargo"] = $i }
        if ($label -match "abono|haber" -and -not $header.ContainsKey("abono")) { $header["abono"] = $i }
      }
      continue
    }
    if ($raw -match "(?i)fecha.*(descripcion|detalle)|saldo anterior|saldo disponible|total") { continue }
    $date = ConvertTo-BiceIsoDate -Text $raw
    if ([string]::IsNullOrWhiteSpace($date)) { continue }

    $withoutDate = $raw -replace "(?<!\d)(20\d{2})[/-](\d{1,2})[/-](\d{1,2})(?!\d)", " "
    $withoutDate = $withoutDate -replace "(?<!\d)(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})(?!\d)", " "
    $amounts = @()
    foreach ($match in [regex]::Matches($withoutDate, '-?\$?\s*\d{1,3}(?:[\.\s]\d{3})+(?:,\d{1,2})?|-?\$?\s*\d{4,}(?:,\d{1,2})?')) {
      $n = [Math]::Abs((ConvertFrom-ClpText -Text $match.Value))
      if ($n -ge 100 -and $n -lt 2000000000) { $amounts += $n }
    }
    $amounts = @($amounts | Select-Object -Unique)
    if ($amounts.Count -eq 0) { continue }

    $cargo = 0.0
    $abono = 0.0
    $direction = "DESCONOCIDO"
    if ($header.Count -gt 0 -and $cols.Count -gt 1) {
      if ($header.ContainsKey("cargo") -and $header["cargo"] -lt $cols.Count) { $cargo = [Math]::Abs((ConvertFrom-ClpText -Text $cols[$header["cargo"]])) }
      if ($header.ContainsKey("abono") -and $header["abono"] -lt $cols.Count) { $abono = [Math]::Abs((ConvertFrom-ClpText -Text $cols[$header["abono"]])) }
      if ($header.ContainsKey("descripcion") -and $header["descripcion"] -lt $cols.Count) { $withoutDate = [string]$cols[$header["descripcion"]] }
    }
    if ($cargo -eq 0 -and $abono -eq 0 -and $cols.Count -ge 4) {
      $tail = @()
      foreach ($c in @($cols | Select-Object -Last 4)) {
        $n = [Math]::Abs((ConvertFrom-ClpText -Text $c))
        if ($n -gt 0) { $tail += $n } else { $tail += 0.0 }
      }
      if ($tail.Count -ge 2) {
        $candidateCargo = [double]$tail[$tail.Count - 2]
        $candidateAbono = [double]$tail[$tail.Count - 1]
        if ($candidateCargo -gt 0 -or $candidateAbono -gt 0) {
          $cargo = $candidateCargo
          $abono = $candidateAbono
        }
      }
    }
    if ($cargo -eq 0 -and $abono -eq 0) {
      $amount = [double]$amounts[-1]
      if ($raw -match "(?i)\babono\b|deposito|deposito|transferencia recibida|haber|credito") {
        $abono = $amount
      } elseif ($raw -match "(?i)\bcargo\b|giro|pago|debito|comision|impuesto|transferencia a") {
        $cargo = $amount
      } else {
        $cargo = 0.0
        $abono = 0.0
      }
    }
    if ($cargo -gt 0) { $direction = "CARGO" }
    if ($abono -gt 0) { $direction = "ABONO" }
    $rows += [pscustomobject]@{
      fecha = $date
      descripcion = ($withoutDate -replace "\s+", " ").Trim()
      cargo = [Math]::Round($cargo, 0)
      abono = [Math]::Round($abono, 0)
      monto = if ($cargo -gt 0) { [Math]::Round($cargo, 0) } elseif ($abono -gt 0) { [Math]::Round($abono, 0) } else { [Math]::Round([double]$amounts[-1], 0) }
      direccion = $direction
      sourceFile = $SourceFile
      lineNumber = $lineNo
      raw = $raw
    }
  }
  @($rows)
}

function Get-SimpleHash {
  param([string]$Text)
  $sha = [System.Security.Cryptography.SHA1]::Create()
  try {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes([string]$Text)
    $hash = $sha.ComputeHash($bytes)
    return (($hash | ForEach-Object { $_.ToString("x2") }) -join "").Substring(0, 12)
  } finally {
    $sha.Dispose()
  }
}

function Normalize-BiceDocument {
  param([object]$Value)
  $raw = ([string]$Value).Trim()
  if ([string]::IsNullOrWhiteSpace($raw)) { return "NA" }
  if ($raw -match "^\d+$") { return ($raw -replace "^0+(?=\d)", "") }
  $raw
}

function Get-BiceMovementAmountKey {
  param([object]$Mov)
  $cargo = [Math]::Round((Get-Number $Mov.cargo), 0)
  $abono = [Math]::Round((Get-Number $Mov.abono), 0)
  if ($cargo -gt 0) { return "C$cargo" }
  "A$abono"
}

function Get-BiceImportSignature {
  param([object]$Mov)
  $desc = (Normalize-MayuText $Mov.descripcion).ToLowerInvariant()
  @(
    [string]$Mov.banco,
    [string]$Mov.cuenta,
    [string]$Mov.fecha,
    (Normalize-BiceDocument $Mov.documento),
    (Get-BiceMovementAmountKey $Mov),
    (Get-SimpleHash $desc)
  ) -join "|"
}

function Get-BiceMovementDocId {
  param([object]$Mov)
  $desc = (Normalize-MayuText $Mov.descripcion).ToLowerInvariant()
  $base = @(
    "bice_mail",
    [string]$Mov.banco,
    [string]$Mov.cuenta,
    [string]$Mov.fecha,
    (Normalize-BiceDocument $Mov.documento),
    (Get-BiceMovementAmountKey $Mov),
    (Get-SimpleHash $desc)
  ) -join "-"
  $base -replace "[^A-Za-z0-9-]", "_"
}

function ConvertTo-BiceBankMovement {
  param([object]$Row, [string]$Account, [long]$NowMs, [string]$SourceFile)
  [pscustomobject]@{
    banco = "BICE"
    cuenta = $Account
    fecha = [string]$Row.fecha
    documento = ""
    descripcion = [string]$Row.descripcion
    cargo = [Math]::Round((Get-Number $Row.cargo), 0)
    abono = [Math]::Round((Get-Number $Row.abono), 0)
    conciliado = $false
    sourceProvider = "BICE_EMAIL"
    apiSync = $true
    biceMailSourceFile = $SourceFile
    biceMailRaw = [string]$Row.raw
    updatedAt = $NowMs
    createdAt = $NowMs
    createdBy = "system-finanzas"
  }
}

function Get-GraphMailboxMessages {
  param([string]$Token, [string]$Mailbox, [int]$Top = 50)
  $uri = "https://graph.microsoft.com/v1.0/users/$Mailbox/mailFolders/inbox/messages?`$top=$Top&`$orderby=receivedDateTime desc&`$select=id,subject,from,bodyPreview,isRead,receivedDateTime,hasAttachments"
  @((Invoke-GraphGet -Token $Token -Uri $uri).value)
}

function Get-GraphMessageAttachments {
  param([string]$Token, [string]$Mailbox, [string]$MessageId)
  $uri = "https://graph.microsoft.com/v1.0/users/$Mailbox/messages/$MessageId/attachments?`$select=id,name,contentType,size,isInline"
  $items = @((Invoke-GraphGet -Token $Token -Uri $uri).value)
  $full = @()
  foreach ($att in @($items)) {
    $attId = [string]$att.id
    if ([string]::IsNullOrWhiteSpace($attId)) {
      $full += $att
      continue
    }
    try {
      $full += (Invoke-GraphGet -Token $Token -Uri "https://graph.microsoft.com/v1.0/users/$Mailbox/messages/$MessageId/attachments/$attId")
    } catch {
      $full += $att
    }
  }
  @($full)
}

function Test-BiceCartolaMessage {
  param([object]$Message)
  $text = "$($Message.subject)`n$($Message.from.emailAddress.address)`n$($Message.bodyPreview)"
  return ($text -match "(?i)bice|cartola|cuenta corriente|estado de cuenta|25-00114")
}

function Find-BiceExistingBankMatch {
  param([object]$Row, [object[]]$BankRows)
  $date = [string]$Row.fecha
  $cargo = [Math]::Round((Get-Number $Row.cargo), 0)
  $abono = [Math]::Round((Get-Number $Row.abono), 0)
  $amount = [Math]::Round((Get-Number $Row.monto), 0)
  foreach ($mov in @($BankRows | Where-Object { [string]$_.fecha -eq $date })) {
    $movCargo = [Math]::Round((Get-Number $mov.cargo), 0)
    $movAbono = [Math]::Round((Get-Number $mov.abono), 0)
    if ($cargo -gt 0 -and $movCargo -eq $cargo) { return $mov }
    if ($abono -gt 0 -and $movAbono -eq $abono) { return $mov }
    if ($cargo -eq 0 -and $abono -eq 0 -and $amount -gt 0 -and ($movCargo -eq $amount -or $movAbono -eq $amount)) { return $mov }
  }
  $null
}

function Render-BiceCartolaMailHtml {
  param([object]$Report)
  $messageRows = (@($Report.messages) | ForEach-Object {
    "<tr><td style='border:1px solid #ddd;padding:6px;'>$(HtmlEscape $_.receivedDateTime)</td><td style='border:1px solid #ddd;padding:6px;'>$(HtmlEscape $_.from)</td><td style='border:1px solid #ddd;padding:6px;'>$(HtmlEscape $_.subject)</td><td style='border:1px solid #ddd;padding:6px;text-align:right;'>$($_.attachments)</td></tr>"
  }) -join ""
  if (-not $messageRows) { $messageRows = "<tr><td colspan='4' style='border:1px solid #ddd;padding:8px;color:#666;'>No se encontraron correos BICE/cartola en la ventana revisada.</td></tr>" }
  $candidateRows = (@($Report.candidates | Select-Object -First 80) | ForEach-Object {
    "<tr><td style='border:1px solid #ddd;padding:6px;'>$(HtmlEscape $_.fecha)</td><td style='border:1px solid #ddd;padding:6px;'>$(HtmlEscape $_.direccion)</td><td style='border:1px solid #ddd;padding:6px;text-align:right;'>$(Format-Clp (Get-Number $_.monto))</td><td style='border:1px solid #ddd;padding:6px;'>$(HtmlEscape $_.estado)</td><td style='border:1px solid #ddd;padding:6px;'>$(HtmlEscape $_.descripcion)</td></tr>"
  }) -join ""
  if (-not $candidateRows) { $candidateRows = "<tr><td colspan='5' style='border:1px solid #ddd;padding:8px;color:#666;'>Sin filas legibles con fecha y monto.</td></tr>" }
  $metrics = @(
    New-MayuEmailMetric -Label "Correos" -Value $Report.summary.messagesFound -Tone "info"
    New-MayuEmailMetric -Label "Adjuntos" -Value $Report.summary.attachmentsSaved -Tone "info"
    New-MayuEmailMetric -Label "Filas leidas" -Value $Report.summary.rowsParsed -Tone "info"
    New-MayuEmailMetric -Label "Nuevos probables" -Value $Report.summary.newCandidates -Tone "amarillo"
    New-MayuEmailMetric -Label "Duplicados" -Value $Report.summary.duplicates -Tone "verde"
  ) -join ""
  $content = @"
<table role="presentation" cellpadding="0" cellspacing="0" style="border-collapse:collapse;margin:0 0 10px 0;"><tr>$metrics</tr></table>
<p style="margin:0 0 12px 0;color:#4b5563;">Modo productivo: este agente lee BICE por correo, guarda respaldo, deduplica y crea solo movimientos bancarios nuevos con direccion clara. No marca correos como leidos.</p>
<h3 style="font-size:16px;margin:18px 0 8px 0;">Correos revisados</h3>
<table style="border-collapse:collapse;width:100%;font-size:13px;"><tr><th style='border:1px solid #ddd;padding:6px;text-align:left;'>Recibido</th><th style='border:1px solid #ddd;padding:6px;text-align:left;'>De</th><th style='border:1px solid #ddd;padding:6px;text-align:left;'>Asunto</th><th style='border:1px solid #ddd;padding:6px;text-align:right;'>Adj.</th></tr>$messageRows</table>
<h3 style="font-size:16px;margin:18px 0 8px 0;">Candidatos de cartola</h3>
<table style="border-collapse:collapse;width:100%;font-size:13px;"><tr><th style='border:1px solid #ddd;padding:6px;text-align:left;'>Fecha</th><th style='border:1px solid #ddd;padding:6px;text-align:left;'>Tipo</th><th style='border:1px solid #ddd;padding:6px;text-align:right;'>Monto</th><th style='border:1px solid #ddd;padding:6px;text-align:left;'>Estado</th><th style='border:1px solid #ddd;padding:6px;text-align:left;'>Descripcion</th></tr>$candidateRows</table>
"@
  New-MayuEmailLayout -Title "BICE Cartola Mail" -Subtitle $Report.date -ContentHtml $content -Footer "Fuente productiva BICE por correo. Los duplicados y filas ambiguas quedan en reporte."
}

function Invoke-BiceCartolaMail {
  param([object]$Config, [string]$GraphToken, [string]$SiteId, [datetime]$Now)
  $mailbox = [string]$Config.mail.sender
  $folder = [string]$Config.sharepoint.bice_cartolas_folder
  $account = [string]$Config.agents.bice_cartola_mail.account
  if ([string]::IsNullOrWhiteSpace($account)) { $account = "25-00114-1" }
  if ([string]::IsNullOrWhiteSpace($folder)) { $folder = "$($Config.sharepoint.base_folder)/bice_cartolas" }
  $dateKey = $Now.ToString("yyyy-MM-dd")
  $runKey = $Now.ToString("yyyyMMdd-HHmmss")
  $nowMs = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
  Write-Output "BICE cartola mail: leyendo inbox de $mailbox en modo importacion productiva."
  Ensure-GraphFolder -Token $GraphToken -SiteId $SiteId -FolderPath $folder
  Ensure-GraphFolder -Token $GraphToken -SiteId $SiteId -FolderPath "$folder/raw/$dateKey"

  $messages = @(Get-GraphMailboxMessages -Token $GraphToken -Mailbox $mailbox -Top 50 | Where-Object { Test-BiceCartolaMessage -Message $_ } | Select-Object -First 10)
  $data = Get-FirestoreData -Config $Config
  $writeToken = Get-FirestoreWriteToken -Config $Config
  $bankRows = @($data.fin_mov_bancarios)
  $saved = @()
  $parsedRows = @()
  $messageSummary = @()

  foreach ($msg in @($messages)) {
    $attachments = @()
    if ($msg.hasAttachments) {
      $attachments = @(Get-GraphMessageAttachments -Token $GraphToken -Mailbox $mailbox -MessageId ([string]$msg.id) | Where-Object { -not $_.isInline })
    }
    $messageSummary += [pscustomobject]@{
      id = [string]$msg.id
      receivedDateTime = [string]$msg.receivedDateTime
      from = [string]$msg.from.emailAddress.address
      subject = [string]$msg.subject
      attachments = @($attachments).Count
      isRead = [bool]$msg.isRead
    }
    foreach ($att in @($attachments)) {
      $name = ConvertTo-SafeFileName -Name ([string]$att.name)
      $contentType = [string]$att.contentType
      if ([string]::IsNullOrWhiteSpace([string]$att.contentBytes)) {
        $saved += [pscustomobject]@{ name = $name; status = "SIN_CONTENT_BYTES"; size = [int64](Get-Number $att.size); contentType = $contentType; filePath = "" }
        continue
      }
      $bytes = [Convert]::FromBase64String([string]$att.contentBytes)
      $filePath = "$folder/raw/$dateKey/$runKey-$name"
      Write-BytesFileToGraph -Token $GraphToken -SiteId $SiteId -FilePath $filePath -Bytes $bytes -ContentType $contentType
      $text = ConvertFrom-BiceAttachmentBytesToText -Bytes $bytes -FileName $name -ContentType $contentType
      $rows = @(ConvertFrom-BiceCartolaText -Text $text -SourceFile $name)
      $parsedRows += $rows
      $saved += [pscustomobject]@{
        name = $name
        status = if ($rows.Count -gt 0) { "PARSE_OK" } else { "SIN_FILAS_LEGIBLES" }
        size = $bytes.Length
        contentType = $contentType
        filePath = $filePath
        rowsParsed = $rows.Count
      }
    }
  }

  $candidates = @()
  $seenInRun = New-Object System.Collections.Generic.HashSet[string]
  $movementsToCreate = @()
  foreach ($row in @($parsedRows)) {
    $match = Find-BiceExistingBankMatch -Row $row -BankRows $bankRows
    $estado = if ($match) { "DUPLICADO_PROBABLE" } else { "NUEVO_PROBABLE" }
    if ([string]$row.direccion -eq "DESCONOCIDO") { $estado = "REVISION_DIRECCION" }
    $mov = ConvertTo-BiceBankMovement -Row $row -Account $account -NowMs $nowMs -SourceFile ([string]$row.sourceFile)
    $signature = Get-BiceImportSignature -Mov $mov
    if ($estado -eq "NUEVO_PROBABLE") {
      if ($seenInRun.Contains($signature)) {
        $estado = "DUPLICADO_EN_ARCHIVOS_BICE"
      } else {
        [void]$seenInRun.Add($signature)
        $docId = Get-BiceMovementDocId -Mov $mov
        $movementsToCreate += [pscustomobject]@{ id = $docId; data = $mov }
      }
    }
    $candidates += [pscustomobject]@{
      fecha = [string]$row.fecha
      descripcion = [string]$row.descripcion
      cargo = [double]$row.cargo
      abono = [double]$row.abono
      monto = [double]$row.monto
      direccion = [string]$row.direccion
      estado = $estado
      firestoreMatchId = if ($match) { [string]$match.id } else { "" }
      firestoreCreateId = if ($estado -eq "NUEVO_PROBABLE") { [string](Get-BiceMovementDocId -Mov $mov) } else { "" }
      sourceFile = [string]$row.sourceFile
      raw = [string]$row.raw
    }
  }

  $created = 0
  foreach ($item in @($movementsToCreate)) {
    Set-FirestoreDocument -Config $Config -Token $writeToken -CollectionName ([string]$Config.collections.fin_mov_bancarios) -DocumentId ([string]$item.id) -Data $item.data
    $created++
  }
  $latestDate = @($parsedRows | ForEach-Object { [string]$_.fecha } | Where-Object { $_ } | Sort-Object | Select-Object -Last 1)
  $periodo = if ($latestDate.Count -gt 0 -and $latestDate[0] -match "^(\d{4})-(\d{2})-") { "$($Matches[1])-$($Matches[2])" } else { $Now.ToString("yyyy-MM") }
  $importId = Add-FirestoreDocument -Config $Config -Token $writeToken -CollectionName ([string]$Config.collections.fin_importaciones) -Data ([pscustomobject]@{
    tipo = "CARTOLA_BICE"
    periodo = $periodo
    archivo = "BICE_EMAIL_${dateKey}_${runKey}"
    filasTotales = @($parsedRows).Count
    filasCreadas = $created
    filasActualizadas = 0
    filasDuplicadas = @($candidates | Where-Object { $_.estado -in @("DUPLICADO_PROBABLE", "DUPLICADO_EN_ARCHIVOS_BICE") }).Count
    errores = @($candidates | Where-Object { $_.estado -eq "REVISION_DIRECCION" }).Count
    createdBy = "system-finanzas"
    createdAt = $nowMs
    origen = "BICE_EMAIL"
    cuenta = $account
  })

  $report = [pscustomobject]@{
    generatedAt = $Now.ToString("o")
    date = $dateKey
    mailbox = $mailbox
    mode = "import"
    importId = $importId
    summary = [pscustomobject]@{
      messagesFound = @($messages).Count
      attachmentsSaved = @($saved | Where-Object { $_.filePath }).Count
      attachmentsWithoutContent = @($saved | Where-Object { $_.status -eq "SIN_CONTENT_BYTES" }).Count
      rowsParsed = @($parsedRows).Count
      newCandidates = @($candidates | Where-Object { $_.estado -eq "NUEVO_PROBABLE" }).Count
      created = $created
      duplicates = @($candidates | Where-Object { $_.estado -eq "DUPLICADO_PROBABLE" }).Count
      duplicatedInRun = @($candidates | Where-Object { $_.estado -eq "DUPLICADO_EN_ARCHIVOS_BICE" }).Count
      reviewDirection = @($candidates | Where-Object { $_.estado -eq "REVISION_DIRECCION" }).Count
      firestoreBankRows = @($bankRows).Count
    }
    messages = $messageSummary
    attachments = $saved
    candidates = $candidates
    notes = @(
      "Escribe solo movimientos nuevos con fecha, monto y direccion clara.",
      "No marca correos como leidos.",
      "Deduplica contra movimientos existentes y dentro del mismo lote BICE."
    )
  }
  $html = Render-BiceCartolaMailHtml -Report $report
  Write-TextFileToGraph -Token $GraphToken -SiteId $SiteId -FilePath "$folder/$dateKey-$runKey.json" -Text ($report | ConvertTo-Json -Depth 80) -ContentType "application/json; charset=utf-8"
  Write-TextFileToGraph -Token $GraphToken -SiteId $SiteId -FilePath "$folder/$dateKey-$runKey.html" -Text $html -ContentType "text/html; charset=utf-8"
  foreach ($msgRow in @($messageSummary)) {
    Write-Output "BICE mensaje: recibido=$($msgRow.receivedDateTime) de=$($msgRow.from) asunto='$($msgRow.subject)' adjuntos=$($msgRow.attachments)."
  }
  foreach ($attRow in @($saved)) {
    Write-Output "BICE adjunto: nombre='$($attRow.name)' status=$($attRow.status) filas=$($attRow.rowsParsed) size=$($attRow.size) contentType='$($attRow.contentType)'."
  }
  Write-Output "BICE cartola mail: correos=$($report.summary.messagesFound) adjuntos=$($report.summary.attachmentsSaved) filas=$($report.summary.rowsParsed) nuevos=$($report.summary.newCandidates) creados=$created duplicados=$($report.summary.duplicates) duplicados_lote=$($report.summary.duplicatedInRun)."
  Write-Output "BICE cartola mail: importId=$importId outputs guardados en $folder. No se envio correo y no se marcaron correos como leidos."
}

function Invoke-Finanzas {
  param([object]$Config, [string]$GraphToken, [string]$SiteId, [datetime]$Now, [bool]$DoSendEmail)
  Write-Output "Finanzas: leyendo Firestore."
  $data = Get-FirestoreData -Config $Config
  $report = Build-FinanzasReport -Config $Config -Data $data -Now $Now
  Write-Output "Finanzas: resumen rojas=$($report.summary.rojas) amarillas=$($report.summary.amarillas) informativas=$($report.summary.informativas) total=$($report.summary.total)."
  foreach ($group in @($report.issues | Group-Object code | Sort-Object Count -Descending | Select-Object -First 20)) {
    Write-Output "Finanzas breakdown: $($group.Name)=$($group.Count)"
  }
  foreach ($issue in @($report.issues | Select-Object -First 25)) {
    Write-Output "Finanzas alerta: [$($issue.severity)] $($issue.code) $($issue.title) | $($issue.ref) | $($issue.action)"
  }
  $html = Render-FinanzasHtmlV2 -Report $report
  $dateKey = $report.date
  Ensure-GraphFolder -Token $GraphToken -SiteId $SiteId -FolderPath $Config.sharepoint.finanzas_folder
  Write-TextFileToGraph -Token $GraphToken -SiteId $SiteId -FilePath "$($Config.sharepoint.finanzas_folder)/$dateKey.json" -Text ($report | ConvertTo-Json -Depth 80) -ContentType "application/json; charset=utf-8"
  Write-TextFileToGraph -Token $GraphToken -SiteId $SiteId -FilePath "$($Config.sharepoint.finanzas_folder)/$dateKey.html" -Text $html -ContentType "text/html; charset=utf-8"
  if ($DoSendEmail) {
    $to = Get-UniqueEmails -Emails @($Config.mail.felix, $Config.mail.valentina)
    $subject = "[Finanzas] Alertas $dateKey - $($report.summary.rojas) ROJAS, $($report.summary.amarillas) AMARILLAS"
    Send-GraphMail -Token $GraphToken -Sender $Config.mail.sender -To $to -Cc @() -Subject $subject -HtmlBody $html
    Write-Output "Finanzas: correo enviado a $($to -join ', ')."
  } else {
    Write-Output "Finanzas: SendEmail=false, no se envia correo."
  }
  Write-Output "Finanzas: outputs guardados en SharePoint."
}

function Get-SafeFileName {
  param([string]$Name)
  $safe = ([string]$Name).Trim()
  if ([string]::IsNullOrWhiteSpace($safe)) { $safe = "archivo" }
  $safe = $safe -replace '[\\/:*?"<>|#%&{}$!''@+`=]', '_'
  if ($safe.Length -gt 120) { $safe = $safe.Substring(0, 120) }
  $safe
}

function Get-SafeDocId {
  param([string]$Value)
  $id = ([string]$Value) -replace '[^A-Za-z0-9_-]', '_'
  if ([string]::IsNullOrWhiteSpace($id)) { $id = "dte_" + ([guid]::NewGuid().ToString("N")) }
  if ($id.Length -gt 140) { $id = $id.Substring(0, 140) }
  $id
}

function Convert-AttachmentBytesToText {
  param([byte[]]$Bytes)
  if ($null -eq $Bytes -or $Bytes.Length -eq 0) { return "" }
  $headLen = [Math]::Min(300, $Bytes.Length)
  $head = [System.Text.Encoding]::ASCII.GetString($Bytes, 0, $headLen)
  if ($head -match 'encoding=["'']ISO-8859-1["'']' -or $head -match 'encoding=["'']Windows-1252["'']') {
    return [System.Text.Encoding]::GetEncoding("ISO-8859-1").GetString($Bytes)
  }
  [System.Text.Encoding]::UTF8.GetString($Bytes)
}

function Get-XmlSingleText {
  param([xml]$Xml, [string]$LocalName, [object]$BaseNode = $null)
  $node = if ($BaseNode) { $BaseNode.SelectSingleNode(".//*[local-name()='$LocalName']") } else { $Xml.SelectSingleNode("//*[local-name()='$LocalName']") }
  if ($node) { return ([string]$node.InnerText).Trim() }
  ""
}

function Parse-DteXmlAttachment {
  param([string]$XmlText)
  try {
    [xml]$xml = $XmlText
    $refs = @()
    foreach ($ref in @($xml.SelectNodes("//*[local-name()='Referencia']"))) {
      $refs += [pscustomobject]@{
        tipoDoc = Get-XmlSingleText -Xml $xml -LocalName "TpoDocRef" -BaseNode $ref
        folio = Get-XmlSingleText -Xml $xml -LocalName "FolioRef" -BaseNode $ref
        razon = Get-XmlSingleText -Xml $xml -LocalName "RazonRef" -BaseNode $ref
      }
    }

    $detalles = @()
    $detalleNodes = @($xml.SelectNodes("//*[local-name()='Detalle']")) | Select-Object -First 20
    foreach ($det in @($detalleNodes)) {
      $detalles += [pscustomobject]@{
        nombre = Get-XmlSingleText -Xml $xml -LocalName "NmbItem" -BaseNode $det
        cantidad = Get-XmlSingleText -Xml $xml -LocalName "QtyItem" -BaseNode $det
        monto = Get-XmlSingleText -Xml $xml -LocalName "MontoItem" -BaseNode $det
      }
    }

    [pscustomobject]@{
      ok = $true
      tipoDte = Get-XmlSingleText -Xml $xml -LocalName "TipoDTE"
      folio = Get-XmlSingleText -Xml $xml -LocalName "Folio"
      fechaEmision = Get-XmlSingleText -Xml $xml -LocalName "FchEmis"
      rutEmisor = Get-XmlSingleText -Xml $xml -LocalName "RUTEmisor"
      razonSocialEmisor = Get-XmlSingleText -Xml $xml -LocalName "RznSoc"
      rutReceptor = Get-XmlSingleText -Xml $xml -LocalName "RUTRecep"
      razonSocialReceptor = Get-XmlSingleText -Xml $xml -LocalName "RznSocRecep"
      montoNeto = Get-XmlSingleText -Xml $xml -LocalName "MntNeto"
      montoExento = Get-XmlSingleText -Xml $xml -LocalName "MntExe"
      iva = Get-XmlSingleText -Xml $xml -LocalName "IVA"
      montoTotal = Get-XmlSingleText -Xml $xml -LocalName "MntTotal"
      referencias = $refs
      detalles = $detalles
    }
  } catch {
    [pscustomobject]@{ ok = $false; error = $_.Exception.Message }
  }
}

function Get-OcFolioFromText {
  param([string]$Text)
  if ([string]::IsNullOrWhiteSpace($Text)) { return "" }
  $m = [regex]::Match($Text, 'MAYU-OC-\d{8}-[A-Z0-9]{3,10}', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
  if ($m.Success) { return $m.Value.ToUpperInvariant() }
  $m = [regex]::Match($Text, '(?i)\b(?:OC|ORDEN\s+DE\s+COMPRA)\s*[:#-]?\s*([0-9]{2,8})\b')
  if ($m.Success) { return $m.Groups[1].Value }
  ""
}

function Get-OcFolioFromDtes {
  param([object[]]$Dtes)
  foreach ($dte in @($Dtes)) {
    foreach ($ref in @($dte.referencias)) {
      $tipoDoc = [string]$ref.tipoDoc
      $folio = ([string]$ref.folio).Trim()
      $razon = [string]$ref.razon
      if ([string]::IsNullOrWhiteSpace($folio)) { continue }
      if ($tipoDoc -eq "801" -or $razon -match '(?i)\bOC\b|orden de compra') {
        return $folio
      }
    }
  }
  ""
}

function Get-DteObjectText {
  param([object]$Item, [string[]]$Names)
  foreach ($name in @($Names)) {
    $prop = $Item.PSObject.Properties[$name]
    if ($prop -and -not [string]::IsNullOrWhiteSpace([string]$prop.Value)) {
      return [string]$prop.Value
    }
  }
  ""
}

function Find-MayuOcByFolio {
  param([object[]]$OcRows, [string]$Folio)
  if ([string]::IsNullOrWhiteSpace($Folio)) { return $null }
  $target = $Folio.Trim().ToUpperInvariant()
  $targetLoose = $target -replace '[^A-Z0-9]', ''
  foreach ($oc in @($OcRows)) {
    $candidates = @(
      (Get-DteObjectText -Item $oc -Names @("folio", "ocFolio", "folioMAYU", "mayuOcFolio")),
      (Get-DteObjectText -Item $oc -Names @("pdfFileName", "ocFileName"))
    )
    foreach ($candidate in @($candidates | Where-Object { $_ })) {
      $candidateText = $candidate.Trim().ToUpperInvariant()
      $candidateLoose = $candidateText -replace '[^A-Z0-9]', ''
      if ($candidateText -eq $target -or $candidateLoose -eq $targetLoose) { return $oc }
      if ($targetLoose.Length -ge 6 -and $candidateLoose.Contains($targetLoose)) { return $oc }
    }
  }
  $null
}

function Get-FinanzasDteOcRows {
  param([object]$Config, [string]$Token)
  $rows = @()
  $collections = @()
  if ($Config.collections.mat_oc_headers) { $collections += [string]$Config.collections.mat_oc_headers }
  if ($Config.collections.mat_ordenes) { $collections += [string]$Config.collections.mat_ordenes }
  foreach ($collection in @($collections | Select-Object -Unique)) {
    try {
      $rows += @(Get-FirestoreCollection -Config $Config -Token $Token -CollectionName $collection)
    } catch {
      Write-Output "DTE inbox: no se pudo leer $collection para match de OC ($($_.Exception.Message))."
    }
  }
  @($rows)
}

function Test-InvoiceEmailCandidate {
  param([object]$Message, [object[]]$Attachments)
  $subject = [string]$Message.subject
  $preview = [string]$Message.bodyPreview
  $names = (@($Attachments | ForEach-Object { [string]$_.name }) -join " ")
  $text = "$subject $preview $names"
  if ($text -match '(?i)\bfactura\b|\bDTE\b|\.xml\b|\.pdf\b|orden de compra|OC MAYU|MAYU-OC') { return $true }
  $false
}

function Invoke-FinanzasDteInbox {
  param([object]$Config, [string]$GraphToken, [string]$SiteId, [datetime]$Now)
  if ($Config.agents.finanzas_dte_inbox -and $Config.agents.finanzas_dte_inbox.enabled -eq $false) {
    Write-Output "DTE inbox: deshabilitado por config."
    return
  }

  $mailbox = [string]$Config.mail.sender
  Disable-GraphMailboxAutoReplies -Token $GraphToken -Mailbox $mailbox
  Write-Output "DTE inbox: leyendo correos recientes."
  $rawMessages = @(Get-GraphMailboxMessages -Token $GraphToken -Mailbox $mailbox -Top 50 | Where-Object {
    ([string]$_.from.emailAddress.address) -ne $mailbox
  })
  if ($rawMessages.Count -eq 0) {
    Write-Output "DTE inbox: sin correos recientes."
    return
  }

  $apiKey = Get-ConfigApiKey -Config $Config
  $idToken = Get-FirebaseIdToken -ApiKey $apiKey
  $ocRows = @(Get-FinanzasDteOcRows -Config $Config -Token $idToken)
  $dteCollection = if ($Config.collections.fin_dte_inbox) { [string]$Config.collections.fin_dte_inbox } else { "fin_dte_inbox" }
  $existingDtes = @(Get-FirestoreCollection -Config $Config -Token $idToken -CollectionName $dteCollection)
  $processedIds = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($item in @($existingDtes)) {
    if ($item.graphMessageId) { [void]$processedIds.Add([string]$item.graphMessageId) }
    if ($item.id) { [void]$processedIds.Add([string]$item.id) }
  }
  $messages = @($rawMessages | Where-Object {
    $messageId = [string]$_.id
    $docId = Get-SafeDocId -Value $messageId
    $_.isRead -eq $false -or -not ($processedIds.Contains($messageId) -or $processedIds.Contains($docId))
  })
  if ($messages.Count -eq 0) {
    Write-Output "DTE inbox: sin correos pendientes."
    return
  }

  $baseFolder = if ($Config.sharepoint.finanzas_dte_inbox_folder) { [string]$Config.sharepoint.finanzas_dte_inbox_folder } else { "agentes_mayu/finanzas/dte_recibidos" }
  $replyToSender = -not ($Config.agents.finanzas_dte_inbox -and $Config.agents.finanzas_dte_inbox.reply_to_sender -eq $false)
  $processed = 0

  foreach ($msg in @($messages)) {
    $attachments = @()
    if ($msg.hasAttachments) {
      $attachments = @(Get-GraphMessageAttachments -Token $GraphToken -Mailbox $mailbox -MessageId ([string]$msg.id) | Where-Object { -not $_.isInline })
    }
    if (-not (Test-InvoiceEmailCandidate -Message $msg -Attachments $attachments)) { continue }

    $processed++
    $datePath = $Now.ToString("yyyy/MM/dd")
    $docId = Get-SafeDocId -Value ([string]$msg.id)
    $folder = "$baseFolder/$datePath/$docId"
    Ensure-GraphFolder -Token $GraphToken -SiteId $SiteId -FolderPath $folder

    $attachmentRows = @()
    $dteDocs = @()
    $searchText = "$($msg.subject)`n$($msg.bodyPreview)"
    $idx = 0
    foreach ($att in @($attachments)) {
      $idx++
      $safeName = "{0:00}-{1}" -f $idx, (Get-SafeFileName -Name ([string]$att.name))
      $contentType = if ($att.contentType) { [string]$att.contentType } else { "application/octet-stream" }
      $searchText += "`n$($att.name)"
      $row = [ordered]@{
        name = [string]$att.name
        contentType = $contentType
        size = [int64]$att.size
        sharepointPath = "$folder/$safeName"
        status = "SIN_BYTES"
      }

      if ($att.contentBytes) {
        $bytes = [Convert]::FromBase64String([string]$att.contentBytes)
        Write-BytesFileToGraph -Token $GraphToken -SiteId $SiteId -FilePath "$folder/$safeName" -Bytes $bytes -ContentType $contentType
        $row.status = "GUARDADO"
        if (([string]$att.name) -match '\.xml$' -or $contentType -match 'xml') {
          $xmlText = Convert-AttachmentBytesToText -Bytes $bytes
          $parsed = Parse-DteXmlAttachment -XmlText $xmlText
          $dteDocs += $parsed
          $searchText += "`n$xmlText"
        }
      }
      $attachmentRows += [pscustomobject]$row
    }

    $ocFolio = Get-OcFolioFromDtes -Dtes $dteDocs
    if ([string]::IsNullOrWhiteSpace($ocFolio)) {
      $ocFolio = Get-OcFolioFromText -Text $searchText
    }
    $oc = Find-MayuOcByFolio -OcRows $ocRows -Folio $ocFolio
    $estado = if ($attachmentRows.Count -eq 0) { "SIN_ADJUNTOS" } elseif ($ocFolio -and $oc) { "OC_MATCH" } elseif ($ocFolio) { "OC_NO_ENCONTRADA" } else { "SIN_OC" }
    $dtePrincipal = @($dteDocs | Where-Object { $_.ok } | Select-Object -First 1)

    $record = [pscustomobject][ordered]@{
      id = $docId
      sourceProvider = "MAIL_NOTIFICACIONES"
      source = $mailbox
      estado = $estado
      validacionRcvEstado = "PENDIENTE_RCV"
      receivedAt = [string]$msg.receivedDateTime
      processedAt = $Now.ToString("o")
      from = [string]$msg.from.emailAddress.address
      fromName = [string]$msg.from.emailAddress.name
      subject = [string]$msg.subject
      bodyPreview = [string]$msg.bodyPreview
      graphMessageId = [string]$msg.id
      sharepointFolder = $folder
      attachmentCount = @($attachmentRows).Count
      attachments = $attachmentRows
      ocFolio = $ocFolio
      ocHeaderId = if ($oc) { [string]$oc.id } else { "" }
      projectId = if ($oc) { Get-DteObjectText -Item $oc -Names @("projectId", "matProjectId", "proyectoId") } else { "" }
      projectName = if ($oc) { Get-DteObjectText -Item $oc -Names @("projectName", "matProjectName", "proyectoNombre") } else { "" }
      destinoFinanciero = if ($oc) { Get-DteObjectText -Item $oc -Names @("destinoFinanciero") } else { "" }
      categoriaContable = if ($oc) { Get-DteObjectText -Item $oc -Names @("categoriaContable") } else { "" }
      tipoCompra = if ($oc) { Get-DteObjectText -Item $oc -Names @("tipoCompra", "tipoOc") } else { "" }
      dte = if ($dtePrincipal.Count -gt 0) { $dtePrincipal[0] } else { [pscustomobject]@{} }
      dtes = $dteDocs
    }

    $json = $record | ConvertTo-Json -Depth 80
    Write-TextFileToGraph -Token $GraphToken -SiteId $SiteId -FilePath "$folder/metadata.json" -Text $json -ContentType "application/json; charset=utf-8"
    Set-FirestoreDocument -Config $Config -Token $idToken -CollectionName $dteCollection -DocumentId $docId -Data $record

    if ($replyToSender) {
      $comment = if ($estado -eq "OC_MATCH") {
        "Hola,`n`nRecibimos la factura/adjuntos y quedo registrada con referencia $ocFolio. La validacion tributaria queda pendiente de cruce con RCV/SII antes del pago.`n`nGracias."
      } elseif ($estado -eq "SIN_OC") {
        "Hola,`n`nRecibimos la factura/adjuntos, pero no encontramos una referencia de OC MAYU. Para procesarla sin demoras, por favor reenviar o confirmar el folio OC MAYU correspondiente.`n`nGracias."
      } else {
        "Hola,`n`nRecibimos la factura/adjuntos y quedo registrada para revision de Finanzas. Si corresponde, por favor confirmar el folio OC MAYU asociado.`n`nGracias."
      }
      try {
        Reply-GraphMail -Token $GraphToken -Mailbox $mailbox -MessageId ([string]$msg.id) -HtmlBody $comment
      } catch {
        $replySubject = if ([string]$msg.subject -match "^(?i)re:") { [string]$msg.subject } else { "RE: $($msg.subject)" }
        try {
          $commentHtml = [System.Net.WebUtility]::HtmlEncode($comment) -replace "\r?\n", "<br>"
          Send-GraphMail -Token $GraphToken -Sender $mailbox -To @([string]$msg.from.emailAddress.address) -Cc @() -Subject $replySubject -HtmlBody $commentHtml
          Write-Output "DTE inbox: respuesta enviada como correo nuevo porque no se pudo responder en hilo."
        } catch {
          Write-Output "DTE inbox: no se pudo enviar respuesta automatica ($($_.Exception.Message))."
        }
      }
    }
    try {
      Set-GraphMailRead -Token $GraphToken -Mailbox $mailbox -MessageId ([string]$msg.id)
    } catch {
      Write-Output "DTE inbox: no se pudo marcar el correo como leido ($($_.Exception.Message))."
    }
    Write-Output "DTE inbox: procesado $estado / $($msg.subject)."
  }

  if ($processed -eq 0) {
    Write-Output "DTE inbox: no habia correos de facturas/DTE."
  } else {
    Write-Output "DTE inbox: $processed correo(s) procesado(s)."
  }
}

function Resolve-FinanzasAdminChoiceSelection {
  param([object]$Case, [string]$Text)
  if ($null -eq $Case -or -not $Case.PSObject.Properties["choices"]) { return $null }
  $clean = ([string]$Text).Trim()
  if ([string]::IsNullOrWhiteSpace($clean)) { return $null }
  foreach ($choice in @($Case.choices)) {
    $key = [string]$choice.key
    $label = [string]$choice.label
    $value = [string]$choice.value
    if ([string]::IsNullOrWhiteSpace($key)) { continue }
    if ($clean -match ("(?i)^\s*(opci.n\s*)?" + [regex]::Escape($key) + "(\)|\.|\s|$)")) { return $choice }
    if (-not [string]::IsNullOrWhiteSpace($value) -and $clean -match ("(?i)\b" + [regex]::Escape($value) + "\b")) { return $choice }
    if (-not [string]::IsNullOrWhiteSpace($label) -and $clean -match [regex]::Escape($label)) { return $choice }
  }
  $null
}

function Resolve-FinanzasAdminReplyDecision {
  param([string]$Text, [object]$Case = $null)
  $choice = Resolve-FinanzasAdminChoiceSelection -Case $Case -Text $Text
  if ($null -ne $choice) {
    $outcome = if ($choice.outcome) { [string]$choice.outcome } else { "CRITERIO_REGISTRADO" }
    return [pscustomobject]@{
      status = $outcome
      label = "opcion $($choice.key): $($choice.label)"
      requiresDetail = $false
      choiceKey = [string]$choice.key
      choiceLabel = [string]$choice.label
      choiceValue = [string]$choice.value
      choiceEffect = [string]$choice.effect
    }
  }
  $clean = ([string]$Text).ToLowerInvariant()
  if ($clean -match "\b(aprobar|aprobado|apruebo|ok|confirmo|vale)\b") {
    return [pscustomobject]@{ status = "APROBADO"; label = "aprobado"; requiresDetail = $false }
  }
  if ($clean -match "\b(bloquear|bloqueado|bloquea|pausar|pausado)\b") {
    return [pscustomobject]@{ status = "BLOQUEADO"; label = "bloqueado"; requiresDetail = $false }
  }
  if ($clean -match "\b(rechazar|rechazado|rechazo|descartar|no corresponde)\b") {
    return [pscustomobject]@{ status = "RECHAZADO"; label = "rechazado"; requiresDetail = $false }
  }
  if ($clean -match "\b(corregir|corregido|correccion|correcciÃ³n|clasificar como|debe ser|es )\b") {
    return [pscustomobject]@{ status = "CORREGIDO"; label = "corregido"; requiresDetail = $true }
  }
  $null
}

function Convert-FinanzasAdminEmailBodyToText {
  param([object]$Message)
  $bodyContent = ""
  if ($Message -and $Message.PSObject.Properties["body"] -and $Message.body -and $Message.body.PSObject.Properties["content"]) {
    $bodyContent = [string]$Message.body.content
  }
  if (-not [string]::IsNullOrWhiteSpace($bodyContent)) {
    $text = $bodyContent -replace "(?i)<br\s*/?>", "`n"
    $text = $text -replace "(?i)</p>|</div>|</li>|</tr>|</h\d>", "`n"
    $text = $text -replace "(?s)<style.*?</style>", " "
    $text = $text -replace "(?s)<script.*?</script>", " "
    $text = $text -replace "<[^>]+>", " "
    return ([System.Net.WebUtility]::HtmlDecode($text) -replace "[ \t]+", " ").Trim()
  }
  ([string]$Message.bodyPreview).Trim()
}

function Get-FinanzasAdminLatestReplyText {
  param([object]$Message)
  $text = Convert-FinanzasAdminEmailBodyToText -Message $Message
  if ([string]::IsNullOrWhiteSpace($text)) { return "" }
  $cut = $text.Length
  $markers = @(
    "(?im)^\s*[-_]{5,}.*$",
    "(?im)^\s*De:\s",
    "(?im)^\s*From:\s",
    "(?im)^\s*Enviado:\s",
    "(?im)^\s*Sent:\s",
    "(?im)^\s*Para:\s",
    "(?im)^\s*To:\s",
    "(?im)^\s*On .+ wrote:\s*$",
    "(?im)^\s*El .+ escribi"
  )
  foreach ($marker in $markers) {
    $m = [regex]::Match($text, $marker)
    if ($m.Success -and $m.Index -lt $cut) { $cut = $m.Index }
  }
  $text.Substring(0, $cut).Trim()
}

function Get-FinanzasAdminReplyMatches {
  param([object[]]$Cases, [string]$Text)
  $matches = @()
  foreach ($case in @($Cases)) {
    $mailCode = [string]$case.mailCode
    $id = [string]$case.id
    if ((-not [string]::IsNullOrWhiteSpace($mailCode) -and $Text -match [regex]::Escape($mailCode)) -or
        (-not [string]::IsNullOrWhiteSpace($id) -and $Text -match [regex]::Escape($id))) {
      $matches += $case
    }
  }
  @($matches)
}

function Get-FinanzasAdminReplyActions {
  param([object[]]$Cases, [string]$Text)
  $caseByCode = @{}
  foreach ($case in @($Cases)) {
    $mailCode = ([string]$case.mailCode).ToUpperInvariant()
    if (-not [string]::IsNullOrWhiteSpace($mailCode)) { $caseByCode[$mailCode] = $case }
  }
  $actions = @()
  $seen = New-Object System.Collections.Generic.HashSet[string]
  foreach ($line in @(([string]$Text) -split "\r?\n")) {
    if ($line -match "(?i)\b(responde|ejemplo|formato|alternativas para responder|para responder por correo)\b") { continue }
    if ($line -match "(?i)^\s*(ADM-[A-F0-9]{8})\s*[:：\-]\s*(.+?)\s*$") {
      $code = $matches[1].ToUpperInvariant()
      $response = ([string]$matches[2]).Trim()
      if ($caseByCode.ContainsKey($code) -and -not [string]::IsNullOrWhiteSpace($response)) {
        $key = "$code|$response"
        if ($seen.Add($key)) {
          $actions += [pscustomobject]@{ case = $caseByCode[$code]; response = $response }
        }
      }
    }
  }
  if ($actions.Count -eq 0) {
    $short = [string]$Text
    if ($short.Length -gt 600) { $short = $short.Substring(0, 600) }
    foreach ($m in [regex]::Matches($short, "(?i)(ADM-[A-F0-9]{8})\s*[:：]\s*([^;\r\n]{1,140})")) {
      $prefixStart = [Math]::Max(0, $m.Index - 45)
      $prefix = $short.Substring($prefixStart, $m.Index - $prefixStart)
      if ($prefix -match "(?i)(responde|ejemplo|formato)") { continue }
      $code = $m.Groups[1].Value.ToUpperInvariant()
      $response = ([string]$m.Groups[2].Value).Trim()
      if ($caseByCode.ContainsKey($code) -and -not [string]::IsNullOrWhiteSpace($response)) {
        $key = "$code|$response"
        if ($seen.Add($key)) {
          $actions += [pscustomobject]@{ case = $caseByCode[$code]; response = $response }
        }
      }
    }
  }
  @($actions)
}

function Render-FinanzasAdminReplyHint {
  param([object]$Case = $null)
  $choicesHtml = ""
  if ($Case -and $Case.PSObject.Properties["choices"] -and @($Case.choices).Count -gt 0) {
    $choicesHtml = (@($Case.choices) | ForEach-Object {
      "<li><strong>$(HtmlEscape $_.key)</strong>: $(HtmlEscape $_.label)</li>"
    }) -join ""
  } else {
    $choicesHtml = "<li><strong>A</strong>: aprobar propuesta</li><li><strong>B</strong>: corregir criterio</li><li><strong>C</strong>: bloquear</li><li><strong>D</strong>: rechazar</li>"
  }
  $code = if ($Case -and $Case.mailCode) { [string]$Case.mailCode } else { "ADM-XXXXXX" }
@"
<div style="font-family:Arial,sans-serif;font-size:14px;color:#202124;line-height:1.45;">
  <p>Recibi la referencia al caso, pero no pude interpretar la decision.</p>
  <p>Responde con una linea asi: <strong>$(HtmlEscape $code): A</strong></p>
  <ul>$choicesHtml</ul>
  <p>Si ninguna alternativa calza, responde: <strong>$(HtmlEscape $code): corregir: [criterio]</strong>.</p>
  <p>No ejecutare cambios financieros solo con esta respuesta; primero registro el criterio.</p>
</div>
"@
}

function Try-ProcessFinanzasAdminReply {
  param(
    [object]$Config,
    [string]$GraphToken,
    [object]$Message,
    [object[]]$Cases,
    [datetime]$Now
  )
  $mailbox = [string]$Config.mail.sender
  $from = [string]$Message.from.emailAddress.address
  if ($from -ne [string]$Config.mail.valentina -and $from -ne [string]$Config.mail.felix) { return $false }
  $subject = [string]$Message.subject
  $latestText = Get-FinanzasAdminLatestReplyText -Message $Message
  $text = "$subject`n$latestText"
  if ($text -notmatch "(?i)ADM-[A-F0-9]{8}|bank-|ap-|ar-") { return $false }
  $actions = @(Get-FinanzasAdminReplyActions -Cases $Cases -Text $latestText)
  if ($actions.Count -eq 0) { return $false }

  $replySubject = if ($subject -match "^(?i)re:") { $subject } else { "RE: $subject" }
  $resolved = @()
  foreach ($action in @($actions)) {
    $decision = Resolve-FinanzasAdminReplyDecision -Text ([string]$action.response) -Case $action.case
    if ($null -eq $decision) {
      $hint = Render-FinanzasAdminReplyHint -Case $action.case
      Send-GraphMail -Token $GraphToken -Sender $mailbox -To @($from) -Cc @() -Subject $replySubject -HtmlBody $hint
      return $true
    }
    $resolved += [pscustomobject]@{ case = $action.case; response = [string]$action.response; decision = $decision }
  }
  if ($resolved.Count -eq 0) {
    $hint = Render-FinanzasAdminReplyHint
    Send-GraphMail -Token $GraphToken -Sender $mailbox -To @($from) -Cc @() -Subject $replySubject -HtmlBody $hint
    return $true
  }

  $writeToken = Get-FirestoreWriteToken -Config $Config
  $collection = if ($Config.collections.fin_admin_cases) { [string]$Config.collections.fin_admin_cases } else { "fin_admin_cases" }
  $updated = @()
  $decisionLabels = @()
  foreach ($item in @($resolved)) {
    $case = $item.case
    $decision = $item.decision
    $case | Add-Member -NotePropertyName status -NotePropertyValue ([string]$decision.status) -Force
    $case | Add-Member -NotePropertyName decision -NotePropertyValue ([string]$item.response) -Force
    $case | Add-Member -NotePropertyName selectedChoiceKey -NotePropertyValue ([string]$decision.choiceKey) -Force
    $case | Add-Member -NotePropertyName selectedChoiceLabel -NotePropertyValue ([string]$decision.choiceLabel) -Force
    $case | Add-Member -NotePropertyName selectedChoiceValue -NotePropertyValue ([string]$decision.choiceValue) -Force
    $case | Add-Member -NotePropertyName selectedChoiceEffect -NotePropertyValue ([string]$decision.choiceEffect) -Force
    $case | Add-Member -NotePropertyName decidedBy -NotePropertyValue $from -Force
    $case | Add-Member -NotePropertyName decidedAt -NotePropertyValue ($Now.ToString("o")) -Force
    $case | Add-Member -NotePropertyName decisionSource -NotePropertyValue "email" -Force
    $case | Add-Member -NotePropertyName replyMessageId -NotePropertyValue ([string]$Message.id) -Force
    Set-FirestoreDocument -Config $Config -Token $writeToken -CollectionName $collection -DocumentId ([string]$case.id) -Data $case
    $updated += "$($case.mailCode) / $($case.title)"
    $decisionLabels += [string]$decision.label
  }

  $items = ($updated | ForEach-Object { "<li>$(HtmlEscape $_)</li>" }) -join ""
  $labelText = (($decisionLabels | Select-Object -Unique) -join "; ")
  $html = @"
<div style="font-family:Arial,sans-serif;font-size:14px;color:#202124;line-height:1.45;">
  <p>Registrado: <strong>$([System.Net.WebUtility]::HtmlEncode($labelText))</strong>.</p>
  <ul>$items</ul>
  <p>Esto queda como aprendizaje/criterio del Administrador Finanzas. En este nivel no ejecuto cambios en caja, pagos, impuestos, cierres ni datos financieros sensibles por correo.</p>
</div>
"@
  $cc = Get-UniqueEmails -Emails @($Config.mail.felix) -Exclude @($from, $mailbox)
  Send-GraphMail -Token $GraphToken -Sender $mailbox -To @($from) -Cc $cc -Subject $replySubject -HtmlBody $html
  $true
}

function Invoke-FinanzasResponder {
  param([object]$Config, [string]$GraphToken, [string]$SiteId, [datetime]$Now)
  Invoke-FinanzasDteInbox -Config $Config -GraphToken $GraphToken -SiteId $SiteId -Now $Now
  $mailbox = [string]$Config.mail.sender
  Disable-GraphMailboxAutoReplies -Token $GraphToken -Mailbox $mailbox
  $stateFile = "$($Config.sharepoint.finanzas_folder)/responder_procesados.json"
  Ensure-GraphFolder -Token $GraphToken -SiteId $SiteId -FolderPath $Config.sharepoint.finanzas_folder
  $stateText = Read-TextFileFromGraph -Token $GraphToken -SiteId $SiteId -FilePath $stateFile
  $processedIds = @()
  if (-not [string]::IsNullOrWhiteSpace($stateText)) {
    try { $processedIds = @($stateText | ConvertFrom-Json) } catch { $processedIds = @() }
  }
  $processedSet = New-Object System.Collections.Generic.HashSet[string]
  foreach ($id in @($processedIds)) { if ($id) { [void]$processedSet.Add([string]$id) } }

  $uri = "https://graph.microsoft.com/v1.0/users/$mailbox/mailFolders/inbox/messages?`$top=25&`$orderby=receivedDateTime desc&`$select=id,subject,from,bodyPreview,body,isRead,receivedDateTime"
  $messages = @((Invoke-GraphGet -Token $GraphToken -Uri $uri).value)
  $adminCollection = if ($Config.collections.fin_admin_cases) { [string]$Config.collections.fin_admin_cases } else { "fin_admin_cases" }
  $adminReadToken = Get-FirestoreWriteToken -Config $Config
  $adminCases = @(Get-FirestoreCollection -Config $Config -Token $adminReadToken -CollectionName $adminCollection)
  $processed = 0
  foreach ($msg in @($messages | Where-Object { $_.isRead -eq $false })) {
    $messageId = [string]$msg.id
    if ($processedSet.Contains($messageId)) { continue }
    $from = [string]$msg.from.emailAddress.address
    if ($from -eq $mailbox) { continue }
    $subject = [string]$msg.subject
    $bodyText = "$subject`n$($msg.bodyPreview)`n$($msg.body.content)"
    if (Try-ProcessFinanzasAdminReply -Config $Config -GraphToken $GraphToken -Message $msg -Cases $adminCases -Now $Now) {
      Set-GraphMailRead -Token $GraphToken -Mailbox $mailbox -MessageId $messageId
      [void]$processedSet.Add($messageId)
      $processed++
      Write-Output "Finanzas responder: decision Administrador Finanzas registrada desde correo."
      continue
    }
    if ($bodyText -notmatch "(?i)Finanzas|F-(D0|B0|CXP|CXC|OC)|cartola|RCV|CxP|CxC|banco|concili") { continue }
    if ($bodyText -notmatch "(?i)como|c.mo|resolver|resuelvo|explica|explicame|problema|que significa|qu. significa|ayuda|hacer") { continue }

    $code = Resolve-FinanzasHelpCode -Text $bodyText
    $help = Get-FinanzasHelpDefinition -Code $code
    $reply = Render-HelpReplyCard -Intro "Hola. Respondo sobre como resolver manualmente alertas de Finanzas." -Help $help -Code $code
    $audience = Get-UniqueEmails -Emails @($Config.mail.felix, $Config.mail.valentina, $from) -Exclude @($mailbox)
    $replySubject = if ($subject -match "^(?i)re:") { $subject } else { "RE: $subject" }
    Send-GraphMail -Token $GraphToken -Sender $mailbox -To $audience -Cc @() -Subject $replySubject -HtmlBody $reply
    Set-GraphMailRead -Token $GraphToken -Mailbox $mailbox -MessageId $messageId
    [void]$processedSet.Add($messageId)
    $processed++
    Write-Output "Finanzas responder: respuesta enviada a audiencia oficial ($($audience -join ', ')) para $code."
  }
  $nextState = @($processedSet.GetEnumerator() | Select-Object -Last 500)
  Write-TextFileToGraph -Token $GraphToken -SiteId $SiteId -FilePath $stateFile -Text ($nextState | ConvertTo-Json -Depth 5) -ContentType "application/json; charset=utf-8"
  Write-Output "Finanzas responder: mensajes procesados=$processed."
}

function Get-CalidadIssues {
  param([object]$Config, [object]$Data, [object[]]$BibliaRows)
  $issues = [System.Collections.ArrayList]::new()
  foreach ($u in @($Data.fab_units)) {
    if ([string]$u.status -eq "M05" -and -not $u.rfFabricaOk) {
      Add-Issue $issues (New-Issue -Severity "amarillo" -Area "Calidad/RF" -Title "POD listo para RF" -Detail "$($u.unitCode) / $($u.matProjectName) esta en M05 sin RF Fabrica OK." -Owner "Felipe / Calidad" -Action "Coordinar RF." -Ref $u.id)
    }
  }
  foreach ($mp in @($Data.mat_projects)) {
    if (-not $mp.chkId) { continue }
    $bib = @($BibliaRows | Where-Object { $_.id -eq $mp.chkId } | Select-Object -First 1)
    if (-not $bib) { continue }
    $cal = @($bib.requirements | Where-Object { $_.key -eq "calidad" } | Select-Object -First 1)
    $desp = @($bib.requirements | Where-Object { $_.key -eq "plan_despachos" } | Select-Object -First 1)
    $trans = @($bib.requirements | Where-Object { $_.key -eq "protocolo_transporte" } | Select-Object -First 1)
    $unitsDespacho = @($Data.fab_units | Where-Object { $_.matProjectId -eq $mp.id -and ([string]$_.status) -in @("RF_FABRICA","QC","TERMINADO") })
    if ($unitsDespacho.Count -gt 0 -and (($cal.estado -ne "verde") -or ($desp.estado -ne "verde") -or ($trans.estado -ne "verde"))) {
      Add-Issue $issues (New-Issue -Severity "rojo" -Area "Calidad/Despacho" -Title "Riesgo de despacho sin documentacion completa" -Detail "$($mp.name) tiene $($unitsDespacho.Count) unidad(es) post RF/QC/PT y docs calidad/despacho/transporte pendientes." -Owner "Carlos / Gabriel / Calidad" -Action "Cerrar calidad, plan despacho y protocolo transporte antes de despacho." -Ref $mp.id)
    }
  }
  @($issues)
}

function Get-CommercialIssues {
  param([object]$Data)
  $issues = [System.Collections.ArrayList]::new()
  $closed = @($Data.crm_projects | Where-Object { $_.estado_comercial -eq "Negocio cerrado" })
  $matCrmIds = New-Object System.Collections.Generic.HashSet[string]
  foreach ($mp in @($Data.mat_projects)) { if ($mp.crmId) { [void]$matCrmIds.Add([string]$mp.crmId) } }
  foreach ($p in $closed) {
    if (-not $matCrmIds.Contains([string]$p.id)) {
      Add-Issue $issues (New-Issue -Severity "rojo" -Area "Comercial" -Title "Cerrado sin operacion" -Detail "$($p.nombre) esta cerrado en CRM y no aparece en Materiales." -Owner "Felix Escudero Vargas" -Action "Traspasar a Control/Materiales." -Ref $p.id)
    }
  }
  $plan = @($Data.crm_plan | Select-Object -First 1)
  if ($plan) {
    $meta = Get-Number $plan.meta_ventas
    $cerrado = 0.0
    foreach ($p in $closed) { $cerrado += Get-Number $p.ingreso_proyectado }
    if ($meta -gt 0) {
      $gap = $meta - $cerrado
      Add-Issue $issues (New-Issue -Severity "info" -Area "Comercial" -Title "Brecha contra meta anual" -Detail "Cerrado: $([Math]::Round($cerrado,0)) CLP / Meta: $([Math]::Round($meta,0)) CLP / Brecha: $([Math]::Round($gap,0)) CLP." -Owner "Felix Escudero Vargas" -Action "Revisar pipeline y cierres prioritarios." -Ref "plan/default")
    }
  }
  @($issues)
}

function Limit-Issues {
  param([object[]]$Issues, [int]$Max)
  @($Issues | Select-Object -First $Max)
}

function Get-DecisionGroupKey {
  param([object]$Issue)
  $ref = [string]$Issue.ref
  if ([string]::IsNullOrWhiteSpace($ref)) {
    return "$($Issue.area)|$($Issue.title)"
  }
  if ($ref -match '^pack-(PRY-[A-Z0-9]+)') { return $matches[1] }
  $ref
}

function Get-EntityDisplayName {
  param([object]$Entity, [string]$Fallback)
  if ($null -eq $Entity) { return $Fallback }
  foreach ($prop in @("name", "nombre", "projectName", "razonSocialContraparte", "title", "titulo", "client", "cliente")) {
    $property = $Entity.PSObject.Properties[$prop]
    if ($property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
      return [string]$property.Value
    }
  }
  $Fallback
}

function Get-InvoiceDisplayName {
  param([object]$Invoice, [string]$Fallback)
  $name = Get-EntityDisplayName -Entity $Invoice -Fallback ""
  $folio = ""
  if ($Invoice -and $Invoice.PSObject.Properties["folio"]) { $folio = [string]$Invoice.folio }
  if (-not [string]::IsNullOrWhiteSpace($name) -and -not [string]::IsNullOrWhiteSpace($folio)) {
    return "$name folio $folio"
  }
  if (-not [string]::IsNullOrWhiteSpace($name)) { return $name }
  if (-not [string]::IsNullOrWhiteSpace($folio)) { return "Folio $folio" }
  $Fallback
}

function Get-DecisionLabels {
  param([object]$Data)
  $labels = @{}
  if ($Data) {
    foreach ($mp in @($Data.mat_projects)) {
      if ($mp.id) { $labels[[string]$mp.id] = Get-EntityDisplayName -Entity $mp -Fallback ([string]$mp.id) }
    }
    foreach ($chk in @($Data.chk_projects)) {
      if ($chk.id) { $labels[[string]$chk.id] = Get-EntityDisplayName -Entity $chk -Fallback ([string]$chk.id) }
    }
    foreach ($crm in @($Data.crm_projects)) {
      if ($crm.id) { $labels[[string]$crm.id] = Get-EntityDisplayName -Entity $crm -Fallback ([string]$crm.id) }
    }
    foreach ($ap in @($Data.fin_facturas_ap)) {
      if ($ap.id) { $labels[[string]$ap.id] = Get-InvoiceDisplayName -Invoice $ap -Fallback ([string]$ap.id) }
    }
    foreach ($ar in @($Data.fin_facturas_ar)) {
      if ($ar.id) { $labels[[string]$ar.id] = Get-InvoiceDisplayName -Invoice $ar -Fallback ([string]$ar.id) }
    }
  }
  $labels["inv_catalogo"] = "Inventario"
  $labels["stock_packs"] = "Stock para packs proximos"
  $labels["fin_facturas_ap"] = "Facturas por pagar"
  $labels["fin_facturas_ar"] = "Facturas por cobrar"
  $labels
}

function Select-DecisionItems {
  param([object[]]$Issues, [object]$Data = $null, [int]$Max = 8)
  $labels = Get-DecisionLabels -Data $Data
  $groups = [ordered]@{}
  foreach ($issue in @($Issues)) {
    if ($null -eq $issue) { continue }
    $key = Get-DecisionGroupKey -Issue $issue
    if (-not $groups.Contains($key)) { $groups[$key] = @() }
    $groups[$key] += $issue
  }
  $items = @()
  foreach ($key in $groups.Keys) {
    $list = @($groups[$key])
    $first = $list[0]
    $titles = @($list | ForEach-Object { $_.title } | Select-Object -Unique)
    $label = if ($labels.ContainsKey($key)) { $labels[$key] } else { $key }
    $text = if ($list.Count -gt 1) {
      "${label}: $($list.Count) temas abiertos ($($titles -join '; ')). $($first.action)"
    } else {
      "${label}: $($first.title). $($first.action)"
    }
    $items += [pscustomobject]@{ text = $text; owner = $first.owner; ref = $first.ref }
    if ($items.Count -ge $Max) { break }
  }
  $items
}

function Build-Pulse {
  param([object]$Config, [object]$Data, [datetime]$Now)
  $bibliaRows = @(Get-BibliaProjectRows -Config $Config -ChkProjects @($Data.chk_projects))
  $traspaso = @(Get-OperationalIssues -Config $Config -Data $Data -BibliaRows $bibliaRows)
  $packIssues = @(Get-PackIssues -Config $Config -Data $Data -Now $Now)
  $abastIssues = @(Get-AbastecimientoIssues -Config $Config -Data $Data)
  $finIssues = @(Get-FinanceIssues -Config $Config -Data $Data -Now $Now)
  $comIssues = @(Get-CommercialIssues -Data $Data)
  $calIssues = @(Get-CalidadIssues -Config $Config -Data $Data -BibliaRows $bibliaRows)

  $max = [int]$Config.thresholds.max_items_per_section
  $all = @(Select-UniqueIssues -Issues @($traspaso + $packIssues + $abastIssues + $finIssues + $comIssues + $calIssues))
  $rojos = @($all | Where-Object { $_.severity -eq "rojo" })
  $amarillos = @($all | Where-Object { $_.severity -eq "amarillo" })
  $infos = @($all | Where-Object { $_.severity -eq "info" })
  $decisiones = @(Select-DecisionItems -Issues $rojos -Data $Data -Max 8)
  [pscustomobject]@{
    generatedAt = $Now.ToString("o")
    date = $Now.ToString("yyyy-MM-dd")
    summary = [pscustomobject]@{
      rojos = @($rojos).Count
      amarillos = @($amarillos).Count
      infos = @($infos).Count
      proyectosBiblia = @($bibliaRows).Count
      bibliaRoja = @($bibliaRows | Where-Object { $_.estado -eq "rojo" }).Count
      bibliaAmarilla = @($bibliaRows | Where-Object { $_.estado -eq "amarillo" }).Count
    }
    decisionesFelixHoy = @($decisiones)
    sections = [pscustomobject]@{
      bibliaProyecto = Limit-Issues -Issues @($bibliaRows | Where-Object { $_.estado -ne "verde" } | ForEach-Object {
        New-Issue -Severity $_.estado -Area "Biblia del Proyecto" -Title "$($_.name) - Documentacion $($_.estado)" -Detail "$($_.rojos) requisito(s) criticos / $($_.amarillos) requisito(s) pendientes." -Owner "Martin / Carlos / Valentina" -Action "Cerrar requisitos pendientes en Control Documental." -Ref $_.id
      }) -Max $max
      traspasoControlOperacion = Limit-Issues -Issues $traspaso -Max $max
      packs = Limit-Issues -Issues $packIssues -Max $max
      abastecimiento = Limit-Issues -Issues $abastIssues -Max $max
      finanzas = Limit-Issues -Issues $finIssues -Max $max
      comercial = Limit-Issues -Issues $comIssues -Max $max
      calidadRfDespacho = Limit-Issues -Issues $calIssues -Max $max
      compromisosReuniones = @()
    }
    raw = [pscustomobject]@{
      biblia = $bibliaRows
      allIssues = $all
    }
  }
}

function Render-IssueList {
  param([object[]]$Items)
  if (@($Items).Count -eq 0) { return "<p style='color:#166534;margin:6px 0;'>Sin alertas relevantes.</p>" }
  $html = "<table style='border-collapse:collapse;width:100%;font-size:13px;'><tr><th style='text-align:left;border:1px solid #ddd;padding:6px;'>Sev</th><th style='text-align:left;border:1px solid #ddd;padding:6px;'>Tema</th><th style='text-align:left;border:1px solid #ddd;padding:6px;'>Accion</th><th style='text-align:left;border:1px solid #ddd;padding:6px;'>Resp.</th></tr>"
  foreach ($it in @($Items)) {
    $color = if ($it.severity -eq "rojo") { "#991b1b" } elseif ($it.severity -eq "amarillo") { "#9a3412" } else { "#2563eb" }
    $codeText = ""
    if ($it.PSObject.Properties["code"] -and -not [string]::IsNullOrWhiteSpace([string]$it.code)) {
      $codeText = "$($it.code) - "
    } elseif ($it.PSObject.Properties["checkId"] -and -not [string]::IsNullOrWhiteSpace([string]$it.checkId)) {
      $codeText = "$($it.checkId) - "
    }
    $html += "<tr>"
    $html += "<td style='border:1px solid #ddd;padding:6px;color:$color;font-weight:bold;'>$(HtmlEscape $it.severity)</td>"
    $html += "<td style='border:1px solid #ddd;padding:6px;'><strong>$(HtmlEscape ($codeText + $it.title))</strong><br><span style='color:#555;'>$(HtmlEscape $it.detail)</span><br><span style='font-size:12px;color:#777;'>$(HtmlEscape $it.ref)</span></td>"
    $html += "<td style='border:1px solid #ddd;padding:6px;'>$(HtmlEscape $it.action)</td>"
    $html += "<td style='border:1px solid #ddd;padding:6px;'>$(HtmlEscape $it.owner)</td>"
    $html += "</tr>"
  }
  $html += "</table>"
  $html
}

function Render-IssueListCards {
  param([object[]]$Items)
  if (@($Items).Count -eq 0) { return (New-MayuEmptyState -Text "Sin alertas relevantes.") }
  $html = ""
  foreach ($it in @($Items)) {
    $rawTone = if ($it.PSObject.Properties["level"]) { [string]$it.level } else { [string]$it.severity }
    $tone = Get-MayuEmailTone -Tone $rawTone
    $codeText = ""
    if ($it.PSObject.Properties["code"] -and -not [string]::IsNullOrWhiteSpace([string]$it.code)) {
      $codeText = "$($it.code) - "
    } elseif ($it.PSObject.Properties["checkId"] -and -not [string]::IsNullOrWhiteSpace([string]$it.checkId)) {
      $codeText = "$($it.checkId) - "
    }
    $html += @"
<div style="border:1px solid #e5e7eb;border-left:4px solid $($tone.accent);background:#ffffff;padding:12px 14px;margin:0 0 10px 0;">
  <div style="margin-bottom:6px;">
    <span style="display:inline-block;background:$($tone.soft);border:1px solid $($tone.accent);color:#202124;font-size:11px;line-height:1;text-transform:uppercase;letter-spacing:.3px;padding:5px 7px;">$(HtmlEscape $rawTone)</span>
    <span style="color:#6b7280;font-size:12px;margin-left:6px;">$(HtmlEscape $it.area)</span>
  </div>
  <div style="font-weight:700;color:#202124;margin-bottom:4px;">$(HtmlEscape ($codeText + $it.title))</div>
  <div style="color:#4b5563;margin-bottom:8px;">$(HtmlEscape $it.detail)</div>
  <div style="background:#f8fafc;border:1px solid #edf2f7;padding:8px 10px;color:#30343b;"><strong>Accion:</strong> $(HtmlEscape $it.action)</div>
  <div style="font-size:12px;color:#6b7280;margin-top:7px;"><strong>Responsable:</strong> $(HtmlEscape $it.owner) &nbsp; <strong>Ref:</strong> $(HtmlEscape $it.ref)</div>
</div>
"@
  }
  $html
}

function Render-BodegaMaterialesHtmlV2 {
  param([object]$Report)
  $s = $Report.summary
  $metrics = @(
    New-MayuEmailMetric -Label "Criticas" -Value $s.criticas -Tone "critico"
    New-MayuEmailMetric -Label "Altas" -Value $s.altas -Tone "alto"
    New-MayuEmailMetric -Label "Medias" -Value $s.medias -Tone "medio"
    New-MayuEmailMetric -Label "Bajas" -Value $s.bajas -Tone "bajo"
  ) -join ""
  $content = @"
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;margin:0 0 12px 0;"><tr>$metrics</tr></table>
<div style="background:#f4f7fb;border-left:4px solid #0078d4;padding:12px 14px;color:#30343b;margin:12px 0 18px 0;">
  Puedes responder este correo con preguntas como <strong>"como resuelvo B-D06"</strong>, <strong>"explicame B-A05"</strong> o copiando una alerta. El agente respondera solo con explicacion operativa y pasos manuales en las apps.
</div>
$(New-MayuEmailSection -Title "Alertas" -Html (Render-IssueListCards -Items @($Report.issues | Select-Object -First 80)))
"@
  New-MayuEmailLayout -Title "Agente Bodega + Materiales - $($Report.date)" -Subtitle "Trazabilidad, catalogo, compras, recepciones y entregas a fabrica." -ContentHtml $content -Footer "Destinatarios productivos esperados: Felix, Valentina, Carlos y Mauricio. El envio queda controlado por SendEmail."
}

function Render-FinanzasHtmlV2 {
  param([object]$Report)
  $s = $Report.summary
  $metrics = @(
    New-MayuEmailMetric -Label "Rojas" -Value $s.rojas -Tone "rojo"
    New-MayuEmailMetric -Label "Amarillas" -Value $s.amarillas -Tone "amarillo"
    New-MayuEmailMetric -Label "Informativas" -Value $s.informativas -Tone "info"
  ) -join ""
  $actionable = Render-IssueListCards -Items @($Report.issues | Where-Object { $_.severity -in @("rojo", "amarillo") })
  $context = Render-IssueListCards -Items @($Report.issues | Where-Object { $_.severity -eq "info" })
  $content = @"
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;margin:0 0 12px 0;"><tr>$metrics</tr></table>
$(New-MayuEmailSection -Title "Alertas accionables" -Html $actionable)
$(New-MayuEmailSection -Title "Contexto gerencial" -Html $context)
<div style="background:#f4f7fb;border-left:4px solid #0078d4;padding:12px 14px;color:#30343b;margin-top:22px;">
  Responde este correo con el codigo de la alerta, por ejemplo <strong>F-B01</strong>, o copia la linea del problema. El respondedor de Finanzas te dira que hacer en la app para resolverlo.
</div>
"@
  New-MayuEmailLayout -Title "Agente Finanzas MAYU - $($Report.date)" -Subtitle "Caja, CxP, CxC, clasificacion, proyectos y fuentes actualizadas." -ContentHtml $content -Footer "Este agente no corrige datos por correo. Solo reporta inconsistencias que afectan la lectura gerencial."
}

function Render-HelpReplyCard {
  param([string]$Intro, [object]$Help, [string]$Code)
  $steps = (@($Help.steps) | ForEach-Object { "<li>$(HtmlEscape $_)</li>" }) -join ""
  $codeText = if ($Code) { "<p><strong>Alerta:</strong> $(HtmlEscape $Code)</p>" } else { "" }
  @"
<div style="font-family:Arial,sans-serif;color:#202124;font-size:14px;line-height:1.45;border:1px solid #e5e7eb;border-left:4px solid #0078d4;background:#ffffff;padding:14px 16px;">
  <p style="margin-top:0;">$(HtmlEscape $Intro)</p>
  $codeText
  <h3 style="margin:0 0 8px 0;color:#202124;">$(HtmlEscape $Help.title)</h3>
  <p><strong>Cual es el problema:</strong> $(HtmlEscape $Help.problem)</p>
  <p><strong>Como resolverlo en la app:</strong></p>
  <ol>$steps</ol>
  <p style="font-size:12px;color:#666;margin-bottom:0;">Si necesitas otra alerta, responde con su codigo o copia la linea exacta del correo del agente.</p>
</div>
"@
}

function Render-PulseHtmlV2 {
  param([object]$Config, [object]$Pulse)
  $s = $Pulse.summary
  $links = $Config.sharepoint.links
  $decisionHtml = if (@($Pulse.decisionesFelixHoy).Count -eq 0) {
    New-MayuEmptyState -Text "No hay decisiones rojas detectadas con la data disponible."
  } else {
    $items = (@($Pulse.decisionesFelixHoy) | ForEach-Object {
      "<li style='margin:0 0 8px 0;'><strong>$(HtmlEscape $_.owner):</strong> $(HtmlEscape $_.text)</li>"
    }) -join "`n"
    "<ol style='margin:0;padding-left:20px;'>$items</ol>"
  }
  $metrics = @(
    New-MayuEmailMetric -Label "Rojos" -Value $s.rojos -Tone "rojo"
    New-MayuEmailMetric -Label "Amarillos" -Value $s.amarillos -Tone "amarillo"
    New-MayuEmailMetric -Label "Biblia roja" -Value $s.bibliaRoja -Tone "rojo"
    New-MayuEmailMetric -Label "Biblia amarilla" -Value $s.bibliaAmarilla -Tone "amarillo"
  ) -join ""
  $apps = "Apps: <a href='$($links.control)' style='color:#0b57d0;'>Control</a> &middot; <a href='$($links.materiales)' style='color:#0b57d0;'>Materiales</a> &middot; <a href='$($links.bodega)' style='color:#0b57d0;'>Bodega</a> &middot; <a href='$($links.fabricacion)' style='color:#0b57d0;'>Fabricacion</a> &middot; <a href='$($links.finanzas)' style='color:#0b57d0;'>Finanzas</a> &middot; <a href='$($links.crm)' style='color:#0b57d0;'>CRM</a>"
  $content = @"
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;margin:0 0 12px 0;"><tr>$metrics</tr></table>
$(New-MayuEmailSection -Title "Decisiones Felix hoy" -Html $decisionHtml)
$(New-MayuEmailSection -Title "Riesgos rojos" -Html (Render-IssueListCards -Items @($Pulse.raw.allIssues | Where-Object { $_.severity -eq "rojo" } | Select-Object -First 12)))
$(New-MayuEmailSection -Title "Amarillos" -Html (Render-IssueListCards -Items @($Pulse.raw.allIssues | Where-Object { $_.severity -eq "amarillo" } | Select-Object -First 12)))
$(New-MayuEmailSection -Title "Biblia del Proyecto" -Html (Render-IssueListCards -Items @($Pulse.sections.bibliaProyecto)))
$(New-MayuEmailSection -Title "Traspaso Control -> Operacion" -Html (Render-IssueListCards -Items @($Pulse.sections.traspasoControlOperacion)))
$(New-MayuEmailSection -Title "Packs" -Html (Render-IssueListCards -Items @($Pulse.sections.packs)))
$(New-MayuEmailSection -Title "Abastecimiento / Compras" -Html (Render-IssueListCards -Items @($Pulse.sections.abastecimiento)))
$(New-MayuEmailSection -Title "Finanzas" -Html (Render-IssueListCards -Items @($Pulse.sections.finanzas)))
$(New-MayuEmailSection -Title "Comercial" -Html (Render-IssueListCards -Items @($Pulse.sections.comercial)))
$(New-MayuEmailSection -Title "Calidad / RF / Despacho" -Html (Render-IssueListCards -Items @($Pulse.sections.calidadRfDespacho)))
<p style="margin-top:24px;font-size:12px;color:#777;">$apps</p>
"@
  New-MayuEmailLayout -Title "Pulso gerencial MAYU - $($Pulse.date)" -Subtitle "Resumen ejecutivo diario desde ERP MAYU." -ContentHtml $content -Footer "Generado por MAYU Agents. Secciones parciales no inventan datos; solo reportan lo que existe en Firestore."
}

function Render-PulseHtml {
  param([object]$Config, [object]$Pulse)
  return Render-PulseHtmlV2 -Config $Config -Pulse $Pulse
  $s = $Pulse.summary
  $links = $Config.sharepoint.links
  $decisionHtml = if (@($Pulse.decisionesFelixHoy).Count -eq 0) {
    "<li>No hay decisiones rojas detectadas con la data disponible.</li>"
  } else {
    (@($Pulse.decisionesFelixHoy) | ForEach-Object { "<li><strong>$(HtmlEscape $_.owner):</strong> $(HtmlEscape $_.text)</li>" }) -join "`n"
  }
  @"
<html>
<body style="font-family:Arial,sans-serif;color:#222;max-width:980px;font-size:14px;">
  <h2 style="margin-bottom:4px;">Pulso gerencial MAYU - $($Pulse.date)</h2>
  <p style="margin-top:0;color:#555;">Rojos: <strong style="color:#991b1b;">$($s.rojos)</strong> · Amarillos: <strong style="color:#9a3412;">$($s.amarillos)</strong> · Biblia roja/amarilla: <strong>$($s.bibliaRoja)/$($s.bibliaAmarilla)</strong></p>

  <h3>Decisiones Felix hoy</h3>
  <ol>
    $decisionHtml
  </ol>

  <h3>Riesgos rojos</h3>
  $(Render-IssueList -Items @($Pulse.raw.allIssues | Where-Object { $_.severity -eq "rojo" } | Select-Object -First 12))

  <h3>Amarillos</h3>
  $(Render-IssueList -Items @($Pulse.raw.allIssues | Where-Object { $_.severity -eq "amarillo" } | Select-Object -First 12))

  <h3>Biblia del Proyecto</h3>
  $(Render-IssueList -Items @($Pulse.sections.bibliaProyecto))

  <h3>Traspaso Control -> Operacion</h3>
  $(Render-IssueList -Items @($Pulse.sections.traspasoControlOperacion))

  <h3>Packs</h3>
  $(Render-IssueList -Items @($Pulse.sections.packs))

  <h3>Abastecimiento / Compras</h3>
  $(Render-IssueList -Items @($Pulse.sections.abastecimiento))

  <h3>Finanzas</h3>
  $(Render-IssueList -Items @($Pulse.sections.finanzas))

  <h3>Comercial</h3>
  $(Render-IssueList -Items @($Pulse.sections.comercial))

  <h3>Calidad / RF / Despacho</h3>
  $(Render-IssueList -Items @($Pulse.sections.calidadRfDespacho))

  <p style="margin-top:24px;font-size:12px;color:#777;">
    Apps: <a href="$($links.control)">Control</a> · <a href="$($links.materiales)">Materiales</a> · <a href="$($links.bodega)">Bodega</a> · <a href="$($links.fabricacion)">Fabricacion</a> · <a href="$($links.finanzas)">Finanzas</a> · <a href="$($links.crm)">CRM</a>
  </p>
  <p style="font-size:12px;color:#777;">Generado por MAYU Agents. Secciones parciales no inventan datos; solo reportan lo que existe en Firestore.</p>
</body>
</html>
"@
}

function Invoke-DailyPulse {
  param([object]$Config, [string]$GraphToken, [string]$SiteId, [datetime]$Now, [bool]$DoSendEmail)
  Write-Output "Pulso: leyendo Firestore."
  $data = Get-FirestoreData -Config $Config
  Write-Output "Pulso: construyendo analisis."
  $pulse = Build-Pulse -Config $Config -Data $data -Now $Now
  $html = Render-PulseHtmlV2 -Config $Config -Pulse $pulse
  $json = $pulse | ConvertTo-Json -Depth 80
  $dateKey = $pulse.date
  Ensure-GraphFolder -Token $GraphToken -SiteId $SiteId -FolderPath $Config.sharepoint.pulso_folder
  Ensure-GraphFolder -Token $GraphToken -SiteId $SiteId -FolderPath $Config.sharepoint.biblia_folder
  Ensure-GraphFolder -Token $GraphToken -SiteId $SiteId -FolderPath $Config.sharepoint.traspaso_folder
  Write-TextFileToGraph -Token $GraphToken -SiteId $SiteId -FilePath "$($Config.sharepoint.pulso_folder)/$dateKey.html" -Text $html -ContentType "text/html; charset=utf-8"
  Write-TextFileToGraph -Token $GraphToken -SiteId $SiteId -FilePath "$($Config.sharepoint.pulso_folder)/$dateKey.json" -Text $json -ContentType "application/json; charset=utf-8"
  Write-TextFileToGraph -Token $GraphToken -SiteId $SiteId -FilePath "$($Config.sharepoint.biblia_folder)/$dateKey.json" -Text ($pulse.raw.biblia | ConvertTo-Json -Depth 60) -ContentType "application/json; charset=utf-8"
  Write-TextFileToGraph -Token $GraphToken -SiteId $SiteId -FilePath "$($Config.sharepoint.traspaso_folder)/$dateKey.json" -Text ($pulse.sections.traspasoControlOperacion | ConvertTo-Json -Depth 30) -ContentType "application/json; charset=utf-8"
  if ($DoSendEmail) {
    $subject = "Pulso gerencial MAYU - $dateKey - R$($pulse.summary.rojos) A$($pulse.summary.amarillos)"
    Send-GraphMail -Token $GraphToken -Sender $Config.mail.sender -To @($Config.mail.felix) -Cc @() -Subject $subject -HtmlBody $html
    Write-Output "Pulso: correo enviado a $($Config.mail.felix)."
  } else {
    Write-Output "Pulso: SendEmail=false, no se envia correo."
  }
  Write-Output "Pulso: outputs guardados en SharePoint."
}

$configJson = Get-RunbookVariable "MayuAgentsConfigJson"
$config = $configJson | ConvertFrom-Json
$now = if ($Date) { [datetime]::Parse($Date) } else { Get-Mayutime -TimeZoneName $config.timezone }
$graphToken = Get-GraphToken
$siteId = Get-SiteId -Token $graphToken -HostName $config.sharepoint.host

Write-Output "MAYU Agents iniciado. Modo=$Mode Fecha=$($now.ToString("yyyy-MM-dd")) SendEmail=$SendEmail"
Ensure-GraphFolder -Token $graphToken -SiteId $siteId -FolderPath $config.sharepoint.base_folder

if ($Mode -eq "test") {
  $body = New-MayuEmailLayout -Title "MAYU Agents operativo" -Subtitle "Prueba de correo del runtime operacional." -ContentHtml "<p>Hora MAYU: $($now.ToString("yyyy-MM-dd HH:mm"))</p>" -Footer "Generado por MAYU Agents."
  if ($SendEmail) {
    Send-GraphMail -Token $graphToken -Sender $config.mail.sender -To @($config.mail.felix) -Cc @() -Subject "Prueba MAYU Agents" -HtmlBody $body
    Write-Output "Prueba enviada a $($config.mail.felix)."
  } else {
    Write-Output "Prueba OK sin correo."
  }
} elseif ($Mode -eq "morning_reports") {
  Write-Output "Reportes manana: Pulso + Bodega/Materiales + Finanzas + Administrador Finanzas."
  Invoke-DailyPulse -Config $config -GraphToken $graphToken -SiteId $siteId -Now $now -DoSendEmail $SendEmail
  Invoke-BodegaMateriales -Config $config -GraphToken $graphToken -SiteId $siteId -Now $now -DoSendEmail $SendEmail
  Invoke-Finanzas -Config $config -GraphToken $graphToken -SiteId $siteId -Now $now -DoSendEmail $SendEmail
  Invoke-FinanzasAdmin -Config $config -GraphToken $graphToken -SiteId $siteId -Now $now -DoSendEmail $SendEmail
} elseif ($Mode -eq "daily_pulse") {
  Invoke-DailyPulse -Config $config -GraphToken $graphToken -SiteId $siteId -Now $now -DoSendEmail $SendEmail
} elseif ($Mode -eq "bodega_materiales") {
  Invoke-BodegaMateriales -Config $config -GraphToken $graphToken -SiteId $siteId -Now $now -DoSendEmail $SendEmail
} elseif ($Mode -eq "bodega_materiales_respuestas") {
  Invoke-BodegaMaterialesResponder -Config $config -GraphToken $graphToken -SiteId $siteId -Now $now
} elseif ($Mode -eq "finanzas") {
  Invoke-Finanzas -Config $config -GraphToken $graphToken -SiteId $siteId -Now $now -DoSendEmail $SendEmail
} elseif ($Mode -eq "finanzas_admin") {
  Invoke-FinanzasAdmin -Config $config -GraphToken $graphToken -SiteId $siteId -Now $now -DoSendEmail $SendEmail
} elseif ($Mode -eq "finanzas_dte_inbox") {
  Invoke-FinanzasDteInbox -Config $config -GraphToken $graphToken -SiteId $siteId -Now $now
} elseif ($Mode -eq "finanzas_respuestas") {
  Invoke-FinanzasResponder -Config $config -GraphToken $graphToken -SiteId $siteId -Now $now
} elseif ($Mode -eq "bice_cartola_mail") {
  Invoke-BiceCartolaMail -Config $config -GraphToken $graphToken -SiteId $siteId -Now $now
}
