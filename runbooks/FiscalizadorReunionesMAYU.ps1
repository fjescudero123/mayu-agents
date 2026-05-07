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

function Get-CoreScriptFromSharePoint {
  param(
    [string]$Token,
    [string]$SiteId,
    [string]$CorePath
  )

  $encodedPath = ConvertTo-DrivePath $CorePath
  $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/root:/$encodedPath`:/content"
  $response = Invoke-WebRequest -Method Get -Uri $uri -Headers @{ Authorization = "Bearer $Token" } -UseBasicParsing
  return [string]$response.Content
}

$configJson = Get-RunbookVariable "MayuFiscalizadorConfigJson"
$config = $configJson | ConvertFrom-Json
$hostName = [string]$config.sharepoint.host
$corePath = "minutas_archivadas/runtime/FiscalizadorReunionesMAYU.core.ps1"

if ($config.sharepoint.runtime_script_path) {
  $corePath = [string]$config.sharepoint.runtime_script_path
}

Write-Output "Fiscalizador MAYU loader iniciado. Modo=$Mode Core=$corePath"

$token = Get-GraphToken
$siteId = Get-SiteId -Token $token -HostName $hostName
$coreScript = Get-CoreScriptFromSharePoint -Token $token -SiteId $siteId -CorePath $corePath

if ([string]::IsNullOrWhiteSpace($coreScript)) {
  throw "No se pudo cargar el script operativo desde SharePoint: $corePath"
}

$scriptBlock = [ScriptBlock]::Create($coreScript)
& $scriptBlock -Mode $Mode -Tipo $Tipo -Date $Date -LookbackDays $LookbackDays
