param(
  [Parameter(Mandatory = $true)]
  [string]$SubscriptionId,
  [string]$ResourceGroup = "rg-mayu-minutas",
  [string]$Location = "eastus",
  [string]$AutomationAccount = "aa-mayu-agent-runtime",
  [string]$RunbookName = "FiscalizadorReunionesMAYU",
  [Parameter(Mandatory = $true)]
  [string]$GraphTenantId,
  [Parameter(Mandatory = $true)]
  [string]$GraphClientId,
  [Parameter(Mandatory = $true)]
  [string]$GraphClientSecret,
  [Parameter(Mandatory = $true)]
  [string]$OpenAiApiKey,
  [string]$OpenAiModel = "gpt-5.4-mini"
)

$ErrorActionPreference = "Stop"
$ApiVersion = "2024-10-23"
$RepoRoot = Split-Path -Parent (Split-Path -Parent $PSCommandPath)

function Invoke-Az {
  param([Parameter(Mandatory = $true)][string[]]$Args)
  & az @Args
  if ($LASTEXITCODE -ne 0) {
    throw "az $($Args -join ' ') fallo con codigo $LASTEXITCODE"
  }
}

function Invoke-AzJson {
  param([Parameter(Mandatory = $true)][string[]]$Args)
  $raw = & az @Args -o json
  if ($LASTEXITCODE -ne 0) {
    throw "az $($Args -join ' ') fallo con codigo $LASTEXITCODE"
  }
  if ([string]::IsNullOrWhiteSpace($raw)) {
    return $null
  }
  return $raw | ConvertFrom-Json
}

function Write-JsonTempFile {
  param([object]$Body, [string]$Prefix)
  $path = Join-Path $env:RUNNER_TEMP ("${Prefix}_$([Guid]::NewGuid().ToString("N")).json")
  if ([string]::IsNullOrWhiteSpace($env:RUNNER_TEMP)) {
    $path = Join-Path $env:TEMP ("${Prefix}_$([Guid]::NewGuid().ToString("N")).json")
  }
  $json = $Body | ConvertTo-Json -Depth 80
  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText($path, $json, $utf8NoBom)
  return $path
}

function Invoke-AzRestBody {
  param([string]$Method, [string]$Url, [object]$Body)
  $path = Write-JsonTempFile -Body $Body -Prefix "mayu_az"
  try {
    Invoke-Az @("rest", "--method", $Method, "--url", $Url, "--body", "@$path") | Out-Null
  } finally {
    if (Test-Path $path) {
      Remove-Item $path -Force
    }
  }
}

function Assert-PowerShellSyntax {
  param([string]$Path)
  $tokens = $null
  $errors = $null
  [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors) | Out-Null
  if ($errors.Count -gt 0) {
    $message = ($errors | ForEach-Object { $_.Message }) -join "`n"
    throw "Error de sintaxis en $Path`n$message"
  }
}

function ConvertTo-DrivePath {
  param([string]$Path)
  (($Path.Trim("/") -split "/") | Where-Object { $_ } | ForEach-Object { [Uri]::EscapeDataString($_) }) -join "/"
}

function Get-GraphToken {
  $token = Invoke-RestMethod `
    -Method Post `
    -Uri "https://login.microsoftonline.com/$GraphTenantId/oauth2/v2.0/token" `
    -ContentType "application/x-www-form-urlencoded" `
    -Body @{
      client_id = $GraphClientId
      client_secret = $GraphClientSecret
      scope = "https://graph.microsoft.com/.default"
      grant_type = "client_credentials"
    }
  return $token.access_token
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

function Ensure-GraphFolder {
  param([string]$Token, [string]$SiteId, [string]$FolderPath)

  $parts = ($FolderPath.Trim("/") -split "/") | Where-Object { $_ }
  $current = ""
  foreach ($part in $parts) {
    $target = if ($current) { "$current/$part" } else { $part }
    $encodedTarget = ConvertTo-DrivePath $target
    try {
      Invoke-GraphGet -Token $Token -Uri "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/root:/$encodedTarget" | Out-Null
    } catch {
      $parentUri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/root/children"
      if ($current) {
        $encodedParent = ConvertTo-DrivePath $current
        $parentUri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/root:/$encodedParent`:/children"
      }
      Invoke-GraphJson -Token $Token -Method Post -Uri $parentUri -Body @{
        name = $part
        folder = @{}
        "@microsoft.graph.conflictBehavior" = "replace"
      } | Out-Null
    }
    $current = $target
  }
}

function Upload-CoreScriptToSharePoint {
  param([object]$Config, [string]$CorePath)

  $token = Get-GraphToken
  $hostName = [string]$Config.sharepoint.host
  $siteId = (Invoke-GraphGet -Token $token -Uri "https://graph.microsoft.com/v1.0/sites/$hostName").id
  $runtimePath = [string]$Config.sharepoint.runtime_script_path
  if ([string]::IsNullOrWhiteSpace($runtimePath)) {
    $runtimePath = "minutas_archivadas/runtime/FiscalizadorReunionesMAYU.core.ps1"
  }

  $folderPath = Split-Path $runtimePath.Replace("\", "/") -Parent
  $folderPath = $folderPath.Replace("\", "/")
  Ensure-GraphFolder -Token $token -SiteId $siteId -FolderPath $folderPath

  $encodedPath = ConvertTo-DrivePath $runtimePath
  $bytes = [System.IO.File]::ReadAllBytes($CorePath)
  Invoke-RestMethod `
    -Method Put `
    -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/drive/root:/$encodedPath`:/content" `
    -Headers @{ Authorization = "Bearer $token" } `
    -ContentType "text/plain; charset=utf-8" `
    -Body $bytes | Out-Null

  Write-Host "Runtime publicado en SharePoint: $runtimePath"
}

function Set-AutomationVariable {
  param(
    [string]$Name,
    [string]$Value,
    [bool]$Encrypted,
    [string]$Description
  )

  if ([string]::IsNullOrWhiteSpace($Value)) {
    throw "Variable $Name viene vacia."
  }

  $serializedValue = $Value | ConvertTo-Json -Compress
  $url = "https://management.azure.com/subscriptions/$SubscriptionId/resourceGroups/$ResourceGroup/providers/Microsoft.Automation/automationAccounts/$AutomationAccount/variables/$Name`?api-version=$ApiVersion"
  Invoke-AzRestBody -Method "put" -Url $url -Body @{
    properties = @{
      description = $Description
      isEncrypted = $Encrypted
      value = $serializedValue
    }
  }
}

function Publish-Runbook {
  param([string]$RunbookPath)

  $base = "https://management.azure.com/subscriptions/$SubscriptionId/resourceGroups/$ResourceGroup/providers/Microsoft.Automation/automationAccounts/$AutomationAccount"
  $runbookUrl = "$base/runbooks/$RunbookName`?api-version=$ApiVersion"
  $contentUrl = "$base/runbooks/$RunbookName/draft/content?api-version=$ApiVersion"
  $publishUrl = "$base/runbooks/$RunbookName/publish?api-version=$ApiVersion"

  Invoke-AzRestBody -Method "put" -Url $runbookUrl -Body @{
    location = $Location
    properties = @{
      runbookType = "PowerShell"
      logProgress = $true
      logVerbose = $true
      description = "Loader estable del fiscalizador autonomo de reuniones MAYU"
    }
  }

  Invoke-Az @(
    "rest",
    "--method", "put",
    "--url", $contentUrl,
    "--headers", "Content-Type=text/powershell",
    "--body", "@$RunbookPath"
  ) | Out-Null

  Start-Sleep -Seconds 5
  Invoke-Az @("rest", "--method", "post", "--url", $publishUrl) | Out-Null
  Write-Host "Runbook publicado: $RunbookName"
}

$runbookPath = Join-Path $RepoRoot "runbooks/FiscalizadorReunionesMAYU.ps1"
$corePath = Join-Path $RepoRoot "runtime/FiscalizadorReunionesMAYU.core.ps1"
$configPath = Join-Path $RepoRoot "config/fiscalizador_config.json"

if (-not (Test-Path $runbookPath)) { throw "No existe runbook: $runbookPath" }
if (-not (Test-Path $corePath)) { throw "No existe runtime: $corePath" }
if (-not (Test-Path $configPath)) { throw "No existe config: $configPath" }

Assert-PowerShellSyntax -Path $runbookPath
Assert-PowerShellSyntax -Path $corePath

$configJson = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8
$config = $configJson | ConvertFrom-Json

Invoke-Az @("account", "set", "--subscription", $SubscriptionId) | Out-Null
Invoke-Az @("group", "create", "--name", $ResourceGroup, "--location", $Location, "--subscription", $SubscriptionId) | Out-Null

$accountUrl = "https://management.azure.com/subscriptions/$SubscriptionId/resourceGroups/$ResourceGroup/providers/Microsoft.Automation/automationAccounts/$AutomationAccount`?api-version=$ApiVersion"
Invoke-AzRestBody -Method "put" -Url $accountUrl -Body @{
  location = $Location
  properties = @{
    publicNetworkAccess = $true
    sku = @{ name = "Basic" }
  }
}

Upload-CoreScriptToSharePoint -Config $config -CorePath $corePath

Set-AutomationVariable -Name "MayuTenantId" -Value $GraphTenantId -Encrypted $false -Description "Tenant Azure AD MAYU para Graph"
Set-AutomationVariable -Name "MayuClientId" -Value $GraphClientId -Encrypted $false -Description "Client ID app MAYU Minutas Bot"
Set-AutomationVariable -Name "MayuClientSecret" -Value $GraphClientSecret -Encrypted $true -Description "Client secret app MAYU Minutas Bot"
Set-AutomationVariable -Name "MayuOpenAiApiKey" -Value $OpenAiApiKey -Encrypted $true -Description "OpenAI API key para reportes inteligentes"
Set-AutomationVariable -Name "MayuOpenAiModel" -Value $OpenAiModel -Encrypted $false -Description "Modelo OpenAI para resumenes"
Set-AutomationVariable -Name "MayuFiscalizadorConfigJson" -Value $configJson -Encrypted $false -Description "Config reuniones fiscalizador MAYU"

Publish-Runbook -RunbookPath $runbookPath

Write-Host "Deploy MAYU completo."
