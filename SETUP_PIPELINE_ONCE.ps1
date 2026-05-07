$ErrorActionPreference = "Stop"

$GitHubOwner = "fjescudero123"
$GitHubRepo = "mayu-agents"
$RepoFullName = "$GitHubOwner/$GitHubRepo"
$SubscriptionId = "ea7a6973-b97f-4c64-a11e-f4afce935d06"
$TenantId = "d63ffa68-a19e-496f-aa33-2e28b2034369"
$ResourceGroup = "rg-mayu-minutas"
$AppName = "mayu-agents-github-deploy"
$EnvironmentName = "mayu-production"

$Root = Split-Path -Parent $PSCommandPath
$MailConfigPath = Join-Path (Split-Path -Parent (Split-Path -Parent $Root)) "resumenes\mayu_mail_config.json"

function Invoke-External {
  param([Parameter(Mandatory = $true)][scriptblock]$Command, [string]$Label)
  & $Command
  if ($LASTEXITCODE -ne 0) {
    throw "$Label fallo con codigo $LASTEXITCODE"
  }
}

function Write-JsonTempFile {
  param([object]$Body, [string]$Prefix)
  $path = Join-Path $env:TEMP ("${Prefix}_$([Guid]::NewGuid().ToString("N")).json")
  $json = $Body | ConvertTo-Json -Depth 20
  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText($path, $json, $utf8NoBom)
  return $path
}

function Get-PlainTextFromSecureString {
  param([securestring]$Secure)
  $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secure)
  try {
    return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
  } finally {
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
  }
}

function Set-GitHubVariable {
  param([string]$Name, [string]$Value)
  if ([string]::IsNullOrWhiteSpace($Value)) {
    throw "Variable GitHub vacia: $Name"
  }
  Write-Host "GitHub variable: $Name" -ForegroundColor Cyan
  Invoke-External -Label "gh variable set $Name" -Command { gh variable set $Name --repo $RepoFullName --body $Value }
}

function Set-GitHubSecret {
  param([string]$Name, [string]$Value)
  if ([string]::IsNullOrWhiteSpace($Value)) {
    throw "Secret GitHub vacio: $Name"
  }
  Write-Host "GitHub secret: $Name" -ForegroundColor Cyan
  Invoke-External -Label "gh secret set $Name" -Command { gh secret set $Name --repo $RepoFullName --body $Value }
}

Write-Host ""
Write-Host "MAYU Agents | Setup unico GitHub Actions + Azure OIDC" -ForegroundColor Cyan
Write-Host "Repo: $RepoFullName"
Write-Host "Azure RG: $ResourceGroup"
Write-Host ""

if (-not (Test-Path $MailConfigPath)) {
  throw "No se encontro mayu_mail_config.json en: $MailConfigPath"
}

try {
  az account show --query id -o tsv | Out-Null
} catch {
  Write-Host "Azure CLI no tiene sesion activa. Se abrira login." -ForegroundColor Yellow
  Invoke-External -Label "az login" -Command { az login --tenant $TenantId }
}
Invoke-External -Label "az account set" -Command { az account set --subscription $SubscriptionId }

try {
  gh auth status -h github.com | Out-Null
} catch {
  Write-Host "GitHub CLI no tiene sesion activa. Se abrira login." -ForegroundColor Yellow
  Invoke-External -Label "gh auth login" -Command { gh auth login -h github.com -p https -w }
}

$mailConfig = Get-Content -LiteralPath $MailConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
$graphTenantId = [string]$mailConfig.azure.tenant_id
$graphClientId = [string]$mailConfig.azure.client_id
$graphClientSecret = [string]$mailConfig.azure.client_secret

if ([string]::IsNullOrWhiteSpace($graphTenantId) -or [string]::IsNullOrWhiteSpace($graphClientId) -or [string]::IsNullOrWhiteSpace($graphClientSecret)) {
  throw "mayu_mail_config.json no trae tenant/client/secret completos."
}

$openAiApiKey = [Environment]::GetEnvironmentVariable("OPENAI_API_KEY")
if ([string]::IsNullOrWhiteSpace($openAiApiKey)) {
  $functionSettingsUrl = "https://management.azure.com/subscriptions/$SubscriptionId/resourceGroups/$ResourceGroup/providers/Microsoft.Web/sites/func-mayu-agent-runtime/config/appsettings/list?api-version=2022-03-01"
  try {
    $raw = az rest --method post --url $functionSettingsUrl -o json 2>$null
    if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace($raw)) {
      $settings = $raw | ConvertFrom-Json
      if ($settings.properties.OPENAI_API_KEY) {
        $openAiApiKey = [string]$settings.properties.OPENAI_API_KEY
      }
    }
  } catch {
    $openAiApiKey = ""
  }
}
if ([string]::IsNullOrWhiteSpace($openAiApiKey)) {
  $secureOpenAi = Read-Host "Pega la OpenAI API Key para GitHub Secret MAYU_OPENAI_API_KEY" -AsSecureString
  $openAiApiKey = Get-PlainTextFromSecureString $secureOpenAi
}

Write-Host ""
Write-Host "1/4 Creando/actualizando app Azure OIDC para GitHub Actions..." -ForegroundColor Cyan
Invoke-External -Label "az group create" -Command { az group create --name $ResourceGroup --location eastus --subscription $SubscriptionId | Out-Null }

$existingRaw = az ad app list --display-name $AppName -o json
if ($LASTEXITCODE -ne 0) { throw "No se pudo consultar Azure AD apps." }
$existing = $existingRaw | ConvertFrom-Json
if (@($existing).Count -gt 0) {
  $app = @($existing)[0]
} else {
  $app = (az ad app create --display-name $AppName -o json) | ConvertFrom-Json
}

$appId = [string]$app.appId
$objectId = [string]$app.id

$spRaw = az ad sp list --filter "appId eq '$appId'" -o json
if ($LASTEXITCODE -ne 0) { throw "No se pudo consultar service principal." }
$sp = $spRaw | ConvertFrom-Json
if (@($sp).Count -eq 0) {
  Invoke-External -Label "az ad sp create" -Command { az ad sp create --id $appId | Out-Null }
  Start-Sleep -Seconds 10
}

$subjectMain = "repo:$RepoFullName`:ref:refs/heads/main"
$subjectEnv = "repo:$RepoFullName`:environment:$EnvironmentName"
$fedRaw = az ad app federated-credential list --id $objectId -o json
if ($LASTEXITCODE -ne 0) { throw "No se pudo consultar federated credentials." }
$fedList = $fedRaw | ConvertFrom-Json
$subjects = @($fedList | ForEach-Object { [string]$_.subject })

if ($subjects -notcontains $subjectMain) {
  $body = @{
    name = "github-main"
    issuer = "https://token.actions.githubusercontent.com"
    subject = $subjectMain
    audiences = @("api://AzureADTokenExchange")
    description = "GitHub Actions main branch MAYU agents"
  }
  $fedPath = Write-JsonTempFile -Body $body -Prefix "mayu_federated_main"
  try {
    Invoke-External -Label "az federated credential main" -Command { az ad app federated-credential create --id $objectId --parameters "@$fedPath" | Out-Null }
  } finally {
    if (Test-Path $fedPath) { Remove-Item $fedPath -Force }
  }
}

if ($subjects -notcontains $subjectEnv) {
  $body = @{
    name = "github-mayu-production"
    issuer = "https://token.actions.githubusercontent.com"
    subject = $subjectEnv
    audiences = @("api://AzureADTokenExchange")
    description = "GitHub Actions protected environment MAYU production"
  }
  $fedPath = Write-JsonTempFile -Body $body -Prefix "mayu_federated_env"
  try {
    Invoke-External -Label "az federated credential environment" -Command { az ad app federated-credential create --id $objectId --parameters "@$fedPath" | Out-Null }
  } finally {
    if (Test-Path $fedPath) { Remove-Item $fedPath -Force }
  }
}

$scope = "/subscriptions/$SubscriptionId/resourceGroups/$ResourceGroup"
try {
  az role assignment create --assignee $appId --role Contributor --scope $scope | Out-Null
} catch {
  Write-Warning "Role assignment Contributor pudo existir ya o tardar en propagarse. Detalle: $($_.Exception.Message)"
}

Write-Host ""
Write-Host "2/4 Configurando variables y secrets de GitHub..." -ForegroundColor Cyan
Set-GitHubVariable -Name "AZURE_CLIENT_ID" -Value $appId
Set-GitHubVariable -Name "AZURE_TENANT_ID" -Value $TenantId
Set-GitHubVariable -Name "AZURE_SUBSCRIPTION_ID" -Value $SubscriptionId
Set-GitHubVariable -Name "MAYU_GRAPH_TENANT_ID" -Value $graphTenantId
Set-GitHubVariable -Name "MAYU_GRAPH_CLIENT_ID" -Value $graphClientId
Set-GitHubVariable -Name "MAYU_OPENAI_MODEL" -Value "gpt-5.4-mini"
Set-GitHubSecret -Name "MAYU_GRAPH_CLIENT_SECRET" -Value $graphClientSecret
Set-GitHubSecret -Name "MAYU_OPENAI_API_KEY" -Value $openAiApiKey

Write-Host ""
Write-Host "3/4 Configurando environment mayu-production con aprobacion de Felix..." -ForegroundColor Cyan
try {
  $user = (gh api user | ConvertFrom-Json)
  $envBody = @{
    wait_timer = 0
    reviewers = @(@{ type = "User"; id = [int64]$user.id })
    deployment_branch_policy = $null
  }
  $envPath = Write-JsonTempFile -Body $envBody -Prefix "mayu_env"
  try {
    Invoke-External -Label "gh api environment" -Command { gh api --method PUT "repos/$RepoFullName/environments/$EnvironmentName" --input $envPath | Out-Null }
  } finally {
    if (Test-Path $envPath) { Remove-Item $envPath -Force }
  }
} catch {
  Write-Warning "No se pudo configurar aprobacion del environment automaticamente. El workflow igual puede correr; revisa Settings > Environments en GitHub. Detalle: $($_.Exception.Message)"
}

Write-Host ""
Write-Host "4/4 Lanzando prueba GitHub Actions modo test..." -ForegroundColor Cyan
try {
  gh workflow run deploy-mayu-agents.yml --repo $RepoFullName -f run_mode=test -f lookback_days=7
  Start-Sleep -Seconds 5
  gh run list --repo $RepoFullName --workflow deploy-mayu-agents.yml --limit 1
  Write-Host ""
  Write-Host "Si el run queda esperando aprobacion, entra a Actions y aprueba el environment '$EnvironmentName'." -ForegroundColor Yellow
} catch {
  Write-Warning "No se pudo lanzar el workflow automaticamente. Puedes hacerlo desde GitHub > Actions > Deploy MAYU Agents > Run workflow. Detalle: $($_.Exception.Message)"
}

Write-Host ""
Write-Host "Setup completo." -ForegroundColor Green
