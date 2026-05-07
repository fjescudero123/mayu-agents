param(
  [Parameter(Mandatory = $true)]
  [string]$GitHubOwner,
  [Parameter(Mandatory = $true)]
  [string]$GitHubRepo,
  [string]$SubscriptionId = "ea7a6973-b97f-4c64-a11e-f4afce935d06",
  [string]$TenantId = "d63ffa68-a19e-496f-aa33-2e28b2034369",
  [string]$ResourceGroup = "rg-mayu-minutas",
  [string]$AppName = "mayu-agents-github-deploy"
)

$ErrorActionPreference = "Stop"

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
  if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
  return $raw | ConvertFrom-Json
}

Invoke-Az @("account", "set", "--subscription", $SubscriptionId) | Out-Null
Invoke-Az @("group", "create", "--name", $ResourceGroup, "--location", "eastus", "--subscription", $SubscriptionId) | Out-Null

$existing = Invoke-AzJson @("ad", "app", "list", "--display-name", $AppName)
if ($existing.Count -gt 0) {
  $app = $existing[0]
} else {
  $app = Invoke-AzJson @("ad", "app", "create", "--display-name", $AppName)
}

$appId = [string]$app.appId
$objectId = [string]$app.id

$sp = Invoke-AzJson @("ad", "sp", "list", "--filter", "appId eq '$appId'")
if ($sp.Count -eq 0) {
  Invoke-Az @("ad", "sp", "create", "--id", $appId) | Out-Null
  Start-Sleep -Seconds 10
}

$subjectMain = "repo:$GitHubOwner/$GitHubRepo:ref:refs/heads/main"
$subjectEnv = "repo:$GitHubOwner/$GitHubRepo:environment:mayu-production"

$fedList = Invoke-AzJson @("ad", "app", "federated-credential", "list", "--id", $objectId)
$subjects = @($fedList | ForEach-Object { [string]$_.subject })

if ($subjects -notcontains $subjectMain) {
  $body = @{
    name = "github-main"
    issuer = "https://token.actions.githubusercontent.com"
    subject = $subjectMain
    audiences = @("api://AzureADTokenExchange")
    description = "GitHub Actions main branch MAYU agents"
  } | ConvertTo-Json -Depth 10 -Compress
  Invoke-Az @("ad", "app", "federated-credential", "create", "--id", $objectId, "--parameters", $body) | Out-Null
}

if ($subjects -notcontains $subjectEnv) {
  $body = @{
    name = "github-mayu-production"
    issuer = "https://token.actions.githubusercontent.com"
    subject = $subjectEnv
    audiences = @("api://AzureADTokenExchange")
    description = "GitHub Actions protected environment MAYU production"
  } | ConvertTo-Json -Depth 10 -Compress
  Invoke-Az @("ad", "app", "federated-credential", "create", "--id", $objectId, "--parameters", $body) | Out-Null
}

$scope = "/subscriptions/$SubscriptionId/resourceGroups/$ResourceGroup"
try {
  Invoke-Az @("role", "assignment", "create", "--assignee", $appId, "--role", "Contributor", "--scope", $scope) | Out-Null
} catch {
  Write-Warning "No se pudo crear role assignment o ya existia: $($_.Exception.Message)"
}

Write-Host ""
Write-Host "Bootstrap Azure OIDC listo."
Write-Host ""
Write-Host "Crea estas GitHub Repository Variables:"
Write-Host "AZURE_CLIENT_ID=$appId"
Write-Host "AZURE_TENANT_ID=$TenantId"
Write-Host "AZURE_SUBSCRIPTION_ID=$SubscriptionId"
Write-Host "MAYU_GRAPH_TENANT_ID=$TenantId"
Write-Host "MAYU_GRAPH_CLIENT_ID=<client_id de MAYU Minutas Bot>"
Write-Host "MAYU_OPENAI_MODEL=gpt-5.4-mini"
Write-Host ""
Write-Host "Crea estos GitHub Secrets:"
Write-Host "MAYU_GRAPH_CLIENT_SECRET=<secret de MAYU Minutas Bot>"
Write-Host "MAYU_OPENAI_API_KEY=<OpenAI API key>"
