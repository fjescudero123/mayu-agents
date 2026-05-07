param(
  [Parameter(Mandatory = $true)]
  [string]$SubscriptionId,
  [string]$ResourceGroup = "rg-mayu-minutas",
  [string]$AutomationAccount = "aa-mayu-agent-runtime",
  [string]$RunbookName = "FiscalizadorReunionesMAYU",
  [ValidateSet("post_reunion", "manual_due_sweep", "weekly_report", "monthly_report", "test")]
  [string]$Mode = "test",
  [string]$Tipo = "",
  [string]$Date = "",
  [int]$LookbackDays = 7,
  [int]$TimeoutSeconds = 900
)

$ErrorActionPreference = "Stop"
$ApiVersion = "2024-10-23"

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
  param([object]$Body)
  $root = if ($env:RUNNER_TEMP) { $env:RUNNER_TEMP } else { $env:TEMP }
  $path = Join-Path $root ("mayu_job_$([Guid]::NewGuid().ToString("N")).json")
  $json = $Body | ConvertTo-Json -Depth 50
  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText($path, $json, $utf8NoBom)
  return $path
}

function Invoke-AzRestBody {
  param([string]$Method, [string]$Url, [object]$Body)
  $path = Write-JsonTempFile -Body $Body
  try {
    Invoke-Az @("rest", "--method", $Method, "--url", $Url, "--body", "@$path") | Out-Null
  } finally {
    if (Test-Path $path) {
      Remove-Item $path -Force
    }
  }
}

Invoke-Az @("account", "set", "--subscription", $SubscriptionId) | Out-Null

$base = "https://management.azure.com/subscriptions/$SubscriptionId/resourceGroups/$ResourceGroup/providers/Microsoft.Automation/automationAccounts/$AutomationAccount"
$jobId = [Guid]::NewGuid().ToString()

$params = @{
  Mode = $Mode
  LookbackDays = [string]$LookbackDays
}
if (-not [string]::IsNullOrWhiteSpace($Tipo)) { $params.Tipo = $Tipo }
if (-not [string]::IsNullOrWhiteSpace($Date)) { $params.Date = $Date }

Invoke-AzRestBody -Method "put" -Url "$base/jobs/$jobId`?api-version=$ApiVersion" -Body @{
  properties = @{
    runbook = @{ name = $RunbookName }
    parameters = $params
  }
}

Write-Host "Job iniciado: $jobId / modo=$Mode"
$deadline = [DateTimeOffset]::UtcNow.AddSeconds($TimeoutSeconds)
$terminal = @("Completed", "Failed", "Stopped", "Suspended")
$job = $null

do {
  Start-Sleep -Seconds 10
  $job = Invoke-AzJson @("rest", "--method", "get", "--url", "$base/jobs/$jobId`?api-version=$ApiVersion")
  $status = [string]$job.properties.status
  Write-Host "Estado: $status"
} while (($terminal -notcontains $status) -and ([DateTimeOffset]::UtcNow -lt $deadline))

try {
  $streams = Invoke-AzJson @("rest", "--method", "get", "--url", "$base/jobs/$jobId/streams?api-version=$ApiVersion")
  if ($streams.value) {
    Write-Host ""
    Write-Host "Streams del job:"
    foreach ($stream in $streams.value) {
      $streamType = [string]$stream.properties.streamType
      $summary = [string]$stream.properties.summary
      if (-not [string]::IsNullOrWhiteSpace($summary)) {
        Write-Host "[$streamType] $summary"
      }
    }
  }
} catch {
  Write-Warning "No se pudieron leer streams del job: $($_.Exception.Message)"
}

if ($null -eq $job) {
  throw "No se pudo leer estado del job."
}

if ([string]$job.properties.status -ne "Completed") {
  $exception = [string]$job.properties.exception
  throw "Job termino en estado $($job.properties.status). $exception"
}

Write-Host "Job completado correctamente."
