$ErrorActionPreference = "Stop"

$RepoUrl = "https://github.com/fjescudero123/mayu-agents.git"
$Root = Split-Path -Parent $PSCommandPath
Set-Location $Root

Write-Host ""
Write-Host "MAYU Agents | Push inicial a GitHub" -ForegroundColor Cyan
Write-Host "Carpeta: $Root"
Write-Host "Repo: $RepoUrl"
Write-Host ""

git config --global --add safe.directory ($Root -replace "\\", "/")

$status = git status --porcelain
if ($status) {
  Write-Host "Hay cambios locales sin commit. Se agregaran al commit actual." -ForegroundColor Yellow
  git add .
  git commit -m "Update MAYU agents deployment pipeline"
}

$remotes = @(git remote)
if ($remotes -notcontains "origin") {
  git remote add origin $RepoUrl
} else {
  $remote = git remote get-url origin
  if ($remote -ne $RepoUrl) {
  git remote set-url origin $RepoUrl
  }
}

Write-Host "Subiendo a GitHub..." -ForegroundColor Cyan
Write-Host "Nota: como el repo remoto fue creado con README inicial, se reemplaza por este paquete local." -ForegroundColor Yellow
git push -u origin main --force
if ($LASTEXITCODE -ne 0) {
  throw "El push a GitHub fallo. Revisa el error rojo anterior."
}

Write-Host ""
Write-Host "Push completado." -ForegroundColor Green
