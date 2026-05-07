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

$remote = git remote get-url origin 2>$null
if ($LASTEXITCODE -ne 0) {
  git remote add origin $RepoUrl
} elseif ($remote -ne $RepoUrl) {
  git remote set-url origin $RepoUrl
}

Write-Host "Sincronizando con main remoto..." -ForegroundColor Cyan
git pull origin main --rebase --allow-unrelated-histories

Write-Host "Subiendo a GitHub..." -ForegroundColor Cyan
git push -u origin main

Write-Host ""
Write-Host "Push completado." -ForegroundColor Green
