# restart.ps1 — Stop then start all services
$base = $PSScriptRoot
Write-Host "Restarting PADMA services..." -ForegroundColor Cyan
Write-Host ""
& "$base\stop.ps1"
Write-Host ""
& "$base\start.ps1"
