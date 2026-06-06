# stop.ps1 - Stop Gateway and UI services
$base    = $PSScriptRoot
$pidFile = "$base\.pids"

function Stop-Svc($name, $pid) {
    $proc = Get-Process -Id $pid -ErrorAction SilentlyContinue
    if ($proc) {
        Stop-Process -Id $pid -Force
        Write-Host "  Stopped $name (PID $pid)" -ForegroundColor Green
    } else {
        Write-Host "  $name (PID $pid) was not running" -ForegroundColor DarkGray
    }
}

if (-not (Test-Path $pidFile)) {
    Write-Host "No .pids file found - trying to kill by port..." -ForegroundColor Yellow
    @(8101, 8000) | ForEach-Object {
        $port = $_
        $conn = Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue
        if ($conn) {
            Stop-Process -Id $conn.OwningProcess -Force -ErrorAction SilentlyContinue
            Write-Host "  Killed process on port $port" -ForegroundColor Green
        } else {
            Write-Host "  Nothing on port $port" -ForegroundColor DarkGray
        }
    }
    exit 0
}

$saved = Get-Content $pidFile | ConvertFrom-Json
Stop-Svc "Gateway" $saved.gateway
Stop-Svc "UI"      $saved.ui

Remove-Item $pidFile -Force
Write-Host ""
Write-Host "All services stopped." -ForegroundColor Cyan
