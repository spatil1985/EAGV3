# start.ps1 — Start Gateway (:8101) and UI (:8000) in background windows
$uv  = "C:\Users\sudip\AppData\Local\Programs\Python\Python314\Scripts\uv.exe"
$base = $PSScriptRoot
$pidFile = "$base\.pids"

# ── Check if already running ────────────────────────────────────────────────
if (Test-Path $pidFile) {
    $existing = Get-Content $pidFile | ConvertFrom-Json
    $gwAlive = $existing.gateway -and (Get-Process -Id $existing.gateway -ErrorAction SilentlyContinue)
    $uiAlive = $existing.ui     -and (Get-Process -Id $existing.ui     -ErrorAction SilentlyContinue)
    if ($gwAlive -or $uiAlive) {
        Write-Host "Services already running. Use restart.ps1 or stop.ps1 first." -ForegroundColor Yellow
        if ($gwAlive) { Write-Host "  Gateway PID: $($existing.gateway)" }
        if ($uiAlive) { Write-Host "  UI      PID: $($existing.ui)" }
        exit 0
    }
}

# ── Start Gateway ────────────────────────────────────────────────────────────
Write-Host "Starting Gateway on :8101 ..." -ForegroundColor Cyan
$gwProc = Start-Process -FilePath $uv `
    -ArgumentList "run uvicorn gateway:app --port 8101" `
    -WorkingDirectory "$base\llm_gatewayV3" `
    -PassThru -NoNewWindow `
    -RedirectStandardOutput "$base\logs\gateway.log" `
    -RedirectStandardError  "$base\logs\gateway_err.log"

# ── Wait for gateway to be ready ─────────────────────────────────────────────
$ready = $false
for ($i = 0; $i -lt 15; $i++) {
    Start-Sleep 1
    try {
        $r = Invoke-WebRequest -Uri "http://localhost:8101/health" -TimeoutSec 2 -ErrorAction Stop
        if ($r.StatusCode -eq 200) { $ready = $true; break }
    } catch {}
    Write-Host "  waiting for gateway... ($($i+1)s)" -ForegroundColor DarkGray
}

if (-not $ready) {
    Write-Host "Gateway did not start. Check logs\gateway_err.log" -ForegroundColor Red
    $gwProc | Stop-Process -Force -ErrorAction SilentlyContinue
    exit 1
}
Write-Host "  Gateway ready (PID $($gwProc.Id))" -ForegroundColor Green

# ── Start UI ─────────────────────────────────────────────────────────────────
Write-Host "Starting UI on :8000 ..." -ForegroundColor Cyan
$uiProc = Start-Process -FilePath $uv `
    -ArgumentList "run python ui_server.py" `
    -WorkingDirectory $base `
    -PassThru -NoNewWindow `
    -RedirectStandardOutput "$base\logs\ui.log" `
    -RedirectStandardError  "$base\logs\ui_err.log"

Start-Sleep 3
if ($uiProc.HasExited) {
    Write-Host "UI failed to start. Check logs\ui_err.log" -ForegroundColor Red
    $gwProc | Stop-Process -Force -ErrorAction SilentlyContinue
    exit 1
}
Write-Host "  UI ready (PID $($uiProc.Id))" -ForegroundColor Green

# ── Save PIDs ─────────────────────────────────────────────────────────────────
New-Item -ItemType Directory -Force "$base\logs" | Out-Null
@{ gateway = $gwProc.Id; ui = $uiProc.Id } | ConvertTo-Json | Set-Content $pidFile

Write-Host ""
Write-Host "All services running." -ForegroundColor Green
Write-Host "  UI:      http://localhost:8000" -ForegroundColor White
Write-Host "  Gateway: http://localhost:8101/health" -ForegroundColor White
Write-Host "  Logs:    $base\logs\" -ForegroundColor DarkGray
Write-Host ""

# Open browser
Start-Process "http://localhost:8000"
