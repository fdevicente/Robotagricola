# Watchdog del Bot Agricola Santa Elisa
# - Reinicia main.py automaticamente si se cae (con backoff ante crash-loop)
# - Instancia unica (Mutex global): evita dos bots peleando por el token de Telegram
# - Log con timestamp en files\logs\watchdog.log
# Pensado para correr 24/7 via Tarea Programada (AtLogon).

$ErrorActionPreference = "Continue"

# --- Instancia unica ---
$createdNew = $false
$mutex = New-Object System.Threading.Mutex($true, "Global\AgricolaBotWatchdog", [ref]$createdNew)
if (-not $createdNew) {
    Write-Host "Ya hay un watchdog corriendo. Saliendo."
    exit 0
}

$root = "C:\Users\Windows\Desktop\Workflow\Agricola Santa Elisa\Robot"
Set-Location $root
$env:PYTHONPATH = $root

$logDir = Join-Path $root "files\logs"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$log = Join-Path $logDir "watchdog.log"

function Log($msg) {
    $line = ("{0} | {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg)
    Add-Content -Path $log -Value $line -Encoding utf8
    Write-Host $line
}

# --- Detectar Python 3.11 del proyecto ---
$py = Join-Path $env:LOCALAPPDATA "Python\bin\python3.11.exe"
if (-not (Test-Path $py)) {
    $py = Join-Path $env:LOCALAPPDATA "Programs\Python\Python311\python.exe"
}
if (-not (Test-Path $py)) {
    $cmd = Get-Command python -ErrorAction SilentlyContinue
    if ($cmd) { $py = $cmd.Source }
}
if (-not (Test-Path $py)) {
    Log "ERROR: no encuentro Python. Watchdog no puede arrancar."
    $mutex.ReleaseMutex(); exit 1
}

Log "==================== Watchdog iniciado (py=$py) ===================="

while ($true) {
    $inicio = Get-Date
    Log "Arrancando bot (main.py)..."
    try {
        $p = Start-Process -FilePath $py -ArgumentList "main.py" `
             -WorkingDirectory $root -NoNewWindow -PassThru -Wait
        $code = $p.ExitCode
        $dur = [int]((Get-Date) - $inicio).TotalSeconds
        Log "Bot termino (codigo=$code, duro ${dur}s)."
    } catch {
        $dur = [int]((Get-Date) - $inicio).TotalSeconds
        Log "Excepcion al lanzar bot: $($_.Exception.Message)"
    }

    # Backoff: si murio en menos de 30s, esperar mas para no entrar en crash-loop
    if ($dur -lt 30) {
        Log "Murio muy rapido (${dur}s). Espero 60s antes de reintentar."
        Start-Sleep -Seconds 60
    } else {
        Log "Reinicio en 10s."
        Start-Sleep -Seconds 10
    }
}
