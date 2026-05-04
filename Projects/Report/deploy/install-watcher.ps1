<#
.SYNOPSIS
    Installa il watcher come servizio Windows tramite NSSM.

.DESCRIPTION
    Crea un servizio chiamato YepponReportsDeployWatcher che:
      - lancia powershell.exe -File watcher.ps1
      - parte automaticamente all'avvio del server
      - si riavvia da solo se crasha
      - redirige stdout/stderr in D:\admin_yeppon_python\deploy\watcher-service-*.log

.PARAMETER Uninstall
    Ferma e rimuove il servizio.

.EXAMPLE
    .\install-watcher.ps1
    Installa e avvia il servizio.

.EXAMPLE
    .\install-watcher.ps1 -Uninstall
    Ferma e rimuove il servizio.

.NOTES
    Da lanciare UNA TANTUM (e di nuovo solo se cambi config o path).
    Lo script deploy.ps1 continua a funzionare anche a mano: il watcher
    e' solo un modo di scatenarlo automaticamente.
#>

[CmdletBinding()]
param(
    [switch]$Uninstall
)

$ErrorActionPreference = "Stop"

# ============================================================================
# CONFIG
# ============================================================================
$CFG = @{
    ServiceName    = "YepponReportsDeployWatcher"
    DisplayName    = "Yeppon Reports Deploy Watcher"
    Description    = "Auto-deploy di Yeppon Reports su modifica della staging D:\admin_yeppon_python\staging\"
    NssmExe        = "D:\admin_yeppon_python\bin\nssm.exe"
    PowerShellExe  = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
    WatcherScript  = "D:\admin_yeppon_python\apps\deploy\watcher.ps1"
    WorkingDir     = "D:\admin_yeppon_python\apps\deploy"
    StdoutLog      = "D:\admin_yeppon_python\deploy\watcher-service-stdout.log"
    StderrLog      = "D:\admin_yeppon_python\deploy\watcher-service-stderr.log"
    # Rotazione log NSSM oltre 10 MB
    RotateBytes    = 10485760
}

# ============================================================================
# HELPERS
# ============================================================================
function Test-Admin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $pr = New-Object Security.Principal.WindowsPrincipal($id)
    return $pr.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Write-Step { param([string]$m) Write-Host "`n==> $m" -ForegroundColor Cyan }
function Write-OK   { param([string]$m) Write-Host "  [OK] $m" -ForegroundColor Green }
function Write-ErrM { param([string]$m) Write-Host "  [ERR] $m" -ForegroundColor Red }
function Write-Info { param([string]$m) Write-Host "    $m" -ForegroundColor Gray }

# ============================================================================
# MAIN
# ============================================================================
if (-not (Test-Admin)) {
    Write-ErrM "Serve PowerShell Administrator."
    exit 1
}

if (-not (Test-Path $CFG.NssmExe)) {
    Write-ErrM "nssm.exe non trovato in $($CFG.NssmExe)"
    exit 1
}

# Controllo esistenza servizio con Get-Service (non inquina stderr come nssm)
$serviceExists = $null -ne (Get-Service -Name $CFG.ServiceName -ErrorAction SilentlyContinue)

if ($Uninstall) {
    Write-Step "Uninstall servizio $($CFG.ServiceName)"
    if (-not $serviceExists) {
        Write-Info "servizio non presente, nulla da fare."
        exit 0
    }
    & $CFG.NssmExe stop   $CFG.ServiceName 2>&1 | Out-Null
    & $CFG.NssmExe remove $CFG.ServiceName confirm 2>&1 | Out-Null
    Write-OK "servizio rimosso."
    exit 0
}

# Install
if ($serviceExists) {
    Write-Step "Servizio gia' presente, aggiorno la configurazione"
    & $CFG.NssmExe stop $CFG.ServiceName 2>&1 | Out-Null
} else {
    Write-Step "Installo servizio $($CFG.ServiceName)"
    if (-not (Test-Path $CFG.WatcherScript)) {
        Write-ErrM "watcher.ps1 non trovato in $($CFG.WatcherScript)"
        exit 1
    }
    $args = "-NoProfile -ExecutionPolicy Bypass -File `"$($CFG.WatcherScript)`""
    & $CFG.NssmExe install $CFG.ServiceName $CFG.PowerShellExe $args | Out-Null
    Write-OK "servizio creato."
}

# Configurazione (idempotente)
Write-Step "Configuro il servizio"
& $CFG.NssmExe set $CFG.ServiceName AppDirectory       $CFG.WorkingDir         | Out-Null
& $CFG.NssmExe set $CFG.ServiceName DisplayName        $CFG.DisplayName        | Out-Null
& $CFG.NssmExe set $CFG.ServiceName Description        $CFG.Description        | Out-Null
& $CFG.NssmExe set $CFG.ServiceName Start              "SERVICE_AUTO_START"    | Out-Null
& $CFG.NssmExe set $CFG.ServiceName AppStdout          $CFG.StdoutLog          | Out-Null
& $CFG.NssmExe set $CFG.ServiceName AppStderr          $CFG.StderrLog          | Out-Null
& $CFG.NssmExe set $CFG.ServiceName AppRotateFiles     1                       | Out-Null
& $CFG.NssmExe set $CFG.ServiceName AppRotateOnline    1                       | Out-Null
& $CFG.NssmExe set $CFG.ServiceName AppRotateBytes     $CFG.RotateBytes        | Out-Null
# Restart: attesa 2s tra un crash e il restart, throttling 5s, azione Restart
& $CFG.NssmExe set $CFG.ServiceName AppExit Default    Restart                 | Out-Null
& $CFG.NssmExe set $CFG.ServiceName AppRestartDelay    2000                    | Out-Null
& $CFG.NssmExe set $CFG.ServiceName AppThrottle        5000                    | Out-Null

# Crea cartella log se non esiste
$logDir = Split-Path $CFG.StdoutLog -Parent
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }

Write-OK "configurazione applicata."

Write-Step "Avvio servizio"
try {
    Start-Service -Name $CFG.ServiceName -ErrorAction Stop
} catch {
    Write-ErrM "Start-Service fallito: $($_.Exception.Message)"
    Write-Info "controlla $($CFG.StderrLog)"
    exit 2
}

# Start-Service aspetta gia' che esca da START_PENDING, ma aggiungiamo un
# piccolo margine per essere sicuri
Start-Sleep -Seconds 1

$svc = Get-Service -Name $CFG.ServiceName -ErrorAction SilentlyContinue
if ($svc -and $svc.Status -eq "Running") {
    Write-OK "servizio in esecuzione."
} else {
    $currentStatus = if ($svc) { $svc.Status } else { "NON TROVATO" }
    Write-ErrM "stato: $currentStatus"
    Write-Info "controlla $($CFG.StderrLog)"
    exit 2
}

Write-Host ""
Write-Host "Fatto. Il watcher ora e' attivo e si riavvia con Windows." -ForegroundColor Green
Write-Host ""
Write-Host "Comandi utili:"
Write-Host "  stato:       & '$($CFG.NssmExe)' status $($CFG.ServiceName)"
Write-Host "  stop:        & '$($CFG.NssmExe)' stop   $($CFG.ServiceName)"
Write-Host "  start:       & '$($CFG.NssmExe)' start  $($CFG.ServiceName)"
Write-Host "  restart:     & '$($CFG.NssmExe)' restart $($CFG.ServiceName)"
Write-Host "  log watcher: Get-Content D:\admin_yeppon_python\deploy\watcher.log -Tail 50 -Wait"
Write-Host "  log NSSM:    Get-Content $($CFG.StderrLog) -Tail 50 -Wait"
