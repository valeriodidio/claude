<#
.SYNOPSIS
    File watcher che lancia automaticamente deploy.ps1 quando la staging
    D:\admin_yeppon_python\staging\ smette di cambiare da N secondi.

.DESCRIPTION
    Pensato per girare come servizio Windows via NSSM (vedi install-watcher.ps1).

    Logica:
      - FileSystemWatcher su SrcRoot (ricorsivo) intercetta create/change/rename/delete
      - Ogni evento aggiorna $state.LastChange
      - Il loop principale controlla ogni 3 secondi:
          * se LastChange e' piu' vecchio di QuietSeconds -> lancia deploy.ps1
          * dopo il deploy, grace period per assorbire gli eventi dell'archiviazione
      - Ogni deploy e' wrappato in try/catch: un fallimento non uccide il watcher
      - Log su file con rotazione semplice (soglia MaxLogMB)

.NOTES
    - Richiede privilegi di Admin (per chiamare deploy.ps1 che tocca NSSM)
    - L'evento "change" di FileSystemWatcher puo' scattare piu' volte per un
      singolo upload FileZilla: il debounce a 30s raggruppa tutto naturalmente.
#>

# ============================================================================
# CONFIG  -- allineato a deploy.ps1
# ============================================================================
$CFG = @{
    SrcRoot        = "D:\admin_yeppon_python\staging"
    DeployScript   = "D:\admin_yeppon_python\apps\deploy\deploy.ps1"
    LogFile        = "D:\admin_yeppon_python\deploy\watcher.log"
    QuietSeconds   = 30          # attesa silenzio prima di triggerare
    PollSeconds    = 3           # intervallo di polling del loop
    GraceSeconds   = 5           # pausa dopo il deploy per assorbire eventi di archiviazione
    MaxLogMB       = 5           # rotazione log oltre questa soglia
}

$ErrorActionPreference = "Continue"  # non vogliamo che il watcher muoia su un errore isolato

# ============================================================================
# LOGGING
# ============================================================================
function Write-WatcherLog {
    param([string]$Level, [string]$Msg)
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "{0} | {1,-5} | {2}" -f $ts, $Level, $Msg

    # su stdout (cattura di NSSM)
    Write-Host $line

    # su file con rotazione naive
    try {
        if (Test-Path $CFG.LogFile) {
            $sizeMB = (Get-Item $CFG.LogFile).Length / 1MB
            if ($sizeMB -gt $CFG.MaxLogMB) {
                $old = $CFG.LogFile + ".old"
                if (Test-Path $old) { Remove-Item $old -Force }
                Move-Item $CFG.LogFile $old -Force
            }
        } else {
            $parent = Split-Path $CFG.LogFile -Parent
            if (-not (Test-Path $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
        }
        Add-Content -Path $CFG.LogFile -Value $line -Encoding UTF8
    } catch {
        # se il log rompe, non abbiamo molto da fare -- continuiamo
    }
}

function Log-Info  { param([string]$m) Write-WatcherLog "INFO"  $m }
function Log-Warn  { param([string]$m) Write-WatcherLog "WARN"  $m }
function Log-Err   { param([string]$m) Write-WatcherLog "ERROR" $m }

# ============================================================================
# HELPERS
# ============================================================================
function Test-StagingHasContent {
    <#  True se la staging contiene almeno un file sotto asp_nuovo\ o python_backend\ #>
    param([string]$Root)
    if (-not (Test-Path $Root)) { return $false }
    foreach ($sub in @("asp_nuovo", "python_backend")) {
        $path = Join-Path $Root $sub
        if (Test-Path $path) {
            $first = Get-ChildItem -Path $path -Recurse -File -ErrorAction SilentlyContinue |
                     Select-Object -First 1
            if ($first) { return $true }
        }
    }
    return $false
}

function Invoke-Deploy {
    param([string]$DeployScript)
    Log-Info "trigger deploy: $DeployScript -AutoConfirm"
    try {
        # Usa & e cattura stdout+stderr. Il deploy gia' logga su console, qui
        # basta inoltrare le righe nel nostro log per avere tutto in un posto.
        $output = & $DeployScript -AutoConfirm 2>&1
        $ec = $LASTEXITCODE
        foreach ($l in $output) { Log-Info "[deploy] $l" }
        if ($ec -eq 0) {
            Log-Info "deploy OK (exit 0)"
        } else {
            Log-Warn "deploy terminato con exit code $ec"
        }
    } catch {
        Log-Err "eccezione durante deploy: $($_.Exception.Message)"
    }
}

# ============================================================================
# MAIN
# ============================================================================
Log-Info "========== watcher start =========="
Log-Info "SrcRoot        = $($CFG.SrcRoot)"
Log-Info "DeployScript   = $($CFG.DeployScript)"
Log-Info "QuietSeconds   = $($CFG.QuietSeconds)"
Log-Info "PollSeconds    = $($CFG.PollSeconds)"
Log-Info "GraceSeconds   = $($CFG.GraceSeconds)"

# Precondizioni
if (-not (Test-Path $CFG.SrcRoot)) {
    Log-Err "SrcRoot non esiste: $($CFG.SrcRoot)"
    Log-Info "creo la cartella..."
    try {
        New-Item -ItemType Directory -Path $CFG.SrcRoot -Force | Out-Null
    } catch {
        Log-Err "impossibile creare SrcRoot: $($_.Exception.Message)"
        exit 1
    }
}
if (-not (Test-Path $CFG.DeployScript)) {
    Log-Err "DeployScript non trovato: $($CFG.DeployScript)"
    exit 1
}

# Stato condiviso fra event handlers (runspace separato) e loop principale
$script:state = [hashtable]::Synchronized(@{
    LastChange       = [datetime]::MinValue
    DeployInProgress = $false
})

# Setup FileSystemWatcher
$fsw = New-Object IO.FileSystemWatcher
$fsw.Path                  = $CFG.SrcRoot
$fsw.IncludeSubdirectories = $true
$fsw.EnableRaisingEvents   = $true
$fsw.NotifyFilter = [IO.NotifyFilters]::FileName `
                 -bor [IO.NotifyFilters]::LastWrite `
                 -bor [IO.NotifyFilters]::DirectoryName `
                 -bor [IO.NotifyFilters]::Size

# Handler per tutti gli eventi. Gira in un runspace separato, quindi NON vede
# $script:state direttamente: passo lo state tramite MessageData e lo riprendo
# via $Event.MessageData dentro l'action.
$onChangeScript = {
    $Event.MessageData.LastChange = Get-Date
}

$events = @()
$events += Register-ObjectEvent -InputObject $fsw -EventName "Created" -Action $onChangeScript -MessageData $script:state
$events += Register-ObjectEvent -InputObject $fsw -EventName "Changed" -Action $onChangeScript -MessageData $script:state
$events += Register-ObjectEvent -InputObject $fsw -EventName "Renamed" -Action $onChangeScript -MessageData $script:state
$events += Register-ObjectEvent -InputObject $fsw -EventName "Deleted" -Action $onChangeScript -MessageData $script:state

Log-Info "FileSystemWatcher attivo. In ascolto..."

# Loop principale
try {
    while ($true) {
        Start-Sleep -Seconds $CFG.PollSeconds

        if ($script:state.DeployInProgress) { continue }
        if ($script:state.LastChange -eq [datetime]::MinValue) { continue }

        $quietFor = (Get-Date) - $script:state.LastChange
        if ($quietFor.TotalSeconds -lt $CFG.QuietSeconds) { continue }

        # Silenzio raggiunto. Staging ha qualcosa?
        if (-not (Test-StagingHasContent -Root $CFG.SrcRoot)) {
            # Probabilmente erano gli eventi di archiviazione post-deploy.
            Log-Info "silenzio per $([int]$quietFor.TotalSeconds)s ma staging vuota, reset"
            $script:state.LastChange = [datetime]::MinValue
            continue
        }

        Log-Info "silenzio per $([int]$quietFor.TotalSeconds)s, staging ha contenuto -> deploy"
        $script:state.DeployInProgress = $true
        try {
            Invoke-Deploy -DeployScript $CFG.DeployScript
        } finally {
            # Assorbi eventi post-deploy (archiviazione svuota la staging)
            Start-Sleep -Seconds $CFG.GraceSeconds
            $script:state.LastChange = [datetime]::MinValue
            $script:state.DeployInProgress = $false
            Log-Info "pronto per il prossimo deploy."
        }
    }
} finally {
    Log-Info "watcher in spegnimento, cleanup..."
    foreach ($e in $events) {
        try { Unregister-Event -SourceIdentifier $e.Name -ErrorAction SilentlyContinue } catch { }
    }
    $fsw.Dispose()
    Log-Info "========== watcher stop =========="
}
