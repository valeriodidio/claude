<#
.SYNOPSIS
    Deploy automatico di Yeppon Reports (ASP admin + backend Python FastAPI).

.DESCRIPTION
    Workflow:
      1) carichi via FileZilla (o altro) i file modificati nella cartella
         di staging D:\admin_yeppon_python\staging\, mantenendo la struttura:
            asp_nuovo\*.asp
            python_backend\app\...
            python_backend\requirements.txt
      2) esegui questo script dal server in PowerShell Administrator
      3) lo script:
            - fa backup dei file di produzione che sta per sovrascrivere
            - copia gli ASP e patcha INTERNAL_TOKEN leggendolo dal .env
            - copia i file Python
            - se requirements.txt e' cambiato: pip install
            - se sono cambiati file Python: restart NSSM
            - fa health check
            - archivia lo staging sotto D:\admin_yeppon_python\deploy\archived\<timestamp>\

.PARAMETER DryRun
    Mostra cosa farebbe senza scrivere niente.

.PARAMETER AutoConfirm
    Non chiede conferma prima di procedere.

.PARAMETER Rollback
    Ripristina i file dall'ultimo backup (o quello specificato in -BackupTag).

.PARAMETER BackupTag
    Nome del backup da ripristinare (es. "20260421_143022"). Usato con -Rollback.

.PARAMETER SkipRestart
    Forza a NON riavviare il servizio anche se sono cambiati file Python.
    Utile se vuoi deployare piu' cose in sequenza e poi riavviare manualmente.

.EXAMPLE
    .\deploy.ps1 -DryRun
    Preview del deploy, niente viene scritto.

.EXAMPLE
    .\deploy.ps1
    Deploy interattivo (chiede conferma).

.EXAMPLE
    .\deploy.ps1 -AutoConfirm
    Deploy senza conferma.

.EXAMPLE
    .\deploy.ps1 -Rollback
    Ripristina l'ultimo backup.

.EXAMPLE
    .\deploy.ps1 -Rollback -BackupTag "20260421_143022"
    Ripristina uno specifico backup.
#>

[CmdletBinding()]
param(
    [switch]$DryRun,
    [switch]$AutoConfirm,
    [switch]$Rollback,
    [string]$BackupTag,
    [switch]$SkipRestart
)

$ErrorActionPreference = "Stop"

# ============================================================================
# CONFIGURAZIONE  -- modificare qui se cambiano i path sul server
# ============================================================================
$CFG = @{
    SrcRoot          = "D:\admin_yeppon_python\staging"
    AspSrcSub        = "asp_nuovo"
    PySrcSub         = "python_backend"

    AspDest          = "D:\www\yeppon.it\admin"
    PyDest           = "D:\admin_yeppon_python\apps\yeppon_reports\python_backend"

    BackupRoot       = "D:\admin_yeppon_python\deploy\backups"
    ArchiveRoot      = "D:\admin_yeppon_python\deploy\archived"
    KeepBackups      = 10

    ServiceName      = "YepponReportsAPI"
    NssmExe          = "D:\admin_yeppon_python\bin\nssm.exe"
    HealthUrl        = "http://127.0.0.1:8000/api/reports/health"
    HealthRetries    = 8
    HealthRetryDelay = 2   # secondi tra un tentativo e l'altro

    EnvFile          = "D:\admin_yeppon_python\apps\yeppon_reports\python_backend\.env"
    VenvActivate     = "D:\admin_yeppon_python\apps\yeppon_reports\python_backend\venv\Scripts\Activate.ps1"

    TokenPlaceholder = "change_me_long_random_string"

    # Pattern da IGNORARE in staging (non verranno copiati)
    IgnorePatterns   = @(
        "__pycache__",
        "*.pyc",
        ".env",
        "venv",
        ".git",
        "*.bak",
        "*.old"
    )
}

# ============================================================================
# HELPER
# ============================================================================
function Write-Step  { param([string]$m) Write-Host "`n==> $m"        -ForegroundColor Cyan }
function Write-Info  { param([string]$m) Write-Host "    $m"          -ForegroundColor Gray }
function Write-OK    { param([string]$m) Write-Host "  [OK]  $m"      -ForegroundColor Green }
function Write-WarnM { param([string]$m) Write-Host "  [WARN] $m"     -ForegroundColor Yellow }
function Write-ErrM  { param([string]$m) Write-Host "  [ERR] $m"      -ForegroundColor Red }
function Write-Act   { param([string]$m) Write-Host "  -> $m"         -ForegroundColor White }

function Test-Admin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $pr = New-Object Security.Principal.WindowsPrincipal($id)
    return $pr.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-InternalToken {
    param([string]$EnvPath)
    if (-not (Test-Path $EnvPath)) {
        throw "File .env non trovato: $EnvPath"
    }
    $line = Get-Content $EnvPath |
            Where-Object { $_ -match '^\s*INTERNAL_TOKEN\s*=' } |
            Select-Object -First 1
    if (-not $line) {
        throw "INTERNAL_TOKEN non definito in $EnvPath"
    }
    $val = ($line -split '=', 2)[1].Trim()
    # togli eventuali virgolette
    if ($val.StartsWith('"') -and $val.EndsWith('"')) { $val = $val.Substring(1, $val.Length - 2) }
    if ($val.StartsWith("'") -and $val.EndsWith("'")) { $val = $val.Substring(1, $val.Length - 2) }
    if ([string]::IsNullOrWhiteSpace($val)) {
        throw "INTERNAL_TOKEN vuoto in $EnvPath"
    }
    return $val
}

function Get-RelativeSourceFiles {
    <#  Ritorna i file sotto $Root (ricorsivo) come path relativi a $Root,
        escludendo i pattern IgnorePatterns. #>
    param([string]$Root)
    if (-not (Test-Path $Root)) { return @() }
    $rootFull = (Resolve-Path $Root).Path.TrimEnd('\') + '\'
    $all = Get-ChildItem -Path $Root -Recurse -File -Force
    $filtered = $all | Where-Object {
        $full = $_.FullName
        foreach ($pat in $CFG.IgnorePatterns) {
            if ($full -like "*\$pat\*" -or $_.Name -like $pat) { return $false }
        }
        return $true
    }
    return $filtered | ForEach-Object { $_.FullName.Substring($rootFull.Length) }
}

function New-Timestamp { Get-Date -Format "yyyyMMdd_HHmmss" }

function Copy-FileWithBackup {
    param(
        [string]$SrcFull,
        [string]$DestFull,
        [string]$BackupDir,
        [string]$BackupLabel,     # sottoradice per backup (es. "asp", "python")
        [string]$BackupRelPath    # path del file relativo alla destinazione
    )

    # 1. backup del file esistente
    if (Test-Path $DestFull) {
        $bkTarget = Join-Path (Join-Path $BackupDir $BackupLabel) $BackupRelPath
        $bkDir = Split-Path $bkTarget -Parent
        if (-not (Test-Path $bkDir)) {
            if (-not $DryRun) { New-Item -ItemType Directory -Path $bkDir -Force | Out-Null }
        }
        if ($DryRun) {
            Write-Act "[dryrun] backup: $DestFull -> $bkTarget"
        } else {
            Copy-Item $DestFull $bkTarget -Force
        }
    }

    # 2. assicura che la cartella di destinazione esista
    $destDir = Split-Path $DestFull -Parent
    if (-not (Test-Path $destDir)) {
        if (-not $DryRun) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
        Write-Act "creo cartella $destDir"
    }

    # 3. copia
    if ($DryRun) {
        Write-Act "[dryrun] copy: $SrcFull -> $DestFull"
    } else {
        Copy-Item $SrcFull $DestFull -Force
    }
}

function Get-FileHashSafe {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return $null }
    return (Get-FileHash -Path $Path -Algorithm SHA256).Hash
}

function Invoke-HealthCheck {
    param([string]$Url, [int]$Retries, [int]$DelaySec)
    for ($i = 1; $i -le $Retries; $i++) {
        try {
            $r = Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec 5 -ErrorAction Stop
            if ($r.StatusCode -eq 200 -and $r.Content -match 'ok') {
                return $true
            }
        } catch { }
        if ($i -lt $Retries) {
            Write-Info "health check tentativo $i/$Retries fallito, attendo $DelaySec s..."
            Start-Sleep -Seconds $DelaySec
        }
    }
    return $false
}

function Remove-OldBackups {
    param([string]$Root, [int]$Keep)
    if (-not (Test-Path $Root)) { return }
    $dirs = Get-ChildItem -Path $Root -Directory | Sort-Object LastWriteTime -Descending
    if ($dirs.Count -le $Keep) { return }
    $toDelete = $dirs | Select-Object -Skip $Keep
    foreach ($d in $toDelete) {
        Write-Info "rimuovo vecchio backup $($d.Name)"
        if (-not $DryRun) {
            Remove-Item $d.FullName -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

# ============================================================================
# ROLLBACK MODE
# ============================================================================
function Invoke-Rollback {
    Write-Step "MODALITA' ROLLBACK"

    if (-not (Test-Path $CFG.BackupRoot)) {
        Write-ErrM "nessuna cartella backup trovata in $($CFG.BackupRoot)"
        exit 1
    }

    if ($BackupTag) {
        $bkDir = Join-Path $CFG.BackupRoot $BackupTag
        if (-not (Test-Path $bkDir)) {
            Write-ErrM "backup non trovato: $bkDir"
            exit 1
        }
    } else {
        $latest = Get-ChildItem -Path $CFG.BackupRoot -Directory |
                  Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if (-not $latest) {
            Write-ErrM "nessun backup disponibile in $($CFG.BackupRoot)"
            exit 1
        }
        $bkDir = $latest.FullName
    }

    Write-Info "ripristino da: $bkDir"
    if (-not $AutoConfirm) {
        $ans = Read-Host "Procedere col rollback? (y/N)"
        if ($ans -notin @("y", "Y", "yes")) { Write-WarnM "annullato."; exit 0 }
    }

    # ripristina ASP
    $aspBackup = Join-Path $bkDir "asp"
    if (Test-Path $aspBackup) {
        Get-ChildItem -Path $aspBackup -Recurse -File | ForEach-Object {
            $rel = $_.FullName.Substring(($aspBackup.TrimEnd('\') + '\').Length)
            $dest = Join-Path $CFG.AspDest $rel
            Write-Act "ASP: $rel"
            if (-not $DryRun) { Copy-Item $_.FullName $dest -Force }
        }
    }

    # ripristina Python
    $pyBackup = Join-Path $bkDir "python"
    $pythonRestored = $false
    if (Test-Path $pyBackup) {
        Get-ChildItem -Path $pyBackup -Recurse -File | ForEach-Object {
            $rel = $_.FullName.Substring(($pyBackup.TrimEnd('\') + '\').Length)
            $dest = Join-Path $CFG.PyDest $rel
            Write-Act "Python: $rel"
            if (-not $DryRun) { Copy-Item $_.FullName $dest -Force }
            $pythonRestored = $true
        }
    }

    if ($pythonRestored -and -not $SkipRestart) {
        Write-Step "restart servizio"
        if (-not $DryRun) {
            # Stessa precauzione su AppExit del deploy: vedi commento piu' sotto.
            $prevAppExit = $null
            try { $prevAppExit = (& $CFG.NssmExe get $CFG.ServiceName AppExit Default) 2>$null } catch { }
            if (-not $prevAppExit) { $prevAppExit = "Restart" }
            $appExitChanged = $false
            if ($prevAppExit -match 'Restart') {
                & $CFG.NssmExe set $CFG.ServiceName AppExit Default Exit | Out-Null
                $appExitChanged = $true
            }
            try {
                & $CFG.NssmExe restart $CFG.ServiceName | Out-Null
                Start-Sleep -Seconds 3
                if (Invoke-HealthCheck -Url $CFG.HealthUrl -Retries $CFG.HealthRetries -DelaySec $CFG.HealthRetryDelay) {
                    Write-OK "servizio OK dopo rollback"
                } else {
                    Write-ErrM "servizio NON risponde dopo rollback. Controlla i log."
                    if ($appExitChanged) {
                        & $CFG.NssmExe set $CFG.ServiceName AppExit Default $prevAppExit | Out-Null
                    }
                    exit 2
                }
            } finally {
                if ($appExitChanged) {
                    & $CFG.NssmExe set $CFG.ServiceName AppExit Default $prevAppExit | Out-Null
                }
            }
        }
    }

    Write-OK "rollback completato."
    exit 0
}

# ============================================================================
# MAIN
# ============================================================================

if (-not (Test-Admin) -and -not $DryRun) {
    Write-ErrM "Questo script richiede privilegi di Administrator."
    Write-Info "Apri PowerShell come amministratore e rilancia."
    exit 1
}

if ($Rollback) { Invoke-Rollback }

# --- 1. Precondizioni ------------------------------------------------------
Write-Step "Verifica precondizioni"

if (-not (Test-Path $CFG.SrcRoot)) {
    Write-ErrM "cartella staging non trovata: $($CFG.SrcRoot)"
    Write-Info "Carica prima i file da deployare in quella cartella."
    exit 1
}

$aspSrc = Join-Path $CFG.SrcRoot $CFG.AspSrcSub
$pySrc  = Join-Path $CFG.SrcRoot $CFG.PySrcSub

$aspSrcExists = Test-Path $aspSrc
$pySrcExists  = Test-Path $pySrc

if (-not $aspSrcExists -and -not $pySrcExists) {
    Write-WarnM "staging vuoto: nessun asp_nuovo\ ne' python_backend\ trovato in $($CFG.SrcRoot)"
    Write-Info "Niente da fare."
    exit 0
}

if (-not (Test-Path $CFG.AspDest)) { Write-ErrM "AspDest non esiste: $($CFG.AspDest)"; exit 1 }
if (-not (Test-Path $CFG.PyDest))  { Write-ErrM "PyDest non esiste: $($CFG.PyDest)";   exit 1 }
if (-not (Test-Path $CFG.EnvFile)) { Write-ErrM "EnvFile non esiste: $($CFG.EnvFile)"; exit 1 }
if (-not (Test-Path $CFG.NssmExe)) { Write-WarnM "nssm.exe non trovato: $($CFG.NssmExe) (ok se non serve restart)" }

Write-OK "tutti i path di base esistono"

# --- 2. Analisi file da deployare ------------------------------------------
Write-Step "Analisi file in staging"

$aspRelFiles = @()
if ($aspSrcExists) {
    $aspRelFiles = Get-RelativeSourceFiles -Root $aspSrc
    Write-Info "ASP: $($aspRelFiles.Count) file (esclusi $($CFG.IgnorePatterns -join ', '))"
    foreach ($rel in $aspRelFiles) { Write-Info "  - $rel" }
}

$pyRelFiles = @()
$requirementsChanged = $false
if ($pySrcExists) {
    $pyRelFiles = Get-RelativeSourceFiles -Root $pySrc
    Write-Info "Python: $($pyRelFiles.Count) file (esclusi __pycache__, venv, .env)"
    foreach ($rel in $pyRelFiles) { Write-Info "  - $rel" }

    # rileva se requirements.txt e' nel set E se differisce dal corrente
    if ($pyRelFiles -contains "requirements.txt") {
        $newReq = Join-Path $pySrc "requirements.txt"
        $curReq = Join-Path $CFG.PyDest "requirements.txt"
        $newH = Get-FileHashSafe $newReq
        $curH = Get-FileHashSafe $curReq
        if ($newH -ne $curH) { $requirementsChanged = $true }
    }
}

$pythonFilesChanged = ($pyRelFiles | Where-Object { $_ -like "app\*.py" -or $_ -like "app/*.py" -or $_ -like "*.py" }).Count -gt 0

if ($aspRelFiles.Count -eq 0 -and $pyRelFiles.Count -eq 0) {
    Write-WarnM "nessun file riconosciuto da deployare. Controlla la struttura di $($CFG.SrcRoot)"
    exit 0
}

# --- 3. Riepilogo + conferma -----------------------------------------------
Write-Step "Riepilogo piano di deploy"
$reqMsg     = if ($requirementsChanged) { "CAMBIATO -> pip install" } else { "invariato" }
$restartMsg = if ($pythonFilesChanged -and -not $SkipRestart) { "SI" } else { "no" }
$modeMsg    = if ($DryRun) { "DRY RUN (niente viene scritto)" } else { "DEPLOY REALE" }
Write-Host "  ASP files:           $($aspRelFiles.Count)"
Write-Host "  Python files:        $($pyRelFiles.Count)"
Write-Host "  requirements.txt:    $reqMsg"
Write-Host "  restart servizio:    $restartMsg"
Write-Host "  modalita':           $modeMsg"

if ($DryRun) { Write-Info "dry-run attivo: proseguo solo in lettura" }
elseif (-not $AutoConfirm) {
    $ans = Read-Host "`nProcedere? (y/N)"
    if ($ans -notin @("y", "Y", "yes")) { Write-WarnM "annullato."; exit 0 }
}

# --- 4. Prepara backup + legge token ---------------------------------------
$backupTs = New-Timestamp
$backupDir = Join-Path $CFG.BackupRoot $backupTs
if (-not $DryRun) {
    New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
}
Write-Info "backup -> $backupDir"

$internalToken = $null
if ($aspRelFiles.Count -gt 0) {
    try {
        $internalToken = Get-InternalToken -EnvPath $CFG.EnvFile
        Write-OK "INTERNAL_TOKEN letto da .env (lunghezza $($internalToken.Length))"
    } catch {
        Write-ErrM $_.Exception.Message
        exit 1
    }
}

# --- 5. Deploy ASP ---------------------------------------------------------
if ($aspRelFiles.Count -gt 0) {
    Write-Step "Deploy ASP"
    foreach ($rel in $aspRelFiles) {
        $srcFull  = Join-Path $aspSrc $rel
        $destFull = Join-Path $CFG.AspDest $rel
        Write-Act "$rel"

        Copy-FileWithBackup -SrcFull $srcFull -DestFull $destFull `
            -BackupDir $backupDir -BackupLabel "asp" -BackupRelPath $rel

        # Patcha il placeholder del token solo nei file .asp
        if ($rel -like "*.asp") {
            if (-not $DryRun) {
                $content = Get-Content $destFull -Raw
                if ($content -match [regex]::Escape($CFG.TokenPlaceholder)) {
                    $patched = $content -replace [regex]::Escape($CFG.TokenPlaceholder), $internalToken
                    # scrivo preservando encoding UTF-8 senza BOM
                    [System.IO.File]::WriteAllText($destFull, $patched, [System.Text.UTF8Encoding]::new($false))
                    Write-OK "token patchato in $rel"
                } else {
                    Write-Info "placeholder non trovato in $rel (ok se il file non usa il token)"
                }
            } else {
                Write-Act "[dryrun] patch token in $destFull"
            }
        }
    }
}

# --- 6. Deploy Python ------------------------------------------------------
if ($pyRelFiles.Count -gt 0) {
    Write-Step "Deploy Python backend"
    foreach ($rel in $pyRelFiles) {
        $srcFull  = Join-Path $pySrc  $rel
        $destFull = Join-Path $CFG.PyDest $rel
        Write-Act "$rel"
        Copy-FileWithBackup -SrcFull $srcFull -DestFull $destFull `
            -BackupDir $backupDir -BackupLabel "python" -BackupRelPath $rel
    }
}

# --- 7. pip install se requirements cambiato -------------------------------
if ($requirementsChanged) {
    Write-Step "pip install requirements.txt"
    if (-not (Test-Path $CFG.VenvActivate)) {
        Write-ErrM "venv non trovato: $($CFG.VenvActivate)"
        Write-Info "pip install saltato. Esegui manualmente."
    } elseif ($DryRun) {
        Write-Act "[dryrun] & $($CFG.VenvActivate) ; pip install -r requirements.txt"
    } else {
        Push-Location $CFG.PyDest
        try {
            & $CFG.VenvActivate
            & pip install -r "requirements.txt" --disable-pip-version-check
            if ($LASTEXITCODE -ne 0) { throw "pip install fallito (exit $LASTEXITCODE)" }
            Write-OK "dipendenze aggiornate"
        } finally {
            Pop-Location
        }
    }
}

# --- 8. Restart servizio se serve ------------------------------------------
$needRestart = ($pythonFilesChanged -or $requirementsChanged) -and -not $SkipRestart

if ($needRestart) {
    Write-Step "Restart servizio $($CFG.ServiceName)"
    if ($DryRun) {
        Write-Act "[dryrun] nssm restart $($CFG.ServiceName)"
    } else {
        # NSSM 2.24: se AppExit Default = Restart, durante il "restart" il
        # service entra in SERVICE_START_PENDING e lo START control fallisce
        # con "Unexpected status SERVICE_START_PENDING". Disattiviamo il
        # restart automatico solo per la durata del restart intenzionale,
        # poi lo ripristiniamo cosi' il service resta protetto al reboot.
        $prevAppExit = $null
        try {
            $prevAppExit = (& $CFG.NssmExe get $CFG.ServiceName AppExit Default) 2>$null
        } catch { }
        if (-not $prevAppExit) { $prevAppExit = "Restart" }

        $appExitChanged = $false
        if ($prevAppExit -match 'Restart') {
            & $CFG.NssmExe set $CFG.ServiceName AppExit Default Exit | Out-Null
            $appExitChanged = $true
            Write-Info "AppExit temporaneamente impostato a Exit (era $prevAppExit)"
        }

        try {
            & $CFG.NssmExe restart $CFG.ServiceName | Out-Null
            Start-Sleep -Seconds 2
            $status = & $CFG.NssmExe status $CFG.ServiceName
            Write-Info "stato: $status"

            Write-Step "Health check"
            if (Invoke-HealthCheck -Url $CFG.HealthUrl -Retries $CFG.HealthRetries -DelaySec $CFG.HealthRetryDelay) {
                Write-OK "servizio risponde su $($CFG.HealthUrl)"
            } else {
                Write-ErrM "servizio NON risponde dopo restart!"
                Write-Info "Controlla D:\admin_yeppon_python\apps\yeppon_reports\logs\stderr.log"
                Write-Info "Per rollback rapido: .\deploy.ps1 -Rollback"
                if ($appExitChanged) {
                    & $CFG.NssmExe set $CFG.ServiceName AppExit Default $prevAppExit | Out-Null
                    Write-Info "AppExit ripristinato a $prevAppExit"
                }
                exit 2
            }
        } finally {
            if ($appExitChanged) {
                & $CFG.NssmExe set $CFG.ServiceName AppExit Default $prevAppExit | Out-Null
                Write-Info "AppExit ripristinato a $prevAppExit"
            }
        }
    }
} else {
    if ($pythonFilesChanged -and $SkipRestart) {
        Write-WarnM "Python cambiato ma -SkipRestart attivo: ricordati di restart manuale"
    } else {
        Write-Info "Nessun restart necessario (solo ASP cambiato)"
    }
}

# --- 9. Archivia staging + pulizia vecchi backup ---------------------------
if (-not $DryRun) {
    Write-Step "Archivio staging"
    if (-not (Test-Path $CFG.ArchiveRoot)) {
        New-Item -ItemType Directory -Path $CFG.ArchiveRoot -Force | Out-Null
    }
    $archiveDir = Join-Path $CFG.ArchiveRoot $backupTs
    # Sposta il contenuto di SrcRoot in ArchiveRoot\<ts>\, lasciando SrcRoot vuoto ma esistente
    if ($aspSrcExists -or $pySrcExists) {
        Move-Item -Path (Join-Path $CFG.SrcRoot '*') -Destination $archiveDir -Force -ErrorAction SilentlyContinue
        # Move-Item con wildcard vuole dir di destinazione esistente
        if (-not (Test-Path $archiveDir)) { New-Item -ItemType Directory -Path $archiveDir -Force | Out-Null }
        # re-try se il primo Move ha fallito per dir mancante
        Get-ChildItem $CFG.SrcRoot | ForEach-Object {
            Move-Item $_.FullName $archiveDir -Force -ErrorAction SilentlyContinue
        }
        Write-OK "staging -> $archiveDir"
    }

    Remove-OldBackups -Root $CFG.BackupRoot -Keep $CFG.KeepBackups
    Remove-OldBackups -Root $CFG.ArchiveRoot -Keep $CFG.KeepBackups
}

# --- 10. Fine --------------------------------------------------------------
Write-Step "Deploy completato"
Write-OK "tag backup: $backupTs"
Write-Info "rollback rapido: .\deploy.ps1 -Rollback"
Write-Info "rollback specifico: .\deploy.ps1 -Rollback -BackupTag $backupTs"
