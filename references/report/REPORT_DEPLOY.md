# Deploy sul Server — Guida Completa

Questo documento descrive come eseguire il deploy dei report sull'ambiente di produzione
Windows Server con IIS + NSSM.

---

## Struttura cartelle sul server

```
D:\
├── www\
│   └── yeppon.it\
│       └── admin\
│           └── asp_nuovo\
│               └── nome_report\          ← file ASP in produzione
│                   ├── nome_report.asp
│                   └── download-nome-xlsx.asp
│
└── admin_yeppon_python\
    ├── apps\
    │   └── yeppon_reports\
    │       └── python_backend\           ← backend Python in produzione
    │           ├── .env                  ← credenziali reali (NON in git)
    │           ├── requirements.txt
    │           ├── venv\
    │           └── app\
    │               ├── main.py
    │               ├── config.py
    │               ├── db.py
    │               ├── routers\
    │               └── services\
    │
    ├── bin\
    │   └── nssm.exe                      ← NSSM service manager
    │
    ├── deploy\
    │   ├── deploy.ps1                    ← script deploy
    │   ├── watcher.ps1                   ← auto-deploy watcher (opzionale)
    │   ├── backups\                      ← backup automatici pre-deploy
    │   └── archived\                     ← staging archiviati post-deploy
    │
    ├── staging\                          ← qui carichi i file nuovi (FileZilla)
    │   ├── asp_nuovo\
    │   │   └── nome_report\
    │   └── python_backend\
    │       └── app\
    │
    └── logs\
        ├── uvicorn_stdout.log
        └── uvicorn_stderr.log
```

---

## Mapping staging → produzione

| Staging (cosa carichi) | Produzione (dove finisce) |
|------------------------|--------------------------|
| `staging\asp_nuovo\*` | `D:\www\yeppon.it\admin\asp_nuovo\*` |
| `staging\python_backend\app\*` | `D:\admin_yeppon_python\apps\yeppon_reports\python_backend\app\*` |
| `staging\python_backend\requirements.txt` | `...\python_backend\requirements.txt` |

**Non caricare mai in staging**: `.env`, `venv\`, `__pycache__\`, file `.pyc`.

---

## Prima installazione (una tantum)

### 1. Crea le cartelle

```powershell
New-Item -ItemType Directory -Force -Path "D:\admin_yeppon_python\apps\yeppon_reports\python_backend"
New-Item -ItemType Directory -Force -Path "D:\admin_yeppon_python\bin"
New-Item -ItemType Directory -Force -Path "D:\admin_yeppon_python\staging\asp_nuovo"
New-Item -ItemType Directory -Force -Path "D:\admin_yeppon_python\staging\python_backend\app"
New-Item -ItemType Directory -Force -Path "D:\admin_yeppon_python\deploy\backups"
New-Item -ItemType Directory -Force -Path "D:\admin_yeppon_python\deploy\archived"
New-Item -ItemType Directory -Force -Path "D:\admin_yeppon_python\logs"
```

### 2. Crea il .env di produzione

Crea manualmente su server (NON via deploy):
```
D:\admin_yeppon_python\apps\yeppon_reports\python_backend\.env
```

Contenuto:
```ini
DB_HOST=<host DB produzione>
DB_PORT=3306
DB_USER=reports_reader
DB_PASSWORD=<password reale>
DB_NAME=yeppon_stats

INTERNAL_TOKEN=<stringa lunga e casuale, es. 64 caratteri hex>

ENABLE_STATIC_TEST_UI=false
LOG_LEVEL=INFO
```

> Genera un token sicuro con: `python -c "import secrets; print(secrets.token_hex(32))"`

### 3. Crea il venv Python

```powershell
cd D:\admin_yeppon_python\apps\yeppon_reports\python_backend
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

### 4. Installa NSSM e registra il servizio

Scarica `nssm.exe` da https://nssm.cc e copialo in `D:\admin_yeppon_python\bin\`.

```powershell
$nssm   = "D:\admin_yeppon_python\bin\nssm.exe"
$python = "D:\admin_yeppon_python\apps\yeppon_reports\python_backend\venv\Scripts\python.exe"
$appDir = "D:\admin_yeppon_python\apps\yeppon_reports\python_backend"

& $nssm install YepponReportsAPI $python
& $nssm set YepponReportsAPI AppParameters "-m uvicorn app.main:app --host 127.0.0.1 --port 8000"
& $nssm set YepponReportsAPI AppDirectory $appDir
& $nssm set YepponReportsAPI AppStdout "D:\admin_yeppon_python\logs\uvicorn_stdout.log"
& $nssm set YepponReportsAPI AppStderr "D:\admin_yeppon_python\logs\uvicorn_stderr.log"
& $nssm set YepponReportsAPI Start SERVICE_AUTO_START
& $nssm start YepponReportsAPI
```

Verifica:
```powershell
Invoke-WebRequest http://127.0.0.1:8000/api/reports/health
# atteso: {"status":"ok"}
```

---

## Workflow deploy normale

### Step 1 — Carica in staging (FileZilla o SFTP)

Carica solo i file modificati mantenendo la struttura:
```
staging\asp_nuovo\nome_report\nome_report.asp
staging\python_backend\app\routers\nome_report.py
staging\python_backend\app\services\nome_query.py
```

### Step 2 — Esegui deploy.ps1

Apri PowerShell come **Amministratore** sul server:

```powershell
cd D:\admin_yeppon_python\deploy

# Preview (niente viene scritto)
.\deploy.ps1 -DryRun

# Deploy interattivo (chiede conferma)
.\deploy.ps1

# Deploy senza conferma
.\deploy.ps1 -AutoConfirm
```

### Cosa fa deploy.ps1 automaticamente

1. Copia i file ASP dallo staging in produzione
2. **Patcha `INTERNAL_TOKEN`**: sostituisce `change_me_long_random_string` con
   il valore reale letto da `.env`
3. Copia i file Python
4. Se `requirements.txt` è cambiato: esegue `pip install -r requirements.txt`
5. Se sono cambiati file Python: riavvia il servizio NSSM
6. Health check: chiama `/api/reports/health` e verifica risposta `{"status":"ok"}`
7. Archivia lo staging sotto `deploy\archived\<timestamp>\`
8. Crea backup dei file sovrascritti in `deploy\backups\<timestamp>\`

---

## Parametri deploy.ps1

```powershell
# Solo preview, niente viene scritto
.\deploy.ps1 -DryRun

# Deploy senza chiedere conferma (utile per automazione)
.\deploy.ps1 -AutoConfirm

# Non riavvia il servizio anche se cambiano file Python
.\deploy.ps1 -SkipRestart

# Ripristina l'ultimo backup
.\deploy.ps1 -Rollback

# Ripristina un backup specifico
.\deploy.ps1 -Rollback -BackupTag "20260421_143022"
```

---

## Rollback manuale

In caso di problemi, ripristina con:

```powershell
.\deploy.ps1 -Rollback
```

Per vedere i backup disponibili:
```powershell
ls D:\admin_yeppon_python\deploy\backups\
```

Per rollback a uno specifico:
```powershell
.\deploy.ps1 -Rollback -BackupTag "20260421_143022"
```

---

## Gestione servizio NSSM

```powershell
$nssm = "D:\admin_yeppon_python\bin\nssm.exe"

# Stato
& $nssm status YepponReportsAPI

# Riavvio manuale
& $nssm restart YepponReportsAPI

# Stop / Start
& $nssm stop YepponReportsAPI
& $nssm start YepponReportsAPI

# Log real-time
Get-Content "D:\admin_yeppon_python\logs\uvicorn_stderr.log" -Tail 50 -Wait
```

---

## Auto-deploy con watcher (opzionale)

`watcher.ps1` monitora la cartella staging e lancia `deploy.ps1 -AutoConfirm`
automaticamente quando rileva nuovi file.

Installa come secondo servizio NSSM:

```powershell
$nssm = "D:\admin_yeppon_python\bin\nssm.exe"
$ps   = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"

& $nssm install YepponReportsWatcher $ps
& $nssm set YepponReportsWatcher AppParameters "-ExecutionPolicy Bypass -File D:\admin_yeppon_python\deploy\watcher.ps1"
& $nssm set YepponReportsWatcher AppDirectory "D:\admin_yeppon_python\deploy"
& $nssm set YepponReportsWatcher AppStdout "D:\admin_yeppon_python\logs\watcher_stdout.log"
& $nssm set YepponReportsWatcher AppStderr "D:\admin_yeppon_python\logs\watcher_stderr.log"
& $nssm set YepponReportsWatcher Start SERVICE_AUTO_START
& $nssm start YepponReportsWatcher
```

Con il watcher attivo il workflow diventa:
1. Carica file in staging (FileZilla)
2. Il watcher rileva la modifica e lancia il deploy automaticamente
3. Controlla i log per verificare l'esito

---

## Checklist deploy nuovo report

- [ ] Crea `.env` sul server con `INTERNAL_TOKEN` reale
- [ ] Carica `staging\asp_nuovo\<nome>\*.asp` (con placeholder token)
- [ ] Carica `staging\python_backend\app\routers\<nome>.py`
- [ ] Carica `staging\python_backend\app\services\<nome>_query.py`
- [ ] Se nuove dipendenze: carica anche `requirements.txt`
- [ ] Esegui `.\deploy.ps1 -DryRun` per verificare cosa verrà fatto
- [ ] Esegui `.\deploy.ps1 -AutoConfirm`
- [ ] Verifica health: `Invoke-WebRequest http://127.0.0.1:8000/api/reports/health`
- [ ] Testa la pagina ASP dal browser admin
