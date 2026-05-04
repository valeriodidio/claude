# Ambiente di Test Locale — Gestione Docker + Seed

Questo documento descrive come configurare e gestire l'ambiente di sviluppo locale
per i report. L'ambiente usa Docker per MySQL e uvicorn per il backend FastAPI.

---

## Prerequisiti

- Docker Desktop installato e in esecuzione
- Python 3.10+ installato
- `venv` creato in `python_backend/venv/`

Crea il venv la prima volta:
```bat
cd python_backend
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

---

## start-local.bat

Lancia tutto con un doppio click: MySQL, seed dati, uvicorn.

**Struttura con `goto` — non usare `if/else` annidati con `EnableDelayedExpansion`** (causa errore `. non atteso` in Windows):

```bat
@echo off
REM Avvia: MySQL Docker → seed dati → uvicorn
REM Argomenti:
REM   --skip-seed     Salta l'import dati
REM   --rows N        Importa N righe (default 20000)
REM   --no-truncate   Aggiunge senza svuotare

setlocal EnableDelayedExpansion

set SKIP_SEED=0
set SEED_ROWS=20000
set SEED_EXTRA=

:parse_args
if "%~1"=="" goto args_done
if /I "%~1"=="--skip-seed"   (set SKIP_SEED=1       & shift & goto parse_args)
if /I "%~1"=="--no-truncate" (set SEED_EXTRA=%SEED_EXTRA% --no-truncate & shift & goto parse_args)
if /I "%~1"=="--rows"        (set SEED_ROWS=%~2      & shift & shift & goto parse_args)
shift & goto parse_args
:args_done

cd /d "%~dp0"

echo.
echo ==== [1/4] Avvio container MySQL ====
docker compose start 1>nul 2>nul
if errorlevel 1 (
    echo    Container non esistente, lo creo...
    docker compose up -d
    if errorlevel 1 (
        echo ERRORE: docker compose non avviato. Docker Desktop e' attivo?
        pause
        exit /b 1
    )
)

echo.
echo ==== [2/4] Attendo MySQL healthy (max 60s) ====
set /a _tries=0
:wait_db
docker inspect -f "{{.State.Health.Status}}" yeppon_reports_mysql 2>nul | findstr /C:"healthy" 1>nul
if errorlevel 1 (
    set /a _tries+=1
    if %_tries% GEQ 60 (
        echo ERRORE: MySQL non healthy dopo 60s.
        echo    docker compose logs mysql
        pause
        exit /b 1
    )
    timeout /t 1 /nobreak 1>nul
    goto wait_db
)
echo    MySQL pronto.

echo.
echo ==== [3/4] Attiva venv Python ====
cd ..
if not exist venv\Scripts\activate.bat (
    echo ERRORE: venv non trovato. Creane uno con:
    echo    python -m venv venv ^&^& venv\Scripts\activate ^&^& pip install -r requirements.txt
    pause
    exit /b 1
)
call venv\Scripts\activate.bat

echo.
echo ==== [3.5/4] Seed dati da produzione ====
if "%SKIP_SEED%"=="1" goto seed_skip
if not exist "dev\.env.prod" goto seed_no_env
echo    Importo le ultime %SEED_ROWS% righe...
python dev\seed_from_prod.py --rows %SEED_ROWS% %SEED_EXTRA%
if errorlevel 1 (
    echo    ATTENZIONE: seed con errore. Continuo comunque.
    pause
)
goto seed_done

:seed_skip
echo    [--skip-seed] Seed saltato.
goto seed_done

:seed_no_env
echo    [SKIP] dev\.env.prod non trovato. Copia dev\.env.prod.example e compila.

:seed_done
echo.
echo ==== [4/4] Avvio FastAPI (uvicorn) ====
echo    API:  http://127.0.0.1:8001/api/reports/health
echo    UI:   http://127.0.0.1:8001/static/test_<nome>.html
echo    Docs: http://127.0.0.1:8001/api/reports/docs
echo    Ctrl+C per fermare.
echo.
python -m uvicorn app.main:app --host 127.0.0.1 --port 8001 --reload

endlocal
```

> **Nota porta**: usa `8001` in locale per non collidere con il servizio di produzione
> che gira su `8000`.

---

## stop-local.bat

```bat
@echo off
cd /d "%~dp0"
echo Fermo il container MySQL...
docker compose stop
echo OK.
pause
```

---

## reset-db.bat

**ATTENZIONE**: distrugge TUTTI i dati locali e riparte da zero.
`init.sql` verrà rieseguito al prossimo `docker compose up -d`.

```bat
@echo off
setlocal
cd /d "%~dp0"

echo ATTENZIONE: questo cancella TUTTI i dati del MySQL di test.
set /p _ok=Continuo? (s/N):
if /I not "%_ok%"=="s" (
    echo Annullato.
    exit /b 0
)

echo Distruggo container e volume...
docker compose down -v
if errorlevel 1 ( echo ERRORE. & pause & exit /b 1 )

echo Ricreo container...
docker compose up -d
if errorlevel 1 ( echo ERRORE. & pause & exit /b 1 )

echo OK. MySQL vuoto e pronto.
echo Ora popola con:
echo    cd .. ^&^& venv\Scripts\activate
echo    python dev\seed_from_prod.py --rows 20000
pause
endlocal
```

---

## docker-compose.yml

```yaml
version: "3.9"
services:
  mysql:
    image: mysql:8.0
    container_name: yeppon_reports_mysql
    environment:
      MYSQL_ROOT_PASSWORD: root_dev
      MYSQL_DATABASE: yeppon_stats
    ports:
      - "3307:3306"   # 3307 locale → 3306 container (non collidere con MySQL di sistema)
    volumes:
      - mysql_data:/var/lib/mysql
      - ./init.sql:/docker-entrypoint-initdb.d/init.sql
    healthcheck:
      test: ["CMD", "mysqladmin", "ping", "-h", "localhost", "-uroot", "-proot_dev"]
      interval: 2s
      timeout: 5s
      retries: 30

volumes:
  mysql_data:
```

> **Importante**: il volume `mysql_data` persiste tra riavvii. `init.sql` viene
> eseguito **solo se il volume non esiste**. Se devi aggiungere tabelle o grant
> a un volume già esistente, usa `docker exec` o fai `reset-db.bat`.

---

## seed_from_prod.py — struttura

```python
#!/usr/bin/env python3
"""
Copia le ultime N righe dal DB di produzione al MySQL locale.
Richiede dev/.env.prod con le credenziali di produzione.
"""
import argparse
import os
import sys
from pathlib import Path
from dotenv import load_dotenv
import sqlalchemy as sa

# Carica credenziali prod
load_dotenv(Path(__file__).parent / ".env.prod")

PROD_URL = (
    f"mysql+pymysql://{os.environ['PROD_DB_USER']}:{os.environ['PROD_DB_PASSWORD']}"
    f"@{os.environ['PROD_DB_HOST']}:{os.environ.get('PROD_DB_PORT', 3306)}"
    f"/{os.environ['PROD_DB_NAME']}?charset=utf8mb4"
)
LOCAL_URL = "mysql+pymysql://root:root_dev@127.0.0.1:3307/yeppon_stats?charset=utf8mb4"

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--rows", type=int, default=20000)
    parser.add_argument("--no-truncate", action="store_true")
    args = parser.parse_args()

    prod_engine  = sa.create_engine(PROD_URL)
    local_engine = sa.create_engine(LOCAL_URL)

    with prod_engine.connect() as prod_conn:
        rows = prod_conn.execute(
            sa.text("SELECT * FROM nome_tabella ORDER BY dataordine DESC LIMIT :n"),
            {"n": args.rows}
        ).fetchall()

    with local_engine.begin() as local_conn:
        if not args.no_truncate:
            local_conn.execute(sa.text("TRUNCATE TABLE nome_tabella"))
        # batch insert
        if rows:
            local_conn.execute(
                sa.text("INSERT INTO nome_tabella (...) VALUES (...)"),
                [dict(r._mapping) for r in rows]
            )

    print(f"Importate {len(rows)} righe.")

if __name__ == "__main__":
    main()
```

---

## .env.prod.example

```ini
# Credenziali DB di produzione per il seed locale.
# Copia in .env.prod e compila. NON committare .env.prod.

PROD_DB_HOST=
PROD_DB_PORT=3306
PROD_DB_USER=
PROD_DB_PASSWORD=
PROD_DB_NAME=yeppon_stats
```

---

## Gestione permessi MySQL

Quando si aggiungono nuove tabelle **dopo** la creazione del volume:

```bash
# Entra nel container
docker exec -it yeppon_reports_mysql mysql -uroot -proot_dev

# Concedi i permessi sull'utente dell'app
GRANT SELECT ON yeppon_stats.nuova_tabella TO 'reports_reader'@'%';
GRANT SELECT ON smart2.ordini_cliente      TO 'reports_reader'@'%';
FLUSH PRIVILEGES;

# Verifica
SHOW GRANTS FOR 'reports_reader'@'%';
```

---

## Errori comuni

| Errore | Causa | Soluzione |
|--------|-------|-----------|
| `. non atteso` in .bat | `if/else` annidati con `EnableDelayedExpansion` | Usa `goto` label (vedi template sopra) |
| `Access denied for user 'reports_reader'` | Password sbagliata nel `.env` o utente inesistente nel volume | Verifica `.env`, ricrea utente via `docker exec` |
| `SELECT command denied for table 'xxx'` | GRANT mancante su quella tabella | `GRANT SELECT ON db.tabella TO 'reports_reader'@'%'; FLUSH PRIVILEGES;` |
| `Table 'db.tabella' doesn't exist` (in GRANT) | Tabella non ancora creata nel volume | Crea la tabella via `docker exec` prima di fare GRANT |
| `Connection refused 3307` | Container non avviato | `docker compose up -d` |
| `init.sql non rieseguito` | Volume già esistente | Fai `reset-db.bat` per distruggere e ricreare il volume |
| Uvicorn non parte | venv non attivato o dipendenze mancanti | `call venv\Scripts\activate.bat` + `pip install -r requirements.txt` |
