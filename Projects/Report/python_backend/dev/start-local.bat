@echo off
REM Avvia l'ambiente di test locale:
REM   1. Fa partire il container MySQL
REM   2. Attende che MySQL sia healthy
REM   3. Importa le ultime 20.000 righe dal DB di produzione (se .env.prod esiste)
REM   4. Attiva il venv Python
REM   5. Avvia uvicorn con --reload
REM
REM Argomenti opzionali:
REM   --skip-seed     Salta l'importazione dati (usa i dati gia' presenti in locale)
REM   --rows N        Importa N righe invece di 20000 (es: --rows 5000)
REM   --no-truncate   Aggiunge righe senza svuotare prima la tabella locale
REM
REM Questo file va lanciato con doppio click oppure dal prompt,
REM da dentro la cartella python_backend\dev\ o da qualsiasi altra
REM (usa path relativi alla propria posizione).

setlocal EnableDelayedExpansion

REM --- Parsing argomenti ---
set SKIP_SEED=0
set SEED_ROWS=20000
set SEED_EXTRA=

:parse_args
if "%~1"=="" goto args_done
if /I "%~1"=="--skip-seed"    (set SKIP_SEED=1 & shift & goto parse_args)
if /I "%~1"=="--no-truncate"  (set SEED_EXTRA=%SEED_EXTRA% --no-truncate & shift & goto parse_args)
if /I "%~1"=="--rows"         (set SEED_ROWS=%~2 & shift & shift & goto parse_args)
shift & goto parse_args
:args_done

REM portati nella cartella dove si trova questo .bat
cd /d "%~dp0"

echo.
echo ==== [1/4] Avvio container MySQL (yeppon_reports_mysql) ====
docker compose start 1>nul 2>nul
if errorlevel 1 (
    echo    Container non ancora creato, lo creo con "up -d"...
    docker compose up -d
    if errorlevel 1 (
        echo.
        echo ERRORE: docker compose non e' partito. Docker Desktop e' avviato?
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
        echo ERRORE: MySQL non healthy dopo 60s. Controlla con:
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
    echo ERRORE: venv non trovato in python_backend\venv
    echo Creane uno con:
    echo    python -m venv venv
    echo    venv\Scripts\activate
    echo    pip install -r requirements.txt
    pause
    exit /b 1
)
call venv\Scripts\activate.bat

echo.
echo ==== [3.5/4] Import dati da produzione ====
if "%SKIP_SEED%"=="1" goto seed_skip
if not exist "dev\.env.prod" goto seed_no_env
echo    Importo le ultime %SEED_ROWS% righe da produzione...
python dev\seed_from_prod.py --rows %SEED_ROWS% %SEED_EXTRA%
if errorlevel 1 (
    echo.
    echo    ATTENZIONE: il seed ha restituito un errore.
    echo    Continuo comunque ^(i dati locali potrebbero essere incompleti^).
    echo    Premi un tasto per continuare o chiudi la finestra per annullare.
    pause
)
goto seed_done

:seed_skip
echo    [--skip-seed] Seed saltato, uso dati gia' presenti in locale.
goto seed_done

:seed_no_env
echo    [SKIP] dev\.env.prod non trovato.
echo    Per importare dati reali copia dev\.env.prod.example in dev\.env.prod
echo    e compila le credenziali di produzione.

:seed_done

echo.
echo ==== [4/4] Avvio FastAPI (uvicorn) ====
echo.
echo    API:   http://127.0.0.1:8001/api/reports/health
echo    UI:    http://127.0.0.1:8001/static/test_turnover.html
echo    Docs:  http://127.0.0.1:8001/api/reports/docs
echo    Ctrl+C per fermare uvicorn.
echo.

python -m uvicorn app.main:app --host 127.0.0.1 --port 8001 --reload

endlocal
