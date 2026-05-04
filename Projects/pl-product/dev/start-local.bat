@echo off
REM ─── Avvio one-click dell'API locale P&L Prodotti ───────────────────────────
REM Sequenza:
REM   1. docker compose up -d (se il container non esiste) o start
REM   2. attesa MySQL healthy
REM   3. attivazione venv
REM   4. seed da prod (se dev\.env.prod esiste e flag non disabilitato)
REM   5. uvicorn --reload
REM ────────────────────────────────────────────────────────────────────────────
setlocal EnableDelayedExpansion

set "ROOT=%~dp0.."
pushd "%ROOT%"

REM ── Argomenti opzionali ────────────────────────────────────────────────────
set "SKIP_SEED="
set "SNAPSHOTS=1"
set "RESI_ROWS=0"
set "NO_TRUNCATE="

:parse_args
if "%~1"=="" goto args_done
if "%~1"=="--skip-seed"   set "SKIP_SEED=1" & shift & goto parse_args
if "%~1"=="--snapshots"   set "SNAPSHOTS=%~2" & shift & shift & goto parse_args
if "%~1"=="--resi-rows"   set "RESI_ROWS=%~2" & shift & shift & goto parse_args
if "%~1"=="--no-truncate" set "NO_TRUNCATE=--no-truncate" & shift & goto parse_args
shift
goto parse_args
:args_done

echo ============================================================
echo   P^&L Prodotti — start-local
echo ============================================================

REM ── 1. Docker ──────────────────────────────────────────────────────────────
echo [1/5] Docker MySQL...
docker compose -f dev\docker-compose.yml ps -q mysql >nul 2>&1
if errorlevel 1 goto compose_up
for /f %%i in ('docker compose -f dev\docker-compose.yml ps -q mysql') do set "MYSQL_CID=%%i"
if "%MYSQL_CID%"=="" goto compose_up
docker compose -f dev\docker-compose.yml start
goto wait_mysql

:compose_up
docker compose -f dev\docker-compose.yml up -d

:wait_mysql
echo [2/5] Attesa MySQL healthy...
set /a WAIT=0
:wait_loop
docker compose -f dev\docker-compose.yml exec -T mysql mysqladmin ping -h 127.0.0.1 -uroot -proot_dev >nul 2>&1
if not errorlevel 1 goto mysql_ok
set /a WAIT+=1
if %WAIT% GEQ 30 goto wait_timeout
timeout /t 2 /nobreak >nul
goto wait_loop
:wait_timeout
echo ERRORE: MySQL non disponibile dopo 30 tentativi
popd & exit /b 1

:mysql_ok
echo    MySQL pronto.

REM ── 3. venv ────────────────────────────────────────────────────────────────
echo [3/5] Virtualenv Python...
if not exist venv (
    python -m venv venv
)
call venv\Scripts\activate.bat
pip install -q -r requirements.txt

REM ── 4. Seed da prod ────────────────────────────────────────────────────────
if defined SKIP_SEED goto skip_seed
if not exist dev\.env.prod (
    echo [4/5] Seed saltato: dev\.env.prod non presente.
    goto run_uvicorn
)
echo [4/5] Seed dati da produzione (snapshots=%SNAPSHOTS%, resi=%RESI_ROWS%)...
python dev\seed_from_prod.py --snapshots %SNAPSHOTS% --resi-rows %RESI_ROWS% %NO_TRUNCATE%
goto run_uvicorn

:skip_seed
echo [4/5] Seed disabilitato (--skip-seed)

:run_uvicorn
echo [5/5] Avvio uvicorn su http://127.0.0.1:8000 ...
echo    Test page:  http://127.0.0.1:8000/static/test_pl_prodotti.html
echo    Swagger:    http://127.0.0.1:8000/api/reports/docs
python -m uvicorn app.main:app --host 127.0.0.1 --port 8000 --reload

popd
endlocal
