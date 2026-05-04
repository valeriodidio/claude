@echo off
REM ATTENZIONE: distrugge il volume MySQL e riparte da zero.
REM Dovrai ri-seedare i dati dopo.

setlocal
cd /d "%~dp0"

echo.
echo ATTENZIONE: questo cancella TUTTI i dati del MySQL di test.
set /p _ok=Continuo? (s/N):
if /I not "%_ok%"=="s" (
    echo Annullato.
    exit /b 0
)

echo.
echo Distruggo container e volume...
docker compose down -v
if errorlevel 1 (
    echo ERRORE durante il down.
    pause
    exit /b 1
)

echo.
echo Ricreo container (esegue init.sql)...
docker compose up -d
if errorlevel 1 (
    echo ERRORE.
    pause
    exit /b 1
)

echo.
echo OK. MySQL vuoto e pronto.
echo Ora popola con:
echo    cd ..
echo    venv\Scripts\activate
echo    python dev\scripts\seed_test_data.py --rows 20000 --days 90
pause
endlocal
