@echo off
REM Ferma il container MySQL senza cancellare i dati.
REM (Uvicorn va fermato manualmente con Ctrl+C nella sua finestra.)

setlocal
cd /d "%~dp0"

echo Fermo il container MySQL...
docker compose stop
if errorlevel 1 (
    echo ERRORE durante lo stop del container.
    pause
    exit /b 1
)

echo.
echo OK. I dati sono conservati nel volume Docker.
echo Per ripartire: start-local.bat
pause
endlocal
