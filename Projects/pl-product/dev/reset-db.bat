@echo off
REM ATTENZIONE: distrugge il volume MySQL e riparte da zero (riesegue init.sql).
REM Usare quando si modifica init.sql e si vuole ripartire da schema pulito.

setlocal
pushd "%~dp0"

echo ATTENZIONE: questo cancella TUTTI i dati del MySQL locale (volume Docker).
set /p _ok=Continuo? (s/N):
if /I not "%_ok%"=="s" (
    echo Annullato.
    popd & exit /b 0
)

echo Fermo e rimuovo il container + volume...
docker compose down -v
if errorlevel 1 (
    echo ERRORE: docker compose down fallito.
    popd & exit /b 1
)

echo Ricreo il container...
docker compose up -d
if errorlevel 1 (
    echo ERRORE: docker compose up fallito. Docker Desktop e' avviato?
    popd & exit /b 1
)

echo.
echo OK. MySQL vuoto (init.sql rieseguito al primo avvio).
echo Lancia start-local.bat per seedare i dati e avviare l'API.
pause
popd
endlocal
