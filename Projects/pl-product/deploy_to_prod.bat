@echo off
:: ============================================================
::  deploy_to_prod.bat  —  P&L Prodotti
::  Carica i file nella staging zone del server.
::  Il deploy.ps1 sul server pensa al resto (token patch, restart, backup).
::
::  USO:
::    1. Modifica STAGING con il percorso UNC della staging zone
::    2. Lancia questo bat
::    3. Sul server: cd D:\admin_yeppon_python\apps\deploy && .\deploy.ps1
::       (oppure aspetta il watcher automatico ~30 secondi)
:: ============================================================

set STAGING=\\SERVER\admin_yeppon_python\staging
set LOCAL=C:\Users\ValerioDidio\Documents\Claude\Projects\pl-product

echo.
echo ============================================================
echo  DEPLOY P^&L Prodotti ^-^> staging zone
echo ============================================================

:: ── Crea le cartelle di staging se non esistono ──────────────
echo.
echo [1/3] Creo struttura staging...
mkdir "%STAGING%\asp_nuovo\pl_prodotti"          2>nul
mkdir "%STAGING%\python_backend\app\routers"     2>nul
mkdir "%STAGING%\python_backend\app\services"    2>nul

:: ── File ASP ─────────────────────────────────────────────────
echo.
echo [2/3] Copio file ASP in staging\asp_nuovo\pl_prodotti\ ...
copy /Y "%LOCAL%\asp\pl_prodotti\index.asp"                      "%STAGING%\asp_nuovo\pl_prodotti\index.asp"
copy /Y "%LOCAL%\asp\pl_prodotti\download-pl-prodotti.xlsx.asp"  "%STAGING%\asp_nuovo\pl_prodotti\download-pl-prodotti.xlsx.asp"

:: ── File Python ───────────────────────────────────────────────
echo.
echo [3/3] Copio file Python in staging\python_backend\ ...
copy /Y "%LOCAL%\app\routers\pl_prodotti.py"         "%STAGING%\python_backend\app\routers\pl_prodotti.py"
copy /Y "%LOCAL%\app\services\pl_prodotti_query.py"  "%STAGING%\python_backend\app\services\pl_prodotti_query.py"

:: ── Riepilogo ─────────────────────────────────────────────────
echo.
echo ============================================================
echo  File caricati in staging. Ora sul server:
echo.
echo    cd D:\admin_yeppon_python\apps\deploy
echo    .\deploy.ps1
echo.
echo  Lo script:
echo    - patcha automaticamente INTERNAL_TOKEN negli ASP
echo    - copia Python backend
echo    - riavvia YepponReportsAPI
echo    - fa health check
echo    - salva backup per rollback
echo.
echo  NOTA: se la cartella admin IIS non esiste ancora, creala prima:
echo    D:\www\yeppon.it\admin\pl_prodotti\
echo ============================================================
echo.
pause
