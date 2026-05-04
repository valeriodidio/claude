@echo off
REM Ferma il container MySQL locale (mantiene il volume).
pushd "%~dp0"
docker compose -f docker-compose.yml stop
popd
