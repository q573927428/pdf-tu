@echo off
setlocal

rem Double-click launcher for the PDF Catalog local server.
set "ROOT=%~dp0"
set "CONFIG=%ROOT%config.yaml"
set "PAGE=http://127.0.0.1:8765/pdf_catalog.html"
set "PYTHON=%ROOT%.venv\Scripts\python.exe"
set "PYTHONPATH=%ROOT%src"

if not exist "%PYTHON%" goto missing_python
if not exist "%CONFIG%" goto missing_config

rem Close stale PDF Catalog server processes so an old version cannot handle requests.
powershell -NoProfile -ExecutionPolicy Bypass -Command "$p=Get-NetTCPConnection -LocalPort 8765 -State Listen -ErrorAction SilentlyContinue | Select-Object -ExpandProperty OwningProcess -Unique; foreach($id in $p){ Stop-Process -Id $id -Force -ErrorAction SilentlyContinue }" >nul 2>&1
timeout /t 1 /nobreak >nul

start "PDF Catalog Server" /min "%PYTHON%" -m pdf_catalog.cli serve --config "%CONFIG%"

echo Waiting for the local server...
timeout /t 3 /nobreak >nul
start "" "%PAGE%"
echo Browser opened: %PAGE%
exit /b 0

:missing_python
echo Python virtual environment not found: "%PYTHON%"
echo Create .venv and install dependencies first.
pause
exit /b 1

:missing_config
echo Configuration file not found: "%CONFIG%"
pause
exit /b 1
