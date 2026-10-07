@echo off
setlocal
cd /d "%~dp0"

set "PYTHON_EXE=%~dp0.venv\Scripts\python.exe"

if not exist "node_modules\playwright" (
  echo Installing dependencies...
  call npm install
  if errorlevel 1 exit /b %errorlevel%
)

call npm run download
if errorlevel 1 exit /b %errorlevel%

if exist "%PYTHON_EXE%" (
  "%PYTHON_EXE%" -X utf8 scripts\build-weekly-report.py --latest-download
) else (
  python scripts\build-weekly-report.py --latest-download
)

pause
