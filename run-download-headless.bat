@echo off
setlocal
cd /d "%~dp0"

if not exist "node_modules\playwright" (
  echo Installing dependencies...
  call npm install
  if errorlevel 1 exit /b %errorlevel%
)

set HEADLESS=1
call npm run download
pause
