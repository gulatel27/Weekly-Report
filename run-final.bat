@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set PYTHON_EXE=C:\Users\JacksonL\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe
set DATE_ENV_FILE=logs\report-dates.cmd

if not exist "logs" mkdir logs

if exist "%PYTHON_EXE%" (
  "%PYTHON_EXE%" -X utf8 scripts\read-report-dates.py --env-file "%DATE_ENV_FILE%"
) else (
  python -X utf8 scripts\read-report-dates.py --env-file "%DATE_ENV_FILE%"
)
if errorlevel 1 exit /b %errorlevel%

call "%DATE_ENV_FILE%"
echo.
echo Report range: %REPORT_START_DATE% ~ %REPORT_END_DATE%
echo Summary mode: %REPORT_SUMMARY_MODE%
echo.

if not exist "node_modules\playwright" (
  echo Installing dependencies...
  call npm install
  if errorlevel 1 exit /b %errorlevel%
)

echo [1/3] Rockpaper weekly report download...
set HEADLESS=1
call npm run download
if errorlevel 1 exit /b %errorlevel%

echo [2/3] Build DS-2 weekly report...
if exist "%PYTHON_EXE%" (
  "%PYTHON_EXE%" -X utf8 scripts\build-weekly-report.py --latest-download
) else (
  python -X utf8 scripts\build-weekly-report.py --latest-download
)
if errorlevel 1 exit /b %errorlevel%

if "%~1"=="" (
  echo [3/3] SharePoint target was not provided. The final report is in the outputs folder.
  echo.
  echo To create a SharePoint-ready workbook, run:
  echo run-final.bat "C:\path\to\sharepoint-weekly-report.xlsx"
  pause
  exit /b 0
)

echo [3/3] Create SharePoint-ready workbook...
if exist "%PYTHON_EXE%" (
  "%PYTHON_EXE%" -X utf8 scripts\apply-to-sharepoint-workbook.py --target "%~1"
) else (
  python -X utf8 scripts\apply-to-sharepoint-workbook.py --target "%~1"
)

pause
