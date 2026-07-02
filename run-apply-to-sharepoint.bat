@echo off
setlocal
cd /d "%~dp0"

set PYTHON_EXE=C:\Users\JacksonL\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe

if "%~1"=="" (
  echo Usage: run-apply-to-sharepoint.bat "C:\path\to\sharepoint-weekly-report.xlsx" ["C:\path\to\generated-report.xlsx"]
  echo.
  echo If the second path is omitted, the latest outputs\락플레이스-DS2_주간보고_*.xlsx file is used.
  exit /b 1
)

set TARGET_PATH=%~1
set SOURCE_ARG=
if not "%~2"=="" set SOURCE_ARG=--source "%~2"

if exist "%PYTHON_EXE%" (
  "%PYTHON_EXE%" -X utf8 scripts\apply-to-sharepoint-workbook.py --target "%TARGET_PATH%" %SOURCE_ARG%
) else (
  python scripts\apply-to-sharepoint-workbook.py --target "%TARGET_PATH%" %SOURCE_ARG%
)

pause
