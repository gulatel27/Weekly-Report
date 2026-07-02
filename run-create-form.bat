@echo off
setlocal
cd /d "%~dp0"

set PYTHON_EXE=C:\Users\JacksonL\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe

if exist "%PYTHON_EXE%" (
  "%PYTHON_EXE%" scripts\create-weekly-report-form.py
) else (
  python scripts\create-weekly-report-form.py
)

pause
