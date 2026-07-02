@echo off
setlocal
cd /d "%~dp0"

set HEADLESS=1
call run-weekly-report.bat
