@echo off
setlocal
cd /d "%~dp0"

set REPORT_START_DATE=2026-06-22
set REPORT_END_DATE=2026-06-28

call run-weekly-report.bat
