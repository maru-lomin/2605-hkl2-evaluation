@echo off
cd /d "%~dp0"

if not defined OUT_DIR set "OUT_DIR=."

uv run result_to_excel.py .\dataset --out-dir result
exit /b %ERRORLEVEL%
