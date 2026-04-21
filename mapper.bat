@echo off
cd /d "%~dp0"

uv run mapper.py result
exit /b %ERRORLEVEL%
