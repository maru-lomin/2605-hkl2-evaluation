@echo off
cd /d "%~dp0"

uv run python evaluation.py result_mapped ground_truth evaluation
exit /b %ERRORLEVEL%
