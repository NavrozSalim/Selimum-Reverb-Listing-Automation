@echo off
cd /d "%~dp0"
set "PY=python"
if exist "..\.venv\Scripts\python.exe" set "PY=..\.venv\Scripts\python.exe"
if "%~1"=="" (
    "%PY%" create_listing.py Create.xlsx
) else (
    "%PY%" create_listing.py %*
)
pause
