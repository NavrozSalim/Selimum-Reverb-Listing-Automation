@echo off
cd /d "%~dp0"
set "PY=python"
if exist "..\.venv\Scripts\python.exe" set "PY=..\.venv\Scripts\python.exe"
if "%~1"=="" (
    "%PY%" sku_replace.py SkuReplace.xlsx
) else (
    "%PY%" sku_replace.py %*
)
pause
