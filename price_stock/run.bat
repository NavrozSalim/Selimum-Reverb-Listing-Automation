@echo off
cd /d "%~dp0"
set "PY=python"
if exist "..\.venv\Scripts\python.exe" set "PY=..\.venv\Scripts\python.exe"
if "%~1"=="" (
    "%PY%" update_price_stock.py PriceStock.xlsx
) else (
    "%PY%" update_price_stock.py %*
)
pause
