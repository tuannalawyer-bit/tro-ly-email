@echo off
cd /d "%~dp0"
".\.venv\Scripts\python.exe" main.py
if errorlevel 1 (
    echo.
    echo Ung dung da dung do loi. Vui long chup lai thong bao phia tren.
    pause
)
