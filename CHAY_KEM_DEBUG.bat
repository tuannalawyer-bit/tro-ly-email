@echo off
cd /d "%~dp0"
set DEBUG=1
start "" ".\.venv\Scripts\python.exe" main.py
