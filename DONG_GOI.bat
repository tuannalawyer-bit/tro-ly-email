@echo off
rem Dung bo cai mot tep: dist\CaiTroLyEmail.exe
rem Can PyInstaller: uv pip install --native-tls -r requirements-build.txt
cd /d "%~dp0"
".\.venv\Scripts\python.exe" dong_goi.py
pause
