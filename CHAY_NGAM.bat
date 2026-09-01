@echo off
rem Chay ngam o khay he thong (khong co cua so den, khong co nut taskbar).
rem Bam vao icon o goc phai thanh tac vu de mo ung dung len.
cd /d "%~dp0"
start "" ".\.venv\Scripts\pythonw.exe" main.py --tray
