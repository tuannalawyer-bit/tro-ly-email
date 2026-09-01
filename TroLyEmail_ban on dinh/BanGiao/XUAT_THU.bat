@echo off
cd /d "%~dp0"
echo Mo Outlook truoc khi chay cong cu nay.
echo Them tham so --nhanh neu chi muon quet thu muc Sent Items.
echo.
".\.venv\Scripts\python.exe" xuat_thu_da_gui.py %*
echo.
pause
