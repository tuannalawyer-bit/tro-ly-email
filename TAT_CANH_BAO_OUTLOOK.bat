@echo off
chcp 65001 >nul
:: Tự động yêu cầu quyền Administrator (UAC) nếu chưa có
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Đang yêu cầu quyền Quản trị viên (Administrator)...
    powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

echo ==================================================================
echo   TẮT CẢNH BÁO TRUY CẬP EMAIL (PROGRAMMATIC ACCESS) TRONG OUTLOOK
echo ==================================================================
echo.

:: 1. Cấu hình HKLM (cấp toàn máy tính)
reg add "HKLM\SOFTWARE\Microsoft\Office\16.0\Outlook\Security" /v "ObjectModelGuard" /t REG_DWORD /d 2 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Microsoft\Office\16.0\Outlook\Security" /v "PromptOOMAddressBookAccess" /t REG_DWORD /d 2 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Microsoft\Office\16.0\Outlook\Security" /v "PromptOOMAddressInformationAccess" /t REG_DWORD /d 2 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Microsoft\Office\16.0\Outlook\Security" /v "PromptOOMSend" /t REG_DWORD /d 2 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Microsoft\Office\16.0\Outlook\Security" /v "PromptOOMSaveAs" /t REG_DWORD /d 2 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Microsoft\Office\16.0\Outlook\Security" /v "PromptSimpleMAPISend" /t REG_DWORD /d 2 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Microsoft\Office\16.0\Outlook\Security" /v "PromptSimpleMAPINameResolve" /t REG_DWORD /d 2 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Microsoft\Office\16.0\Outlook\Security" /v "PromptSimpleMAPIOpenMessage" /t REG_DWORD /d 2 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Microsoft\Office\16.0\Outlook\Security" /v "AdminSecurityMode" /t REG_DWORD /d 3 /f >nul 2>&1

:: 2. Cấu hình HKLM Policies (chính sách bảo mật)
reg add "HKLM\SOFTWARE\Policies\Microsoft\Office\16.0\Outlook\Security" /v "ObjectModelGuard" /t REG_DWORD /d 2 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Policies\Microsoft\Office\16.0\Outlook\Security" /v "PromptOOMAddressBookAccess" /t REG_DWORD /d 2 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Policies\Microsoft\Office\16.0\Outlook\Security" /v "PromptOOMAddressInformationAccess" /t REG_DWORD /d 2 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Policies\Microsoft\Office\16.0\Outlook\Security" /v "PromptOOMSend" /t REG_DWORD /d 2 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Policies\Microsoft\Office\16.0\Outlook\Security" /v "PromptOOMSaveAs" /t REG_DWORD /d 2 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Policies\Microsoft\Office\16.0\Outlook\Security" /v "PromptSimpleMAPISend" /t REG_DWORD /d 2 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Policies\Microsoft\Office\16.0\Outlook\Security" /v "PromptSimpleMAPINameResolve" /t REG_DWORD /d 2 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Policies\Microsoft\Office\16.0\Outlook\Security" /v "PromptSimpleMAPIOpenMessage" /t REG_DWORD /d 2 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Policies\Microsoft\Office\16.0\Outlook\Security" /v "AdminSecurityMode" /t REG_DWORD /d 3 /f >nul 2>&1

:: 3. Cấu hình HKCU Policies nếu ghi được
reg add "HKCU\Software\Policies\Microsoft\Office\16.0\Outlook\Security" /v "ObjectModelGuard" /t REG_DWORD /d 2 /f >nul 2>&1
reg add "HKCU\Software\Policies\Microsoft\Office\16.0\Outlook\Security" /v "PromptOOMAddressBookAccess" /t REG_DWORD /d 2 /f >nul 2>&1
reg add "HKCU\Software\Policies\Microsoft\Office\16.0\Outlook\Security" /v "PromptOOMAddressInformationAccess" /t REG_DWORD /d 2 /f >nul 2>&1
reg add "HKCU\Software\Policies\Microsoft\Office\16.0\Outlook\Security" /v "PromptOOMSend" /t REG_DWORD /d 2 /f >nul 2>&1
reg add "HKCU\Software\Policies\Microsoft\Office\16.0\Outlook\Security" /v "AdminSecurityMode" /t REG_DWORD /d 3 /f >nul 2>&1

:: 4. Hỗ trợ trường hợp Office 32-bit trên Windows 64-bit
reg add "HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Outlook\Security" /v "ObjectModelGuard" /t REG_DWORD /d 2 /f >nul 2>&1
reg add "HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Outlook\Security" /v "PromptOOMAddressInformationAccess" /t REG_DWORD /d 2 /f >nul 2>&1
reg add "HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Outlook\Security" /v "PromptOOMAddressBookAccess" /t REG_DWORD /d 2 /f >nul 2>&1
reg add "HKLM\SOFTWARE\WOW6432Node\Policies\Microsoft\Office\16.0\Outlook\Security" /v "ObjectModelGuard" /t REG_DWORD /d 2 /f >nul 2>&1
reg add "HKLM\SOFTWARE\WOW6432Node\Policies\Microsoft\Office\16.0\Outlook\Security" /v "PromptOOMAddressInformationAccess" /t REG_DWORD /d 2 /f >nul 2>&1

echo [THÀNH CÔNG] Đã tắt vĩnh viễn cảnh báo truy cập thông tin email trong Outlook.
echo.
echo Hãy khởi động lại Outlook để thiết lập mới có hiệu lực.
echo.
pause
