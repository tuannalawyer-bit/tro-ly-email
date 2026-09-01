"""Bộ cài tự bung: một tệp .exe duy nhất mang sẵn cả ứng dụng ở phần đuôi.

Vì sao không dùng thẳng PyInstaller --onefile cho cả ứng dụng: đo thực tế trên máy đích
cho thấy onefile mất **20,5 giây mỗi lần chạy** vì phải bung ~800 tệp ra %TEMP% và phần
mềm bảo vệ quét lại từng tệp. Bản --onedir chỉ mất 6,9 giây.

Cách làm ở đây: bộ cài là một exe onefile RẤT NHỎ (chỉ thư viện chuẩn, khởi động ~2 giây),
gói ứng dụng dạng .zip được nối thẳng vào đuôi tệp exe. Bung một lần vào %LOCALAPPDATA%
rồi từ đó chạy bản nhanh. Người dùng vẫn chỉ nhận đúng một tệp.

zipfile của Python đọc được zip nối sau phần dữ liệu khác (nó dò End Of Central Directory
từ cuối tệp), còn bootloader của PyInstaller vẫn tìm ra archive của nó — đã kiểm chứng.
"""
from __future__ import annotations

import ctypes
import os
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

APP_NAME = "Trợ lý Email"
APP_FOLDER = "TroLyEmail"
EXE_NAME = "TroLyEmail.exe"

# Dấu phiên bản do dong_goi.py ghi vào gói. Bản cài trước 2.1.0 không có tệp này, khi đó
# coi như "không rõ" và vẫn cho nâng cấp.
VERSION_FILE = "phien_ban.txt"

MB_ICONINFORMATION = 0x40
MB_ICONERROR = 0x10
MB_YESNO = 0x04
IDYES = 6


def message(text: str, flags: int = MB_ICONINFORMATION) -> int:
    return ctypes.windll.user32.MessageBoxW(None, text, APP_NAME, flags)


def install_root() -> Path:
    base = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    return Path(base) / APP_FOLDER


def extract(dest: Path) -> None:
    """Bung phần zip nối ở đuôi chính tệp exe này."""
    with zipfile.ZipFile(sys.executable) as z:
        z.extractall(dest)


def find_exe(root: Path) -> Path | None:
    hit = list(root.rglob(EXE_NAME))
    return hit[0] if hit else None


def installed_version(exe: Path) -> str:
    try:
        return (exe.parent / VERSION_FILE).read_text(encoding="utf-8").strip()
    except OSError:
        return ""


def payload_version() -> str:
    """Phiên bản nằm trong gói đính ở đuôi chính tệp exe này."""
    try:
        with zipfile.ZipFile(sys.executable) as z:
            for name in z.namelist():
                if name.endswith("/" + VERSION_FILE):
                    return z.read(name).decode("utf-8").strip()
    except Exception:
        pass
    return ""


def stop_running() -> None:
    """Ứng dụng đang chạy sẽ khoá exe khiến việc ghi đè thất bại."""
    subprocess.run(["taskkill", "/IM", EXE_NAME, "/F"],
                   capture_output=True, creationflags=0x08000000)   # CREATE_NO_WINDOW


def main() -> int:
    root = install_root()
    app_dir = root / "app"
    exe = find_exe(app_dir) if app_dir.is_dir() else None

    upgrading = False
    if exe and exe.is_file():
        cu, moi = installed_version(exe), payload_version()
        if cu and moi and cu == moi:
            # Đúng bản này rồi: chạy thẳng, không bung lại.
            if message(f"{APP_NAME} v{cu} đã được cài.\n\nMở ứng dụng ngay?",
                       MB_ICONINFORMATION | MB_YESNO) == IDYES:
                subprocess.Popen([str(exe)], cwd=str(exe.parent))
            return 0

        hien_co = f"v{cu}" if cu else "một bản cũ"
        if message(f"Máy đang có {APP_NAME} {hien_co}.\n\n"
                   f"Cài đè bằng v{moi or '(mới)'}?\n\n"
                   f"Khoá API, chứng chỉ và kho tri thức của bạn được giữ nguyên.",
                   MB_ICONINFORMATION | MB_YESNO) != IDYES:
            return 0
        stop_running()
        upgrading = True

    try:
        app_dir.parent.mkdir(parents=True, exist_ok=True)
        if app_dir.exists():
            shutil.rmtree(app_dir, ignore_errors=True)
        extract(app_dir)
    except Exception as e:
        message(f"Không bung được gói cài đặt:\n\n{e}\n\n"
                f"Thử chép tệp này ra ổ đĩa cục bộ rồi chạy lại.", MB_ICONERROR)
        return 1

    exe = find_exe(app_dir)
    if not exe:
        message(f"Bung xong nhưng không tìm thấy {EXE_NAME}.", MB_ICONERROR)
        return 1

    # Ứng dụng tự lo phần còn lại: hỏi khoá API, sinh và cài chứng chỉ, tạo lối tắt.
    # Cả khi nâng cấp cũng chạy bước này: manifest.xml mang số phiên bản nên phải xuất
    # lại, còn lối tắt thì trỏ vào exe vừa bung. Khoá API và chứng chỉ đã có thì các bước
    # đó tự nhận ra và bỏ qua, không hỏi lại người dùng.
    subprocess.Popen([str(exe), "--thiet-lap-lan-dau"], cwd=str(exe.parent))
    if upgrading:
        message(f"Đã nâng cấp xong.\n\nNếu bạn đang mở Outlook, hãy đóng và mở lại "
                f"để add-in nhận bản mới.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
