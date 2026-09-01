"""Tách đường dẫn tài nguyên (chỉ đọc) khỏi đường dẫn dữ liệu (ghi được).

Chạy từ mã nguồn thì hai thứ nằm chung một chỗ nên trước giờ dùng chung `BASE_DIR`.
Sau khi đóng gói thì không còn đúng nữa:

- **Tài nguyên** (`frontend/`, `addin/`, `kien_thuc_mau/`) nằm trong gói PyInstaller,
  tức thư mục `sys._MEIPASS`. Thư mục này CHỈ ĐỌC và bị xoá khi thoát.
- **Dữ liệu** (`data/`, `certs/`, `kien_thuc/`, `xuat_thu/`, `.env`) phải sống lâu dài
  ở hồ sơ người dùng, nếu không thì tắt máy là mất sạch khoá API lẫn kho tri thức.

Module này KHÔNG import config để tránh vòng lặp — config mới là bên import nó.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

FROZEN = getattr(sys, "frozen", False)

_SRC_DIR = Path(__file__).resolve().parent


def harden_streams() -> None:
    """Cho phép ghi tiếng Việt ra stdout/stderr trong MỌI cách khởi chạy.

    Khi tiến trình không gắn với console thật — chạy tách rời bằng start, chạy bằng
    pythonw.exe, hay là tiến trình con có stdout chuyển hướng vào tệp — Python dùng bảng
    mã theo locale, cp1252 trên máy tiếng Việt. Mọi lệnh print có dấu sẽ ném
    UnicodeEncodeError và giết luôn tác vụ đang chạy.

    Đã cắn hai lần: v1.9.0 làm ứng dụng chết im lặng lúc khởi động (stderr), và lần này
    làm hỏng việc xuất thư ngay ở dòng "Tìm thấy … thư mục" (stdout). Nên đặt ở đây và
    gọi lúc import: paths là module mà mọi lối vào đều nạp, sửa một chỗ là xong cả.
    """
    for name in ("stdout", "stderr"):
        stream = getattr(sys, name, None)
        if stream is None:
            continue
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, OSError, ValueError):
            pass


harden_streams()

# Tài nguyên chỉ đọc. Đóng băng thì PyInstaller đặt sẵn _MEIPASS (cả onefile lẫn onedir).
RES_DIR = Path(getattr(sys, "_MEIPASS", _SRC_DIR)) if FROZEN else _SRC_DIR

APP_FOLDER_NAME = "TroLyEmail"


def _data_root() -> Path:
    if not FROZEN:
        return _SRC_DIR          # chạy từ mã nguồn: giữ nguyên nếp cũ, kiểm thử không đổi
    base = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    return Path(base) / APP_FOLDER_NAME


DATA_ROOT = _data_root()

ENV_FILE = DATA_ROOT / ".env"
CERT_DIR = DATA_ROOT / "certs"
LOG_DIR = DATA_ROOT / "data" / "logs"

# Thư mục addin trong gói (chỉ đọc). Bản manifest.xml người dùng thêm vào Outlook thì
# phải là bản ghi ra đĩa ở MANIFEST_OUT: Outlook cài add-in bằng "Add from file" nên
# cần đường dẫn cố định, mà _MEIPASS thì đổi tên mỗi lần chạy.
ADDIN_RES_DIR = RES_DIR / "addin"
MANIFEST_OUT = DATA_ROOT / "manifest.xml"


def app_exe() -> str:
    """Đường dẫn dùng cho lối tắt: chính exe khi đã đóng gói, pythonw khi chạy mã nguồn."""
    if FROZEN:
        return sys.executable
    exe = Path(sys.executable)
    quiet = exe.with_name("pythonw.exe")
    return str(quiet if quiet.is_file() else exe)


def app_args() -> str:
    """Tham số kèm theo lối tắt. Bản đóng gói chạy thẳng, không cần trỏ tới main.py."""
    return "" if FROZEN else f'"{_SRC_DIR / "main.py"}" --tray'


def app_workdir() -> Path:
    """Thư mục làm việc của lối tắt.

    KHÔNG lấy thư mục chứa app_exe(): chạy từ mã nguồn thì đó là .venv\\Scripts, không
    phải thư mục dự án.
    """
    return Path(sys.executable).parent if FROZEN else _SRC_DIR
