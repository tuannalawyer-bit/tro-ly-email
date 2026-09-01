"""Thiết lập lần đầu cho bản đóng gói: khoá API, chứng chỉ, kho tri thức, lối tắt.

Chạy tự động khi thiếu `.env` hoặc chứng chỉ, hoặc gọi tay bằng `TroLyEmail.exe --thiet-lap`.
Giao diện dựng bằng WinForms qua pythonnet — đã có sẵn vì pywebview kéo theo, không thêm
phụ thuộc nào.
"""
from __future__ import annotations

import logging
import os
import shutil
from pathlib import Path

import clr  # noqa: F401

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

import System.Windows.Forms as WinForms  # noqa: E402
from System.Drawing import Font, FontStyle, Point, Size  # noqa: E402

from config import APP_NAME, KNOWLEDGE_SAMPLE_DIR, VERSION  # noqa: E402
from paths import (ADDIN_RES_DIR, CERT_DIR, DATA_ROOT, ENV_FILE,  # noqa: E402
                   MANIFEST_OUT)

logger = logging.getLogger(__name__)

KEY_URL = "https://aistudio.google.com/apikey"


def can_skip() -> bool:
    """Đã thiết lập xong thì không hỏi lại."""
    return ENV_FILE.is_file() and (CERT_DIR / "localhost.crt").is_file()


# ------------------------------------------------------------------ hộp thoại

def ask_api_key() -> str:
    """Hộp thoại nhập khoá. Trả chuỗi rỗng nếu người dùng bỏ qua."""
    form = WinForms.Form()
    form.Text = f"{APP_NAME} — thiết lập lần đầu"
    form.Size = Size(560, 300)
    form.StartPosition = WinForms.FormStartPosition.CenterScreen
    form.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog
    form.MaximizeBox = form.MinimizeBox = False

    title = WinForms.Label()
    title.Text = "Nhập khoá API Google Gemini"
    title.Font = Font(title.Font, FontStyle.Bold)
    title.Location = Point(20, 18)
    title.Size = Size(500, 22)
    form.Controls.Add(title)

    note = WinForms.Label()
    note.Text = ("Khoá dùng để soạn thư. Lấy miễn phí tại aistudio.google.com/apikey — "
                 "mỗi người nên dùng khoá riêng vì hạn mức tính theo từng khoá.\n\n"
                 "Bỏ trống cũng được, sau này nhập trong phần Cài đặt của ứng dụng.")
    note.Location = Point(20, 46)
    note.Size = Size(500, 70)
    form.Controls.Add(note)

    box = WinForms.TextBox()
    box.Location = Point(20, 126)
    box.Size = Size(500, 26)
    box.UseSystemPasswordChar = True
    form.Controls.Add(box)

    show = WinForms.CheckBox()
    show.Text = "Hiện khoá"
    show.Location = Point(20, 158)
    show.Size = Size(120, 24)
    show.CheckedChanged += lambda *_: setattr(box, "UseSystemPasswordChar", not show.Checked)
    form.Controls.Add(show)

    link = WinForms.LinkLabel()
    link.Text = "Mở trang lấy khoá"
    link.Location = Point(150, 160)
    link.Size = Size(160, 22)
    link.LinkClicked += lambda *_: os.startfile(KEY_URL)   # noqa: S606
    form.Controls.Add(link)

    ok = WinForms.Button()
    ok.Text = "Tiếp tục"
    ok.Location = Point(410, 200)
    ok.Size = Size(110, 32)
    ok.DialogResult = WinForms.DialogResult.OK
    form.Controls.Add(ok)
    form.AcceptButton = ok

    form.ShowDialog()
    value = box.Text.strip()
    form.Dispose()
    return value


def show_summary(lines: list) -> None:
    WinForms.MessageBox.Show(
        "\n".join(lines), f"{APP_NAME} v{VERSION} — thiết lập xong",
        WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information)


# ------------------------------------------------------------------ các bước

def has_api_key() -> bool:
    if os.environ.get("GEMINI_API_KEY", "").strip():
        return True
    if not ENV_FILE.is_file():
        return False
    return any(ln.startswith("GEMINI_API_KEY=") and ln.split("=", 1)[1].strip()
               for ln in ENV_FILE.read_text(encoding="utf-8").splitlines())


def save_api_key(key: str) -> str:
    from dotenv import set_key
    ENV_FILE.parent.mkdir(parents=True, exist_ok=True)
    ENV_FILE.touch(exist_ok=True)
    if key:
        set_key(str(ENV_FILE), "GEMINI_API_KEY", key)
        os.environ["GEMINI_API_KEY"] = key
    set_key(str(ENV_FILE), "GEMINI_MODEL", os.environ.get("GEMINI_MODEL", "gemini-3.5-flash-lite"))
    if not key:
        return "Chưa nhập khoá API — nhập sau trong phần Cài đặt."
    return "Đã lưu khoá API."


def make_cert() -> str:
    if (CERT_DIR / "localhost.crt").is_file():
        return "Chứng chỉ đã có sẵn."
    import tao_chung_chi_addin
    tao_chung_chi_addin.create()
    return "Đã sinh chứng chỉ cho backend add-in."


def trust_cert() -> str:
    """Cài chứng chỉ vào kho Trusted Root của NGƯỜI DÙNG (không cần quyền Admin).

    Windows sẽ hiện hộp xác nhận — đó là hành vi đúng, người dùng phải thấy mình đang
    tin một chứng chỉ tự ký.
    """
    try:
        import cai_chung_chi
        code = cai_chung_chi.main()
        return ("Đã cài chứng chỉ vào Trusted Root." if code == 0
                else "Chưa cài được chứng chỉ — chạy lại phần thiết lập nếu add-in lỗi.")
    except Exception as e:
        logger.exception("Không cài được chứng chỉ")
        return f"Không cài được chứng chỉ: {e}"


def copy_knowledge() -> str:
    dest = DATA_ROOT / "kien_thuc"
    if dest.exists():
        return "Kho tri thức đã có sẵn."
    if not KNOWLEDGE_SAMPLE_DIR.is_dir():
        return "Không tìm thấy kho tri thức mẫu."
    shutil.copytree(KNOWLEDGE_SAMPLE_DIR, dest)
    return "Đã tạo kho tri thức từ bản mẫu."


def export_manifest() -> str:
    """Ghi manifest ra đĩa: Outlook cài add-in bằng Add from file nên cần đường dẫn cố
    định, mà thư mục tài nguyên trong gói thì đổi tên mỗi lần chạy."""
    src = ADDIN_RES_DIR / "manifest.xml"
    if not src.is_file():
        return "Không tìm thấy manifest.xml trong gói."
    MANIFEST_OUT.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(src, MANIFEST_OUT)
    return f"Đã xuất manifest.xml ra {MANIFEST_OUT}"


def make_shortcuts() -> list:
    from tray import AutoStart, DesktopShortcut
    out = []
    desktop = DesktopShortcut()
    err = desktop.enable()
    out.append(err or f"Đã tạo lối tắt ngoài Desktop ({desktop.dir}).")

    auto = AutoStart()
    err = auto.enable()
    out.append(err or "Đã bật khởi động cùng Windows (tắt được trong menu khay).")
    return out


def check_webview2() -> str:
    """WebView2 Runtime thiếu thì cửa sổ và task pane trắng trơn mà không báo lỗi."""
    import winreg
    keys = [
        (winreg.HKEY_LOCAL_MACHINE,
         r"SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients"
         r"\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"),
        (winreg.HKEY_CURRENT_USER,
         r"SOFTWARE\Microsoft\EdgeUpdate\Clients"
         r"\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"),
    ]
    for root, path in keys:
        try:
            with winreg.OpenKey(root, path) as k:
                return f"WebView2 Runtime: có (bản {winreg.QueryValueEx(k, 'pv')[0]})."
        except OSError:
            continue
    return ("THIẾU WebView2 Runtime — cửa sổ ứng dụng sẽ trắng trơn. "
            "Cài Microsoft Edge WebView2 Runtime rồi chạy lại.")


# ------------------------------------------------------------------ chạy

def run() -> None:
    DATA_ROOT.mkdir(parents=True, exist_ok=True)
    results = []

    # Chạy lại phần thiết lập thì không hỏi lại khoá đã có.
    if has_api_key():
        results.append("Khoá API đã có sẵn, giữ nguyên.")
    else:
        results.append(save_api_key(ask_api_key()))
    for step in (make_cert, trust_cert, copy_knowledge, export_manifest, check_webview2):
        try:
            results.append(step())
        except Exception as e:
            logger.exception("Bước thiết lập %s thất bại", step.__name__)
            results.append(f"Lỗi ở bước {step.__name__}: {e}")
    results.extend(make_shortcuts())

    results.append("")
    results.append("Bước cuối, làm trong Outlook:")
    results.append("  Get Add-ins → My add-ins → Add a custom add-in → Add from file")
    results.append(f"  rồi chọn: {MANIFEST_OUT}")

    logger.info("Thiết lập xong: %s", " | ".join(r for r in results if r))
    show_summary(results)
    try:
        os.startfile(str(DATA_ROOT))       # noqa: S606 — mở Explorer để người dùng lấy manifest
    except OSError:
        pass


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    run()
