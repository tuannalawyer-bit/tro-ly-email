"""Icon khay hệ thống: chạy ngầm như UniKey, gọi lên mới hiện cửa sổ.

Dựng trên System.Windows.Forms qua pythonnet — gói này đã có sẵn vì pywebview kéo
theo, nên KHÔNG thêm phụ thuộc mới. NotifyIcon được tạo trên chính GUI thread của
pywebview (qua sự kiện before_show) nên dùng luôn vòng lặp thông điệp có sẵn, không
cần Application.Run thứ hai.
"""
from __future__ import annotations

import ctypes
import logging
import os
import socket
import subprocess
import sys
import threading
from pathlib import Path
from typing import Callable, Optional

import clr  # noqa: F401  — pythonnet, phải import trước AddReference

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

import System.Windows.Forms as WinForms  # noqa: E402
from System.Drawing import Bitmap, Font, FontStyle, Icon  # noqa: E402

from config import APP_NAME  # noqa: E402
from paths import (CERT_DIR, DATA_ROOT, FROZEN, LOG_DIR, RES_DIR,  # noqa: E402
                   app_args, app_exe, app_workdir)

logger = logging.getLogger(__name__)

TRAY_ICON = RES_DIR / "addin" / "assets" / "icon-16.png"
WINDOW_ICON = RES_DIR / "addin" / "assets" / "icon-32.png"
SHORTCUT_NAME = "Tro ly Email.lnk"

BACKEND_HOST = "127.0.0.1"
BACKEND_PORT = 8765

CREATE_NO_WINDOW = 0x08000000


def load_icon(path: Path) -> Optional[Icon]:
    """PNG -> Icon. Không sinh tệp .ico vì Icon của .NET Framework xử lý kém các
    khung ICO nén PNG, còn Bitmap.GetHicon thì chạy đúng ở mọi kích thước.

    Clone rồi huỷ handle gốc: Icon.FromHandle KHÔNG sở hữu handle, để nguyên là rò
    HICON. Đây cũng là cách pywebview tự làm khi nạp icon cửa sổ.
    """
    if not path.is_file():
        logger.warning("Không tìm thấy icon %s", path)
        return None
    try:
        handle = Bitmap(str(path)).GetHicon()
        try:
            return Icon.FromHandle(handle).Clone()
        finally:
            ctypes.windll.user32.DestroyIcon(handle.ToInt64())
    except Exception:
        logger.exception("Không dựng được icon từ %s", path)
        return None


# --------------------------------------------------------------- backend add-in

class BackendProcess:
    """Chạy backend add-in."""

    def __init__(self, port: int = BACKEND_PORT) -> None:
        self.port = port
        self._lock = threading.Lock()
        self._proc: Optional[subprocess.Popen] = None
        self._log = None
        self._server = None                    # chế độ đóng gói: đối tượng Server
        self._thread: Optional[threading.Thread] = None

    def is_port_busy(self) -> bool:
        for host in ("127.0.0.1", "localhost", "::1"):
            try:
                with socket.create_connection((host, self.port), timeout=0.3):
                    return True
            except OSError:
                pass
        return False

    def is_running(self) -> bool:
        with self._lock:
            return self._thread is not None and self._thread.is_alive()

    def start(self) -> str:
        """Trả thông báo lỗi cho người dùng; chuỗi rỗng nghĩa là đã khởi động."""
        with self._lock:
            if self._thread is not None and self._thread.is_alive():
                return ""
            if self.is_port_busy():
                logger.info("Cổng %s đang được phục vụ.", self.port)
                return ""
            if not (CERT_DIR / "localhost.crt").is_file():
                return "Thiếu chứng chỉ. Chạy lại phần thiết lập để sinh chứng chỉ."

            LOG_DIR.mkdir(parents=True, exist_ok=True)
            return self._start_thread()

    def _start_thread(self) -> str:
        try:
            from addin_server import create_server
            self._server = create_server()
        except SystemExit as e:                # _build_ssl_context ném khi thiếu cert
            return str(e).strip() or "Không dựng được backend."
        except OSError as e:
            if getattr(e, "winerror", None) == 10048 or "10048" in str(e):
                logger.info("Cổng %s đã được tiến trình khác chiếm, sử dụng sẵn.", self.port)
                return ""
            logger.exception("Không dựng được backend add-in")
            return f"Không khởi động được backend: {e}"
        except Exception as e:
            logger.exception("Không dựng được backend add-in")
            return f"Không khởi động được backend: {e}"

        self._thread = threading.Thread(target=self._server.serve_forever,
                                        name="addin-server", daemon=True)
        self._thread.start()
        logger.info("Đã khởi động backend add-in (thread nội bộ).")
        return ""

    def _start_process(self) -> str:
        try:
            self._log = open(LOG_DIR / "addin_server.log", "a", encoding="utf-8")
            self._proc = subprocess.Popen(
                [_pythonw(), str(Path(__file__).resolve().parent / "addin_server.py")],
                cwd=str(Path(__file__).resolve().parent),
                stdout=self._log, stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL,
                creationflags=CREATE_NO_WINDOW)
        except OSError as e:
            self._close_log()
            logger.exception("Không khởi động được backend add-in")
            return f"Không khởi động được backend: {e}"

        logger.info("Đã khởi động backend add-in (PID %s).", self._proc.pid)
        return ""

    def stop(self) -> None:
        with self._lock:
            if self._server is not None:
                logger.info("Đang dừng backend add-in (thread nội bộ).")
                try:
                    self._server.shutdown()
                    self._server.server_close()
                except Exception:
                    logger.exception("Lỗi khi dừng backend")
                self._server = None
                self._thread = None
            if self._proc is not None and self._proc.poll() is None:
                logger.info("Đang dừng backend add-in (PID %s).", self._proc.pid)
                self._proc.terminate()
                try:
                    self._proc.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    self._proc.kill()
            self._proc = None
            self._close_log()

    def restart(self) -> str:
        self.stop()
        return self.start()

    def _close_log(self) -> None:
        if self._log:
            try:
                self._log.close()
            except OSError:
                pass
            self._log = None


def _pythonw() -> str:
    """pythonw.exe cạnh python.exe đang chạy — spawn không kèm cửa sổ đen."""
    exe = Path(sys.executable)
    quiet = exe.with_name("pythonw.exe")
    return str(quiet if quiet.is_file() else exe)


# ------------------------------------------------------ khởi động cùng Windows

class Shortcut:
    """Lối tắt .lnk trỏ tới ứng dụng. Không cần quyền Admin.

    Dùng cho cả thư mục Startup lẫn Desktop — hai chỗ chỉ khác thư mục đích. Chọn lối
    tắt thay vì khoá Run trong registry để người dùng nhìn thấy và tự xoá được.
    """

    def __init__(self, folder: Path, label: str) -> None:
        self.dir = Path(folder)
        self.label = label
        self.path = self.dir / SHORTCUT_NAME

    def enabled(self) -> bool:
        return self.path.is_file()

    def enable(self) -> str:
        try:
            self.dir.mkdir(parents=True, exist_ok=True)
            import win32com.client        # nạp muộn: chỉ cần khi người dùng bật
            shell = win32com.client.Dispatch("WScript.Shell")
            link = shell.CreateShortCut(str(self.path))
            link.TargetPath = app_exe()
            link.Arguments = app_args()
            link.WorkingDirectory = str(app_workdir())
            link.IconLocation = str(WINDOW_ICON)
            link.Description = f"{APP_NAME} — chạy ngầm ở khay hệ thống"
            link.Save()
        except Exception as e:
            logger.exception("Không tạo được lối tắt %s", self.label)
            return f"Không tạo được lối tắt {self.label}: {e}"
        return ""

    def disable(self) -> str:
        try:
            self.path.unlink(missing_ok=True)
        except OSError as e:
            logger.exception("Không xoá được lối tắt %s", self.label)
            return f"Không xoá được lối tắt {self.label}: {e}"
        return ""


def startup_dir() -> Path:
    appdata = os.environ.get("APPDATA") or str(Path.home() / "AppData" / "Roaming")
    return Path(appdata) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"


def desktop_dir() -> Path:
    """Hỏi Windows chứ KHÔNG ghép %USERPROFILE%\\Desktop.

    Máy dùng Microsoft 365 rất hay chuyển hướng Desktop sang OneDrive; ghép tay sẽ tạo
    lối tắt vào một thư mục người dùng không bao giờ nhìn thấy.
    """
    try:
        import win32com.client
        shell = win32com.client.Dispatch("WScript.Shell")
        return Path(str(shell.SpecialFolders("Desktop")))
    except Exception:
        logger.warning("Không hỏi được đường dẫn Desktop, dùng mặc định.", exc_info=True)
        return Path.home() / "Desktop"


def AutoStart(startup_folder: Optional[Path] = None) -> Shortcut:
    return Shortcut(startup_folder or startup_dir(), "khởi động cùng Windows")


def DesktopShortcut(folder: Optional[Path] = None) -> Shortcut:
    return Shortcut(folder or desktop_dir(), "ngoài Desktop")


# ------------------------------------------------------------------ icon khay

class TrayApp:
    """NotifyIcon + menu. Mọi thao tác chạy trên GUI thread của pywebview."""

    def __init__(self, window, backend: Optional[BackendProcess] = None,
                 autostart: Optional[Shortcut] = None) -> None:
        self.window = window
        self.backend = backend or BackendProcess()
        self.autostart = autostart or AutoStart()
        self.quitting = False
        self._icon: Optional[WinForms.NotifyIcon] = None
        self._status_item = None
        self._toggle_item = None
        self._autostart_item = None

    # ------------------------------------------------------------- dựng menu

    def attach(self) -> None:
        """Gọi từ window.events.before_show / loaded — chạy trên GUI thread."""
        try:
            if self._icon is None:
                self._build()
        except Exception:
            logger.exception("Không dựng được icon khay; ứng dụng vẫn chạy bình thường")

    def _build(self) -> None:
        menu = WinForms.ContextMenuStrip()

        open_item = self._item(menu, "Mở Trợ lý Email", lambda *_: self.show_window())
        open_item.Font = Font(open_item.Font, FontStyle.Bold)   # hành động mặc định
        menu.Items.Add(WinForms.ToolStripSeparator())

        self._status_item = WinForms.ToolStripMenuItem("Backend add-in: …")
        self._status_item.Enabled = False
        menu.Items.Add(self._status_item)
        self._toggle_item = self._item(menu, "Khởi động", lambda *_: self._toggle_backend())
        self._item(menu, "Khởi động lại", lambda *_: self._notify(self.backend.restart()))

        menu.Items.Add(WinForms.ToolStripSeparator())
        self._item(menu, "Xuất thư đã gửi để phân tích", lambda *_: self._export())
        self._item(menu, "Kiểm tra add-in", lambda *_: self._check_addin())
        self._item(menu, "Mở thư mục dữ liệu", lambda *_: self._open_data_dir())

        menu.Items.Add(WinForms.ToolStripSeparator())
        self._autostart_item = self._item(menu, "Khởi động cùng Windows",
                                          lambda *_: self._toggle_autostart())
        self._item(menu, "Tạo lối tắt ngoài Desktop", lambda *_: self._make_desktop())
        menu.Items.Add(WinForms.ToolStripSeparator())
        self._item(menu, "Thoát", lambda *_: self.quit())

        menu.Opening += lambda *_: self._refresh()

        self._icon = WinForms.NotifyIcon()
        self._icon.Text = APP_NAME            # tooltip, tối đa 63 ký tự
        icon = load_icon(TRAY_ICON)
        if icon:
            self._icon.Icon = icon
        self._icon.ContextMenuStrip = menu
        self._icon.DoubleClick += lambda *_: self.show_window()
        self._icon.Visible = True

        window_icon = load_icon(WINDOW_ICON)
        if window_icon and getattr(self.window, "native", None):
            self.window.native.Icon = window_icon

        logger.info("Đã hiện icon khay hệ thống.")

    @staticmethod
    def _item(menu, text: str, handler: Callable):
        item = WinForms.ToolStripMenuItem(text)
        item.Click += handler
        menu.Items.Add(item)
        return item

    def _refresh(self) -> None:
        """Cập nhật nhãn ngay lúc menu bật ra — khỏi cần bộ đếm giờ."""
        ours = self.backend.is_running()
        running = ours or self.backend.is_port_busy()
        self._status_item.Text = (
            "Backend add-in: đang chạy" if ours else
            "Backend add-in: đang chạy (do tiến trình khác)" if running else
            "Backend add-in: đã dừng")
        self._toggle_item.Text = "Dừng" if running else "Khởi động"
        # Backend do CHAY_ADDIN_BACKEND.bat chạy thì không phải của mình, không dừng hộ.
        self._toggle_item.Enabled = ours or not running
        self._autostart_item.Checked = self.autostart.enabled()

    # -------------------------------------------------------------- thao tác

    def show_window(self) -> None:
        self.window.show()
        if getattr(self.window, "native", None):
            try:
                native = self.window.native
                native.WindowState = WinForms.FormWindowState.Normal
                native.BringToFront()
                native.Activate()
            except Exception:
                pass

    def hide_window(self) -> None:
        self.window.hide()

    def _toggle_backend(self) -> None:
        if self.backend.is_running():
            self.backend.stop()
            self._notify("", "Đã dừng backend add-in.")
        else:
            self._notify(self.backend.start(), "Đã khởi động backend add-in.")

    def _toggle_autostart(self) -> None:
        if self.autostart.enabled():
            self._notify(self.autostart.disable(), "Đã tắt khởi động cùng Windows.")
        else:
            self._notify(self.autostart.enable(), "Đã bật khởi động cùng Windows.")

    def _make_desktop(self) -> None:
        self._notify(DesktopShortcut().enable(), "Đã tạo lối tắt ngoài Desktop.")

    def _open_data_dir(self) -> None:
        DATA_ROOT.mkdir(parents=True, exist_ok=True)
        os.startfile(str(DATA_ROOT))       # noqa: S606 — mở Explorer, đường dẫn nội bộ

    def _export(self) -> None:
        """Chạy nền: quét cả hộp thư có thể mất hàng chục phút, không được khoá giao diện."""
        def work():
            try:
                from xuat_thu_da_gui import export
                res = export(deep=False)
                self._notify("", f"Đã xuất {res['unique']} thư vào {res['dir']}")
            except Exception as e:
                logger.exception("Xuất thư thất bại")
                self._notify(f"Xuất thư thất bại: {e}")

        self._notify("", "Đang xuất thư đã gửi, sẽ báo khi xong…")
        threading.Thread(target=work, name="xuat-thu", daemon=True).start()

    def _check_addin(self) -> None:
        def work():
            try:
                import kiem_tra_addin
                code = kiem_tra_addin.main()
                self._notify("" if code == 0 else "Có mục chưa đạt, xem log để biết chi tiết.",
                             "Kiểm tra add-in: mọi điều kiện bắt buộc đều đạt.")
            except Exception as e:
                logger.exception("Kiểm tra add-in thất bại")
                self._notify(f"Kiểm tra thất bại: {e}")

        threading.Thread(target=work, name="kiem-tra", daemon=True).start()

    def _notify(self, error: str, done: str = "") -> None:
        if not self._icon:
            return
        if error:
            self._icon.ShowBalloonTip(6000, APP_NAME, error,
                                      WinForms.ToolTipIcon.Warning)
        elif done:
            self._icon.ShowBalloonTip(3000, APP_NAME, done, WinForms.ToolTipIcon.Info)

    def on_closing(self) -> bool:
        """Nút X thu về khay thay vì thoát. Trả False để huỷ lệnh đóng."""
        if self.quitting:
            return True
        self.hide_window()
        return False

    def quit(self) -> None:
        self.quitting = True
        self.backend.stop()
        self.dispose()
        self.window.destroy()

    def dispose(self) -> None:
        if self._icon:
            self._icon.Visible = False
            self._icon.Dispose()
            self._icon = None
