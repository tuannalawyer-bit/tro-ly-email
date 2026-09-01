import logging
import logging.handlers
import os
import sys

import webview  # LƯU Ý: gói pip tên "pywebview" nhưng module import là "webview"

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from backend.api import EmailAssistantAPI
from config import APP_NAME
from paths import FROZEN, LOG_DIR, RES_DIR
# Local\ chứ không phải Global\: mỗi phiên đăng nhập được chạy bản riêng, và tạo đối
# tượng trong Global\ đòi quyền SeCreateGlobalPrivilege mà tài khoản thường không có.
MUTEX_NAME = "Local\\TroLyEmail_SingleInstance"

# Giữ tham chiếu suốt vòng đời tiến trình: PyHANDLE tự đóng khi bị thu gom, mất handle
# là mất luôn mutex và cơ chế chống chạy trùng thành vô dụng.
_instance_lock = None

logger = logging.getLogger(__name__)


def setup_logging() -> None:
    """Luôn ghi ra tệp: chạy bằng pythonw.exe thì lỗi không có chỗ nào hiện ra.

    Khi chạy tách rời (start "" pythonw.exe ...), stderr KHÔNG phải None mà là luồng
    dùng bảng mã theo locale — cp1252 trên máy này. Ghi log tiếng Việt vào đó ném
    UnicodeEncodeError, rồi logging lại in traceback vào chính luồng hỏng ấy nên ngoại
    lệ thứ hai làm chết ứng dụng. Phải ép errors="replace" trước khi dùng.
    """
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    handlers: list[logging.Handler] = [
        logging.handlers.RotatingFileHandler(
            LOG_DIR / "app.log", maxBytes=1_000_000, backupCount=3, encoding="utf-8")
    ]
    if sys.stderr is not None:          # python.exe: giữ luôn log ra console
        handlers.append(logging.StreamHandler())   # paths.harden_streams đã lo bảng mã
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        handlers=handlers,
    )


def claim_single_instance():
    """Trả handle mutex, hoặc None nếu đã có bản khác đang chạy.

    Ứng dụng nền bắt buộc phải chống chạy trùng: mở hai lần sẽ ra hai icon khay,
    hai tiến trình cùng tranh Outlook COM.
    """
    try:
        import win32api
        import win32event
        import winerror
        handle = win32event.CreateMutex(None, False, MUTEX_NAME)
        if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
            return None
        return handle
    except ImportError:                 # thiếu pywin32 thì cứ chạy, đừng chặn người dùng
        logger.warning("Không kiểm tra được bản đang chạy (thiếu pywin32).")
        return True


def warn_already_running() -> None:
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(
            None,
            f"{APP_NAME} đang chạy sẵn ở khay hệ thống.\n"
            "Bấm vào icon ở góc phải thanh tác vụ để mở lên.",
            APP_NAME, 0x40)             # MB_ICONINFORMATION
    except Exception:
        logger.warning("%s đang chạy sẵn.", APP_NAME)


def attach_console() -> None:
    """Cấp cửa sổ console cho các lệnh chẩn đoán.

    exe dựng ở chế độ --windowed nên KHÔNG có console: mọi thứ in ra biến mất. Các lệnh
    như --kiem-tra hay --xuat-thu chỉ có giá trị khi người dùng đọc được kết quả.
    """
    if not FROZEN:
        return
    import ctypes
    if not ctypes.windll.kernel32.AllocConsole():
        return
    for name, stream in (("stdout", sys.stdout), ("stderr", sys.stderr)):
        try:
            setattr(sys, name, open("CONOUT$", "w", encoding="utf-8", errors="replace"))
        except OSError:
            pass


def pause_console() -> None:
    if FROZEN:
        try:
            input("\nNhan Enter de dong cua so nay...")
        except (EOFError, OSError):
            pass


def dispatch() -> bool:
    """Các lối vào phụ, thay cho những tệp .bat khi đã đóng gói.

    Trả True nếu đã xử lý xong và không cần mở ứng dụng.
    """
    args = set(sys.argv[1:])

    if "--thiet-lap" in args:
        from thiet_lap import run
        run()
        return True

    if "--addin-server" in args:
        attach_console()
        import addin_server
        addin_server.main()
        return True

    if "--xuat-thu" in args:
        attach_console()
        from xuat_thu_da_gui import main as export_main
        code = export_main()
        pause_console()
        sys.exit(code)

    if "--kiem-tra" in args:
        attach_console()
        import kiem_tra_addin
        code = kiem_tra_addin.main()
        pause_console()
        sys.exit(code)

    return False


def main() -> None:
    setup_logging()

    if dispatch():
        return

    # Bản đóng gói chạy lần đầu: thiết lập trước khi mở ứng dụng, nếu không sẽ không có
    # khoá API lẫn chứng chỉ mà chẳng nói gì. Bộ cài SFX truyền cờ này ngay sau khi bung.
    if FROZEN:
        from thiet_lap import can_skip, run as setup_run
        if "--thiet-lap-lan-dau" in sys.argv or not can_skip():
            setup_run()

    frontend_path = RES_DIR / "frontend" / "index.html"
    if not frontend_path.exists():
        raise FileNotFoundError(f"Không tìm thấy giao diện tại: {frontend_path}")

    global _instance_lock
    # Bản đóng gói mặc định chạy ngầm ở khay; chạy từ mã nguồn thì giữ nếp cũ.
    tray_mode = "--tray" in sys.argv or (FROZEN and "--cua-so" not in sys.argv)
    if tray_mode:
        _instance_lock = claim_single_instance()
        if _instance_lock is None:
            warn_already_running()
            return

    api = EmailAssistantAPI()
    window = webview.create_window(
        title=APP_NAME,
        url=str(frontend_path),
        js_api=api,
        width=1400,
        height=900,
        min_size=(1100, 700),
        background_color="#0a0a1a",
        hidden=tray_mode,
    )

    tray = None
    if tray_mode:
        from tray import TrayApp
        tray = TrayApp(window)
        # before_show chạy ĐỒNG BỘ trên GUI thread và window.native đã sẵn sàng, nên
        # NotifyIcon dùng luôn vòng lặp thông điệp của pywebview.
        window.events.before_show += tray.attach
        window.events.loaded += tray.attach
        window.events.closing += tray.on_closing

    debug = os.getenv("DEBUG", "").strip() in ("1", "true", "True")
    logger.info("Khởi động ứng dụng (debug=%s, khay=%s)...", debug, tray_mode)
    try:
        webview.start(tray.attach if tray else None, debug=debug)
    finally:
        if tray:
            tray.backend.stop()
            tray.dispose()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        # pythonw.exe không có stderr: không bắt ở đây thì lỗi biến mất không dấu vết.
        logging.getLogger(__name__).exception("Ứng dụng dừng do lỗi")
        raise
