import sys
sys.coinit_flags = 2  # Bắt buộc khởi tạo STA apartment cho COM và WinForms trên Windows

import logging
import logging.handlers
import os

import webview  # LƯU Ý: gói pip tên "pywebview" nhưng module import là "webview"

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from backend.api import EmailAssistantAPI
from config import APP_NAME
from paths import FROZEN, LOG_DIR, RES_DIR
# Local\ chứ không phải Global\: mỗi phiên đăng nhập được chạy bản riêng, và tạo đối
# tượng trong Global\ đòi quyền SeCreateGlobalPrivilege mà tài khoản thường không có.
MUTEX_NAME_TRAY = "Local\\TroLyEmail_Tray_Instance"
MUTEX_NAME_GUI = "Local\\TroLyEmail_GUI_Instance"

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


def claim_single_instance(mutex_name: str = MUTEX_NAME_GUI):
    """Trả handle mutex, hoặc None nếu đã có bản khác đang chạy.
    
    Dùng trực tiếp ctypes gọi Win32 kernel32 API để đảm bảo 100% tin cậy, không phụ thuộc
    pywin32 DLL vốn có thể lỗi nạp động trong bản đóng gói.
    """
    try:
        import ctypes
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.CreateMutexW(None, False, mutex_name)
        if kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
            if handle:
                kernel32.CloseHandle(handle)
            return None
        return handle
    except Exception:
        logger.exception("Lỗi khi kiểm tra tiến trình đang chạy")
        return None


def bring_existing_to_front() -> bool:
    try:
        import ctypes
        user32 = ctypes.windll.user32
        hwnd = user32.FindWindowW(None, APP_NAME)
        if hwnd:
            user32.ShowWindow(hwnd, 9)       # SW_RESTORE
            user32.SetForegroundWindow(hwnd)
            return True
    except Exception:
        pass
    return False


def warn_already_running() -> None:
    if bring_existing_to_front():
        return
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(
            None,
            f"{APP_NAME} đang chạy ở khay hệ thống.\n"
            "Bấm vào icon ở góc phải thanh tác vụ (gần đồng hồ) để mở lên.",
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

    # Khởi động cùng Windows ở chế độ khay hệ thống (--tray):
    # Chạy ngầm 100% bằng WinForms.ApplicationContext thuần túy, KHÔNG tạo bất kỳ cửa sổ nào.
    # Tuyệt đối tránh tải pywebview / WebView2 / COM lúc khởi động máy để không bao giờ bị treo xám.
    if "--tray" in sys.argv:
        _instance_lock = claim_single_instance(MUTEX_NAME_TRAY)
        if _instance_lock is None:
            logger.info("Trợ lý Email (khay hệ thống) đã đang chạy sẵn.")
            return

        from tray import run_tray_standalone
        logger.info("Khởi động Trợ lý Email ở chế độ khay hệ thống (chạy ngầm hoàn toàn)...")
        run_tray_standalone()
        return

    # Chế độ mở cửa sổ giao diện chính (Desktop shortcut hoặc mở thủ công):
    _instance_lock = claim_single_instance(MUTEX_NAME_GUI)
    if _instance_lock is None:
        warn_already_running()
        return

    # Đảm bảo tiến trình khay hệ thống (--tray) chạy ngầm để phục vụ Outlook 24/7 và giữ icon khay
    tray_lock = claim_single_instance(MUTEX_NAME_TRAY)
    if tray_lock is not None:
        try:
            import ctypes
            ctypes.windll.kernel32.CloseHandle(tray_lock)
        except Exception:
            pass
        import subprocess
        from paths import app_exe, app_workdir
        exe = app_exe()
        workdir = app_workdir()
        if FROZEN:
            subprocess.Popen([exe, "--tray"], cwd=str(workdir))
        else:
            python = sys.executable
            main_py = str(os.path.abspath(__file__))
            subprocess.Popen([python, main_py, "--tray"], cwd=str(workdir))
    else:
        logger.info("Tiến trình khay hệ thống đang chạy sẵn.")

    api = EmailAssistantAPI()
    window = webview.create_window(
        title=APP_NAME,
        url=frontend_path.as_uri(),
        js_api=api,
        width=1400,
        height=900,
        min_size=(1100, 700),
        background_color="#0a0a1a",
        hidden=False,
    )

    debug = os.getenv("DEBUG", "").strip() in ("1", "true", "True")
    logger.info("Khởi động cửa sổ giao diện chính (debug=%s)...", debug)
    try:
        webview.start(debug=debug)
        logger.info("Cửa sổ giao diện đã đóng bình thường.")
    except Exception:
        logger.exception("Lỗi khi mở cửa sổ giao diện")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        # pythonw.exe không có stderr: không bắt ở đây thì lỗi biến mất không dấu vết.
        logging.getLogger(__name__).exception("Ứng dụng dừng do lỗi")
        raise
