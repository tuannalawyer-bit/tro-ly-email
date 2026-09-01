"""Backend HTTPS cục bộ cho Outlook Web Add-in.

Mô hình bảo mật
---------------
Task pane được Outlook nạp TỪ CHÍNH https://localhost:8765, nên origin của nó
đúng bằng origin của server này -> mọi fetch từ task pane là same-origin.
Vì vậy server KHÔNG phát bất kỳ header Access-Control-Allow-* nào: trang web ở
origin khác sẽ bị trình duyệt chặn đọc phản hồi.

Thêm một lớp token sinh ngẫu nhiên lúc khởi động, nhúng vào taskpane.html và bắt
buộc ở header X-Addin-Token. Header tuỳ biến không thể thêm vào request chéo
origin mà không kích hoạt preflight, nên CSRF dạng text/plain cũng bị chặn trước
khi kịp gọi Gemini.

Rủi ro còn lại: tiến trình cục bộ chạy dưới cùng tài khoản người dùng vẫn đọc
được token từ taskpane.html. Đây là giới hạn không thể khắc phục — task pane
buộc phải lấy được bí mật đó bằng cách nào đó. Tiến trình như vậy cũng đọc được
thẳng tệp .env.
"""
from __future__ import annotations

import hmac
import json
import logging
import secrets
import ssl
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Optional
from urllib.parse import unquote, urlsplit

from dotenv import load_dotenv

from backend.api import EmailAssistantAPI, fail
from config import GEMINI_MODEL, VERSION
from paths import ADDIN_RES_DIR, CERT_DIR, ENV_FILE

HOST = "127.0.0.1"
PORT = 8765
ADDIN_DIR = ADDIN_RES_DIR          # chỉ đọc: nằm trong gói khi đã đóng gói
MAX_BODY_BYTES = 25 * 1024 * 1024

ADDIN_TOKEN = secrets.token_urlsafe(32)
ALLOWED_HOSTS = {f"localhost:{PORT}", f"127.0.0.1:{PORT}"}
ALLOWED_ORIGINS = {f"https://localhost:{PORT}", f"https://127.0.0.1:{PORT}"}

CONTENT_TYPES = {
    ".html": "text/html; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".js": "text/javascript; charset=utf-8",
    ".svg": "image/svg+xml",
    ".png": "image/png",
    ".ico": "image/x-icon",
}

logger = logging.getLogger(__name__)


class AddinService:
    """Dựng MỘT lần lúc khởi động và dùng lại cho mọi request.

    EmailAssistantAPI.__init__ không chạm COM (OutlookClient chỉ gán self.app=None),
    nên khởi tạo ở đây rẻ và không mở kết nối Outlook thứ hai.
    """

    def __init__(self) -> None:
        self.api = EmailAssistantAPI()
        self.kb = self.api.kb                 # EmailAssistantAPI đã dựng sẵn, dùng lại
        # Một người dùng, một luồng suy nghĩ: nháy đúp nút không được nhân đôi
        # lượt gọi Gemini và đốt hạn mức hai lần.
        self.ai_lock = threading.Lock()
        self.job = ExportJob(self.api)


class ExportJob:
    """Chạy việc xuất thư đã gửi ở luồng nền và báo tiến độ qua polling.

    Bắt buộc phải là tác vụ nền: quét mọi kho thư trong Outlook có thể mất 10-30 phút.
    Gọi đồng bộ sẽ làm task pane treo cứng và Outlook có thể tự đóng nó.

    Đây là endpoint DUY NHẤT của add-in được phép chạm Outlook COM. Vì addin_server
    là tiến trình riêng với main.py nên nó mở kết nối Outlook thứ hai — Outlook chấp
    nhận, nhưng lần đầu có thể hiện hộp thoại xin quyền truy cập tự động.

    Việc xuất KHÔNG gọi Gemini lần nào.
    """

    def __init__(self, api: EmailAssistantAPI) -> None:
        self.api = api
        self._lock = threading.Lock()
        self._thread: Optional[threading.Thread] = None
        self._state = self._blank()

    @staticmethod
    def _blank() -> dict:
        return {"running": False, "phase": "", "current": 0, "total": 0,
                "detail": "", "done": False, "error": None, "result": None}

    def status(self) -> dict:
        with self._lock:
            return dict(self._state)

    def start(self, deep: bool = True) -> tuple:
        with self._lock:
            if self._state["running"]:
                return False, "Đang chạy rồi."
            self._state = self._blank()
            self._state.update(running=True, phase="Đang khởi động")
        self._thread = threading.Thread(target=self._run, args=(deep,), daemon=True,
                                        name="xuat-thu")
        self._thread.start()
        return True, ""

    def _report(self, phase: str, current: int = 0, total: int = 0,
                detail: str = "") -> None:
        with self._lock:
            self._state.update(phase=phase, current=current, total=total,
                               detail=detail)

    def _run(self, deep: bool) -> None:
        try:
            res = self.api.export_sent_emails(progress=self._report, deep=deep)
            with self._lock:
                if res.get("success"):
                    d = res.get("data") or {}
                    self._state.update(
                        phase="Hoàn tất", done=True,
                        result={"total_raw": d.get("total_raw", 0),
                                "unique": d.get("unique", 0),
                                "short": d.get("short", 0),
                                "files": d.get("files", 0),
                                "analyzed_count": d.get("analyzed_count", 0),
                                "dir": d.get("dir", "")})
                else:
                    self._state.update(phase="Lỗi", done=True, error=res.get("error"))
        except Exception as e:                        # luồng nền không được ném ra ngoài
            logger.exception("Xuat thu that bai")
            with self._lock:
                self._state.update(phase="Lỗi", done=True, error=str(e))
        finally:
            with self._lock:
                self._state["running"] = False


service: Optional[AddinService] = None


class Handler(BaseHTTPRequestHandler):
    server_version = f"TroLyEmail/{VERSION}"

    def log_message(self, fmt, *args):
        logger.info("%s - %s", self.address_string(), fmt % args)

    # ------------------------------------------------------------- tiện ích

    def _send(self, status: int, body: bytes, content_type: str,
              extra_headers: Optional[dict] = None) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("X-Content-Type-Options", "nosniff")
        for key, value in (extra_headers or {}).items():
            self.send_header(key, value)
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)

    def _json(self, status: int, data) -> None:
        self._send(status, json.dumps(data, ensure_ascii=False).encode("utf-8"),
                   "application/json; charset=utf-8", {"Cache-Control": "no-store"})

    def _authorized(self) -> bool:
        supplied = self.headers.get("X-Addin-Token", "")
        if not hmac.compare_digest(supplied, ADDIN_TOKEN):
            return False
        if self.headers.get("Host", "") not in ALLOWED_HOSTS:   # chặn DNS rebinding
            return False
        origin = self.headers.get("Origin")
        if origin and origin not in ALLOWED_ORIGINS:
            return False
        site = self.headers.get("Sec-Fetch-Site")               # WebView2 là Chromium
        if site is not None and site != "same-origin":
            return False
        return True

    def _resolve_static(self, url_path: str) -> Optional[Path]:
        """Ánh xạ đường dẫn URL sang tệp trong addin/, chặn thoát thư mục.

        Bản cũ dùng ADDIN_DIR / path.removeprefix("/addin/") nên "/addin/../.env"
        trả về nguyên tệp .env kèm GEMINI_API_KEY cho bất kỳ tiến trình cục bộ nào.
        """
        rel = unquote(urlsplit(url_path).path)      # bỏ luôn ?_host_Info=... của Office
        if rel in ("/", "/addin", "/addin/", "/taskpane.html"):
            rel = "/addin/taskpane.html"
        if not rel.startswith("/addin/"):
            return None
        root = ADDIN_DIR.resolve()
        try:
            candidate = (root / rel[len("/addin/"):]).resolve()
        except (OSError, ValueError):
            return None
        if root != candidate and root not in candidate.parents:
            return None
        return candidate if candidate.is_file() else None

    def _read_json_body(self) -> Optional[dict]:
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            self._json(400, fail("Content-Length không hợp lệ."))
            return None
        if length > MAX_BODY_BYTES:                 # chặn TRƯỚC khi đọc, tránh cạn RAM
            self._json(413, fail(
                f"Dữ liệu gửi lên quá lớn ({length / 1048576:.1f} MB). "
                f"Giới hạn {MAX_BODY_BYTES // 1048576} MB."))
            return None
        try:
            return json.loads(self.rfile.read(length) or b"{}")
        except (json.JSONDecodeError, UnicodeDecodeError) as e:
            self._json(400, fail(f"Dữ liệu gửi lên không phải JSON hợp lệ: {e}"))
            return None

    # ------------------------------------------------------------------ GET

    def do_GET(self):
        path = urlsplit(self.path).path

        if path.startswith("/api/"):
            if not self._authorized():
                self._json(401, fail(
                    "Phiên không hợp lệ. Backend có thể đã khởi động lại — "
                    "hãy đóng và mở lại task pane."))
                return
            if path == "/api/health":
                self._json(200, {"success": True, "error": None, "data": {
                    "version": VERSION,
                    "model": GEMINI_MODEL,
                    "api_key_configured": service.api.ai_engine.is_ready,
                    "knowledge": service.kb.status(),
                }})
                return
            if path == "/api/knowledge":
                self._json(200, service.api.knowledge_status())
                return
            if path == "/api/email-types":
                self._json(200, service.api.list_email_types())
                return
            if path in ("/api/export-emails/status", "/api/learn-style/status"):
                self._json(200, {"success": True, "error": None,
                                 "data": service.job.status()})
                return
            self._json(404, fail("Không tìm thấy API."))
            return

        target = self._resolve_static(path)
        if target is None:
            self._json(404, fail("Không tìm thấy tài nguyên."))
            return

        if target.name == "taskpane.html":
            html = target.read_text(encoding="utf-8").replace("{{ADDIN_TOKEN}}", ADDIN_TOKEN)
            self._send(200, html.encode("utf-8"), CONTENT_TYPES[".html"],
                       {"Cache-Control": "no-store"})   # đừng để WebView2 ghi token ra đĩa
            return

        self._send(200, target.read_bytes(),
                   CONTENT_TYPES.get(target.suffix, "application/octet-stream"))

    def do_HEAD(self):
        self.do_GET()

    # ----------------------------------------------------------------- POST

    def do_POST(self):
        path = urlsplit(self.path).path
        if not self._authorized():
            self._json(401, fail(
                "Phiên không hợp lệ. Backend có thể đã khởi động lại — "
                "hãy đóng và mở lại task pane."))
            return

        payload = self._read_json_body()
        if payload is None:
            return

        try:
            if path == "/api/generate-reply":
                with service.ai_lock:
                    self._json(200, service.api.generate_reply_from_payload(
                        payload.get("email") or {},
                        payload.get("instruction", ""),
                        payload.get("attachments"),
                        payload.get("me"),
                        payload.get("email_type", "")))
                return

            if path == "/api/refine-draft":
                with service.ai_lock:
                    self._json(200, service.api.refine_draft(
                        payload.get("draft_html", ""), payload.get("feedback", "")))
                return

            if path == "/api/classify-email":
                email = payload.get("email") or {}
                with service.ai_lock:
                    result = service.api.classify_emails([{
                        "entry_id": "addin-current",
                        "sender_name": email.get("sender_name", ""),
                        "subject": email.get("subject", ""),
                        "preview": (email.get("body") or "")[:1200],
                    }])
                if not result.get("success"):
                    self._json(200, result)
                    return
                self._json(200, {"success": True, "error": None,
                                 "data": next(iter(result["data"].values()), {})})
                return

            if path in ("/api/export-emails", "/api/learn-style"):
                # Trả về ngay lập tức; task pane theo dõi qua .../status.
                started, why = service.job.start(deep=bool(payload.get("deep", True)))
                if not started:
                    self._json(200, fail(why))
                    return
                self._json(200, {"success": True, "error": None,
                                 "data": service.job.status()})
                return

            self._json(404, fail("Không tìm thấy API."))
        except Exception as exc:                  # lưới an toàn cuối, không để sập server
            logger.exception("Lỗi không lường trước tại %s", path)
            self._json(500, fail(str(exc) or exc.__class__.__name__))


class Server(ThreadingHTTPServer):
    daemon_threads = True
    allow_reuse_address = True


def _build_ssl_context() -> ssl.SSLContext:
    cert_file = CERT_DIR / "localhost.crt"
    key_file = CERT_DIR / "localhost.key"
    if not (cert_file.exists() and key_file.exists()):
        raise SystemExit(
            "\n[LOI] Thieu certs/localhost.crt hoac certs/localhost.key.\n"
            "      Chay TAO_CHUNG_CHI_ADDIN.bat mot lan roi thu lai.\n"
            "      KHONG chay add-in qua HTTP thuong: Outlook se tu choi nap task pane\n"
            "      va du lieu email se di qua kenh khong ma hoa.\n")
    context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    context.minimum_version = ssl.TLSVersion.TLSv1_2
    context.load_cert_chain(certfile=str(cert_file), keyfile=str(key_file))
    return context


def create_server() -> Server:
    """Dựng server đã bọc TLS, chưa chạy.

    Tách khỏi main() để bản đóng gói chạy được backend ngay trong tiến trình bằng một
    thread: gọi lại chính exe onefile sẽ bung nén toàn bộ gói lần thứ hai.
    """
    global service
    load_dotenv(ENV_FILE if ENV_FILE.is_file() else None)
    context = _build_ssl_context()          # dừng hẳn trước khi làm gì khác nếu thiếu cert
    service = AddinService()

    server = Server((HOST, PORT), Handler)
    server.socket = context.wrap_socket(server.socket, server_side=True)
    return server


def main():
    logging.basicConfig(level=logging.INFO,
                        format="%(asctime)s [%(levelname)s] %(message)s")

    server = create_server()
    status = service.kb.status()

    print("=" * 64)
    print(f"  Tro ly Email - backend add-in v{VERSION}")
    print(f"  Dia chi     : https://localhost:{PORT}")
    print(f"  Model       : {GEMINI_MODEL}")
    print(f"  API key     : {'da cau hinh' if service.api.ai_engine.is_ready else 'CHUA CAU HINH'}")
    print(f"  Kho tri thuc: {status['file_count']} tep, {status['total_chars']} ky tu"
          f" ({'co' if status['dir_exists'] else 'THIEU THU MUC kien_thuc/'})")
    print(f"  Van phong   : {'da hoc, ' + str(status['sample_count']) + ' mau' if status['has_learned'] else 'chua hoc'}")
    print("-" * 64)
    print("  Luu y: moi lan khoi dong lai, token doi -> phai dong va mo lai task pane.")
    print("  Dung server: Ctrl+C")
    print("=" * 64, flush=True)

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nDa dung backend add-in.")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
