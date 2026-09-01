from __future__ import annotations

import logging
import traceback
from typing import Any, Dict, List, Optional

from dotenv import set_key

from config import (CACHE_DIR, DRAFT_FONT_FAMILY, DRAFT_FONT_SIZE,
                    DRAFT_TEXT_COLOR, GEMINI_API_KEY, GEMINI_MODEL,
                    KNOWLEDGE_DIR, STYLE_PROFILES_DIR)
from paths import ENV_FILE

from .ai_engine import AIEngine, AIEngineError
from .attachment_reader import read_attachments
from .classification_cache import ClassificationCache
from .com_worker import com_call
from .knowledge_base import KnowledgeBase
from .outlook_client import OutlookClient, strip_quoted_text
from .style_analyzer import StyleAnalyzer

logger = logging.getLogger(__name__)


def ok(data: Any = None) -> Dict:
    return {"success": True, "data": data, "error": None}


def fail(error: str) -> Dict:
    return {"success": False, "data": None, "error": error}


def _guard(fn):
    """Bọc mọi phương thức API: ghi log stack trace, trả lỗi dạng chuỗi cho UI."""
    def wrapper(*args, **kwargs):
        try:
            return fn(*args, **kwargs)
        except AIEngineError as e:
            logger.warning("Lỗi AI: %s", e)
            return fail(str(e))
        except Exception as e:
            logger.error("%s", traceback.format_exc())
            return fail(str(e) or e.__class__.__name__)
    wrapper.__name__ = fn.__name__
    return wrapper


class EmailAssistantAPI:
    @staticmethod
    def _build_engine(api_key: str) -> AIEngine:
        return AIEngine(api_key=api_key, model_name=GEMINI_MODEL,
                        font_family=DRAFT_FONT_FAMILY, font_size=DRAFT_FONT_SIZE,
                        text_color=DRAFT_TEXT_COLOR)

    def __init__(self) -> None:
        self.outlook = OutlookClient()
        self.ai_engine = self._build_engine(GEMINI_API_KEY)
        self.style_analyzer = StyleAnalyzer(data_dir=STYLE_PROFILES_DIR)
        self.cache = ClassificationCache(CACHE_DIR)
        self.kb = KnowledgeBase(KNOWLEDGE_DIR, STYLE_PROFILES_DIR / "default.json")

    # ------------------------------------------------------------------ Outlook

    @_guard
    def connect_outlook(self) -> Dict:
        if com_call(self.outlook.connect):
            return ok({"status": "connected"})
        return fail("Không kết nối được Outlook. Hãy mở Outlook rồi thử lại.")

    @_guard
    def get_folders(self) -> Dict:
        return ok(com_call(self.outlook.get_folders))

    @_guard
    def get_emails(self, folder_id: Optional[str] = None, limit: int = 30,
                   offset: int = 0) -> Dict:
        return ok(com_call(self.outlook.get_emails, folder_id, limit, offset))

    @_guard
    def search_emails(self, query: str, folder_id: Optional[str] = None) -> Dict:
        return ok(com_call(self.outlook.search_emails, query, folder_id))

    @_guard
    def get_email_detail(self, entry_id: str) -> Dict:
        detail = com_call(self.outlook.get_email_detail, entry_id)
        return ok(detail) if detail else fail("Không tìm thấy email.")

    @_guard
    def get_conversation(self, entry_id: str) -> Dict:
        return ok(com_call(self.outlook.get_conversation_thread, entry_id))

    @_guard
    def mark_as_read(self, entry_id: str) -> Dict:
        return ok({"marked": com_call(self.outlook.mark_as_read, entry_id)})

    # ------------------------------------------------------------------ AI soạn thư

    def _type_info(self, subject: str, body: str, email_type: str) -> Dict:
        """Cho giao diện biết cuối cùng đã dùng mẫu thư nào — tự khớp hay do chọn tay."""
        name = self.kb.match_email_type(subject, body, email_type) or ""
        return {"email_type": name,
                "email_type_title": self.kb.type_title(name) if name else "",
                "email_type_auto": not (email_type or "").strip()}

    @_guard
    def list_email_types(self) -> Dict:
        return ok({"types": self.kb.list_types()})

    @_guard
    def generate_reply(self, entry_id: str, instruction: str,
                       reply_all: bool = False, email_type: str = "") -> Dict:
        detail = com_call(self.outlook.get_email_detail, entry_id)
        if not detail:
            return fail("Không tìm thấy email để tạo phản hồi.")
        thread = com_call(self.outlook.get_conversation_thread, entry_id)
        profile = self.style_analyzer.load_profile("default")
        subject, body = detail.get("subject", ""), detail.get("body", "")
        html = self.ai_engine.generate_reply(
            detail, thread, instruction, profile,
            knowledge=self.kb.build_prompt_block(
                detail.get("sender_email", ""), subject, body, email_type))
        return ok({"html_body": html} | self._type_info(subject, body, email_type))

    @_guard
    def generate_reply_from_payload(self, email: Dict, instruction: str = "",
                                    attachments: Optional[List[Dict]] = None,
                                    me: Optional[Dict] = None,
                                    email_type: str = "") -> Dict:
        """Đường dành cho add-in: nội dung thư do Office.js cung cấp, KHÔNG dùng COM.

        Task pane gửi lên toàn bộ thân thư (gồm cả phần trích dẫn), nên phải tách ở
        đây — nếu không, generate_reply cắt ở 4000 ký tự và thư dài sẽ chỉ còn trích dẫn.
        """
        email = dict(email or {})
        full_body = email.get("body") or ""
        email["body"] = strip_quoted_text(full_body)
        if len(full_body) > len(email["body"]):
            email["quoted_tail"] = full_body[len(email["body"]):].strip()

        parsed = read_attachments(attachments)
        subject = email.get("subject", "")
        html = self.ai_engine.generate_reply(
            email, [], instruction, self.kb.style_profile(),
            knowledge=self.kb.build_prompt_block(
                email.get("sender_email", ""), subject, email["body"], email_type),
            attachment_text=parsed.text,
            attachment_blobs=parsed.blobs,
            me=me)
        return ok({"html_body": html,
                   "notes": parsed.notes,
                   "attachments_used": parsed.used}
                  | self._type_info(subject, email["body"], email_type))

    @_guard
    def knowledge_status(self) -> Dict:
        return ok(self.kb.status())

    @_guard
    def refine_draft(self, draft_html: str, feedback: str) -> Dict:
        return ok({"html_body": self.ai_engine.refine_draft(draft_html, feedback)})

    @_guard
    def save_draft(self, entry_id: str, html_body: str,
                   reply_all: bool = False) -> Dict:
        result = com_call(self.outlook.create_draft_reply, entry_id, html_body, reply_all)
        return ok(result) if result else fail("Không lưu được thư nháp.")

    # ------------------------------------------------------------------ Phân loại

    @_guard
    def classify_emails(self, emails: List[Dict]) -> Dict:
        """Nhận danh sách summary (đã có sẵn ở frontend) -> trả map entry_id -> nhãn.

        Frontend gửi lại summary thay vì entry_id để tránh phải đọc lại Outlook.
        """
        if not emails:
            return ok({})

        ids = [e.get("entry_id", "") for e in emails if e.get("entry_id")]
        cached = self.cache.get_many(ids)
        pending = [e for e in emails if e.get("entry_id") not in cached]

        if pending:
            fresh = self.ai_engine.classify_emails(pending)
            self.cache.put_many(fresh)
            cached.update(fresh)
        return ok(cached)

    @_guard
    def classify_all_emails(self, folder_id: Optional[str] = None) -> Dict:
        """Quét và phân loại toàn bộ thư trong thư mục, dùng cache theo EntryID."""
        all_emails = []
        offset = 0
        while True:
            page = com_call(self.outlook.get_emails, folder_id, 100, offset)
            all_emails.extend(page.get("items", []))
            if not page.get("has_more"):
                break
            next_offset = page.get("next_offset", offset)
            if next_offset <= offset:
                break
            offset = next_offset
        result = self.classify_emails(all_emails)
        if not result.get("success"):
            return result
        return ok({"classifications": result.get("data", {}), "scanned": len(all_emails)})

    # ------------------------------------------------------------------ Văn phong

    @_guard
    def export_sent_emails(self, progress=None, deep: bool = True) -> Dict:
        """Xuất toàn bộ thư đã gửi ra Markdown để phân tích bằng công cụ AI ngoài.

        KHÔNG gọi AI lần nào. Thay cho analyze_style cũ vốn tốn 16 lượt gọi Gemini mà
        vẫn cho ra hồ sơ khuôn trống, vì bước hợp nhất nhiều hồ sơ con làm mất chi tiết.

        Đồng thời cập nhật hồ sơ văn phong bằng SỐ LIỆU ĐẾM ĐƯỢC (Python, chính xác
        tuyệt đối) — hồ sơ này vẫn được nạp vào prompt mỗi lần soạn thư.
        """
        from xuat_thu_da_gui import export        # nạp muộn: tránh import vòng

        def report(phase, current=0, total=0, detail=""):
            if progress:
                try:
                    progress(phase, current, total, detail)
                except Exception:      # lỗi vẽ tiến độ không được làm hỏng việc xuất
                    logger.debug("Bỏ qua lỗi callback tiến độ", exc_info=True)

        # export() tự ghi hồ sơ văn phong (schema v3, thuần Python) nên chạy
        # XUAT_THU.bat trực tiếp cũng cập nhật được.
        result = export(deep=deep, progress=report)
        self.kb.reload(force=True)

        stats = result.get("stats") or {}
        return ok({k: v for k, v in result.items() if k != "stats"} |
                  {"analyzed_count": stats.get("analyzed_count", 0)})

    @_guard
    def get_style_profile(self) -> Dict:
        profile = self.style_analyzer.load_profile("default")
        return ok(profile) if profile else fail("Chưa có hồ sơ văn phong.")

    # ------------------------------------------------------------------ Cài đặt

    @_guard
    def get_settings(self) -> Dict:
        profile = self.style_analyzer.load_profile("default")
        return ok({
            "api_key_configured": self.ai_engine.is_ready,
            "model": self.ai_engine.model_name,
            "style_analyzed_at": (profile or {}).get("analyzed_at"),
        })

    @_guard
    def save_api_key(self, api_key: str) -> Dict:
        api_key = (api_key or "").strip()
        if not api_key:
            return fail("API key trống.")
        ENV_FILE.parent.mkdir(parents=True, exist_ok=True)
        ENV_FILE.touch(exist_ok=True)       # set_key cần tệp có sẵn mới ghi được
        set_key(str(ENV_FILE), "GEMINI_API_KEY", api_key)
        self.ai_engine = self._build_engine(api_key)
        return ok({"api_key_configured": self.ai_engine.is_ready})
