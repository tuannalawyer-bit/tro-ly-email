from __future__ import annotations

import logging
import os
import re
from typing import Any, Callable, Dict, List, Optional, Tuple

from dotenv import set_key

from config import ENV_FILE, GEMINI_API_KEY
from backend.ai_engine import AIEngine, AIEngineError
from backend.attachment_reader import read_attachments
from backend.classification_cache import ClassificationCache
from backend.com_worker import com_call
from backend.knowledge_base import KnowledgeBase
from backend.outlook_client import OutlookClient, strip_quoted_text
from backend.web_search import WebSearcher

logger = logging.getLogger(__name__)


def ok(data: Any = None) -> Dict[str, Any]:
    return {"success": True, "data": data}


def fail(error: str) -> Dict[str, Any]:
    return {"success": False, "error": error}


def _guard(fn: Callable) -> Callable:
    """Decorator bọc xử lý lỗi chung cho các hàm API."""
    def wrapper(*args, **kwargs):
        try:
            return fn(*args, **kwargs)
        except AIEngineError as e:
            return fail(str(e))
        except Exception as e:
            logger.exception("API Error in %s: %s", fn.__name__, e)
            return fail(f"Lỗi: {e}")
    return wrapper


class EmailAssistantAPI:
    def _build_engine(self, key: str, model_name: str = "") -> AIEngine:
        from config import (DRAFT_FONT_FAMILY, DRAFT_FONT_SIZE,
                            DRAFT_TEXT_COLOR, GEMINI_MODEL)
        target_model = model_name or GEMINI_MODEL or "gemini-3.5-flash-lite"
        return AIEngine(
            api_key=key,
            model_name=target_model,
            font_family=DRAFT_FONT_FAMILY,
            font_size=DRAFT_FONT_SIZE,
            text_color=DRAFT_TEXT_COLOR,
        )

    def __init__(self, api_key: str = "", serper_key: str = "", google_cse_key: str = "", google_cse_id: str = "") -> None:
        self.outlook = OutlookClient()
        self.kb = KnowledgeBase()
        effective_key = api_key or GEMINI_API_KEY
        self.ai_engine = self._build_engine(effective_key)
        self.web_searcher = WebSearcher(serper_key=serper_key, google_cse_key=google_cse_key, google_cse_id=google_cse_id)
        self.classifier_cache = ClassificationCache()

    def _get_engine(self) -> AIEngine:
        from dotenv import load_dotenv
        from paths import ENV_FILE
        if ENV_FILE.is_file():
            load_dotenv(ENV_FILE, override=True)
        key = os.getenv("GEMINI_API_KEY", "").strip()
        model = os.getenv("GEMINI_MODEL", "gemini-3.5-flash-lite").strip()
        if (
            self.ai_engine is None
            or not self.ai_engine.is_ready
            or key != getattr(self.ai_engine, "api_key", "")
            or model != getattr(self.ai_engine, "model_name", "")
        ):
            self.ai_engine = self._build_engine(key, model_name=model)
        return self.ai_engine

    # ------------------------------------------------------------------ outlook

    def connect_outlook(self) -> Dict[str, Any]:
        """Thử kết nối tới Outlook COM."""
        success = com_call(self.outlook.connect)
        if success:
            return ok("Đã kết nối Outlook thành công.")
        return fail("Không thể kết nối Outlook. Vui lòng mở Outlook trước.")

    def get_folders(self) -> Dict[str, Any]:
        """Lấy danh sách thư mục email."""
        folders = com_call(self.outlook.get_folders)
        return ok(folders)

    def get_emails(self, folder_id: str, offset: int = 0, limit: int = 30) -> Dict[str, Any]:
        """Lấy danh sách email tóm tắt từ một thư mục."""
        result = com_call(self.outlook.get_emails, folder_id, offset, limit)
        return ok(result)

    def search_emails(self, query: str, folder_id: str = "inbox", limit: int = 30) -> Dict[str, Any]:
        """Tìm kiếm email."""
        result = com_call(self.outlook.search_emails, query, folder_id, limit)
        return ok(result)

    def get_email_detail(self, entry_id: str) -> Dict[str, Any]:
        """Lấy chi tiết email bao gồm nội dung đầy đủ và danh sách đính kèm."""
        result = com_call(self.outlook.get_email_detail, entry_id)
        if not result:
            return fail("Không tìm thấy email hoặc không thể đọc nội dung.")
        return ok(result)

    def get_conversation(self, entry_id: str) -> Dict[str, Any]:
        """Lấy toàn bộ chuỗi hội thoại của email."""
        thread = com_call(self.outlook.get_conversation_thread, entry_id)
        return ok(thread)

    def mark_as_read(self, entry_id: str) -> Dict[str, Any]:
        """Đánh dấu email là đã đọc."""
        success = com_call(self.outlook.mark_as_read, entry_id)
        return ok(success)

    # ------------------------------------------------------------------ soạn thư

    def _type_info(self, subject: str, body: str, email_type: str = "") -> dict:
        """Cho giao diện biết cuối cùng đã dùng mẫu thư nào — tự khớp hay do chọn tay."""
        type_to_match = email_type or "tham-dinh-mat-bang"
        name = self.kb.match_email_type(subject, body, type_to_match) or "tham-dinh-mat-bang"
        return {
            "email_type": name,
            "email_type_title": self.kb.type_title(name) if name else "Thẩm định mặt bằng & mở điểm mới",
            "email_type_auto": not bool(email_type and email_type.strip()),
        }

    def list_email_types(self) -> Dict[str, Any]:
        """Danh sách các mẫu thư trong kho tri thức để người dùng tự chọn."""
        return ok(self.kb.list_types())

    @_guard
    def generate_reply(self, entry_id: str, instruction: str = "",
                       reply_all: bool = False,
                       email_type: str = "") -> Dict[str, Any]:
        """Soạn thư phản hồi bằng AI dựa trên ngữ cảnh email, đính kèm và kho tri thức."""
        detail = com_call(self.outlook.get_email_detail, entry_id)
        if not detail:
            return fail("Không tìm thấy email.")

        thread = com_call(self.outlook.get_conversation_thread, entry_id)
        profile = self.kb.style_profile()

        subject = detail.get("subject", "")
        body = detail.get("body", "")

        knowledge, rental_table_html, coords = self._build_knowledge_with_presearch(
            sender_email=detail.get("sender_email", ""),
            subject=subject,
            body=body,
            attachment_text="",
            email_type=email_type,
            instruction=instruction,
        )

        html = self._get_engine().generate_reply(
            email_data=detail,
            thread=thread,
            instruction=instruction,
            style_profile=profile,
            knowledge=knowledge,
            attachment_text="",
            attachment_blobs=[],
        )

        if coords:
            html = self._inject_planning_links(html, coords[0], coords[1])

        if rental_table_html:
            html = self._inject_rental_table(html, rental_table_html)

        return ok({"html_body": html} | self._type_info(subject, body, email_type))

    @_guard
    def generate_reply_from_payload(self, email: Dict[str, Any],
                                    instruction: str = "",
                                    attachments: Optional[List[Dict[str, Any]]] = None,
                                    me: Optional[Dict[str, Any]] = None,
                                    email_type: str = "") -> Dict[str, Any]:
        """Soạn thư phản hồi từ payload do Outlook Add-in gửi lên."""
        email = dict(email or {})
        full_body = email.get("body") or ""
        email["body"] = strip_quoted_text(full_body)
        if len(full_body) > len(email["body"]):
            email["quoted_tail"] = full_body[len(email["body"]):].strip()

        parsed = read_attachments(attachments)
        subject = email.get("subject", "")

        knowledge, rental_table_html, coords = self._build_knowledge_with_presearch(
            sender_email=email.get("sender_email", ""),
            subject=subject,
            body=email["body"],
            attachment_text=parsed.text,
            email_type=email_type,
            instruction=instruction,
        )

        html = self._get_engine().generate_reply(
            email_data=email,
            thread=[],
            instruction=instruction,
            style_profile=self.kb.style_profile(),
            knowledge=knowledge,
            attachment_text=parsed.text,
            attachment_blobs=parsed.blobs,
            me=me,
        )

        if coords:
            html = self._inject_planning_links(html, coords[0], coords[1])

        if rental_table_html:
            html = self._inject_rental_table(html, rental_table_html)

        return ok({
            "html_body": html,
            "notes": parsed.notes,
            "attachments_used": parsed.used,
        } | self._type_info(subject, email["body"], email_type))

    # ------------------------------------------------------------------ hỗ trợ tra cứu mặt bằng & quy hoạch

    def _build_knowledge_with_presearch(self, sender_email: str, subject: str, body: str,
                                        attachment_text: str, email_type: str,
                                        instruction: str = "") -> Tuple[str, str, Optional[Tuple[float, float]]]:
        """Tạo khối knowledge + bảng HTML giá thuê để inject sau vào response.

        Returns:
            (knowledge_str, rental_table_html, coords)
        """
        knowledge_block = self.kb.build_prompt_block(
            sender_email=sender_email,
            subject=subject,
            body=body,
            email_type=email_type or "tham-dinh-mat-bang",
        )

        parts: List[str] = []
        rental_table_html = ""

        combined_text = f"{subject} {body} {attachment_text} {instruction}"
        coords = self._extract_coordinates(combined_text)

        if not coords:
            hints_pre = self._extract_location_hints(subject, body, attachment_text, coords=None)
            for hint in hints_pre:
                geo = self._forward_geocode(hint)
                if geo:
                    coords = geo
                    logger.info("Forward geocoded '%s' -> (%.6f, %.6f)", hint, coords[0], coords[1])
                    break

        if coords:
            lat, lng = coords[0], coords[1]
            guland_url = f"https://guland.vn/soi-quy-hoach?lat={lat:.6f}&lng={lng:.6f}&zoom=18"
            meeymap_url = f"https://meeymap.com/?lat={lat:.6f}&lng={lng:.6f}"
            gmaps_url = f"https://www.google.com/maps/search/?api=1&query={lat:.6f},{lng:.6f}"

            map_links_block = (
                f"### Bản đồ soi quy hoạch trực tiếp theo tọa độ GPS ({lat:.6f}, {lng:.6f})\n"
                f"- [Soi quy hoạch Guland]({guland_url}): Bản đồ quy hoạch phân khu & chỉ giới mở rộng đường tại vị trí mặt bằng.\n"
                f"- [Bản đồ quy hoạch MeeyMap]({meeymap_url}): Lớp quy hoạch kế hoạch sử dụng đất và giao thông.\n"
                f"- [Định vị vệ tinh Google Maps]({gmaps_url}): Xem không gian thực địa và mật độ dân cư xung quanh.\n\n"
                f"*HƯỚNG DẪN AI: Tại mục 'Kiểm tra quy hoạch' trong email, BẮT BUỘC chèn liên kết <a href=\"{guland_url}\">Soi quy hoạch Guland</a> và <a href=\"{meeymap_url}\">Bản đồ MeeyMap</a> để người thẩm định có thể bấm 1-click mở ngay bản đồ kiểm tra chỉ giới mở đường thực tế.*"
            )
            parts.append(map_links_block)

            try:
                from backend.map_vision import capture_map_screenshot, analyze_map_with_vision
                img_bytes = capture_map_screenshot(lat, lng, timeout_sec=15)
                if img_bytes:
                    vision_analysis = analyze_map_with_vision(
                        api_key=self.ai_engine.api_key,
                        model_name=self.ai_engine.model_name,
                        img_bytes=img_bytes,
                        lat=lat,
                        lng=lng,
                    )
                    if vision_analysis:
                        logger.info("AI Vision đọc ảnh bản đồ thành công (%d chars)", len(vision_analysis))
                        parts.append(
                            f"### Kết quả phân tích ảnh chụp bản đồ quy hoạch bằng AI Vision (Tọa độ: {lat:.6f}, {lng:.6f})\n"
                            f"{vision_analysis}\n\n"
                            f"*HƯỚNG DẪN AI: Sử dụng thông tin trên ảnh bản đồ quy hoạch này để bổ sung chi tiết vào mục 'Kiểm tra quy hoạch'.*"
                        )
            except Exception as exc:
                logger.warning("Map vision error: %s", exc)

        location_hints = self._extract_location_hints(subject, body, attachment_text, coords=coords)

        if self.web_searcher.enabled and location_hints:
            all_rental: List[Dict[str, Any]] = []
            all_planning: List[Dict[str, Any]] = []
            all_project: List[Dict[str, Any]] = []

            for hint in location_hints[:2]:
                r = self.web_searcher.search_rental_prices(hint, num=6)
                all_rental.extend(r)
                p = self.web_searcher.search_planning_info(hint, num=5)
                all_planning.extend(p)
                pj = self.web_searcher.search_project_info(hint, num=3)
                all_project.extend(pj)

            def dedup(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
                seen = set()
                out = []
                for it in items:
                    link = it.get("link", "")
                    if link and link not in seen:
                        seen.add(link)
                        out.append(it)
                return out

            all_rental = dedup(all_rental)[:5]
            all_planning = dedup(all_planning)[:5]
            all_project = dedup(all_project)[:3]

            hints_str = location_hints[0] if location_hints else ""

            if all_rental:
                rental_table_html = self.web_searcher.build_rental_table_html(all_rental, location_hint=hints_str)
                parts.append(self.web_searcher.format_rental_context_for_prompt(all_rental, label=f"Giá thuê mặt bằng thực tế tại '{hints_str}' (link cụ thể từ Google Search)"))

            if all_planning:
                parts.append(self.web_searcher.format_results_for_prompt(all_planning, label=f"Quy hoạch mở rộng đường / Chỉ giới đường đỏ tại '{hints_str}'"))

            if all_project:
                parts.append(self.web_searcher.format_results_for_prompt(all_project, label=f"Thông tin dự án / khu vực '{hints_str}'"))

        elif self.web_searcher.enabled:
            search_query = self._extract_search_query(subject, body, instruction)
            if search_query:
                results = self.web_searcher.search_general(search_query, num=4)
                if results:
                    parts.append(self.web_searcher.format_results_for_prompt(results, label=f"Kết quả tra cứu thông tin cho '{search_query}' (link từ Google Search)"))

        if (email_type == "tham-dinh-mat-bang" or not email_type) and not rental_table_html:
            hints_str = location_hints[0] if (location_hints and len(location_hints) > 0) else (subject or "Đông Anh, Hà Nội")
            rental_table_html = self.web_searcher.build_rental_table_html([], location_hint=hints_str)

        if parts:
            extra_search_context = "\n\n".join(parts)
            logger.info("Pre-search/Planning context OK — length: %d chars", len(extra_search_context))
            knowledge = (
                f"{knowledge_block}\n\n---\n## KẾT QUẢ TÌM KIẾM THỰC TẾ & BẢN ĐỒ QUY HOẠCH (DÙNG CHO BÁO CÁO)\n"
                f"Các đường link dưới đây được hệ thống trích xuất trực tiếp từ bản đồ quy hoạch và Google Search. "
                f"AI BẮT BUỘC chèn chính xác các URL này vào báo cáo bằng thẻ liên kết <a href=\"...\">, "
                f"TUYỆT ĐỐI KHÔNG tự tạo link ảo khác:\n\n{extra_search_context}"
            )
            return knowledge, rental_table_html, coords

        return knowledge_block, rental_table_html, coords

    @classmethod
    def _inject_rental_table(cls, html: str, table_html: str) -> str:
        """Chèn bảng HTML giá thuê vào ngay dưới mục - Giá thuê trong email."""
        if not table_html or not html:
            return html
        if "<table" in html:
            return html

        patterns = [
            r'(<(?:p|li|div)[^>]*>[\s\S]*?-\s*Giá thuê:[\s\S]*?</(?:p|li|div)>)',
            r'(<(?:p|li|div)[^>]*>[\s\S]*?-\s*giá thuê:[\s\S]*?</(?:p|li|div)>)',
            r'(-\s*Giá thuê:[^\n<]*?(?:</p>|<br\s*/?>))',
            r'(Giá thuê:[^\n<]*?(?:</p>|<br\s*/?>))',
            r'(<(?:p|li|div)[^>]*>.*?Bảng khảo sát.*?</(?:p|li|div)>)',
        ]
        for pat in patterns:
            m = re.search(pat, html, re.IGNORECASE)
            if m:
                idx = m.end()
                return html[:idx] + "\n" + table_html + "\n" + html[idx:]

        # Nếu không khớp mẫu trên, chèn trước mục Diễn giải tiến trình hoặc Đánh giá
        m_next = re.search(r'(<(?:p|li|div)[^>]*>.*?(?:Diễn giải tiến trình|Đánh giá &amp; Trách nhiệm|Đánh giá & Trách nhiệm))', html, re.IGNORECASE)
        if m_next:
            idx = m_next.start()
            return html[:idx] + table_html + "\n" + html[idx:]

        # Chèn trước thẻ đóng div cuối cùng
        if html.rstrip().endswith("</div>"):
            idx = html.rfind("</div>")
            return html[:idx] + table_html + html[idx:]

        return html + "\n" + table_html

    @classmethod
    def _inject_planning_links(cls, html: str, lat: float, lng: float) -> str:
        """Tự động chèn link Soi quy hoạch Guland và MeeyMap vào mục Quy hoạch trong HTML."""
        if not lat or not lng or not html:
            return html

        guland_url = f"https://guland.vn/soi-quy-hoach?lat={lat:.6f}&lng={lng:.6f}&zoom=18"
        meeymap_url = f"https://meeymap.com/?lat={lat:.6f}&lng={lng:.6f}"
        gmaps_url = f"https://www.google.com/maps/search/?api=1&query={lat:.6f},{lng:.6f}"

        guland_link = f'<a href="{guland_url}" style="color:#1155CC;font-weight:bold;">Soi quy hoạch Guland</a>'
        meeymap_link = f'<a href="{meeymap_url}" style="color:#1155CC;font-weight:bold;">Bản đồ MeeyMap</a>'
        gmaps_link = f'<a href="{gmaps_url}" style="color:#1155CC;">Định vị Google Maps</a>'

        links_inline = f" (Tra cứu trực tiếp: {guland_link} | {meeymap_link} | {gmaps_link})"

        if "guland.vn" in html and "meeymap.com" in html:
            return html

        m_planning = re.search(r'(<(?:li|p|strong|b)[^>]*>.*?quy hoạch.*?)(</(?:li|p|strong|b)>)', html, re.IGNORECASE)
        if m_planning:
            idx = m_planning.start(2)
            return html[:idx] + links_inline + html[idx:]

        m_header = re.search(r'(Kiểm tra quy hoạch|Quy hoạch & Xây dựng|Quy hoạch)', html, re.IGNORECASE)
        if m_header:
            idx = m_header.end(1)
            return html[:idx] + links_inline + html[idx:]

        return html

    @classmethod
    def _extract_coordinates(cls, text: str) -> Optional[Tuple[float, float]]:
        """Trích xuất cặp tọa độ GPS (Vĩ độ, Kinh độ) từ mọi định dạng văn bản hoặc link Google Maps."""
        if not text:
            return None

        m_short = re.search(r'https?://(?:maps\.app\.goo\.gl|goo\.gl/maps)/[a-zA-Z0-9_\-]+', text)
        if m_short:
            try:
                import urllib.request
                req = urllib.request.Request(m_short.group(0), headers={"User-Agent": "Mozilla/5.0"})
                with urllib.request.urlopen(req, timeout=4) as resp:
                    final_url = resp.geturl()
                    m_coords = re.search(r'[@?&]q?=?([+-]?\d{1,2}\.\d{4,16})[,\s]+([+-]?\d{2,3}\.\d{4,16})', final_url)
                    if m_coords:
                        v1 = float(m_coords.group(1))
                        v2 = float(m_coords.group(2))
                        if 8.0 <= v1 <= 24.0 and 102.0 <= v2 <= 110.0:
                            return (v1, v2)
                        elif 8.0 <= v2 <= 24.0 and 102.0 <= v1 <= 110.0:
                            return (v2, v1)
            except Exception:
                pass

        m_url = re.search(r'(?:google\.com/maps[^\s"\'<>]*[@?&]q?=?|maps\.app\.goo\.gl[^\s"\'<>]*|maps\.google\.com[^\s"\'<>]*[?&]q=)([+-]?\d{1,2}\.\d{4,16})[,\s]+([+-]?\d{2,3}\.\d{4,16})', text)
        if m_url:
            try:
                v1 = float(m_url.group(1))
                v2 = float(m_url.group(2))
                if 8.0 <= v1 <= 24.0 and 102.0 <= v2 <= 110.0:
                    return (v1, v2)
                elif 8.0 <= v2 <= 24.0 and 102.0 <= v1 <= 110.0:
                    return (v2, v1)
            except Exception:
                pass

        m_lat = re.search(r'(?:vĩ\s*độ|lat|latitude)[\s:=]*([+-]?\d{1,2}\.\d{4,16})', text, re.IGNORECASE)
        m_lng = re.search(r'(?:kinh\s*độ|lng|long|longitude)[\s:=]*([+-]?\d{2,3}\.\d{4,16})', text, re.IGNORECASE)
        if m_lat and m_lng:
            try:
                lat = float(m_lat.group(1))
                lng = float(m_lng.group(1))
                if 8.0 <= lat <= 24.0 and 102.0 <= lng <= 110.0:
                    return (lat, lng)
            except Exception:
                pass

        m_dms = re.search(r'(\d{1,2})[°\s]+(\d{1,2})[\'\s]+(\d{1,2}(?:\.\d+)?)\"?\s*([NS])[\s,]+(\d{2,3})[°\s]+(\d{1,2})[\'\s]+(\d{1,2}(?:\.\d+)?)\"?\s*([EW])', text, re.IGNORECASE)
        if m_dms:
            try:
                lat = float(m_dms.group(1)) + float(m_dms.group(2))/60 + float(m_dms.group(3))/3600
                if m_dms.group(4).upper() == 'S':
                    lat = -lat
                lng = float(m_dms.group(5)) + float(m_dms.group(6))/60 + float(m_dms.group(7))/3600
                if m_dms.group(8).upper() == 'W':
                    lng = -lng
                if 8.0 <= lat <= 24.0 and 102.0 <= lng <= 110.0:
                    return (lat, lng)
            except Exception:
                pass

        m_kw = re.search(r'(?:tọa độ|tọa độ map|tọa độ gps|lat|gps|vị trí map|vị trí gps)[\s:=]*([+-]?\d{1,3}\.\d{4,16})[,\s;]+([+-]?\d{1,3}\.\d{4,16})', text, re.IGNORECASE)
        if m_kw:
            try:
                v1 = float(m_kw.group(1))
                v2 = float(m_kw.group(2))
                if 8.0 <= v1 <= 24.0 and 102.0 <= v2 <= 110.0:
                    return (v1, v2)
                elif 8.0 <= v2 <= 24.0 and 102.0 <= v1 <= 110.0:
                    return (v2, v1)
            except Exception:
                pass

        m_direct = re.search(r'\b(2[0-3]\.\d{4,16})[,\s;]+(10[2-9]\.\d{4,16})\b', text)
        if m_direct:
            try:
                lat = float(m_direct.group(1))
                lng = float(m_direct.group(2))
                return (lat, lng)
            except Exception:
                pass

        m_rev = re.search(r'\b(10[2-9]\.\d{4,16})[,\s;]+(2[0-3]\.\d{4,16})\b', text)
        if m_rev:
            try:
                lng = float(m_rev.group(1))
                lat = float(m_rev.group(2))
                return (lat, lng)
            except Exception:
                pass

        return None

    @classmethod
    def _forward_geocode(cls, query: str) -> Optional[Tuple[float, float]]:
        """Tra cứu tọa độ GPS từ địa chỉ hành chính nếu không có GPS sẵn trong thư."""
        if not query or len(query.strip()) < 4:
            return None
        import urllib.parse
        import urllib.request
        import json

        parts = [p.strip() for p in query.split(",") if p.strip()]
        search_terms = [query]
        if len(parts) >= 2:
            search_terms.append(", ".join(parts[-2:]))

        for term in search_terms:
            try:
                url = f"https://nominatim.openstreetmap.org/search?format=json&q={urllib.parse.quote(term)}&limit=1"
                req = urllib.request.Request(url, headers={"User-Agent": "EmailAssistant/2.0"})
                with urllib.request.urlopen(req, timeout=3) as resp:
                    data = json.loads(resp.read().decode("utf-8"))
                    if data:
                        lat = float(data[0]["lat"])
                        lon = float(data[0]["lon"])
                        if 8.0 <= lat <= 24.0 and 102.0 <= lon <= 110.0:
                            return (lat, lon)
            except Exception:
                pass
        return None

    @classmethod
    def _reverse_geocode(cls, lat: float, lng: float) -> Optional[str]:
        """Tra cứu địa danh hành chính (xã/phường, quận/huyện, tỉnh/tp) từ tọa độ GPS."""
        try:
            import urllib.request
            import json

            url = f"https://nominatim.openstreetmap.org/reverse?format=json&lat={lat}&lon={lng}"
            req = urllib.request.Request(url, headers={"User-Agent": "EmailAssistant/2.0"})
            with urllib.request.urlopen(req, timeout=3) as resp:
                data = json.loads(resp.read().decode("utf-8"))
                addr = data.get("address", {})
                road = addr.get("road", "")
                ward = addr.get("village") or addr.get("suburb") or addr.get("quarter") or ""
                district = addr.get("county") or addr.get("district") or addr.get("town") or addr.get("city_district") or ""
                state = addr.get("state") or addr.get("city") or ""

                parts = [p.strip() for p in (road, ward, district, state) if p.strip()]
                if parts:
                    return ", ".join(parts)
        except Exception:
            pass
        return None

    @classmethod
    def _extract_location_hints(cls, subject: str, body: str, attachment_text: str = "",
                               coords: Optional[Tuple[float, float]] = None) -> List[str]:
        """Trích xuất địa điểm/khu vực chính xác của mặt bằng.

        QUY TẮC BẮT BUỘC:
        1. ƯU TIÊN SỐ 1 TUYỆT ĐỐI là TỌA ĐỘ GPS (nếu có):
           - Dùng tọa độ GPS tra cứu chính xác địa danh thực tế (Tên đường, Xã/Phường, Huyện/Quận, Tỉnh/TP).
           - Toàn bộ khảo sát giá thuê và kiểm tra quy hoạch sẽ thực hiện theo vị trí thực tế của tọa độ GPS này.
        2. ƯU TIÊN 2 (Chỉ khi KHÔNG có GPS): Mới quét địa chỉ trong Body / Subject / Attachment.
        """
        import re

        def _clean_addr(raw: str) -> str:
            h = re.sub(
                r'^(?:thẩm định|an giá|thực địa|nso|mbmm|mb\s*r?\d+[\w_]*|wm\+?|hni|hcm|địa chỉ|đ/c|tại|vị trí|mặt bằng)[\s:=_-]*',
                '', raw, flags=re.IGNORECASE
            ).strip()
            h = re.split(
                r'\s+(?:tọa độ|tọa độ map|lat|gps|tkmb|cụ thể|diện tích|dtt|mã mb|tên mb|giá thuê|trưởng nhóm|thông tin|chủ nhà|bct)\b',
                h, flags=re.IGNORECASE
            )[0].strip(' ,.-')
            return h

        primary_hints: List[str] = []

        if coords:
            geo_loc = EmailAssistantAPI._reverse_geocode(coords[0], coords[1])
            if geo_loc:
                clean_geo = _clean_addr(geo_loc)
                if clean_geo:
                    primary_hints.append(clean_geo)
                    parts = [p.strip() for p in clean_geo.split(',') if p.strip()]
                    if len(parts) >= 3:
                        broad = ", ".join(parts[1:])
                        if broad not in primary_hints:
                            primary_hints.append(broad)
                    return primary_hints

        pattern_addr = (
            r'((?:thôn|xóm|phố|đường|xã|phường|thị trấn|quận|huyện|tp|tỉnh|p\.|q\.|tx\.|tt\.|h\.|x\.)\s+[^,;\n\r]{2,30}'
            r'(?:,\s*(?:thôn|xóm|phố|đường|xã|phường|thị trấn|quận|huyện|tp|tỉnh|p\.|q\.|tx\.|tt\.|h\.|x\.)?\s*[^,;\n\r]{2,30}){1,3})'
        )

        for text_source in (body, subject, attachment_text):
            if not text_source:
                continue
            for m in re.finditer(pattern_addr, text_source, re.IGNORECASE):
                c = _clean_addr(m.group(0))
                if c and len(c) > 6 and c not in primary_hints:
                    primary_hints.append(c)
                    if len(primary_hints) >= 2:
                        return primary_hints

        return primary_hints

    @classmethod
    def _extract_search_query(cls, subject: str, body: str, instruction: str = "") -> str:
        """Trích xuất từ khóa tìm kiếm động khi người dùng yêu cầu tra cứu."""
        import re

        full = f"{instruction} {subject}".strip()
        m = re.search(r'(?:tra cứu|tìm kiếm|kiểm tra|tìm hiểu|xem thông tin)\s+(?:về\s+)?([^,.?!;]+)', full, re.IGNORECASE)
        if m:
            return m.group(1).strip()
        if instruction.strip() and len(instruction.strip()) < 80:
            return instruction.strip()
        if subject.strip():
            return subject.strip()
        return ""

    # ------------------------------------------------------------------ kho tri thức & học văn phong

    def knowledge_status(self) -> Dict[str, Any]:
        """Tình trạng kho tri thức."""
        return ok(self.kb.status())

    @_guard
    def refine_draft(self, draft_html: str, feedback: str) -> Dict[str, Any]:
        """Tinh chỉnh thư nháp theo phản hồi của người dùng."""
        profile = self.kb.style_profile()
        new_html = self._get_engine().refine_draft(draft_html, feedback, style_profile=profile)
        return ok({"html_body": new_html})

    def save_draft(self, entry_id: str, html_body: str, reply_all: bool = False) -> Dict[str, Any]:
        """Lưu nội dung vào thư nháp trong Outlook."""
        result = com_call(self.outlook.create_draft_reply, entry_id, html_body, reply_all)
        if result:
            return ok(result)
        return fail("Không lưu được thư nháp.")

    # ------------------------------------------------------------------ phân loại & học văn phong

    def classify_emails(self, emails: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Nhận danh sách summary (đã có sẵn ở frontend) -> trả map entry_id -> nhãn.

        Frontend gửi lại summary thay vì entry_id để tránh phải đọc lại Outlook.
        """
        if not emails:
            return ok({})

        ids = [e["entry_id"] for e in emails if e.get("entry_id")]
        cached = self.classifier_cache.get_many(ids)

        pending = [e for e in emails if e.get("entry_id") and e["entry_id"] not in cached]
        engine = self._get_engine()
        if pending and engine.is_ready:
            fresh = engine.classify_email_batch(pending)
            self.classifier_cache.set_many(fresh)
            cached.update(fresh)

        return ok(cached)

    def classify_all_emails(self, folder_id: str = "inbox") -> Dict[str, Any]:
        """Quét và phân loại toàn bộ thư trong thư mục, dùng cache theo EntryID."""
        all_emails: List[Dict[str, Any]] = []
        offset = 0
        while True:
            page = com_call(self.outlook.get_emails, folder_id, offset, 50)
            items = page.get("items", [])
            all_emails.extend(items)
            if not page.get("has_more"):
                break
            offset = page.get("next_offset", offset + len(items))

        result = self.classify_emails(all_emails)
        if result.get("success"):
            return ok(result.get("data", {}))
        return result

    def export_sent_emails(self, progress: Optional[Callable[[Dict[str, Any]], None]] = None,
                           deep: bool = True) -> Dict[str, Any]:
        """Xuất toàn bộ thư đã gửi ra Markdown để phân tích bằng công cụ AI ngoài.

        KHÔNG gọi AI lần nào. Thay cho analyze_style cũ vốn tốn 16 lượt gọi Gemini mà
        vẫn cho ra hồ sơ khuôn trống, vì bước hợp nhất nhiều hồ sơ con làm mất chi tiết.

        Đồng thời cập nhật hồ sơ văn phong bằng SỐ LIỆU ĐẾM ĐƯỢC (Python, chính xác
        tuyệt đối) — hồ sơ này vẫn được nạp vào prompt mỗi lần soạn thư.
        """
        from xuat_thu_da_gui import export_sent_emails as run_export
        from backend.style_stats import update_profile_from_stats

        def report(phase: str, current: int, total: int, detail: str = ""):
            if progress:
                try:
                    progress({
                        "phase": phase,
                        "current": current,
                        "total": total,
                        "detail": detail,
                    })
                except Exception:
                    pass

        result = com_call(run_export, progress=report, deep=deep)
        stats = result.get("stats") or {}
        if stats.get("analyzed_count", 0) > 0:
            update_profile_from_stats(stats)
            self.kb.invalidate_cache()

        return ok({k: v for k, v in result.items() if k != "stats"})

    def get_style_profile(self) -> Dict[str, Any]:
        """Đọc hồ sơ văn phong hiện tại."""
        profile = self.kb.style_profile()
        if not profile or profile.get("name") == "default":
            return fail("Chưa có hồ sơ văn phong.")
        return ok(profile)

    # ------------------------------------------------------------------ cài đặt

    def get_settings(self) -> Dict[str, Any]:
        """Lấy thông tin cấu hình hiện tại."""
        profile = self.kb.style_profile()
        engine = self._get_engine()
        return ok({
            "api_key_configured": engine.is_ready,
            "serper_key_configured": self.web_searcher.enabled,
            "model": engine.model_name,
            "style_profile_name": profile.get("name", "default"),
            "style_analyzed_at": profile.get("analyzed_at", None),
        })

    def save_api_key(self, api_key: str = "", serper_key: str = "") -> Dict[str, Any]:
        """Lưu khóa API Gemini và/hoặc Serper vào .env và cập nhật runtime."""
        api_key = (api_key or "").strip()
        serper_key = (serper_key or "").strip()

        if not api_key and not serper_key:
            return fail("Chưa nhập khóa API nào.")

        ENV_FILE.parent.mkdir(parents=True, exist_ok=True)
        ENV_FILE.touch(exist_ok=True)

        if api_key:
            set_key(str(ENV_FILE), "GEMINI_API_KEY", api_key)
            os.environ["GEMINI_API_KEY"] = api_key
            self.ai_engine = self._build_engine(api_key)

        if serper_key:
            set_key(str(ENV_FILE), "SERPER_API_KEY", serper_key)
            os.environ["SERPER_API_KEY"] = serper_key
            self.web_searcher = WebSearcher(serper_key=serper_key)

        engine = self._get_engine()
        return ok({
            "api_key_configured": engine.is_ready,
            "serper_key_configured": self.web_searcher.enabled,
        })

    def open_reply_all(self, html_body: str = "", subject: str = "",
                       sender_email: str = "") -> Dict[str, Any]:
        """Mở cửa sổ Reply All trực tiếp trong Outlook."""
        result = com_call(self.outlook.open_reply_all_in_outlook, html_body, subject, sender_email)
        if result and result.get("success"):
            return ok(result)
        return fail((result.get("error") if result else None) or "Không mở được Reply All trong Outlook.")
