from __future__ import annotations

import json
import logging
import re
from typing import Dict, List, Optional

logger = logging.getLogger(__name__)

_FENCE_OPEN_RE = re.compile(r"^```[a-zA-Z]*\s*")
_FENCE_CLOSE_RE = re.compile(r"\s*```$")
_JSON_BLOCK_RE = re.compile(r"[\[{].*[\]}]", re.DOTALL)

# Bóc chi tiết hạn mức từ thông báo 429 của Google để báo lỗi cho đúng.
# "gemini-flash-latest" là bí danh nên model thật trong lỗi có thể khác hẳn.
_QUOTA_LIMIT_RE = re.compile(r"limit:\s*(\d+)")
_QUOTA_MODEL_RE = re.compile(r"model:\s*([\w.\-]+)")

VALID_CATEGORIES = ("Cần trả lời", "Việc cần làm", "Chỉ để biết", "Quảng cáo/Rác")
VALID_PRIORITIES = ("Cao", "Trung bình", "Thấp")


class AIEngineError(RuntimeError):
    """Lỗi thuộc về tầng AI — được api.py bắt và trả cho UI dưới dạng toast đỏ."""


def _strip_fence(text: str) -> str:
    text = (text or "").strip()
    if text.startswith("```"):
        text = _FENCE_CLOSE_RE.sub("", _FENCE_OPEN_RE.sub("", text))
    return text.strip()


# Nội dung Gemini sinh ra sẽ được chèn thẳng vào form trả lời thật của Outlook,
# nên phải lọc trắng danh sách trước — prompt chỉ là lời dặn, không phải bảo đảm.
_ALLOWED_TAGS = {"p", "br", "strong", "b", "em", "i", "u",
                 "ul", "ol", "li", "a", "blockquote", "span", "div"}
_ALLOWED_ATTRS = {"a": {"href", "style"}}
_DROP_ENTIRELY = {"script", "style", "iframe", "object", "embed", "img",
                  "form", "input", "link", "meta"}
_SAFE_HREF_RE = re.compile(r"^(https?:|mailto:)", re.I)

# Thuộc tính style ĐƯỢC giữ nhưng chỉ với các thuộc tính CSS vô hại dưới đây.
# Loại bỏ position/behavior/expression/url() vì đó là các vector tấn công cũ của
# HTML email. Không có style thì thư nháp không thể mang định dạng nào.
_ALLOWED_CSS = {
    "font-family", "font-size", "font-weight", "font-style", "color",
    "margin", "margin-top", "margin-bottom", "margin-left", "margin-right",
    "padding", "padding-left", "line-height", "text-align", "text-decoration",
}
_CSS_UNSAFE_RE = re.compile(r"(expression|javascript:|url\s*\(|@import)", re.I)

# Kiểu chèn vào từng thẻ để Outlook không áp khoảng cách đoạn của Word.
_TAG_BASE_STYLE = {
    "p": "margin:0 0 10pt 0;",
    "ul": "margin:0 0 10pt 0; padding-left:22pt;",
    "ol": "margin:0 0 10pt 0; padding-left:22pt;",
    "li": "margin:0 0 4pt 0;",
    "blockquote": "margin:0 0 10pt 12pt; padding-left:10pt;",
}


def _parse_style(value: str) -> dict:
    """Tách chuỗi style thành dict, bỏ thuộc tính ngoài danh sách trắng."""
    out = {}
    for decl in (value or "").split(";"):
        if ":" not in decl:
            continue
        prop, _, val = decl.partition(":")
        prop, val = prop.strip().lower(), val.strip()
        if prop in _ALLOWED_CSS and val and not _CSS_UNSAFE_RE.search(val):
            out[prop] = val
    return out


def _clean_style(value: str) -> str:
    """Lọc thuộc tính style theo danh sách trắng thuộc tính CSS."""
    return "; ".join(f"{p}:{v}" for p, v in _parse_style(value).items())


def _merge_style(base: str, existing: str) -> str:
    """Gộp style mặc định với style của model, MỖI thuộc tính chỉ xuất hiện một lần.

    Nối chuỗi thẳng sẽ khiến style phình ra sau mỗi lượt tinh chỉnh
    ("margin:...; margin:...; margin:...") vì refine_draft đưa lại bản đã định dạng.
    """
    merged = _parse_style(base)
    merged.update(_parse_style(existing))       # style của model được ưu tiên
    return "; ".join(f"{p}:{v}" for p, v in merged.items())


def _sanitize_html(html: str) -> str:
    """Giữ lại thẻ trong danh sách trắng, gỡ bỏ phần còn lại nhưng giữ chữ."""
    if not html:
        return ""
    try:
        from bs4 import BeautifulSoup
    except ImportError:                       # thiếu thư viện thì thà mất định dạng
        logger.warning("Thiếu beautifulsoup4 — bỏ qua bước làm sạch HTML.")
        return html

    soup = BeautifulSoup(html, "html.parser")
    for tag in soup.find_all(_DROP_ENTIRELY):
        tag.decompose()                       # xoá cả nội dung bên trong
    for tag in soup.find_all(True):
        if tag.name not in _ALLOWED_TAGS:
            tag.unwrap()                      # thẻ lạ: bỏ thẻ, giữ chữ
            continue
        allowed = _ALLOWED_ATTRS.get(tag.name, set()) | {"style"}
        for attr in list(tag.attrs):
            if attr not in allowed:
                del tag[attr]
        if tag.get("style"):
            cleaned = _clean_style(tag["style"])
            if cleaned:
                tag["style"] = cleaned
            else:
                del tag["style"]
        href = tag.get("href")
        if href and not _SAFE_HREF_RE.match(href.strip()):
            del tag["href"]                   # chặn javascript:, data:, file:
    return str(soup).strip()


def _paragraphize(soup) -> None:
    """Bọc các đoạn văn bản trần ở cấp cao nhất vào thẻ <p>.

    Model nhỏ (flash-lite) hay trả về văn bản thuần ngăn bằng ký tự xuống dòng thay
    vì thẻ <p>. Trong HTML, xuống dòng trần KHÔNG tạo ngắt đoạn, nên cả thư sẽ dính
    liền thành một khối khi hiển thị trong Outlook. Chuẩn hoá ở đây để kết quả không
    phụ thuộc vào việc model có tuân thủ prompt hay không.
    """
    from bs4 import NavigableString

    for node in list(soup.contents):
        if not isinstance(node, NavigableString) or not node.strip():
            continue
        blocks = [b.strip() for b in re.split(r"\n\s*\n", str(node)) if b.strip()]
        if not blocks:
            continue
        replacements = []
        for block in blocks:
            para = soup.new_tag("p")
            lines = block.split("\n")
            for i, line in enumerate(lines):
                if i:
                    para.append(soup.new_tag("br"))
                para.append(NavigableString(line.strip()))
            replacements.append(para)
        node.replace_with(replacements[0])
        for extra in reversed(replacements[1:]):
            replacements[0].insert_after(extra)


def _format_email_html(html: str, font_family: str, font_size: str,
                       color: str) -> str:
    """Gắn định dạng inline để thư nháp hòa vào ngữ cảnh Outlook.

    Ba vấn đề phải xử lý:
      1. HTML không khai báo font sẽ được Outlook dựng bằng font mặc định của trình
         duyệt (thường là Times New Roman), lệch hẳn với chữ ký và phần trích dẫn.
      2. Thẻ <p> trần bị Outlook áp khoảng cách đoạn của Word, giãn rất rộng.
      3. Model nhỏ hay trả text trần thay vì thẻ <p> -> cả thư dính liền một khối.
    Email không đọc được CSS ngoài, nên mọi thứ phải là style inline.
    """
    if not html:
        return ""
    try:
        from bs4 import BeautifulSoup
    except ImportError:
        return html

    soup = BeautifulSoup(html, "html.parser")

    # Idempotent: refine_draft đưa lại bản nháp ĐÃ bọc, không gỡ thì mỗi lượt tinh
    # chỉnh lại chồng thêm một lớp div.
    while True:
        nodes = [n for n in soup.contents if not (isinstance(n, str) and not n.strip())]
        if len(nodes) == 1 and getattr(nodes[0], "name", None) == "div" \
                and "font-family" in (nodes[0].get("style") or ""):
            nodes[0].unwrap()
        else:
            break

    _paragraphize(soup)

    _sig_re = re.compile(
        r"^(?:trân trọng|kính thư|thân ái|thanks\s*(?:&|and)?\s*regards|best\s*regards|regards|cảm ơn|kstt|thân mến|chúc anh/chị)\b",
        re.IGNORECASE,
    )
    while True:
        nodes = [tag for tag in soup.find_all(["p", "div", "li"]) if tag.get_text(strip=True)]
        if not nodes:
            break
        last_block = nodes[-1]
        text = last_block.get_text(strip=True)
        if _sig_re.search(text) or (len(text) < 50 and any(kw in text.lower() for kw in ("trân trọng", "thanks &", "best regards", "kính thư", "kstt"))):
            last_block.decompose()
        else:
            break

    for tag in soup.find_all(list(_TAG_BASE_STYLE)):
        # Hợp nhất theo thuộc tính; style của model ghi đè mặc định của ta.
        tag["style"] = _merge_style(_TAG_BASE_STYLE[tag.name], tag.get("style") or "")

    wrapper = soup.new_tag("div")
    wrapper["style"] = (f"font-family:{font_family}; font-size:{font_size}; "
                        f"color:{color};")
    for node in list(soup.contents):
        wrapper.append(node.extract())
    return str(wrapper)


def _parse_json(raw: str):
    """Bóc JSON chịu lỗi: gỡ code fence, nếu vẫn hỏng thì tìm khối [...] / {...}."""
    text = _strip_fence(raw)
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        m = _JSON_BLOCK_RE.search(text)
        if not m:
            raise AIEngineError("Gemini trả về dữ liệu không phải JSON hợp lệ.")
        return json.loads(m.group(0))


class AIEngine:
    def __init__(self, api_key: str, model_name: str = "gemini-3.5-flash-lite",
                 font_family: str = "Calibri, 'Segoe UI', sans-serif",
                 font_size: str = "11pt", text_color: str = "#000000") -> None:
        self.api_key = (api_key or "").strip()
        self.model_name = model_name
        self.font_family = font_family
        self.font_size = font_size
        self.text_color = text_color
        self.model = object()          # dummy flag để tương thích
        self.json_model = {"json": True}

        if not self.api_key:
            logger.warning("Chưa có GEMINI_API_KEY — các tính năng AI sẽ bị tắt.")

    @property
    def is_ready(self) -> bool:
        return bool(self.api_key)

    def _require_key(self) -> None:
        if not self.is_ready:
            import os
            from paths import ENV_FILE
            from dotenv import load_dotenv
            if ENV_FILE.is_file():
                load_dotenv(ENV_FILE, override=True)
            k = os.getenv("GEMINI_API_KEY", "").strip()
            if k:
                self.api_key = k
        if not self.is_ready:
            raise AIEngineError("Chưa cấu hình API key Gemini. Mở Cài đặt để nhập khóa.")

    def _call_gemini_api(self, contents, temperature: float = 0.7,
                          response_mime_type: Optional[str] = None, timeout: int = 90) -> str:
        import base64
        import json
        import urllib.error
        import urllib.request

        self._require_key()
        parts = []
        if isinstance(contents, str):
            parts.append({"text": contents})
        elif isinstance(contents, list):
            for item in contents:
                if isinstance(item, str):
                    parts.append({"text": item})
                elif isinstance(item, dict):
                    if "text" in item:
                        parts.append({"text": item["text"]})
                    elif "mime_type" in item and "data" in item:
                        data = item["data"]
                        if isinstance(data, bytes):
                            data = base64.b64encode(data).decode("ascii")
                        parts.append({
                            "inline_data": {
                                "mime_type": item["mime_type"],
                                "data": data,
                            }
                        })

        payload = {
            "contents": [{"parts": parts}],
            "generationConfig": {
                "temperature": temperature,
            },
        }
        if response_mime_type:
            payload["generationConfig"]["responseMimeType"] = response_mime_type

        # Chuẩn hóa model: nếu có prefix "models/" thì giữ, không thì dùng thẳng
        model_endpoint = self.model_name.replace("models/", "")
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_endpoint}:generateContent"
        headers = {
            "Content-Type": "application/json; charset=utf-8",
            "x-goog-api-key": self.api_key,
        }

        try:
            req = urllib.request.Request(
                url,
                data=json.dumps(payload).encode("utf-8"),
                headers=headers,
            )
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                res_data = json.loads(resp.read().decode("utf-8"))
                candidates = res_data.get("candidates", [])
                if not candidates:
                    raise AIEngineError("Gemini không trả về nội dung (có thể bị bộ lọc an toàn chặn).")
                parts = candidates[0].get("content", {}).get("parts", [])
                text = "".join(p.get("text", "") for p in parts)
                if not text.strip():
                    raise AIEngineError("Gemini trả về phản hồi rỗng.")
                return text
        except urllib.error.HTTPError as e:
            err_body = e.read().decode("utf-8", errors="replace")
            err_str = f"HTTP {e.code}: {err_body}"
            if e.code == 400 or "API_KEY_INVALID" in err_body or "API key not valid" in err_body:
                raise AIEngineError(
                    "Khoá API Gemini không hợp lệ. Kiểm tra lại API Key trong ⚙️ Cài đặt."
                ) from e
            if e.code == 401 or "ACCESS_TOKEN_TYPE_UNSUPPORTED" in err_body:
                raise AIEngineError(
                    "Khóa API Gemini chưa đúng hoặc đã hết hạn. "
                    "Hãy mở Cài đặt (⚙️) hoặc sửa tệp .env để nhập khóa API Gemini từ https://aistudio.google.com/app/apikey"
                ) from e
            if e.code == 429 or "RESOURCE_EXHAUSTED" in err_body or "Quota exceeded" in err_body:
                limit = _QUOTA_LIMIT_RE.search(err_body)
                actual = _QUOTA_MODEL_RE.search(err_body)
                shown = actual.group(1) if actual else self.model_name
                if "PerDay" in err_body:
                    raise AIEngineError(
                        f"Đã dùng hết hạn mức MIỄN PHÍ TRONG NGÀY của model '{shown}'"
                        + (f" ({limit.group(1)} lượt/ngày)" if limit else "")
                        + ". Chờ thêm không giải quyết được. Hãy đổi GEMINI_MODEL trong "
                          "tệp .env sang model khác (ví dụ gemini-3.5-flash hoặc "
                          "gemini-3.1-flash-lite — mỗi model có hạn mức riêng), "
                          "hoặc bật thanh toán tại https://ai.dev/rate-limit."
                    ) from e
                raise AIEngineError(
                    f"Gọi Gemini quá nhanh (429, model '{shown}'). "
                    "Chờ khoảng 1 phút rồi thử lại."
                ) from e
            if e.code == 404 or "not found" in err_body.lower():
                raise AIEngineError(
                    f"Model '{self.model_name}' không khả dụng với API key này. Hãy đổi "
                    "GEMINI_MODEL trong tệp .env sang gemini-3.5-flash hoặc "
                    "gemini-3.1-flash-lite."
                ) from e
            raise AIEngineError(f"Gọi Gemini thất bại: {err_str}") from e
        except Exception as e:
            if isinstance(e, AIEngineError):
                raise
            raise AIEngineError(f"Lỗi kết nối tới máy chủ Gemini: {e}") from e

    def _generate(self, model, contents, timeout: int = 90) -> str:
        """contents: chuỗi prompt, hoặc list [prompt, {mime_type, data}, ...] khi
        cần gửi kèm tệp đính kèm dạng nhị phân (PDF/ảnh) cho Gemini đọc trực tiếp."""
        is_json = (model is self.json_model) or (isinstance(model, dict) and model.get("json"))
        temp = 0.2 if is_json else 0.7
        mime = "application/json" if is_json else None
        return self._call_gemini_api(contents, temperature=temp, response_mime_type=mime, timeout=timeout)

    # Việc học văn phong ĐÃ CHUYỂN sang backend/style_stats.py (đếm bằng Python) và
    # kien_thuc/loai_thu/ (hướng dẫn do công cụ AI ngoài viết). Hai phương thức cũ
    # analyze_writing_style + merge_writing_styles đã bị xoá ở v1.6.0: bước hợp nhất
    # nhiều hồ sơ con làm LLM trừu tượng hoá mất chi tiết, biến "Dear CH," thành
    # "Dear [Name],", và càng nhiều dữ liệu thì kết quả càng nhạt.

    def generate_reply(self, original_email: Optional[Dict] = None,
                       conversation_thread: Optional[List[Dict]] = None,
                       user_instruction: str = "",
                       style_profile: Optional[Dict] = None,
                       *,
                       email_data: Optional[Dict] = None,
                       thread: Optional[List[Dict]] = None,
                       instruction: str = "",
                       knowledge: str = "",
                       attachment_text: str = "",
                       attachment_blobs: Optional[List[Dict]] = None,
                       me: Optional[Dict] = None,
                       timeout: int = 90) -> str:
        self._require_key()
        email_obj = original_email if original_email is not None else (email_data or {})
        thread_obj = conversation_thread if conversation_thread is not None else (thread or [])
        instr_str = user_instruction or instruction or ""

        prompt = f"""Bạn là trợ lý soạn email. Hãy viết nội dung thư phản hồi.

## Người viết thư trả lời (chính là bạn)
{self._format_me(me)}

## Email cần trả lời
Người gửi: {email_obj.get('sender_name', '')} <{email_obj.get('sender_email', '')}>
Chủ đề: {email_obj.get('subject', '')}
Nội dung:
{(email_obj.get('body') or '')[:4000]}

## Lịch sử hội thoại
{self._format_thread(thread_obj) if thread_obj
 else (email_obj.get('quoted_tail') or '(Không có)')[:3000]}

## Chỉ dẫn của người dùng
{instr_str}

## Phong cách bắt buộc tuân theo
{self._format_style(style_profile)}

## Kiến thức nội bộ do người dùng biên soạn (ƯU TIÊN CAO NHẤT)
{knowledge or '(Chưa có)'}

## Nội dung tệp đính kèm (đã trích xuất)
{attachment_text or '(Không có)'}

## Quy tắc đầu ra
- Viết bằng ĐÚNG ngôn ngữ của email gốc.
- Nếu "Kiến thức nội bộ" mâu thuẫn với "Phong cách bắt buộc tuân theo", LUÔN theo Kiến thức nội bộ.
- Chỉ dùng thẻ HTML cơ bản: <p>, <br>, <strong>, <em>, <ul>, <li>.
- KHÔNG kèm <html>, <head>, <body>, KHÔNG kèm markdown code fence.
- KHÔNG bịa thông tin không có trong email gốc, tệp đính kèm hoặc chỉ dẫn.
- Chỉ trả về nội dung HTML của thư, không lời dẫn, không giải thích."""

        # Tệp nhị phân (PDF/ảnh) đi kèm prompt dưới dạng inline part để Gemini tự đọc,
        # kể cả PDF scan — không cần OCR cục bộ.
        contents = [prompt]
        for blob in (attachment_blobs or []):
            contents.append({"mime_type": blob["mime"], "data": blob["data"]})

        # KHÔNG bọc try/except trả chuỗi lỗi ở đây: lỗi phải nổi lên tới UI dưới dạng
        # toast đỏ, tuyệt đối không được chèn vào ô soạn thảo như thể là nội dung thư (lỗi L10).
        return self._clean_html(self._generate(
            self.model,
            contents if len(contents) > 1 else prompt,
            timeout=max(timeout, 180) if attachment_blobs else timeout,
        ))

    def refine_draft(self, current_draft: str, feedback: str) -> str:
        self._require_key()
        prompt = f"""Đây là bản nháp email dạng HTML:

{current_draft}

Yêu cầu chỉnh sửa của người dùng: {feedback}

Hãy chỉnh sửa bản nháp theo đúng yêu cầu, giữ nguyên những phần không liên quan.
Chỉ trả về HTML đã sửa, không markdown, không giải thích."""
        return self._clean_html(self._generate(self.model, prompt))

    # ------------------------------------------------------------------ phân loại

    def classify_emails(self, emails: List[Dict], batch_size: int = 15) -> Dict[str, Dict]:
        """Trả về {entry_id: {summary, category, priority, needs_reply}}."""
        self._require_key()
        results: Dict[str, Dict] = {}
        for start in range(0, len(emails), batch_size):
            batch = emails[start:start + batch_size]
            try:
                results.update(self._classify_batch(batch))
            except Exception as e:
                logger.error("Phân loại lô %s thất bại: %s", start // batch_size, e)
        return results

    def _classify_batch(self, batch: List[Dict]) -> Dict[str, Dict]:
        # Gửi chỉ số thay vì entry_id (chuỗi hex ~140 ký tự) để tiết kiệm token
        lines = []
        for idx, e in enumerate(batch):
            lines.append(
                f"[{idx}] Từ: {e.get('sender_name', '')} | Chủ đề: {e.get('subject', '')}\n"
                f"    {(e.get('preview') or '')[:300]}"
            )

        prompt = (
            "Bạn là trợ lý phân loại hộp thư. Với MỖI email dưới đây, hãy tóm tắt và gán nhãn.\n\n"
            + "\n".join(lines)
            + "\n\nTrả về DUY NHẤT một mảng JSON, mỗi phần tử:\n"
              '{"index": <số trong ngoặc vuông>, "summary": "<tóm tắt 1 câu tiếng Việt, tối đa 20 từ>", '
              f'"category": <một trong {list(VALID_CATEGORIES)}>, '
              f'"priority": <một trong {list(VALID_PRIORITIES)}>, '
              '"needs_reply": <true|false>}\n'
              "Mảng phải có đúng số phần tử bằng số email. Không thêm trường nào khác."
        )

        parsed = _parse_json(self._generate(self.json_model, prompt, timeout=120))
        if isinstance(parsed, dict):
            parsed = parsed.get("results") or parsed.get("items") or []

        out: Dict[str, Dict] = {}
        for row in parsed if isinstance(parsed, list) else []:
            try:
                idx = int(row.get("index", -1))
                if not 0 <= idx < len(batch):
                    continue
                entry_id = batch[idx].get("entry_id")
                if not entry_id:
                    continue
                category = row.get("category")
                priority = row.get("priority")
                out[entry_id] = {
                    "summary": str(row.get("summary", ""))[:200],
                    "category": category if category in VALID_CATEGORIES else "Chỉ để biết",
                    "priority": priority if priority in VALID_PRIORITIES else "Trung bình",
                    "needs_reply": bool(row.get("needs_reply", False)),
                }
            except Exception:
                continue
        return out

    # ------------------------------------------------------------------ tiện ích

    def _clean_html(self, text: str) -> str:
        """Gỡ code fence -> lọc trắng danh sách -> gắn định dạng inline cho Outlook."""
        return _format_email_html(_sanitize_html(_strip_fence(text)),
                                  self.font_family, self.font_size, self.text_color)

    @staticmethod
    def _format_me(me: Optional[Dict]) -> str:
        """Model cần biết nó đang viết từ phía nào của cuộc hội thoại."""
        if not me:
            return "(Không rõ — viết ở ngôi thứ nhất, không tự xưng tên)"
        name = (me.get("name") or "").strip()
        email = (me.get("email") or "").strip()
        return f"{name} <{email}>".strip() or "(Không rõ)"

    @staticmethod
    def _fmt_patterns(items, limit: int = 8) -> str:
        """Nhận CẢ HAI định dạng: list chuỗi (hồ sơ v1/v2) và list dict kèm tần suất (v3).

        Hồ sơ v3 do style_stats đếm ra có dạng {"text","count","ratio"} — nếu cứ join
        thẳng như bản cũ sẽ ném TypeError. Có tần suất thì in kèm để model biết mẫu nào
        là thói quen chính, mẫu nào chỉ dùng hoạ hoằn.
        """
        out = []
        for item in (items or [])[:limit]:
            if isinstance(item, dict):
                text = item.get("text", "")
                ratio = item.get("ratio")
                out.append(f'"{text}" ({ratio:.0%})' if isinstance(ratio, (int, float))
                           else f'"{text}"')
            elif item:
                out.append(str(item))
        return ", ".join(out) or "(không rõ)"

    @staticmethod
    def _format_style(profile: Optional[Dict]) -> str:
        if not profile:
            return "Chưa có hồ sơ văn phong — hãy viết chuyên nghiệp, lịch sự, ngắn gọn."
        level = profile.get("formality_level", 3)
        desc = "trung lập"
        if isinstance(level, (int, float)):
            if level <= 2:
                desc = "thân mật, gần gũi"
            elif level >= 4:
                desc = "trang trọng, chuyên nghiệp"
        parts = [
            f"- Mức độ trang trọng: {desc}",
            f"- Câu chào thường dùng: {AIEngine._fmt_patterns(profile.get('greeting_patterns'))}",
            f"- Câu kết thường dùng: {AIEngine._fmt_patterns(profile.get('closing_patterns'))}",
        ]
        if profile.get("common_phrases"):
            # Đây là các mảnh n-gram đếm được, KHÔNG phải câu mẫu. Nói rõ để model
            # dùng làm vốn từ chứ không nhồi nguyên văn vào câu — thực tế nó từng viết
            # "...không tuân thủ với quy định hiện hành cần được kiểm soát và xử lý đối
            # với hành vi không tuân thủ" vì hiểu nhầm là phải chèn đủ.
            parts.append(
                "- Vốn từ nghiệp vụ hay xuất hiện (dùng làm tham khảo về cách gọi tên "
                "sự việc, TUYỆT ĐỐI không chèn nguyên văn cho đủ): "
                + AIEngine._fmt_patterns(profile['common_phrases']))
        length = profile.get("length") or {}
        if length.get("median_words"):
            parts.append(
                f"- Độ dài mục tiêu: khoảng {length['median_words']} từ, "
                f"{length.get('median_sentences', 0)} câu (bằng độ dài thư người dùng thường viết)")
        if profile.get("tone_notes"):
            parts.append(f"- Giọng văn: {profile['tone_notes']}")
        if profile.get("signature"):
            parts.append(f"- Kết thư bằng chữ ký:\n{profile['signature']}")
        return "\n".join(parts)

    @staticmethod
    def _format_thread(thread: List[Dict]) -> str:
        if not thread:
            return "(Không có)"
        rows = []
        for msg in thread[-8:]:          # chỉ 8 lượt gần nhất để không phình prompt
            body = " ".join((msg.get("body") or "")[:400].split())
            rows.append(f"[{msg.get('sent_time', '')}] {msg.get('sender_name', '')}: {body}")
        return "\n".join(rows)
