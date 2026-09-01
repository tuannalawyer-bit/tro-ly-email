# ĐẶC TẢ TRIỂN KHAI — Trợ lý Email v1.1.0

> **Tài liệu này dành cho AI thực thi (Gemini).** Mọi đoạn mã dưới đây là mã hoàn chỉnh, sao chép
> nguyên văn. Không tự ý thêm tính năng, không refactor vùng không được nhắc tới, không đổi tên biến.
> Thư mục gốc dự án: `d:\AI\Xu ly email`. Python: `.venv\Scripts\python.exe`.

---

## 1. BỐI CẢNH — Tại sao phải sửa

Dự án là desktop app pywebview đọc Outlook qua COM, dùng Gemini để học văn phong và soạn thư trả lời.
Kiểm tra mã nguồn hiện tại cho thấy **ứng dụng chưa từng chạy được lần nào**. Danh sách lỗi đã kiểm chứng:

| # | Lỗi | Vị trí | Hậu quả |
|---|-----|--------|---------|
| L1 | `import pywebview as webview` — module thực tế tên là `webview` | `main.py:3` | `ModuleNotFoundError`, app không khởi động |
| L2 | Phân tích văn phong chạy trên `preview` 100 ký tự thay vì nội dung thư | `api.py:120` + `ai_engine.py:24` | Tính năng cốt lõi vô hiệu |
| L3 | `sanitizeHtml()` là hàm rỗng, HTML email đổ vào `innerHTML` | `utils.js:97-101`, `email-viewer.js:91` | XSS — email độc hại chiếm được `window.pywebview.api` (toàn quyền Outlook + hệ thống tệp) |
| L4 | COM object tạo ở thread A, dùng ở thread B | `api.py` (8 chỗ `CoInitialize`) | Lỗi *"marshalled for a different thread"* ngắt quãng |
| L5 | Gán cứng `tzinfo=timezone.utc` cho giờ local | `outlook_client.py:275-279` | Mọi email lệch **7 giờ** ở VN |
| L6 | Gọi `msg.HTMLBody` + BeautifulSoup cho **mỗi** email trong danh sách | `outlook_client.py:269` | Tải danh sách rất chậm |
| L7 | Badge thư mục dùng `Items.Count` (tổng thư) nhưng gọi là `updateUnreadCount` | `outlook_client.py:44`, `app.js:126` | Số chưa đọc sai |
| L8 | `btn-ai-reply`, `btn-forward`, `theme-toggle` không có event listener | `index.html` | 3 nút chết, gồm cả nút chủ đạo |
| L9 | Thiếu **toàn bộ** CSS cho `.thread-*`, `.attachments-*`, `.avatar.large`, `.sender-details`… | `styles.css` | Panel chi tiết / đính kèm / hội thoại vỡ giao diện |
| L10 | Lỗi AI được trả về **như thể là nội dung thư** | `ai_engine.py:102` | Người dùng lưu nhầm thông báo lỗi vào Drafts |
| L11 | Query tìm kiếm nối thẳng vào DASL filter | `outlook_client.py:212-215` | Query chứa `'` làm vỡ filter |
| L12 | `email_models.py` (84 dòng) không được import ở đâu; `CACHE_DIR` tạo ra nhưng không dùng | — | Mã chết |
| L13 | Nội dung mới chèn **trước** thẻ `<html>` của thư nháp | `outlook_client.py:189` | Hỏng chữ ký, HTML dị dạng |
| L14 | Click email không đánh dấu đã đọc trong Outlook thật | `app.js:239` | Chỉ đổi state JS |
| L15 | `@import` Google Fonts từ internet trong app desktop | `styles.css:1` | Chậm khởi động, hỏng khi offline |

**Phạm vi đã chốt:** giữ kiến trúc pywebview + COM; hai nhóm tính năng: **(A) soạn thư trả lời bằng AI**
và **(B) tóm tắt & phân loại tự động**. KHÔNG làm hành động hàng loạt trên Outlook, KHÔNG xuất Excel.

**Thay đổi hợp đồng API quan trọng (đọc kỹ trước khi code):**
1. Thư mục được định danh bằng **`entry_id`** thay cho tên — chống lỗi Outlook giao diện tiếng Việt
   (`"Hộp thư đến"` ≠ `"Inbox"`).
2. `get_emails` trả về `{items, next_offset, has_more}` thay cho mảng phẳng — sửa dứt điểm lỗi phân trang.

---

## 2. THỨ TỰ THỰC HIỆN

Làm tuần tự Bước 0 → 8. **Sau Bước 4 phải dừng lại chạy thử và báo cáo** trước khi làm tiếp.

---

## BƯỚC 0 — Chuẩn bị kho mã

**0.1** Tạo `.gitignore` ở thư mục gốc:
```gitignore
.venv/
__pycache__/
*.pyc
.env
data/
.backups/
```

**0.2** Chạy `git init` rồi commit toàn bộ trạng thái hiện tại với message `chore: snapshot trước khi khắc phục v1.0.0`.
Đây là điểm quay lui bắt buộc.

**0.3** Tạo `.env.example`:
```
GEMINI_API_KEY=
```

**0.4** Tạo `development_log.csv` ở thư mục gốc, **encoding UTF-8 có BOM (`utf-8-sig`)**, nội dung khởi tạo:
```csv
Yêu cầu sửa đổi,Giải pháp thực hiện,Tên phiên bản,Thời gian xây dựng
"Khởi tạo nhật ký phát triển","Tạo tệp development_log.csv theo quy tắc chung",1.0.0,<dd-mm-yyyy HH:mm:ss>
```
Sau **mỗi bước** từ 1→8, thêm một dòng mới ghi lại việc đã làm.

**0.5** Xóa tệp mã chết `backend/email_models.py` (không tệp nào import nó — đã kiểm chứng bằng grep).

---

## BƯỚC 1 — Gỡ chặn khởi động

### 1.1 Ghi đè toàn bộ `main.py`

```python
import logging
import os
import sys
from pathlib import Path

import webview  # LƯU Ý: gói pip tên "pywebview" nhưng module import là "webview"

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from backend.api import EmailAssistantAPI
from config import APP_NAME, BASE_DIR

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger(__name__)


def main() -> None:
    frontend_path = BASE_DIR / "frontend" / "index.html"
    if not frontend_path.exists():
        raise FileNotFoundError(f"Không tìm thấy giao diện tại: {frontend_path}")

    api = EmailAssistantAPI()
    webview.create_window(
        title=APP_NAME,
        url=str(frontend_path),
        js_api=api,
        width=1400,
        height=900,
        min_size=(1100, 700),
        background_color="#0a0a1a",
    )

    debug = os.getenv("DEBUG", "").strip() in ("1", "true", "True")
    logger.info("Khởi động ứng dụng (debug=%s)...", debug)
    webview.start(debug=debug)


if __name__ == "__main__":
    main()
```

> Khối tự ghi `index.html` placeholder (dòng 20-24 cũ) bị xóa vì nó **che giấu** lỗi thiếu tệp giao diện.

### 1.2 Sửa `config.py` — chỉ thêm 2 dòng cuối, giữ nguyên phần còn lại

```python
VERSION = "1.1.0"          # sửa dòng VERSION đang có (1.0.0 -> 1.1.0)

# ... giữ nguyên toàn bộ phần giữa ...

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-2.0-flash")   # THÊM MỚI
```

### 1.3 Kiểm chứng ngay
```powershell
.\.venv\Scripts\python.exe -c "import webview; print('OK', webview.__version__)"
```
Phải in `OK 6.2.1`. Nếu vẫn lỗi, **dừng lại và báo cáo** — không đi tiếp.

---

## BƯỚC 2 — Tệp mới `backend/com_worker.py` (sửa lỗi L4)

Tạo mới, nội dung đầy đủ:

```python
"""Chạy toàn bộ lời gọi Outlook COM trên MỘT thread STA duy nhất.

pywebview gọi mỗi phương thức js_api trên một thread khác nhau. COM object của
Outlook thuộc về apartment (STA) nơi nó được tạo ra; dùng nó từ thread khác mà
không marshalling sẽ ném lỗi "The application called an interface that was
marshalled for a different thread".

Giải pháp chuẩn: một ThreadPoolExecutor đúng 1 worker, gọi CoInitialize() ngay
khi thread khởi động. Mọi thao tác Outlook đều đi qua com_call().
"""
from __future__ import annotations

import atexit
import logging
from concurrent.futures import ThreadPoolExecutor
from typing import Any, Callable

import pythoncom

logger = logging.getLogger(__name__)


def _initializer() -> None:
    pythoncom.CoInitialize()
    logger.info("Thread COM đã khởi tạo apartment STA.")


_executor = ThreadPoolExecutor(
    max_workers=1,                      # BẮT BUỘC là 1
    initializer=_initializer,
    thread_name_prefix="outlook-com",
)


def com_call(fn: Callable[..., Any], *args: Any, **kwargs: Any) -> Any:
    """Thực thi fn trên đúng thread sở hữu các COM object.

    Chặn cho tới khi có kết quả; ngoại lệ được ném lại nguyên vẹn cho phía gọi.
    """
    return _executor.submit(fn, *args, **kwargs).result()


@atexit.register
def _shutdown() -> None:
    _executor.shutdown(wait=False)
```

---

## BƯỚC 3 — Ghi đè toàn bộ `backend/outlook_client.py`

Sửa các lỗi L5, L6, L7, L11, L13, L14 + đổi sang định danh thư mục bằng `entry_id`.

```python
"""Lớp truy cập Outlook qua COM.

QUAN TRỌNG: mọi phương thức public của lớp này PHẢI được gọi thông qua
backend.com_worker.com_call(). Không gọi trực tiếp từ thread khác.
"""
from __future__ import annotations

import logging
import re
from datetime import datetime
from typing import Dict, List, Optional

import win32com.client

logger = logging.getLogger(__name__)

OL_MAIL_ITEM = 43          # MailItem.Class
OL_FOLDER_SENT = 5
OL_FOLDER_INBOX = 6
OL_FOLDER_DRAFTS = 16

# Marker bắt đầu phần trích dẫn thư cũ trong body
_QUOTE_MARKERS = (
    "-----Original Message-----",
    "________________________________",
    "\nFrom: ",
    "\nTừ: ",
    "\nSent: ",
    "\nĐã gửi: ",
)
_ON_WROTE_RE = re.compile(r"\nOn .{0,120}?\bwrote:", re.IGNORECASE)
_BODY_TAG_RE = re.compile(r"<body[^>]*>", re.IGNORECASE)


def _to_iso(value) -> str:
    """Chuyển thời gian Outlook sang ISO-8601 ĐÚNG múi giờ.

    pywin32 >= 224 trả về pywintypes.datetime (kế thừa datetime) đã kèm tzinfo.
    Tuyệt đối KHÔNG gán cứng timezone.utc như mã cũ — làm lệch 7 giờ ở Việt Nam.
    """
    if not isinstance(value, datetime):
        return str(value or "")
    if value.tzinfo is None:
        value = value.astimezone()        # gán múi giờ local của máy
    return value.isoformat()


def strip_quoted_text(body: str, max_len: int = 3000) -> str:
    """Bỏ phần trích dẫn thư cũ, chỉ giữ nội dung người dùng thực sự viết."""
    if not body:
        return ""
    text = body.replace("\r\n", "\n")
    cut = len(text)
    for marker in _QUOTE_MARKERS:
        idx = text.find(marker)
        if 0 <= idx < cut:
            cut = idx
    m = _ON_WROTE_RE.search(text)
    if m and m.start() < cut:
        cut = m.start()
    return text[:cut].strip()[:max_len]


class OutlookClient:
    def __init__(self) -> None:
        self.app = None
        self.namespace = None
        self._default_ids: Dict[str, str] = {}

    # ------------------------------------------------------------------ kết nối

    def connect(self) -> bool:
        """Kết nối tới Outlook. KHÔNG gọi CoInitialize ở đây — com_worker đã lo."""
        try:
            self.app = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.app.GetNamespace("MAPI")
            self._default_ids = {}
            for kind, code in (("inbox", OL_FOLDER_INBOX),
                               ("sent", OL_FOLDER_SENT),
                               ("drafts", OL_FOLDER_DRAFTS)):
                try:
                    self._default_ids[kind] = self.namespace.GetDefaultFolder(code).EntryID
                except Exception:
                    pass
            return True
        except Exception as e:
            logger.error("Không kết nối được Outlook: %s", e)
            self.app = None
            self.namespace = None
            return False

    def _ensure(self) -> bool:
        return bool(self.namespace) or self.connect()

    # ------------------------------------------------------------------ thư mục

    def get_folders(self, max_depth: int = 2) -> List[Dict]:
        """Trả về danh sách thư mục thư (đệ quy tối đa max_depth cấp)."""
        if not self._ensure():
            return []
        try:
            root = self.namespace.GetDefaultFolder(OL_FOLDER_INBOX).Parent
            out: List[Dict] = []
            self._walk_folders(root, out, 0, max_depth)
            # Sắp xếp: Inbox -> Sent -> Drafts -> còn lại
            order = {"inbox": 0, "sent": 1, "drafts": 2, "other": 3}
            out.sort(key=lambda f: (order[f["kind"]], f["depth"], f["name"].lower()))
            return out
        except Exception as e:
            logger.error("Lỗi lấy danh sách thư mục: %s", e)
            return []

    def _walk_folders(self, parent, out: List[Dict], depth: int, max_depth: int) -> None:
        if depth > max_depth:
            return
        try:
            children = parent.Folders
        except Exception:
            return
        for folder in children:
            try:
                if getattr(folder, "DefaultItemType", -1) != 0:   # 0 = thư mục thư
                    continue
                entry_id = folder.EntryID
                kind = "other"
                for k, fid in self._default_ids.items():
                    if fid == entry_id:
                        kind = k
                        break
                out.append({
                    "entry_id": entry_id,
                    "name": folder.Name,
                    "kind": kind,
                    "depth": depth,
                    "item_count": int(getattr(folder, "Items").Count),
                    "unread_count": int(getattr(folder, "UnReadItemCount", 0)),
                })
                self._walk_folders(folder, out, depth + 1, max_depth)
            except Exception as e:
                logger.warning("Bỏ qua thư mục lỗi: %s", e)

    def _resolve_folder(self, folder_id: Optional[str]):
        """Lấy folder theo EntryID; None -> Inbox mặc định."""
        if not self._ensure():
            return None
        try:
            if folder_id:
                return self.namespace.GetFolderFromID(folder_id)
            return self.namespace.GetDefaultFolder(OL_FOLDER_INBOX)
        except Exception as e:
            logger.error("Không mở được thư mục %s: %s", folder_id, e)
            return None

    # ------------------------------------------------------------------ danh sách thư

    def get_emails(self, folder_id: Optional[str] = None, limit: int = 30,
                   offset: int = 0) -> Dict:
        """Trả về {items, next_offset, has_more}.

        offset là chỉ số THÔ trong collection đã sắp xếp (không phải số email đã
        trả về) — nhờ vậy các item không phải thư bị bỏ qua không gây lệch trang.
        """
        empty = {"items": [], "next_offset": offset, "has_more": False}
        folder = self._resolve_folder(folder_id)
        if folder is None:
            return empty
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)     # mới nhất trước
            total = int(items.Count)

            collected: List[Dict] = []
            i = max(offset, 0) + 1                 # collection COM đánh chỉ số từ 1
            while i <= total and len(collected) < limit:
                try:
                    msg = items.Item(i)
                    if getattr(msg, "Class", 0) == OL_MAIL_ITEM:
                        summary = self._email_to_summary(msg)
                        if summary:
                            collected.append(summary)
                except Exception as e:
                    logger.warning("Bỏ qua email tại vị trí %s: %s", i, e)
                i += 1

            return {"items": collected, "next_offset": i - 1, "has_more": i <= total}
        except Exception as e:
            logger.error("Lỗi đọc danh sách email: %s", e)
            return empty

    def search_emails(self, query: str, folder_id: Optional[str] = None,
                      limit: int = 50) -> List[Dict]:
        folder = self._resolve_folder(folder_id)
        if folder is None or not query.strip():
            return []
        safe = query.strip().replace("'", "''")    # ESCAPE — sửa lỗi L11
        try:
            dasl = (
                f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{safe}%' OR "
                f"\"urn:schemas:httpmail:fromname\" LIKE '%{safe}%'"
            )
            items = folder.Items.Restrict(dasl)
            items.Sort("[ReceivedTime]", True)
            return self._take_summaries(items, limit)
        except Exception as e:
            logger.warning("Restrict thất bại (%s), chuyển sang quét tuyến tính.", e)
            return self._linear_search(folder, query.strip().lower(), limit)

    def _linear_search(self, folder, needle: str, limit: int) -> List[Dict]:
        """Dự phòng khi store không hỗ trợ DASL filter — quét tối đa 500 thư gần nhất."""
        out: List[Dict] = []
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)
            total = min(int(items.Count), 500)
            for i in range(1, total + 1):
                if len(out) >= limit:
                    break
                try:
                    msg = items.Item(i)
                    if getattr(msg, "Class", 0) != OL_MAIL_ITEM:
                        continue
                    haystack = f"{getattr(msg, 'Subject', '')} {getattr(msg, 'SenderName', '')}".lower()
                    if needle in haystack:
                        summary = self._email_to_summary(msg)
                        if summary:
                            out.append(summary)
                except Exception:
                    continue
        except Exception as e:
            logger.error("Quét tuyến tính thất bại: %s", e)
        return out

    def _take_summaries(self, items, limit: int) -> List[Dict]:
        out: List[Dict] = []
        for item in items:
            if len(out) >= limit:
                break
            try:
                if getattr(item, "Class", 0) == OL_MAIL_ITEM:
                    summary = self._email_to_summary(item)
                    if summary:
                        out.append(summary)
            except Exception:
                continue
        return out

    # ------------------------------------------------------------------ chi tiết thư

    def get_email_detail(self, entry_id: str) -> Dict:
        if not self._ensure():
            return {}
        try:
            msg = self.namespace.GetItemFromID(entry_id)
            if getattr(msg, "Class", 0) != OL_MAIL_ITEM:
                return {}

            detail = self._email_to_summary(msg)
            attachments = []
            try:
                for att in msg.Attachments:
                    attachments.append({
                        "filename": getattr(att, "FileName", ""),
                        "size": int(getattr(att, "Size", 0) or 0),
                    })
            except Exception as e:
                logger.warning("Không đọc được đính kèm: %s", e)

            detail.update({
                "body": getattr(msg, "Body", "") or "",
                "html_body": getattr(msg, "HTMLBody", "") or "",
                "to": getattr(msg, "To", "") or "",
                "cc": getattr(msg, "CC", "") or "",
                "attachments": attachments,
                "conversation_topic": getattr(msg, "ConversationTopic", "") or "",
            })
            return detail
        except Exception as e:
            logger.error("Lỗi đọc chi tiết email: %s", e)
            return {}

    def mark_as_read(self, entry_id: str) -> bool:
        """Đánh dấu đã đọc TRONG OUTLOOK THẬT (sửa lỗi L14)."""
        if not self._ensure():
            return False
        try:
            msg = self.namespace.GetItemFromID(entry_id)
            if getattr(msg, "UnRead", False):
                msg.UnRead = False
                msg.Save()
            return True
        except Exception as e:
            logger.warning("Không đánh dấu được đã đọc: %s", e)
            return False

    def get_conversation_thread(self, entry_id: str) -> List[Dict]:
        if not self._ensure():
            return []
        try:
            msg = self.namespace.GetItemFromID(entry_id)
            if getattr(msg, "Class", 0) != OL_MAIL_ITEM:
                return []

            thread: List[Dict] = []
            seen: set = set()
            try:
                conv = msg.GetConversation()
                if conv:
                    for item in conv.GetRootItems():
                        self._walk_conversation(item, conv, thread, seen)
            except Exception as e:
                logger.warning("GetConversation thất bại (%s), dùng ConversationTopic.", e)
                self._thread_by_topic(msg, thread, seen)

            thread.sort(key=lambda x: x.get("sent_time", ""))
            return thread
        except Exception as e:
            logger.error("Lỗi lấy chuỗi hội thoại: %s", e)
            return []

    def _walk_conversation(self, item, conv, thread: List[Dict], seen: set) -> None:
        try:
            if getattr(item, "Class", 0) == OL_MAIL_ITEM:
                eid = item.EntryID
                if eid not in seen:
                    seen.add(eid)
                    thread.append(self._to_thread_message(item))
        except Exception:
            pass
        try:
            for child in conv.GetChildren(item):
                self._walk_conversation(child, conv, thread, seen)
        except Exception:
            pass

    def _thread_by_topic(self, msg, thread: List[Dict], seen: set) -> None:
        topic = (getattr(msg, "ConversationTopic", "") or "").replace("'", "''")
        if not topic:
            return
        dasl = ("@SQL=\"http://schemas.microsoft.com/mapi/proptag/0x0070001E\" = "
                f"'{topic}'")
        for code in (OL_FOLDER_INBOX, OL_FOLDER_SENT):
            try:
                folder = self.namespace.GetDefaultFolder(code)
                for item in folder.Items.Restrict(dasl):
                    if getattr(item, "Class", 0) != OL_MAIL_ITEM:
                        continue
                    eid = item.EntryID
                    if eid in seen:
                        continue
                    seen.add(eid)
                    thread.append(self._to_thread_message(item))
            except Exception as e:
                logger.warning("Tìm theo topic thất bại: %s", e)

    # ------------------------------------------------------------------ soạn thư

    def create_draft_reply(self, entry_id: str, html_body: str,
                           reply_all: bool = False) -> Dict:
        if not self._ensure():
            return {}
        try:
            msg = self.namespace.GetItemFromID(entry_id)
            if getattr(msg, "Class", 0) != OL_MAIL_ITEM:
                raise ValueError("Mục gốc không phải email.")

            reply = msg.ReplyAll() if reply_all else msg.Reply()
            original = reply.HTMLBody or ""

            # Chèn NGAY SAU thẻ <body> để không phá cấu trúc HTML + chữ ký (sửa L13)
            m = _BODY_TAG_RE.search(original)
            if m:
                reply.HTMLBody = original[:m.end()] + html_body + "<br>" + original[m.end():]
            else:
                reply.HTMLBody = html_body + "<br><br>" + original

            reply.Save()
            return {"entry_id": reply.EntryID, "subject": reply.Subject or ""}
        except Exception as e:
            logger.error("Lỗi tạo thư nháp: %s", e)
            return {}

    def get_sent_emails_for_style(self, limit: int = 40) -> List[Dict]:
        """Lấy NỘI DUNG ĐẦY ĐỦ thư đã gửi để phân tích văn phong (sửa lỗi L2).

        Khác get_emails: trả về body thật đã lược bỏ trích dẫn, KHÔNG phải preview.
        """
        if not self._ensure():
            return []
        try:
            folder = self.namespace.GetDefaultFolder(OL_FOLDER_SENT)
            items = folder.Items
            items.Sort("[SentOn]", True)
            total = int(items.Count)

            out: List[Dict] = []
            i = 1
            while i <= total and len(out) < limit:
                try:
                    msg = items.Item(i)
                    if getattr(msg, "Class", 0) == OL_MAIL_ITEM:
                        body = strip_quoted_text(getattr(msg, "Body", "") or "")
                        if len(body) >= 40:            # bỏ thư quá ngắn, không đủ tín hiệu
                            out.append({
                                "subject": getattr(msg, "Subject", "") or "",
                                "body": body,
                            })
                except Exception:
                    pass
                i += 1
            return out
        except Exception as e:
            logger.error("Lỗi đọc thư đã gửi: %s", e)
            return []

    # ------------------------------------------------------------------ tiện ích

    def _attachment_count(self, msg) -> int:
        try:
            return int(msg.Attachments.Count)
        except Exception:
            return 0

    def _resolve_sender_email(self, msg) -> str:
        """Xử lý địa chỉ Exchange (X500) và SMTP."""
        try:
            sender = msg.Sender
            if getattr(sender, "AddressEntryUserObjectType", None) in (0, 30):
                exch = sender.GetExchangeUser()
                if exch:
                    return exch.PrimarySmtpAddress
        except Exception:
            pass
        return getattr(msg, "SenderEmailAddress", "") or ""

    def _email_to_summary(self, msg) -> Dict:
        """Chuyển MailItem -> dict cho DANH SÁCH.

        TUYỆT ĐỐI không đụng tới msg.HTMLBody ở đây: nó tải toàn bộ thân HTML qua
        COM cho từng thư và là nguyên nhân chính gây chậm (lỗi L6).
        """
        try:
            body = getattr(msg, "Body", "") or ""
            preview = " ".join(body[:400].split())[:160]
            return {
                "entry_id": msg.EntryID,
                "subject": getattr(msg, "Subject", "") or "",
                "sender_name": getattr(msg, "SenderName", "") or "",
                "sender_email": self._resolve_sender_email(msg),
                "received_time": _to_iso(getattr(msg, "ReceivedTime", None)),
                "is_unread": bool(getattr(msg, "UnRead", False)),
                "has_attachments": self._attachment_count(msg) > 0,
                "categories": getattr(msg, "Categories", "") or "",
                "preview": preview,
            }
        except Exception as e:
            logger.error("Lỗi chuyển email sang summary: %s", e)
            return {}

    def _to_thread_message(self, msg) -> Dict:
        try:
            sent = getattr(msg, "SentOn", None) or getattr(msg, "ReceivedTime", None)
            return {
                "entry_id": msg.EntryID,
                "subject": getattr(msg, "Subject", "") or "",
                "sender_name": getattr(msg, "SenderName", "") or "",
                "sender_email": self._resolve_sender_email(msg),
                "sent_time": _to_iso(sent),
                "body": strip_quoted_text(getattr(msg, "Body", "") or "", max_len=1500),
                "direction": "sent" if getattr(msg, "Sent", False) else "received",
            }
        except Exception as e:
            logger.error("Lỗi chuyển thư sang định dạng hội thoại: %s", e)
            return {}
```

> `beautifulsoup4` không còn cần thiết ở tầng này (HTML được render trong iframe ở frontend).
> Giữ nguyên trong `requirements.txt` — không gỡ, tránh phá `.venv` đang chạy.

---

## BƯỚC 4 — Ghi đè `backend/api.py`

Mọi lời gọi Outlook đi qua `com_call`; xóa sạch 8 lời gọi `pythoncom.CoInitialize()` rải rác.

```python
from __future__ import annotations

import logging
import traceback
from typing import Any, Dict, List, Optional

from dotenv import set_key

from config import BASE_DIR, CACHE_DIR, GEMINI_API_KEY, GEMINI_MODEL, STYLE_PROFILES_DIR

from .ai_engine import AIEngine, AIEngineError
from .classification_cache import ClassificationCache
from .com_worker import com_call
from .outlook_client import OutlookClient
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
    def __init__(self) -> None:
        self.outlook = OutlookClient()
        self.ai_engine = AIEngine(api_key=GEMINI_API_KEY, model_name=GEMINI_MODEL)
        self.style_analyzer = StyleAnalyzer(data_dir=STYLE_PROFILES_DIR)
        self.cache = ClassificationCache(CACHE_DIR)

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

    @_guard
    def generate_reply(self, entry_id: str, instruction: str,
                       reply_all: bool = False) -> Dict:
        detail = com_call(self.outlook.get_email_detail, entry_id)
        if not detail:
            return fail("Không tìm thấy email để tạo phản hồi.")
        thread = com_call(self.outlook.get_conversation_thread, entry_id)
        profile = self.style_analyzer.load_profile("default")
        html = self.ai_engine.generate_reply(detail, thread, instruction, profile)
        return ok({"html_body": html})

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

    # ------------------------------------------------------------------ Văn phong

    @_guard
    def analyze_style(self, limit: int = 40) -> Dict:
        # DÙNG get_sent_emails_for_style — KHÔNG dùng get_emails (sửa lỗi L2)
        samples = com_call(self.outlook.get_sent_emails_for_style, limit)
        if not samples:
            return fail("Không tìm thấy thư đã gửi nào đủ dài để phân tích.")
        profile = self.style_analyzer.analyze_and_save(samples, self.ai_engine, "default")
        return ok(profile)

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
        set_key(str(BASE_DIR / ".env"), "GEMINI_API_KEY", api_key)
        self.ai_engine = AIEngine(api_key=api_key, model_name=GEMINI_MODEL)
        return ok({"api_key_configured": self.ai_engine.is_ready})
```

### ⏸ ĐIỂM DỪNG BẮT BUỘC

Sau Bước 4, chạy kiểm chứng dưới đây và **báo cáo kết quả cho người dùng trước khi làm tiếp**:

```powershell
.\.venv\Scripts\python.exe -c "from backend.com_worker import com_call; from backend.outlook_client import OutlookClient; c = OutlookClient(); r = com_call(c.get_emails, None, 5, 0); [print(e['received_time'], '|', e['subject'][:50]) for e in r['items']]"
```

Tiêu chí đạt:
- In ra 5 email.
- `received_time` **khớp đúng giờ hiển thị trong Outlook** (đây là bằng chứng lỗi L5 đã hết —
  nếu lệch 7 tiếng là chưa sửa đúng).
- Chạy xong dưới 2 giây (bằng chứng lỗi L6 đã hết).

---

## BƯỚC 5 — Ghi đè `backend/ai_engine.py`

Sửa lỗi L10 + thêm phân loại tự động.

```python
from __future__ import annotations

import json
import logging
import re
from typing import Dict, List, Optional

import google.generativeai as genai

logger = logging.getLogger(__name__)

_FENCE_OPEN_RE = re.compile(r"^```[a-zA-Z]*\s*")
_FENCE_CLOSE_RE = re.compile(r"\s*```$")
_JSON_BLOCK_RE = re.compile(r"[\[{].*[\]}]", re.DOTALL)

VALID_CATEGORIES = ("Cần trả lời", "Việc cần làm", "Chỉ để biết", "Quảng cáo/Rác")
VALID_PRIORITIES = ("Cao", "Trung bình", "Thấp")


class AIEngineError(RuntimeError):
    """Lỗi thuộc về tầng AI — được api.py bắt và trả cho UI dưới dạng toast đỏ."""


def _strip_fence(text: str) -> str:
    text = (text or "").strip()
    if text.startswith("```"):
        text = _FENCE_CLOSE_RE.sub("", _FENCE_OPEN_RE.sub("", text))
    return text.strip()


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
    def __init__(self, api_key: str, model_name: str = "gemini-2.0-flash") -> None:
        self.api_key = (api_key or "").strip()
        self.model_name = model_name
        self.model = None          # model sinh văn bản (temperature cao)
        self.json_model = None     # model trả JSON (temperature thấp)

        if not self.api_key:
            logger.warning("Chưa có GEMINI_API_KEY — các tính năng AI sẽ bị tắt.")
            return

        genai.configure(api_key=self.api_key)
        self.model = genai.GenerativeModel(
            model_name, generation_config={"temperature": 0.7})
        self.json_model = genai.GenerativeModel(
            model_name,
            generation_config={"temperature": 0.2,
                               "response_mime_type": "application/json"},
        )

    @property
    def is_ready(self) -> bool:
        return self.model is not None

    def _require_key(self) -> None:
        if not self.is_ready:
            raise AIEngineError("Chưa cấu hình API key Gemini. Mở Cài đặt để nhập khóa.")

    def _generate(self, model, prompt: str, timeout: int = 90) -> str:
        self._require_key()
        try:
            resp = model.generate_content(prompt, request_options={"timeout": timeout})
        except Exception as e:
            raise AIEngineError(f"Gọi Gemini thất bại: {e}") from e
        text = getattr(resp, "text", None)
        if not text:
            raise AIEngineError("Gemini không trả về nội dung (có thể bị bộ lọc an toàn chặn).")
        return text

    # ------------------------------------------------------------------ văn phong

    def analyze_writing_style(self, sent_emails: List[Dict]) -> Dict:
        """sent_emails PHẢI chứa key 'body' với nội dung đầy đủ (xem
        OutlookClient.get_sent_emails_for_style), không phải preview 100 ký tự."""
        self._require_key()
        if not sent_emails:
            raise AIEngineError("Không có thư mẫu để phân tích.")

        blocks = []
        for e in sent_emails[:40]:
            body = (e.get("body") or "").strip()
            if not body:
                continue
            blocks.append(f"### Chủ đề: {e.get('subject', '')}\n{body[:2000]}")
        if not blocks:
            raise AIEngineError("Các thư mẫu đều rỗng sau khi lược bỏ trích dẫn.")

        prompt = (
            "Bạn là chuyên gia phân tích ngôn ngữ. Dưới đây là các email do MỘT người dùng "
            "tự tay viết và gửi đi. Hãy trích xuất phong cách viết đặc trưng của họ.\n\n"
            + "\n\n".join(blocks)
            + "\n\nTrả về DUY NHẤT một object JSON với các trường:\n"
              '- "greeting_patterns": mảng tối đa 5 câu chào thật sự xuất hiện\n'
              '- "closing_patterns": mảng tối đa 5 câu kết thật sự xuất hiện\n'
              '- "formality_level": số nguyên 1-5 (1 rất suồng sã, 5 rất trang trọng)\n'
              '- "common_phrases": mảng tối đa 8 cụm từ đặc trưng\n'
              '- "signature": khối chữ ký xuất hiện nhiều nhất, chuỗi rỗng nếu không có\n'
              '- "tone_notes": 1-2 câu mô tả giọng văn (dài/ngắn, trực tiếp/vòng vo, dùng emoji…)\n'
              '- "language": "Vietnamese" | "English" | "Bilingual"\n'
              "Chỉ dựa vào bằng chứng có trong các thư trên, không bịa."
        )

        data = _parse_json(self._generate(self.json_model, prompt))
        if not isinstance(data, dict):
            raise AIEngineError("Kết quả phân tích văn phong không đúng định dạng.")
        data["sample_count"] = len(blocks)
        return data

    # ------------------------------------------------------------------ soạn thư

    def generate_reply(self, original_email: Dict, conversation_thread: List[Dict],
                       user_instruction: str, style_profile: Optional[Dict] = None) -> str:
        self._require_key()
        prompt = f"""Bạn là trợ lý soạn email. Hãy viết nội dung thư phản hồi.

## Email cần trả lời
Người gửi: {original_email.get('sender_name', '')} <{original_email.get('sender_email', '')}>
Chủ đề: {original_email.get('subject', '')}
Nội dung:
{(original_email.get('body') or '')[:4000]}

## Lịch sử hội thoại
{self._format_thread(conversation_thread)}

## Chỉ dẫn của người dùng
{user_instruction}

## Phong cách bắt buộc tuân theo
{self._format_style(style_profile)}

## Quy tắc đầu ra
- Viết bằng ĐÚNG ngôn ngữ của email gốc.
- Chỉ dùng thẻ HTML cơ bản: <p>, <br>, <strong>, <em>, <ul>, <li>.
- KHÔNG kèm <html>, <head>, <body>, KHÔNG kèm markdown code fence.
- KHÔNG bịa thông tin không có trong email gốc hoặc chỉ dẫn.
- Chỉ trả về nội dung HTML của thư, không lời dẫn, không giải thích."""

        # KHÔNG bọc try/except trả chuỗi lỗi ở đây: lỗi phải nổi lên tới UI dưới dạng
        # toast đỏ, tuyệt đối không được chèn vào ô soạn thảo như thể là nội dung thư (lỗi L10).
        return self._clean_html(self._generate(self.model, prompt))

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

    @staticmethod
    def _clean_html(text: str) -> str:
        return _strip_fence(text)

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
            f"- Câu chào thường dùng: {', '.join(profile.get('greeting_patterns') or []) or '(không rõ)'}",
            f"- Câu kết thường dùng: {', '.join(profile.get('closing_patterns') or []) or '(không rõ)'}",
        ]
        if profile.get("common_phrases"):
            parts.append(f"- Cụm từ quen dùng: {', '.join(profile['common_phrases'][:8])}")
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
```

---

## BƯỚC 6 — Tệp mới `backend/classification_cache.py` + sửa `style_analyzer.py`

### 6.1 Tạo `backend/classification_cache.py`

`CACHE_DIR` hiện được `config.py` tạo ra nhưng không dòng code nào dùng (lỗi L12) — đây là chỗ dùng nó.

```python
"""Bộ nhớ đệm kết quả phân loại email, tránh gọi lại Gemini cho thư đã xử lý."""
from __future__ import annotations

import json
import logging
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List

logger = logging.getLogger(__name__)


class ClassificationCache:
    def __init__(self, cache_dir, max_entries: int = 2000) -> None:
        self.path = Path(cache_dir) / "classifications.json"
        self.max_entries = max_entries
        self._data: Dict[str, Dict] = self._load()

    def _load(self) -> Dict[str, Dict]:
        if not self.path.exists():
            return {}
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, dict) else {}
        except Exception as e:
            logger.warning("Không đọc được cache phân loại: %s", e)
            return {}

    def _save(self) -> None:
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.path, "w", encoding="utf-8") as f:
                json.dump(self._data, f, ensure_ascii=False)
        except Exception as e:
            logger.warning("Không ghi được cache phân loại: %s", e)

    def get_many(self, entry_ids: List[str]) -> Dict[str, Dict]:
        return {eid: self._data[eid] for eid in entry_ids if eid in self._data}

    def put_many(self, mapping: Dict[str, Dict]) -> None:
        if not mapping:
            return
        now = datetime.now(timezone.utc).isoformat()
        for entry_id, value in mapping.items():
            self._data[entry_id] = {**value, "cached_at": now}
        self._evict()
        self._save()

    def _evict(self) -> None:
        """Giữ lại max_entries mục mới nhất theo cached_at."""
        if len(self._data) <= self.max_entries:
            return
        ordered = sorted(self._data.items(),
                         key=lambda kv: kv[1].get("cached_at", ""), reverse=True)
        self._data = dict(ordered[:self.max_entries])
```

### 6.2 Sửa `backend/style_analyzer.py` — chỉ thay phương thức `analyze_and_save`

```python
    def analyze_and_save(self, sent_emails: List[Dict], ai_engine,
                         profile_name: str = 'default') -> Dict:
        """Phân tích văn phong rồi lưu ra JSON.

        LƯU Ý: sent_emails phải có key 'body' chứa nội dung ĐẦY ĐỦ.
        Ngoại lệ được để nổi lên api.py để UI báo lỗi rõ ràng — KHÔNG nuốt lỗi.
        """
        from datetime import datetime

        logger.info("Phân tích %d thư để dựng hồ sơ '%s'", len(sent_emails), profile_name)
        profile = ai_engine.analyze_writing_style(sent_emails)
        profile['name'] = profile_name
        profile['analyzed_at'] = datetime.now().astimezone().isoformat()

        filepath = self.data_dir / f"{profile_name}.json"
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(profile, f, ensure_ascii=False, indent=2)
        logger.info("Đã lưu hồ sơ văn phong: %s", filepath)
        return profile
```
Các phương thức `load_profile`, `list_profiles`, `delete_profile`: **giữ nguyên**, không sửa.

---

## BƯỚC 7 — Frontend

### 7.1 Ghi đè `frontend/index.html`

Gỡ 2 nút chết (`btn-forward`, `theme-toggle`), thêm bộ lọc phân loại + nút phân tích hộp thư,
để `folder-list` rỗng cho JS render động.

```html
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Trợ lý Email</title>
    <link rel="stylesheet" href="css/styles.css">
</head>
<body>
    <div id="app">
        <aside id="sidebar" class="glass-panel">
            <div class="sidebar-header">
                <h1 class="logo-text">✨ Trợ lý Email</h1>
            </div>
            <nav class="folder-list" id="folder-list">
                <div class="folder-loading">Đang tải thư mục…</div>
            </nav>
            <div class="sidebar-footer">
                <div class="profile-status" id="profile-status">
                    <div class="status-indicator"></div>
                    <span>Đang kiểm tra…</span>
                </div>
                <button id="btn-settings" class="btn-icon-text"><span>⚙️</span> Cài đặt</button>
            </div>
        </aside>

        <main id="email-list-panel" class="glass-panel">
            <div class="panel-header">
                <div class="search-container">
                    <span class="search-icon">🔍</span>
                    <input type="text" id="search-input" placeholder="Tìm kiếm email...">
                </div>
                <button id="btn-classify" class="btn-ai-action btn-classify" title="AI tóm tắt & phân loại các thư đang hiển thị">✨ Phân tích hộp thư</button>
                <div class="filter-bar" id="filter-bar">
                    <button class="filter-chip active" data-filter="all">Tất cả</button>
                    <button class="filter-chip" data-filter="Cần trả lời">Cần trả lời</button>
                    <button class="filter-chip" data-filter="Việc cần làm">Việc cần làm</button>
                    <button class="filter-chip" data-filter="Chỉ để biết">Chỉ để biết</button>
                </div>
            </div>
            <div id="email-list-container" class="scrollable"></div>
            <div class="pagination" id="pagination">
                <button id="btn-load-more" class="btn-secondary" style="display: none;">Tải thêm</button>
            </div>
        </main>

        <section id="detail-panel" class="glass-panel">
            <div id="email-empty-state" class="empty-state">
                <div class="empty-icon">✉️</div>
                <p>Chọn một email để đọc</p>
            </div>

            <div id="email-view" class="hidden">
                <div class="email-header" id="email-header"></div>
                <div class="email-actions">
                    <button class="btn-ai-action" id="btn-ai-reply">✨ Trả lời bằng AI</button>
                    <button class="btn-secondary" id="btn-ai-reply-all">✨ Trả lời tất cả</button>
                </div>
                <div class="detail-scroll scrollable">
                    <div class="email-body" id="email-body"></div>
                    <div class="email-attachments" id="email-attachments"></div>
                    <div class="conversation-thread" id="conversation-thread"></div>
                </div>
            </div>

            <div id="draft-panel" class="hidden glass-card sliding-panel">
                <div class="draft-header">
                    <h3>✨ Trợ lý AI Soạn thảo</h3>
                    <button class="btn-close" id="btn-close-draft">✖</button>
                </div>
                <div class="draft-content scrollable">
                    <div class="instruction-group">
                        <label for="ai-instruction">Chỉ dẫn cho AI:</label>
                        <textarea id="ai-instruction" placeholder="Ví dụ: Đồng ý tham gia cuộc họp sáng thứ 3, hỏi thêm về tài liệu cần chuẩn bị..."></textarea>
                    </div>
                    <button id="btn-generate-draft" class="btn-gradient w-100">
                        <span class="btn-text">Tạo nháp</span>
                        <span class="loader hidden"></span>
                    </button>
                    <div id="draft-result-area" class="hidden mt-4">
                        <label>Bản nháp:</label>
                        <div id="draft-editor" contenteditable="true" class="editable-draft glass-input"></div>
                        <div class="refine-group mt-3">
                            <input type="text" id="refine-instruction" class="glass-input w-100" placeholder="Yêu cầu chỉnh sửa thêm (ví dụ: Viết trang trọng hơn)">
                            <button id="btn-refine-draft" class="btn-secondary w-100 mt-2">Tinh chỉnh</button>
                        </div>
                        <div class="draft-actions mt-4">
                            <button id="btn-save-draft" class="btn-success w-100">Lưu vào Thư nháp Outlook</button>
                        </div>
                    </div>
                </div>
            </div>
        </section>
    </div>

    <div id="settings-modal" class="modal-overlay hidden">
        <div class="modal-card glass-card">
            <h2>Cài đặt</h2>
            <div class="form-group">
                <label>Gemini API Key</label>
                <input type="password" id="api-key-input" class="glass-input" placeholder="Nhập khóa API...">
                <small class="text-muted mt-1 d-block" id="api-key-status"></small>
            </div>
            <div class="form-group mt-4">
                <button id="btn-analyze-style" class="btn-secondary w-100">Phân tích văn phong</button>
                <small class="text-muted mt-1 d-block text-center" id="style-status">AI sẽ đọc thư đã gửi để học cách bạn viết.</small>
            </div>
            <div class="modal-actions">
                <button id="btn-close-settings" class="btn-secondary">Đóng</button>
                <button id="btn-save-settings" class="btn-gradient">Lưu</button>
            </div>
        </div>
    </div>

    <div id="global-loading" class="modal-overlay hidden"><div class="spinner"></div></div>
    <div id="toast-container"></div>

    <script src="js/utils.js"></script>
    <script src="js/email-list.js"></script>
    <script src="js/email-viewer.js"></script>
    <script src="js/draft-editor.js"></script>
    <script src="js/app.js"></script>
</body>
</html>
```

### 7.2 `frontend/js/utils.js` — 2 thay đổi

**(a) XÓA HOÀN TOÀN** khối `sanitizeHtml` (dòng 97-101) — nó là hàm rỗng trả về nguyên HTML, và
việc còn tồn tại khiến người sau tưởng nội dung đã an toàn.

**(b) XÓA HOÀN TOÀN** khối `api: { call: ... }` (dòng 109-123) — không nơi nào dùng, mọi chỗ gọi
thẳng `window.pywebview.api`.

**(c) THÊM** hàm bọc HTML email cho iframe:

```js
    /**
     * Bọc HTML thư vào một tài liệu độc lập để render trong iframe sandbox.
     * CSP chặn mọi tài nguyên ngoài (kể cả tracking pixel); sandbox chặn script.
     */
    buildEmailDocument: (html) => {
        const csp = "default-src 'none'; img-src data: cid:; style-src 'unsafe-inline'; font-src data:";
        return `<!DOCTYPE html><html><head><meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="${csp}">
<base target="_blank">
<style>
  html,body{margin:0;padding:16px;background:#12122a;color:#dcdce6;
    font-family:'Segoe UI Variable','Segoe UI',system-ui,sans-serif;
    font-size:14px;line-height:1.6;word-wrap:break-word}
  a{color:#7c5cfc}
  img{max-width:100%;height:auto}
  table{max-width:100%!important}
  blockquote{border-left:3px solid #7c5cfc55;margin:8px 0;padding-left:12px;color:#8888a8}
</style></head><body>${html || '<p style="color:#55556a">Không có nội dung</p>'}</body></html>`;
    },
```

### 7.3 `frontend/js/email-viewer.js` — thay 3 phương thức

**(a) `bindEvents`** — nối lại theo ID mới trong HTML:
```js
    bindEvents() {
        document.getElementById('btn-ai-reply')?.addEventListener('click',
            () => window.DraftEditorUI?.showDraftPanel(false));
        document.getElementById('btn-ai-reply-all')?.addEventListener('click',
            () => window.DraftEditorUI?.showDraftPanel(true));
    }
```

**(b) `renderEmailBody`** — thay `innerHTML` bằng iframe sandbox (**sửa lỗi XSS L3**):
```js
    renderEmailBody(htmlContent) {
        if (!this.bodyContainer) return;
        this.bodyContainer.innerHTML = '';

        const iframe = document.createElement('iframe');
        iframe.className = 'email-iframe';
        // sandbox KHÔNG có allow-scripts và KHÔNG có allow-same-origin
        // => nội dung thư không thể chạm tới window.pywebview.api của trang cha.
        // allow-popups để link trong thư vẫn mở được ra trình duyệt.
        iframe.setAttribute('sandbox', 'allow-popups allow-popups-to-escape-sandbox');
        iframe.setAttribute('referrerpolicy', 'no-referrer');
        iframe.srcdoc = window.Utils.buildEmailDocument(htmlContent);
        this.bodyContainer.appendChild(iframe);
    }
```

**(c) `renderConversationThread`** — bỏ `onclick` inline, dùng event delegation:
```js
    renderConversationThread(messages) {
        if (!this.threadContainer) return;
        if (!messages || messages.length <= 1) {
            this.threadContainer.innerHTML = '';
            return;
        }

        const rows = messages.map((msg, index) => {
            const senderName = window.Utils.escapeHtml(msg.sender_name || 'Không rõ');
            const timeStr = window.Utils.formatDate(msg.sent_time);
            const body = window.Utils.escapeHtml(
                window.Utils.truncate((msg.body || '').replace(/\s+/g, ' '), 400));
            const icon = msg.direction === 'sent' ? '📤' : '📥';
            const expanded = index === messages.length - 1 ? 'expanded' : '';
            return `<div class="thread-item ${expanded}">
                <div class="thread-dot"></div>
                <div class="thread-card">
                    <div class="thread-card-header">
                        <div class="thread-sender">${icon} <strong>${senderName}</strong></div>
                        <span class="thread-time">${timeStr}</span>
                    </div>
                    <div class="thread-body">${body}</div>
                </div></div>`;
        }).join('');

        this.threadContainer.innerHTML =
            `<div class="thread-section"><h4 class="thread-title">💬 Lịch sử hội thoại</h4>
             <div class="thread-timeline">${rows}</div></div>`;

        this.threadContainer.onclick = (e) => {
            e.target.closest('.thread-item')?.classList.toggle('expanded');
        };
    }
```
Xóa phương thức `toggleThread` (không còn được gọi). Các phương thức khác giữ nguyên.

### 7.4 `frontend/js/email-list.js` — 4 thay đổi

**(a)** Trong `constructor`, thêm `this.classifications = {};` và `this.activeFilter = 'all';`

**(b)** Trong `bindEvents`, thêm xử lý bộ lọc + nút phân loại:
```js
        document.getElementById('filter-bar')?.addEventListener('click', (e) => {
            const chip = e.target.closest('.filter-chip');
            if (!chip) return;
            document.querySelectorAll('.filter-chip').forEach(c => c.classList.remove('active'));
            chip.classList.add('active');
            this.activeFilter = chip.dataset.filter;
            window.App?.rerenderList();
        });

        document.getElementById('btn-classify')?.addEventListener('click',
            () => window.App?.classifyVisible());
```

**(c)** Thay `renderEmailList` — nhận thêm `hasMore`, hỗ trợ lọc, bỏ ngưỡng cứng `>= 20`:
```js
    renderEmailList(emails, selectedId = null, hasMore = false) {
        this.container.innerHTML = '';

        const visible = this.activeFilter === 'all'
            ? emails
            : emails.filter(e => this.classifications[e.entry_id]?.category === this.activeFilter);

        if (!visible.length) {
            const isSearch = this.searchInput?.value.trim() !== '';
            this.renderEmptyState(isSearch);
            return;
        }

        visible.forEach((email, index) => {
            const el = this.createEmailItem(email, email.entry_id === selectedId);
            el.style.animationDelay = `${Math.min(index, 12) * 0.03}s`;
            this.container.appendChild(el);
        });

        if (this.btnLoadMore) {
            const canPage = hasMore && this.activeFilter === 'all'
                            && !this.searchInput?.value.trim();
            this.btnLoadMore.style.display = canPage ? 'block' : 'none';
        }
    }
```

**(d)** Trong `createEmailItem`, chèn chip phân loại và ưu tiên `summary` của AI:
```js
        const cls = this.classifications[email.entry_id];
        const summaryText = cls?.summary || email.preview || '';
        const preview = window.Utils.truncate(window.Utils.escapeHtml(summaryText), 120);

        const CAT_SLUG = {
            'Cần trả lời': 'reply', 'Việc cần làm': 'todo',
            'Chỉ để biết': 'info', 'Quảng cáo/Rác': 'spam'
        };
        const PRIO_SLUG = { 'Cao': 'high', 'Trung bình': 'mid', 'Thấp': 'low' };
        const chip = cls
            ? `<span class="cat-chip cat-${CAT_SLUG[cls.category] || 'info'}">
                 <i class="prio-dot prio-${PRIO_SLUG[cls.priority] || 'mid'}"></i>
                 ${window.Utils.escapeHtml(cls.category)}</span>`
            : '';
```
rồi thêm `${chip}` vào cuối `innerHTML`, ngay sau `<div class="email-preview">…</div>`.

**(e)** Đổi `updateUnreadCount` thành `renderFolders` (danh sách động, badge dùng `unread_count` — sửa L7):
```js
    renderFolders(folders, activeId) {
        const list = document.getElementById('folder-list');
        if (!list) return;
        const ICONS = { inbox: '📥', sent: '📤', drafts: '📝', other: '📁' };

        list.innerHTML = folders.map(f => `
            <div class="folder-item ${f.entry_id === activeId ? 'active' : ''}"
                 data-folder-id="${f.entry_id}" style="padding-left:${12 + f.depth * 14}px">
                <span class="folder-icon">${ICONS[f.kind] || ICONS.other}</span>
                <span class="folder-name">${window.Utils.escapeHtml(f.name)}</span>
                <span class="badge" ${f.unread_count > 0 ? '' : 'style="display:none"'}>${f.unread_count}</span>
            </div>`).join('');
    }
```

### 7.5 `frontend/js/app.js` — ghi đè các phần sau

Giữ nguyên `bindEvents` phần Settings modal; thay phần còn lại:

```js
class AppController {
    constructor() {
        this.state = {
            folders: [],
            currentFolderId: null,
            currentEmailId: null,
            emails: [],
            offset: 0,
            limit: 30,
            hasMore: false,
            searchQuery: '',
        };
        this.init();
    }

    init() {
        window.addEventListener('pywebviewready', () => this.onReady());
        this.bindEvents();
    }

    async onReady() {
        window.Utils.showLoading(true);
        try {
            const conn = await window.pywebview.api.connect_outlook();
            if (!conn?.success) {
                window.Utils.showToast(conn?.error || 'Không kết nối được Outlook', 'error');
                return;
            }
            window.Utils.showToast('Đã kết nối Outlook', 'success');
            await this.loadFolders();
            await this.loadEmails(true);
            await this.checkStyleProfile();
        } catch (e) {
            window.Utils.showToast('Lỗi khởi tạo: ' + (e.message || e), 'error');
        } finally {
            window.Utils.showLoading(false);
        }
    }

    async loadFolders() {
        const res = await window.pywebview.api.get_folders();
        if (!res?.success) return;
        this.state.folders = res.data || [];
        const inbox = this.state.folders.find(f => f.kind === 'inbox') || this.state.folders[0];
        this.state.currentFolderId = inbox?.entry_id || null;
        window.EmailListUI.renderFolders(this.state.folders, this.state.currentFolderId);
    }

    async switchFolder(folderId) {
        if (this.state.currentFolderId === folderId) return;
        this.state.currentFolderId = folderId;
        this.state.searchQuery = '';
        this.state.currentEmailId = null;
        const s = document.getElementById('search-input');
        if (s) s.value = '';
        window.EmailListUI.renderFolders(this.state.folders, folderId);
        window.EmailViewerUI?.showEmptyDetail();
        window.DraftEditorUI?.hideDraftPanel();
        await this.loadEmails(true);
    }

    async loadEmails(reset = false) {
        if (reset) { this.state.offset = 0; this.state.emails = []; this.state.hasMore = false; }
        window.EmailListUI.renderSkeletonLoading();
        try {
            let items = [], hasMore = false, nextOffset = this.state.offset;
            if (this.state.searchQuery) {
                const res = await window.pywebview.api.search_emails(
                    this.state.searchQuery, this.state.currentFolderId);
                if (res?.success) items = res.data || [];
            } else {
                const res = await window.pywebview.api.get_emails(
                    this.state.currentFolderId, this.state.limit, this.state.offset);
                if (res?.success) {
                    items = res.data.items || [];
                    hasMore = !!res.data.has_more;
                    nextOffset = res.data.next_offset;
                }
            }
            this.state.emails = reset ? items : this.state.emails.concat(items);
            this.state.hasMore = hasMore;
            this.state.offset = nextOffset;
            this.rerenderList();
        } catch (e) {
            window.Utils.showToast('Lỗi tải email: ' + (e.message || e), 'error');
            window.EmailListUI.renderEmptyState();
        }
    }

    rerenderList() {
        window.EmailListUI.renderEmailList(
            this.state.emails, this.state.currentEmailId, this.state.hasMore);
    }

    async loadMoreEmails() { await this.loadEmails(false); }

    async handleSearch(query) {
        this.state.searchQuery = query;
        await this.loadEmails(true);
    }

    async selectEmail(entryId) {
        this.state.currentEmailId = entryId;
        window.DraftEditorUI?.hideDraftPanel();
        try {
            const res = await window.pywebview.api.get_email_detail(entryId);
            if (!res?.success) {
                window.Utils.showToast(res?.error || 'Không tải được email', 'error');
                return;
            }
            window.EmailViewerUI.renderEmailDetail(res.data);

            // Đánh dấu đã đọc trong Outlook THẬT, không chỉ đổi state JS
            window.pywebview.api.mark_as_read(entryId);
            const item = this.state.emails.find(e => e.entry_id === entryId);
            if (item) item.is_unread = false;

            const th = await window.pywebview.api.get_conversation(entryId);
            if (th?.success) window.EmailViewerUI.renderConversationThread(th.data);
        } catch (e) {
            window.Utils.showToast('Lỗi tải email: ' + (e.message || e), 'error');
        }
    }

    async classifyVisible() {
        const pending = this.state.emails
            .filter(e => !window.EmailListUI.classifications[e.entry_id])
            .slice(0, 45);          // giới hạn 3 lô mỗi lần bấm
        if (!pending.length) {
            window.Utils.showToast('Tất cả thư đang hiển thị đã được phân loại', 'info');
            return;
        }
        window.Utils.showToast(`Đang phân tích ${pending.length} thư…`, 'info');
        try {
            const res = await window.pywebview.api.classify_emails(pending);
            if (!res?.success) {
                window.Utils.showToast(res?.error || 'Lỗi phân loại', 'error');
                return;
            }
            Object.assign(window.EmailListUI.classifications, res.data || {});
            this.rerenderList();
            window.Utils.showToast('✅ Đã phân tích xong', 'success');
        } catch (e) {
            window.Utils.showToast('Lỗi phân loại: ' + (e.message || e), 'error');
        }
    }
```

Trong `bindEvents`, đổi phần click thư mục sang `data-folder-id` và mở modal thì nạp cài đặt:
```js
        document.getElementById('folder-list')?.addEventListener('click', (e) => {
            const item = e.target.closest('.folder-item');
            if (item?.dataset.folderId) this.switchFolder(item.dataset.folderId);
        });

        document.getElementById('btn-settings')?.addEventListener('click', async () => {
            document.getElementById('settings-modal')?.classList.remove('hidden');
            const res = await window.pywebview.api.get_settings();   // API này trước đây KHÔNG được gọi ở đâu
            if (res?.success) {
                document.getElementById('api-key-status').textContent =
                    res.data.api_key_configured
                        ? `✅ Đã cấu hình (model: ${res.data.model})`
                        : '⚠️ Chưa cấu hình API key';
                document.getElementById('style-status').textContent =
                    res.data.style_analyzed_at
                        ? `Lần phân tích gần nhất: ${window.Utils.formatDate(res.data.style_analyzed_at)}`
                        : 'AI sẽ đọc thư đã gửi để học cách bạn viết.';
            }
        });
```

Các phương thức `generateDraft` / `refineDraft` / `saveDraft` / `analyzeStyle` / `checkStyleProfile`:
**giữ nguyên logic cũ**. **XÓA HOÀN TOÀN** `mockInit()` và khối `setTimeout` 2 giây gọi nó trong
`init()` — dữ liệu giả che lấp lỗi kết nối thật.

### 7.6 `frontend/css/styles.css`

**(a) Sửa dòng 1** — bỏ phụ thuộc internet (lỗi L15):
```css
/* XÓA: @import url('https://fonts.googleapis.com/...'); */
```
và sửa `body { font-family: ... }` thành:
```css
    font-family: 'Segoe UI Variable', 'Segoe UI', system-ui, -apple-system, sans-serif;
```

**(b) Thêm khối sau vào CUỐI tệp** — bổ sung toàn bộ CSS đang thiếu (lỗi L9):

```css
/* ============ Thư mục động ============ */
.folder-loading { padding: 16px; color: var(--text-muted); font-size: 0.85rem; }
.folder-item { display: flex; align-items: center; gap: 10px; }
.folder-name { flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }

/* ============ Chip phân loại ============ */
.filter-bar { display: flex; gap: 6px; margin-top: 10px; flex-wrap: wrap; }
.filter-chip {
    padding: 4px 12px; border-radius: 999px; font-size: 0.75rem; cursor: pointer;
    background: var(--bg-glass); border: 1px solid var(--border-glass);
    color: var(--text-secondary); transition: all var(--transition-fast);
}
.filter-chip:hover { background: var(--bg-glass-hover); color: var(--text-primary); }
.filter-chip.active { background: var(--accent-gradient); color: #fff; border-color: transparent; }
.btn-classify { width: 100%; margin-top: 10px; font-size: 0.8rem; padding: 8px; }

.cat-chip {
    display: inline-flex; align-items: center; gap: 6px; margin-top: 8px;
    padding: 3px 10px; border-radius: 999px; font-size: 0.7rem; font-weight: 500;
}
.cat-reply { background: rgba(124, 92, 252, 0.18); color: #b3a0ff; }
.cat-todo  { background: rgba(255, 165, 2, 0.18); color: #ffc46b; }
.cat-info  { background: rgba(136, 136, 168, 0.16); color: var(--text-secondary); }
.cat-spam  { background: rgba(255, 71, 87, 0.15); color: #ff8a94; }
.prio-dot { width: 6px; height: 6px; border-radius: 50%; display: inline-block; }
.prio-high { background: var(--danger); }
.prio-mid  { background: var(--warning); }
.prio-low  { background: var(--text-muted); }

.unread-dot {
    width: 7px; height: 7px; border-radius: 50%;
    background: var(--accent-primary); flex-shrink: 0;
    box-shadow: 0 0 8px var(--accent-primary);
}
.attach-icon { font-size: 0.8rem; opacity: 0.7; }
.email-item-meta { display: flex; align-items: center; gap: 6px; }

/* ============ Chi tiết email ============ */
.detail-scroll { flex: 1; min-height: 0; display: flex; flex-direction: column; }
.avatar.large { width: 48px; height: 48px; font-size: 1rem; }
.sender-details { display: flex; flex-direction: column; gap: 2px; }
.sender-details .sender-name { font-size: 1rem; font-weight: 600; }
.sender-email { font-size: 0.8rem; color: var(--text-muted); }
.email-date { font-size: 0.8rem; color: var(--text-secondary); white-space: nowrap; }
.email-recipients { font-size: 0.8rem; color: var(--text-secondary); }
.email-meta-line { margin-top: 4px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }

.email-body { padding: 0; min-height: 340px; display: flex; }
.email-iframe {
    width: 100%; min-height: 340px; flex: 1;
    border: 0; background: var(--bg-secondary);
}

/* ============ Đính kèm ============ */
.attachments-section { padding: 16px 24px; border-top: 1px solid var(--border-glass); }
.attachments-title { font-size: 0.85rem; color: var(--text-secondary); margin-bottom: 12px; }
.attachments-grid { display: flex; flex-wrap: wrap; gap: 10px; }
.attachment-item {
    display: flex; align-items: center; gap: 8px; padding: 8px 14px;
    background: var(--bg-glass); border: 1px solid var(--border-glass);
    border-radius: var(--radius-md); font-size: 0.82rem;
    transition: all var(--transition-fast);
}
.attachment-item:hover { background: var(--bg-glass-hover); transform: translateY(-1px); }
.att-icon { font-size: 1.1rem; }
.att-name { max-width: 200px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.att-size { color: var(--text-muted); font-size: 0.75rem; }

/* ============ Lịch sử hội thoại ============ */
.thread-section { padding: 16px 24px 32px; border-top: 1px solid var(--border-glass); }
.thread-title { font-size: 0.85rem; color: var(--text-secondary); margin-bottom: 16px; }
.thread-timeline { position: relative; padding-left: 20px; }
.thread-timeline::before {
    content: ''; position: absolute; left: 5px; top: 6px; bottom: 6px;
    width: 2px; background: linear-gradient(180deg, var(--accent-primary), transparent);
    opacity: 0.35;
}
.thread-item { position: relative; margin-bottom: 12px; }
.thread-dot {
    position: absolute; left: -19px; top: 14px; width: 8px; height: 8px;
    border-radius: 50%; background: var(--accent-primary);
    box-shadow: 0 0 0 3px var(--bg-primary);
}
.thread-card {
    background: var(--bg-glass); border: 1px solid var(--border-glass);
    border-radius: var(--radius-md); padding: 12px 14px; cursor: pointer;
    transition: background var(--transition-fast);
}
.thread-card:hover { background: var(--bg-glass-hover); }
.thread-card-header {
    display: flex; justify-content: space-between; align-items: center;
    margin-bottom: 6px; font-size: 0.82rem;
}
.thread-sender { color: var(--text-primary); }
.thread-time { color: var(--text-muted); font-size: 0.75rem; }
.thread-body {
    font-size: 0.82rem; color: var(--text-secondary); line-height: 1.5;
    max-height: 2.9em; overflow: hidden;
    transition: max-height var(--transition-slow);
}
.thread-item.expanded .thread-body { max-height: 400px; overflow-y: auto; }

/* ============ Bảng soạn nháp ============ */
.sliding-panel { display: flex; flex-direction: column; }
.draft-content { flex: 1; min-height: 0; }
.refine-group { display: flex; flex-direction: column; }
.draft-actions { padding-bottom: 8px; }
```

### 7.7 `frontend/js/draft-editor.js` — 1 thay đổi

Trong `setLoading`, thêm khóa/mở nút Lưu để tránh lưu bản nháp dở dang:
```js
        this.btnSave.disabled = isLoading;
        this.btnRefine.disabled = isLoading;
```
Phần còn lại giữ nguyên.

---

## BƯỚC 8 — Tài liệu & bàn giao

**8.1** Ghim phiên bản trong `requirements.txt` (chạy `.\.venv\Scripts\pip.exe freeze` để lấy số thật):
```
pywebview==6.2.1
pywin32==<số thật>
google-generativeai==0.8.6
python-dotenv==<số thật>
beautifulsoup4==<số thật>
```

**8.2** Tạo `README.md`: yêu cầu hệ thống (Windows + Outlook desktop đã đăng nhập), cách tạo `.venv`,
cách lấy Gemini API key, cách chạy (`python main.py`, `DEBUG=1` để bật DevTools).

**8.3** Tạo `HUONG_DAN_SU_DUNG.md` (tiếng Việt) kèm ảnh chụp màn hình thật của 4 luồng chính:
đọc thư → phân tích hộp thư → phân tích văn phong → soạn thư trả lời bằng AI.

**8.4** Ghi dòng cuối vào `development_log.csv` với phiên bản `1.1.0`, rồi commit git.

---

## 9. NGHIỆM THU — Chạy thử bắt buộc trước khi báo hoàn thành

Mở Outlook trước, chạy `$env:DEBUG="1"; .\.venv\Scripts\python.exe main.py`.

| # | Thao tác | Tiêu chí ĐẠT | Chứng minh đã sửa lỗi |
|---|----------|--------------|----------------------|
| N1 | Khởi động app | Cửa sổ mở trong < 3 giây | L1 |
| N2 | Xem sidebar | Liệt kê thư mục thật (kể cả thư mục con), badge = **số thư chưa đọc** khớp Outlook | L7 |
| N3 | Xem cột giữa | Thời gian mỗi thư **khớp đúng** với Outlook (không lệch 7h); tải 30 thư < 2 giây | L5, L6 |
| N4 | Mở 1 thư có ảnh + đính kèm | Thân thư hiển thị trong iframe, đính kèm có style đẹp, thread hiện bên dưới, cuộn mượt | L9 |
| N5 | **Kiểm tra XSS** — tự gửi cho mình thư chứa `<img src=x onerror="parent.pywebview.api.get_folders()">` rồi mở | Console báo lỗi sandbox/CSP; **không** có lời gọi API nào; ảnh remote bị chặn | **L3** |
| N6 | Mở thư chưa đọc | Thư chuyển sang đã đọc **trong Outlook thật** | L14 |
| N7 | Bấm "✨ Phân tích hộp thư" | Mỗi thư có chip phân loại + tóm tắt 1 câu. Đổi thư mục rồi quay lại → hiện **tức thì**, không gọi lại AI. Kiểm tra `data/cache/classifications.json` đã có nội dung | Tính năng B, L12 |
| N8 | Bấm các chip lọc | Danh sách lọc đúng, nút "Tải thêm" tự ẩn khi đang lọc | — |
| N9 | Cài đặt → "Phân tích văn phong" | Mở `data/style_profiles/default.json`: `greeting_patterns`, `signature`, `tone_notes` **thật sự khớp** cách bạn viết; `sample_count` ≥ 20 | **L2** |
| N10 | "✨ Trả lời bằng AI" → tạo → tinh chỉnh → lưu | Mở Outlook: thư nháp **giữ nguyên chữ ký và trích dẫn thư gốc**, HTML không vỡ | L13 |
| N11 | Sửa `.env` thành key rác → tạo nháp | Hiện **toast đỏ**; ô soạn thảo **trống** (không bị chèn thông báo lỗi như thể là nội dung thư) | **L10** |
| N12 | Gõ `it's` vào ô tìm kiếm | Không lỗi, trả kết quả hoặc rỗng bình thường | L11 |
| N13 | Chuyển thư mục liên tục 20 lần thật nhanh, mở/đóng thư liên tục | Console **không** xuất hiện `marshalled for a different thread` | **L4** |
| N14 | Ngắt mạng internet, khởi động lại app | Giao diện vẫn hiển thị đầy đủ font, đọc thư bình thường (chỉ tính năng AI báo lỗi) | L15 |

**Không được báo cáo hoàn thành nếu bất kỳ mục nào từ N1–N14 chưa đạt.** Với mục không đạt, ghi rõ
hiện tượng và lỗi trong console, không tự ý đổi tiêu chí nghiệm thu.
