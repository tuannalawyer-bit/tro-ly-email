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

# Nhận diện thư mục kiểu "đã gửi" — trong đó MỌI thư đều là thư mình gửi đi.
# Dùng so khớp tiền tố chứ không phải danh sách cứng: kho lưu trữ thực tế có cả
# "Sent Items", "sent items", "Sent item" (số ít) — danh sách cứng bỏ sót ngay.
def is_sent_folder_name(name: str) -> bool:
    low = (name or "").strip().lower()
    return low.startswith("sent") or "đã gửi" in low


# Inbox chỉ chứa thư NHẬN. Quét nó khi tìm thư đã gửi là đọc thừa hàng nghìn mục
# rồi loại hết — trên hộp thư thật đây là hơn một phần ba tổng khối lượng.
INBOX_FOLDER_NAMES = {"inbox", "hộp thư đến"}

# Không quét: thư rác, thùng rác, nháp, hàng chờ gửi, lịch sử chat.
SKIP_FOLDER_NAMES = {
    "deleted items", "junk email", "junk e-mail", "drafts", "outbox",
    "conversation history", "rss feeds", "sync issues", "quarantine",
    "thùng rác", "thư nháp", "hộp thư đi", "thư rác",
}


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


def split_quoted_text(body: str) -> tuple:
    """Tách thư thành (phần người dùng viết, phần trích dẫn thư cũ).

    Phần trích dẫn CHÍNH LÀ thư mà người ta gửi tới — dùng để dựng cặp
    "thư đến → cách bạn trả lời" khi xuất thư mà không tốn thêm lần gọi COM nào.
    """
    if not body:
        return "", ""
    text = body.replace("\r\n", "\n")
    cut = len(text)
    for marker in _QUOTE_MARKERS:
        idx = text.find(marker)
        if 0 <= idx < cut:
            cut = idx
    m = _ON_WROTE_RE.search(text)
    if m and m.start() < cut:
        cut = m.start()
    return text[:cut].strip(), text[cut:].strip()


def strip_quoted_text(body: str, max_len: int = 3000) -> str:
    """Bỏ phần trích dẫn thư cũ, chỉ giữ nội dung người dùng thực sự viết.

    max_len=0 nghĩa là KHÔNG cắt — bắt buộc dùng khi xuất thư để phân tích văn phong,
    vì cắt ở 3000 ký tự sẽ làm cụt đúng những thư dài, tức là những thư giàu thông tin
    nhất về cách người dùng triển khai lập luận.
    """
    own, _ = split_quoted_text(body)
    return own if max_len <= 0 else own[:max_len]


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

    def get_sent_emails_for_style(self, limit: int = 0) -> List[Dict]:
        """Lấy NỘI DUNG ĐẦY ĐỦ thư đã gửi để phân tích văn phong (sửa lỗi L2).

        Khác get_emails: trả về body thật đã lược bỏ trích dẫn, KHÔNG phải preview.
        limit=0 nghĩa là quét toàn bộ thư đủ dài trong Sent Items.
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
            while i <= total and (limit <= 0 or len(out) < limit):
                try:
                    msg = items.Item(i)
                    if getattr(msg, "Class", 0) == OL_MAIL_ITEM:
                        body = strip_quoted_text(getattr(msg, "Body", "") or "")
                        if len(body) >= 40:            # bỏ thư quá ngắn, không đủ tín hiệu
                            out.append({
                                "subject": getattr(msg, "Subject", "") or "",
                                "body": body,
                                "sent_time": _to_iso(getattr(msg, "SentOn", None)),
                                "to": getattr(msg, "To", "") or "",
                            })
                except Exception:
                    pass
                i += 1
            return out
        except Exception as e:
            logger.error("Lỗi đọc thư đã gửi: %s", e)
            return []

    # -------------------------------------------------- quét toàn bộ kho thư

    def get_my_addresses(self) -> set:
        """Mọi địa chỉ SMTP thuộc về người dùng, để nhận ra thư do họ gửi."""
        out = set()
        if not self._ensure():
            return out
        try:
            user = self.namespace.CurrentUser
            exch = user.AddressEntry.GetExchangeUser()
            if exch and exch.PrimarySmtpAddress:
                out.add(exch.PrimarySmtpAddress.strip().lower())
        except Exception:
            pass
        try:                                    # tài khoản thứ hai, tài khoản POP/IMAP
            for account in self.namespace.Accounts:
                addr = (getattr(account, "SmtpAddress", "") or "").strip().lower()
                if addr:
                    out.add(addr)
        except Exception:
            pass
        logger.info("Địa chỉ của bạn: %s", ", ".join(sorted(out)) or "(không xác định)")
        return out

    def iter_mail_folders(self) -> List[Dict]:
        """Liệt kê mọi thư mục thư trong MỌI kho đang gắn vào Outlook.

        _walk_folders sẵn có chỉ duyệt kho mặc định. Hàm này đi qua namespace.Stores
        nên với tới được các tệp PST lưu trữ đã gắn thêm.
        """
        if not self._ensure():
            return []
        out: List[Dict] = []
        try:
            stores = list(self.namespace.Stores)
        except Exception as e:
            logger.warning("Không đọc được danh sách kho: %s", e)
            return out

        for store in stores:
            store_name = "(không tên)"
            try:
                store_name = store.DisplayName or store_name
                root = store.GetRootFolder()
            except Exception as e:
                logger.warning("Bỏ qua kho %s: %s", store_name, e)
                continue
            self._collect_folders(root, store_name, out, 0)
        return out

    def _collect_folders(self, parent, store_name: str, out: List[Dict],
                         depth: int) -> None:
        if depth > 8:                            # chặn cây thư mục lồng bất thường
            return
        try:
            children = list(parent.Folders)
        except Exception:
            return
        for folder in children:
            try:
                name = folder.Name or ""
                low = name.strip().lower()
                if low in SKIP_FOLDER_NAMES:
                    continue
                if getattr(folder, "DefaultItemType", -1) == 0:   # 0 = thư mục thư
                    out.append({
                        "folder": folder,
                        "name": name,
                        "store": store_name,
                        "path": f"{store_name} / {name}",
                        "count": int(folder.Items.Count),
                        "is_sent": is_sent_folder_name(name)
                                   or folder.EntryID == self._default_ids.get("sent"),
                        "is_inbox": low in INBOX_FOLDER_NAMES,
                    })
                self._collect_folders(folder, store_name, out, depth + 1)
            except Exception as e:
                logger.warning("Bỏ qua thư mục lỗi: %s", e)

    def collect_sent_emails(self, progress=None, deep: bool = True) -> List[Dict]:
        """Gom MỌI thư do người dùng gửi, trên mọi kho đang gắn vào Outlook.

        deep=False: chỉ quét thư mục kiểu Sent (nhanh).
        deep=True : quét thêm thư mục lưu trữ trộn lẫn, giữ thư có người gửi là bạn.

        Trả về thư CHƯA cắt ngắn, kèm phần trích dẫn (chính là thư đến).
        """
        if not self._ensure():
            return []
        folders = self.iter_mail_folders()
        targets = [f for f in folders
                   if f["is_sent"] or (deep and not f.get("is_inbox"))]
        my = self.get_my_addresses()

        total_items = sum(f["count"] for f in targets)
        logger.info("Sẽ quét %d thư mục (%d thư) trên %d kho",
                    len(targets), total_items, len({f["store"] for f in targets}))
        if progress:
            progress("Bắt đầu quét", 0, total_items, "")

        out: List[Dict] = []
        done = 0
        for meta in targets:
            if progress:
                progress("Đang quét", done, total_items, meta["path"])
            out.extend(self._scan_folder(meta, my))
            done += meta["count"]
        if progress:
            progress("Quét xong", total_items, total_items, "")
        logger.info("Thu được %d thư do bạn gửi", len(out))
        return out

    def _scan_folder(self, meta: Dict, my_addresses: set) -> List[Dict]:
        """Đọc một thư mục. Thư mục Sent lấy tất cả; thư mục khác lọc theo người gửi."""
        out: List[Dict] = []
        try:
            items = meta["folder"].Items
            total = int(items.Count)
        except Exception as e:
            logger.warning("Không đọc được %s: %s", meta["path"], e)
            return out

        for i in range(1, total + 1):
            try:
                msg = items.Item(i)
                if getattr(msg, "Class", 0) != OL_MAIL_ITEM:
                    continue
                if not meta["is_sent"]:
                    sender = (self._resolve_sender_email(msg) or "").strip().lower()
                    if not sender or (my_addresses and sender not in my_addresses):
                        continue
                own, quoted = split_quoted_text(getattr(msg, "Body", "") or "")
                if not own and not quoted:
                    continue
                out.append({
                    "subject": getattr(msg, "Subject", "") or "",
                    "body": own,                 # KHÔNG cắt ngắn
                    "quoted": quoted,            # chính là thư đến
                    "sent_time": _to_iso(getattr(msg, "SentOn", None)),
                    "to": getattr(msg, "To", "") or "",
                    "entry_id": getattr(msg, "EntryID", "") or "",
                    "source": meta["path"],
                })
            except Exception:
                continue
        return out

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
