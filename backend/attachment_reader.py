"""Đọc tệp đính kèm do Office.js gửi lên (base64) để đưa vào prompt Gemini.

Hai đường xử lý:
  - PDF và ảnh  -> giữ nguyên bytes, gửi thẳng cho Gemini dạng inline part.
                   Gemini tự đọc được cả PDF scan nên không cần OCR cục bộ.
  - .xlsx/.docx -> trích thành văn bản tại chỗ vì Gemini không đọc được định dạng này.

Mọi ngưỡng đều được kiểm lại ở đây, KHÔNG tin ngưỡng phía client.
Tệp hỏng luôn trở thành một dòng ghi chú, không bao giờ ném exception làm chết request.
"""
from __future__ import annotations

import base64
import binascii
import io
import logging
import zipfile
from dataclasses import dataclass, field
from typing import Dict, List, Optional
from xml.etree import ElementTree as ET

logger = logging.getLogger(__name__)

MAX_FILE_BYTES = 8 * 1024 * 1024
MAX_TOTAL_BLOB_BYTES = 12 * 1024 * 1024
MAX_FILES = 5
MAX_TEXT_CHARS_FILE = 6_000
MAX_TEXT_CHARS_TOTAL = 20_000

MAX_SHEETS, MAX_ROWS, MAX_COLS = 10, 200, 30

# MIME suy từ phần mở rộng, KHÔNG dùng contentType của Exchange —
# Exchange rất hay khai application/octet-stream và Gemini từ chối MIME đó.
EXT_MIME = {
    "pdf": "application/pdf",
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "gif": "image/gif",
    "webp": "image/webp",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
}
INLINE_MIMES = {"application/pdf", "image/png", "image/jpeg",
                "image/gif", "image/webp"}

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


@dataclass
class ReadResult:
    blobs: List[Dict] = field(default_factory=list)   # [{mime, data: bytes, name}]
    text: str = ""                                     # văn bản đã trích, có tiêu đề tệp
    notes: List[str] = field(default_factory=list)     # thông báo tiếng Việt cho UI
    used: List[str] = field(default_factory=list)      # tên tệp thực sự đã dùng


def _ext_of(name: str) -> str:
    return (name or "").rsplit(".", 1)[-1].lower() if "." in (name or "") else ""


def _decode(data_b64: str) -> Optional[bytes]:
    try:
        return base64.b64decode(data_b64 or "", validate=True)
    except (binascii.Error, ValueError):
        return None


def _read_xlsx(data: bytes) -> tuple:
    """Trả (văn bản TSV, ghi chú). Cần openpyxl."""
    try:
        import openpyxl
    except ImportError:
        return "", "Chưa cài openpyxl nên không đọc được tệp .xlsx."

    wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    parts, empty_cells, total_cells = [], 0, 0
    try:
        for ws in wb.worksheets[:MAX_SHEETS]:
            rows = []
            for row in ws.iter_rows(max_row=MAX_ROWS, max_col=MAX_COLS,
                                    values_only=True):
                cells = ["" if v is None else str(v) for v in row]
                total_cells += len(cells)
                empty_cells += sum(1 for c in cells if not c)
                while cells and not cells[-1]:
                    cells.pop()
                if cells:
                    rows.append("\t".join(cells))
            if rows:
                parts.append(f"#### Sheet: {ws.title}\n" + "\n".join(rows))
    finally:
        wb.close()

    note = ""
    # data_only=True chỉ trả kết quả công thức ĐÃ được Excel lưu cache. Tệp sinh
    # tự động mà chưa từng mở bằng Excel sẽ cho ra toàn ô rỗng.
    if total_cells and empty_cells / total_cells > 0.7:
        note = ("Bảng tính có rất nhiều ô rỗng — nếu tệp chứa công thức chưa từng "
                "được mở bằng Excel thì giá trị không đọc được.")
    return "\n\n".join(parts), note


def _read_docx(data: bytes) -> tuple:
    """Trích văn bản .docx bằng stdlib (zipfile + ElementTree).

    Cố ý KHÔNG dùng python-docx: nó kéo theo lxml (C extension), thứ dễ hỏng nhất
    khi cài trên máy công ty có proxy chặn wheel nhị phân.
    """
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    lines = []
    for para in root.iter(f"{_W}p"):
        txt = "".join(t.text or "" for t in para.iter(f"{_W}t")).strip()
        if txt:
            lines.append(txt)
    note = "" if lines else "Tệp .docx không có nội dung văn bản đọc được."
    return "\n".join(lines), note


def read_attachments(items: Optional[List[Dict]]) -> ReadResult:
    """items: [{name, content_type, size, data_b64}] từ task pane."""
    result = ReadResult()
    if not items:
        return result

    blob_bytes = 0
    text_chars = 0
    text_parts = []

    for item in items[:MAX_FILES]:
        name = (item.get("name") or "(không tên)").strip()
        ext = _ext_of(name)
        mime = EXT_MIME.get(ext)
        if not mime:
            result.notes.append(f'Bỏ qua "{name}": chưa hỗ trợ định dạng .{ext}.')
            continue

        data = _decode(item.get("data_b64", ""))
        if data is None:
            result.notes.append(f'Bỏ qua "{name}": dữ liệu base64 hỏng.')
            continue
        if len(data) > MAX_FILE_BYTES:
            result.notes.append(
                f'Bỏ qua "{name}": {len(data) / 1048576:.1f} MB, vượt giới hạn '
                f'{MAX_FILE_BYTES // 1048576} MB.')
            continue

        if mime in INLINE_MIMES:
            if blob_bytes + len(data) > MAX_TOTAL_BLOB_BYTES:
                result.notes.append(f'Bỏ qua "{name}": đã chạm giới hạn tổng dung lượng.')
                continue
            result.blobs.append({"mime": mime, "data": data, "name": name})
            blob_bytes += len(data)
            result.used.append(name)
            continue

        try:
            extracted, note = _read_xlsx(data) if ext == "xlsx" else _read_docx(data)
        except (zipfile.BadZipFile, ET.ParseError, KeyError) as e:
            result.notes.append(f'Bỏ qua "{name}": tệp hỏng hoặc sai định dạng ({e}).')
            continue
        except Exception as e:                       # thư viện lỗi bất ngờ
            logger.warning("Lỗi đọc đính kèm %s: %s", name, e)
            result.notes.append(f'Bỏ qua "{name}": không đọc được ({e}).')
            continue

        if note:
            result.notes.append(f'"{name}": {note}')
        if not extracted:
            continue
        extracted = extracted[:MAX_TEXT_CHARS_FILE]
        if text_chars + len(extracted) > MAX_TEXT_CHARS_TOTAL:
            extracted = extracted[:max(0, MAX_TEXT_CHARS_TOTAL - text_chars)]
            if not extracted:
                result.notes.append(f'Bỏ qua "{name}": đã chạm giới hạn tổng văn bản.')
                continue
            result.notes.append(f'"{name}": đã cắt bớt do tổng văn bản quá dài.')
        text_parts.append(f"### Tệp đính kèm: {name}\n{extracted}")
        text_chars += len(extracted)
        result.used.append(name)

    if len(items) > MAX_FILES:
        result.notes.append(
            f"Chỉ xử lý {MAX_FILES} tệp đầu tiên trong tổng số {len(items)} tệp đính kèm.")

    result.text = "\n\n".join(text_parts)
    return result
