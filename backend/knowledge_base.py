"""Kho tri thức offline: quy tắc viết tay (Markdown) + dữ liệu máy học (JSON).

Thư mục `kien_thuc/` chứa Markdown thuần do người dùng biên soạn, có thể sửa bằng
bất kỳ công cụ AI nào bên ngoài (Claude Desktop, ChatGPT Desktop, Antigravity).
Tệp được nạp lại theo mtime ở mỗi lần dựng prompt, nên sửa xong là thư kế tiếp
đã tuân theo — không cần khởi động lại backend.
"""
from __future__ import annotations

import json
import logging
import re
from pathlib import Path
from typing import Dict, List, Optional

from .style_stats import strip_accents
from config import KNOWLEDGE_DIR, STYLE_PROFILES_DIR

logger = logging.getLogger(__name__)

# Thứ tự tệp = thứ tự nạp = thứ tự ưu tiên trong prompt.
# Khi phải cắt bớt vì quá dài, cắt từ cuối danh sách ngược lên.
KNOWLEDGE_FILES = (
    "00_HUONG_DAN.md",
    "01_ho_so_ca_nhan.md",
    "02_quy_tac_tra_loi.md",
    "03_doi_tac_va_xung_ho.md",
    "04_thuat_ngu_va_viet_tat.md",
    "05_mau_cau_thuong_dung.md",
    "06_khong_duoc_lam.md",
)

# Tệp xưng hô phình to theo số đối tác nên chỉ nạp mục liên quan người gửi.
PER_SENDER_FILE = "03_doi_tac_va_xung_ho.md"
ALWAYS_SECTION = "chung"

# Hướng dẫn viết riêng cho từng loại thư, do công cụ AI ngoài (Antigravity/Claude
# Desktop) sinh ra sau khi đọc kho thư xuất bởi XUAT_THU.bat.
TYPE_DIR_NAME = "loai_thu"
FALLBACK_TYPE = "khac"

# Dòng khai từ khoá ở đầu mỗi tệp loại thư. Cố ý dùng văn bản thường thay vì chú
# thích HTML hay YAML: công cụ AI sửa tệp mà không làm hỏng, người đọc cũng thấy được.
_KEYWORD_RE = re.compile(r"^\s*\*\*Từ khoá nhận diện:\*\*\s*(.+)$", re.M | re.I)

# Tiêu đề "# ..." đầu tệp dùng làm nhãn hiển thị trong ô chọn mẫu thư.
_TITLE_RE = re.compile(r"^#\s+(.+?)\s*$", re.M)

DEFAULT_MAX_CHARS = 12_000

# Các khoá của hồ sơ văn phong (xem backend/style_stats.py::analyze_corpus).
_STYLE_KEYS = ("greeting_patterns", "closing_patterns", "formality_level",
               "common_phrases", "signature", "signature_confidence",
               "length", "tone_notes", "language")
_STATS_KEYS = ("source_count", "unique_count", "group_count",
               "sample_count", "batch_count", "analyzed_count", "short_count")

_HEADING_RE = re.compile(r"^##\s+(.+?)\s*$", re.M)


def _split_sections(text: str) -> List[tuple]:
    """Tách Markdown thành [(tiêu đề mục, nội dung cả mục)] theo heading '## '.

    Phần đứng trước heading đầu tiên được gán tiêu đề rỗng và luôn được giữ.
    """
    matches = list(_HEADING_RE.finditer(text))
    if not matches:
        return [("", text)]
    sections = []
    if matches[0].start() > 0:
        preamble = text[:matches[0].start()].strip()
        if preamble:
            sections.append(("", preamble))
    for i, m in enumerate(matches):
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        sections.append((m.group(1).strip(), text[m.start():end].strip()))
    return sections


def _filter_by_sender(text: str, sender_email: str) -> str:
    """Giữ mục '## Chung' và các mục có tiêu đề khớp email người gửi."""
    sender = (sender_email or "").strip().lower()
    kept = []
    for heading, block in _split_sections(text):
        key = heading.strip().lower()
        if not key or key == ALWAYS_SECTION:
            kept.append(block)
        elif sender and key.lstrip("@") and key.lstrip("@") in sender:
            kept.append(block)
    return "\n\n".join(kept).strip()


def normalize_learned(raw: Optional[Dict]) -> Optional[Dict]:
    """Nâng hồ sơ văn phong phẳng (v1) lên dạng chuẩn hoá v2 trong bộ nhớ.

    KHÔNG ghi ngược ra đĩa: StyleAnalyzer vẫn giữ định dạng cũ để desktop app
    và dữ liệu đã học sẵn không bị ảnh hưởng.
    """
    if not raw:
        return None
    version = raw.get("schema_version")
    stats = dict(raw.get("stats") or {})
    # v1 lưu các chỉ số ở cấp cao nhất; v3 đã có sẵn khối "stats".
    for key in _STATS_KEYS:
        if key not in stats and raw.get(key) is not None:
            stats[key] = raw[key]
    return {
        "schema_version": version if version and version >= 2 else 2,
        "name": raw.get("name", "default"),
        "analyzed_at": raw.get("analyzed_at"),
        "stats": stats,
        "style": {k: raw.get(k) for k in _STYLE_KEYS if raw.get(k) is not None},
    }


class KnowledgeBase:
    def __init__(self, md_dir: Optional[Path] = None, learned_path: Optional[Path] = None,
                 max_chars: int = DEFAULT_MAX_CHARS) -> None:
        self.md_dir = Path(md_dir) if md_dir is not None else KNOWLEDGE_DIR
        self.learned_path = Path(learned_path) if learned_path is not None else (STYLE_PROFILES_DIR / "current_profile.json")
        self.max_chars = max_chars
        self._docs: Dict[str, str] = {}
        self._types: Dict[str, Dict] = {}      # {tên loại: {text, keywords}}
        self._learned: Optional[Dict] = None
        self._mtimes: Dict[str, float] = {}
        self._truncated = False
        if not self.md_dir.exists():
            logger.warning(
                "Chưa có thư mục %s. Sao chép kien_thuc_mau/ thành kien_thuc/ "
                "rồi sửa nội dung cho đúng thông tin của bạn.", self.md_dir)
        self.reload(force=True)

    # ------------------------------------------------------------------ đĩa

    def _current_mtimes(self) -> Dict[str, float]:
        stamps = {}
        for name in KNOWLEDGE_FILES:
            path = self.md_dir / name
            try:
                stamps[name] = path.stat().st_mtime
            except OSError:
                pass
        try:
            stamps["__learned__"] = self.learned_path.stat().st_mtime
        except OSError:
            pass
        type_dir = self.md_dir / TYPE_DIR_NAME
        if type_dir.is_dir():
            for path in sorted(type_dir.glob("*.md")):
                try:
                    stamps[f"loai:{path.name}"] = path.stat().st_mtime
                except OSError:
                    pass
        return stamps

    def reload(self, force: bool = False) -> bool:
        """Nạp lại nếu mtime đổi. Trả True khi thực sự đọc lại đĩa."""
        stamps = self._current_mtimes()
        if not force and stamps == self._mtimes:
            return False
        self._mtimes = stamps

        self._docs = {}
        for name in KNOWLEDGE_FILES:
            path = self.md_dir / name
            if not path.is_file():
                continue
            try:
                content = path.read_text(encoding="utf-8").strip()
            except (OSError, UnicodeError) as e:
                logger.warning("Không đọc được %s: %s", path, e)
                continue
            if content:
                self._docs[name] = content

        self._types = {}
        type_dir = self.md_dir / TYPE_DIR_NAME
        if type_dir.is_dir():
            for path in sorted(type_dir.glob("*.md")):
                # Tệp có tiền tố số là tài liệu hướng dẫn, không phải một loại thư.
                # Bỏ qua chúng, nếu không dòng ví dụ nằm trong code fence sẽ bị bắt
                # nhầm thành từ khoá thật.
                if path.stem[:1].isdigit():
                    continue
                try:
                    content = path.read_text(encoding="utf-8").strip()
                except (OSError, UnicodeError) as e:
                    logger.warning("Không đọc được %s: %s", path, e)
                    continue
                if not content:
                    continue
                m = _KEYWORD_RE.search(content)
                keywords = []
                if m:
                    keywords = [strip_accents(k).strip().lower()
                                for k in m.group(1).split(",") if k.strip()]
                title = _TITLE_RE.search(content)
                self._types[path.stem.lower()] = {
                    "text": content, "keywords": keywords, "file": path.name,
                    "title": title.group(1).strip() if title else path.stem,
                }

        raw = None
        if self.learned_path.is_file():
            try:
                raw = json.loads(self.learned_path.read_text(encoding="utf-8"))
            except (OSError, UnicodeError, json.JSONDecodeError) as e:
                logger.warning("Không đọc được hồ sơ văn phong %s: %s",
                               self.learned_path, e)
        self._learned = normalize_learned(raw)
        logger.info("Kho tri thức: %d tệp quy tắc, hồ sơ văn phong: %s",
                    len(self._docs), "có" if self._learned else "chưa có")
        return True

    # ------------------------------------------------------------------ đọc

    def learned(self) -> Optional[Dict]:
        """Hồ sơ máy học đã chuẩn hoá sang schema v2."""
        return self._learned

    def style_profile(self) -> Optional[Dict]:
        """Dạng phẳng, tương thích AIEngine._format_style sẵn có."""
        if not self._learned:
            return None
        flat = dict(self._learned.get("style", {}))
        flat["analyzed_at"] = self._learned.get("analyzed_at")
        return flat

    def rules_markdown(self, sender_email: str = "") -> str:
        """Ghép Markdown quy tắc, đã lọc mục theo người gửi và cắt bớt nếu dài."""
        blocks = []
        for name in KNOWLEDGE_FILES:
            content = self._docs.get(name)
            if not content:
                continue
            if name == PER_SENDER_FILE:
                content = _filter_by_sender(content, sender_email)
                if not content:
                    continue
            blocks.append((name, f"### Từ tệp {name}\n{content}"))

        total = sum(len(b) for _, b in blocks)
        self._truncated = False
        # Cắt nguyên tệp từ ưu tiên thấp nhất ngược lên, không cắt giữa chừng.
        while blocks and total > self.max_chars:
            _, dropped = blocks.pop()
            total -= len(dropped)
            self._truncated = True

        text = "\n\n".join(b for _, b in blocks)
        if self._truncated:
            text += "\n\n[… đã cắt bớt một số tệp quy tắc do vượt giới hạn độ dài …]"
        return text

    # ------------------------------------------------------- loại thư

    def list_types(self) -> List[Dict]:
        """Danh sách mẫu thư cho ô chọn ở giao diện. Loại dự phòng xếp cuối."""
        self.reload()
        items = [{"name": n, "title": m.get("title") or n,
                  "keywords": len(m["keywords"])}
                 for n, m in sorted(self._types.items())]
        items.sort(key=lambda t: (t["name"] == FALLBACK_TYPE, t["title"].lower()))
        return items

    def match_email_type(self, subject: str = "", body: str = "",
                         forced: str = "") -> Optional[str]:
        """Chọn loại thư khớp nhất theo từ khoá. So khớp chuỗi cục bộ — KHÔNG gọi AI.

        Bỏ dấu trước khi so để "thong ke mat bang" vẫn khớp "thống kê mặt bằng".
        Từ khoá dài được ưu tiên khi hai loại cùng số điểm, vì nó cụ thể hơn.

        `forced`: tên mẫu thư người dùng tự chọn ở giao diện. Chọn tên không tồn tại
        thì bỏ qua và tự khớp như thường, không báo lỗi — tệp loại thư có thể vừa bị
        xoá hay đổi tên trong lúc giao diện còn giữ danh sách cũ.
        """
        if not self._types:
            return None
        forced = (forced or "").strip().lower()
        if forced and forced in self._types:
            return forced
        haystack = strip_accents(f"{subject}\n{(body or '')[:1000]}").lower()

        best, best_score, best_len = None, 0, 0
        for name, meta in self._types.items():
            if name == FALLBACK_TYPE:
                continue
            hits = [k for k in meta["keywords"] if k and k in haystack]
            if not hits:
                continue
            longest = max(len(k) for k in hits)
            if len(hits) > best_score or (len(hits) == best_score and longest > best_len):
                best, best_score, best_len = name, len(hits), longest

        if best:
            return best
        return FALLBACK_TYPE if FALLBACK_TYPE in self._types else None

    def type_markdown(self, subject: str = "", body: str = "",
                      forced: str = "") -> tuple:
        """Trả (tên loại, khối hướng dẫn) — chuỗi rỗng nếu không có tệp loại nào."""
        name = self.match_email_type(subject, body, forced)
        if not name:
            return "", ""
        meta = self._types.get(name) or {}
        return name, meta.get("text", "")

    def type_title(self, name: str) -> str:
        """Nhãn hiển thị của một mẫu thư, để giao diện cho biết đã dùng mẫu nào."""
        return (self._types.get((name or "").lower()) or {}).get("title", "") or name

    def build_prompt_block(self, sender_email: str = "", subject: str = "",
                           body: str = "", email_type: str = "") -> str:
        """Khối kiến thức chèn vào prompt. Tự nạp lại khi tệp trên đĩa đổi."""
        self.reload()
        blocks = [self.rules_markdown(sender_email)]

        type_name, type_text = self.type_markdown(subject, body, email_type)
        if type_text:
            # Nhãn rõ ràng để khi thư ra sai kiểu còn biết nó đã chọn nhầm loại nào.
            blocks.append(
                f"### Hướng dẫn riêng cho loại thư này ({type_name})\n"
                f"Thư đang trả lời được nhận diện thuộc loại này. Tuân thủ nghiêm "
                f"cấu trúc và cách lập luận dưới đây.\n\n{type_text}")
        return "\n\n".join(b for b in blocks if b)

    def status(self) -> Dict:
        self.reload()
        files = []
        for name in KNOWLEDGE_FILES:
            content = self._docs.get(name)
            files.append({
                "name": name,
                "exists": content is not None,
                "chars": len(content) if content else 0,
            })
        learned = self._learned or {}
        stats = learned.get("stats", {})
        return {
            "dir": str(self.md_dir),
            "dir_exists": self.md_dir.exists(),
            "files": files,
            "file_count": len(self._docs),
            "total_chars": sum(len(c) for c in self._docs.values()),
            "truncated": self._truncated,
            "types": [{"name": n, "keywords": len(m["keywords"]), "chars": len(m["text"])}
                      for n, m in sorted(self._types.items())],
            "type_count": len(self._types),
            "has_learned": bool(self._learned),
            "learned_at": learned.get("analyzed_at"),
            # v3 dùng analyzed_count (số thư thực sự thống kê); v1/v2 dùng sample_count.
            "sample_count": stats.get("analyzed_count") or stats.get("sample_count", 0),
        }
