"""Xuất toàn bộ thư đã gửi ra Markdown để phân tích văn phong bằng công cụ AI ngoài.

Vì sao cần công cụ này: Antigravity / Claude Desktop chạy bằng gói thuê bao riêng, KHÔNG
đụng tới GEMINI_API_KEY, nên khâu phân tích (tốn kém nhất) thành miễn phí. Nhưng chúng
không đọc thẳng được Outlook — chúng đọc tệp trên đĩa. Công cụ này bắc cầu, và chỉ dùng
COM nên KHÔNG gọi AI lần nào.

Mỗi thư xuất ra dưới dạng CẶP "thư đến → cách bạn trả lời". Thư đến lấy từ chính phần
trích dẫn nằm sẵn trong thư bạn gửi đi, nên không tốn thêm lần đọc Outlook nào.
"""
from __future__ import annotations

import logging
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List

from backend.com_worker import com_call
from backend.outlook_client import OutlookClient
from backend.style_stats import (FALLBACK_TOPIC, TOPIC_RULES, analyze_corpus,
                                 dedupe_emails, group_by_topic, strip_accents)
from config import BASE_DIR

OUT_DIR = BASE_DIR / "xuat_thu"
SHORT_BODY_CHARS = 40          # dưới ngưỡng này coi là "thư trả lời ngắn"
MAX_PER_FILE = 80              # nhóm lớn hơn thì tách phần, để AI đọc hết trong một lượt

FALLBACK_NOTE = ("Thư không khớp nhóm mục đích nào. Vẫn đáng đọc vì cho thấy người "
                 "dùng xoay xở thế nào với việc không thường gặp.")

# Phần trích dẫn chỉ để biết ĐANG TRẢ LỜI CÁI GÌ. Chuỗi chuyển tiếp lồng nhau ngoài
# thực tế có thể lên tới 17.000 dòng cho MỘT thư — để nguyên thì tệp phình lên nửa MB
# và công cụ AI không đọc nổi. Phần người dùng tự viết thì KHÔNG cắt.
MAX_QUOTED_CHARS = 3000

logger = logging.getLogger(__name__)


def slugify(text: str, max_len: int = 30) -> str:
    """Cắt ở ranh giới từ — cắt giữa từ sinh ra tên kiểu 'ho-tro-noi' khó đọc."""
    s = strip_accents(text or "").lower()
    s = re.sub(r"[^a-z0-9]+", "-", s).strip("-")
    if len(s) > max_len:
        s = s[:max_len].rsplit("-", 1)[0]
    return s.strip("-") or "khac"


def split_chunks(items: List[Dict], size: int) -> List[List[Dict]]:
    """Chia đều thành các phần xấp xỉ bằng nhau, không để phần đuôi lẻ vài thư."""
    parts = max(1, -(-len(items) // size))
    step = -(-len(items) // parts)
    return [items[i:i + step] for i in range(0, len(items), step)]


def fmt_email(index: int, item: Dict) -> str:
    sent = (item.get("sent_time") or "")[:10]
    to = (item.get("to") or "").strip()
    parts = [
        f"### Thư {index} — {sent or 'không rõ ngày'}",
        f"**Chủ đề:** {item.get('subject') or '(không có)'}",
        f"**Gửi tới:** {to or '(không rõ)'}",
        f"**Nguồn:** {item.get('source') or '(không rõ)'}",
        "",
    ]
    quoted = (item.get("quoted") or "").strip()
    if quoted:
        cut = len(quoted) > MAX_QUOTED_CHARS
        if cut:
            quoted = quoted[:MAX_QUOTED_CHARS].rstrip()
        parts.append("#### Thư đến")
        parts.extend("> " + ln for ln in quoted.split("\n"))
        if cut:
            parts.append("> …[đã cắt bớt phần trích dẫn lồng nhau phía sau]…")
        parts.append("")
    parts.append("#### Bạn đã trả lời")
    parts.append(item.get("body") or "(trống)")
    parts.append("")
    return "\n".join(parts)


def write_group(path: Path, title: str, items: List[Dict], note: str = "") -> None:
    lines = [f"# {title}", "", f"Tổng: **{len(items)} thư**.", ""]
    if note:
        lines += [note, ""]
    lines.append("---")
    lines.append("")
    for i, item in enumerate(items, 1):
        lines.append(fmt_email(i, item))
        lines.append("---")
        lines.append("")
    path.write_text("\n".join(lines), encoding="utf-8")


def fmt_stats(stats: Dict) -> str:
    """Bảng số liệu đã đếm chính xác — để AI khỏi phải đếm bằng mắt."""
    if not stats.get("analyzed_count"):
        return "_Không đủ dữ liệu để thống kê._"

    def table(rows, header):
        if not rows:
            return "_(không có)_\n"
        out = [f"| {header} | Số lần | Tỷ lệ |", "|---|---:|---:|"]
        for r in rows:
            text = r["text"].replace("|", "\\|")
            out.append(f"| `{text}` | {r['count']} | {r['ratio']:.0%} |")
        return "\n".join(out) + "\n"

    length = stats.get("length", {})
    return "\n".join([
        f"Đã thống kê trên **{stats['analyzed_count']} thư** (đếm bằng Python, chính xác tuyệt đối).",
        "",
        "### Câu chào thường dùng",
        table(stats.get("greeting_patterns"), "Câu chào"),
        "### Câu kết thường dùng",
        table(stats.get("closing_patterns"), "Câu kết"),
        "### Cụm từ đặc trưng",
        table(stats.get("common_phrases"), "Cụm từ"),
        "### Chữ ký",
        f"Độ tin cậy: **{stats.get('signature_confidence', 0):.0%}** số thư dùng khối này.",
        "",
        "```",
        stats.get("signature") or "(không phát hiện)",
        "```",
        "",
        "### Độ dài thư (trung vị)",
        f"- {length.get('median_words', 0)} từ mỗi thư",
        f"- {length.get('median_sentences', 0)} câu mỗi thư",
        f"- {length.get('words_per_sentence', 0)} từ mỗi câu",
        "",
        f"Ngôn ngữ chính: **{stats.get('language', 'không rõ')}**",
    ])


DE_BAI = """\
## Đề bài cho công cụ AI (Antigravity / Claude Desktop)

Bạn đang đọc kho thư đã gửi của một người dùng. Nhiệm vụ: rút ra **hướng dẫn viết thư
theo từng loại thư**, để một trợ lý AI khác có thể soạn thư thay người này.

**Các bước:**

1. Đọc lần lượt các tệp nhóm trong thư mục này (`01_*.md`, `02_*.md`, …). Thư **đã được
   gom sẵn theo mục đích** (rà soát, xử lý vi phạm, báo cáo, thẩm định mặt bằng…) — xem
   mục lục ở cuối tệp này. Mỗi thư có cặp "Thư đến" và "Bạn đã trả lời"; hãy chú ý
   **cách người này chuyển từ yêu cầu sang phản hồi**, không chỉ chú ý câu chữ.
2. Lấy các nhóm đó làm điểm xuất phát cho bộ **loại thư**. Được phép tách một nhóm thành
   nhiều loại nếu đọc thấy hai lối viết khác hẳn nhau, hoặc gộp hai nhóm nếu cách viết
   giống nhau. Đặt tên tệp ngắn, không dấu.
3. Với **mỗi loại**, tạo một tệp `kien_thuc/loai_thu/<ten-loai>.md` theo đúng khuôn:

```markdown
# <Tên loại thư đầy đủ>

**Từ khoá nhận diện:** <các từ/cụm từ xuất hiện trong chủ đề hoặc đầu thư, cách nhau bởi dấu phẩy>

## Cấu trúc thư
1. <câu mở đầu thường làm gì>
2. <thân thư trình bày theo trình tự nào>
3. <kết thư yêu cầu gì>

## Cách lập luận
- <người này viện dẫn quy định, số liệu, mốc thời gian ra sao>
- <khi nào cứng rắn, khi nào mềm mỏng>

## Câu mẫu
- <trích nguyên văn 3-5 câu tiêu biểu từ thư thật>

## Tránh
- <những gì người này không bao giờ làm trong loại thư này>
```

**Yêu cầu quan trọng:**

- Dòng `**Từ khoá nhận diện:**` được chương trình đọc để tự chọn đúng hướng dẫn khi soạn
  thư. Giữ đúng định dạng đó, đừng đổi thành chú thích HTML hay YAML.
- Phần "Câu mẫu" phải **trích nguyên văn** từ thư thật, không viết lại.
- Dựa vào bảng số liệu bên dưới cho các con số (câu chào, tần suất, độ dài) — đó là kết
  quả đếm chính xác, **đừng tự đếm lại bằng mắt**.
- Nếu một loại thư có quá ít mẫu (dưới 5 thư), đừng tạo tệp riêng — gộp vào `khac.md`.
- Viết bằng tiếng Việt.
"""


def build_overview(stats: Dict, groups: List, short_count: int, total_raw: int,
                   total_unique: int, folders: int, stores: int) -> str:
    rows = "\n".join(f"| [{name}]({name}) | {topic} | {count} |"
                     for name, topic, count in groups)
    return "\n".join([
        "# Kho thư đã gửi — dữ liệu để phân tích văn phong",
        "",
        f"_Xuất lúc {datetime.now():%d-%m-%Y %H:%M}_",
        "",
        f"- Quét **{stores} kho thư**, **{folders} thư mục**",
        f"- Thu được **{total_raw} thư**, còn **{total_unique} thư** sau khi lọc trùng",
        f"- Trong đó **{short_count} thư trả lời ngắn** để riêng ở `99_thu_ngan.md`",
        "",
        "> Thư mục này chứa thư công việc thật. Không chia sẻ ra ngoài, không đưa lên Git.",
        "",
        "---",
        "",
        DE_BAI,
        "",
        "---",
        "",
        "## Số liệu đã đếm",
        "",
        fmt_stats(stats),
        "",
        "---",
        "",
        "## Mục lục các nhóm",
        "",
        f"Thư được gom theo **mục đích viết** ({len(TOPIC_RULES)} nhóm + nhóm còn lại), "
        "dựa trên từ khoá trong chủ đề. Muốn đổi cách gom: sửa bảng `TOPIC_RULES` "
        "trong `backend/style_stats.py` rồi chạy lại `XUAT_THU.bat`.",
        "",
        "| Tệp | Nhóm mục đích | Số thư |",
        "|---|---|---:|",
        rows,
        "",
    ])


def export(deep: bool = True, progress=None) -> Dict:
    client = OutlookClient()
    if not com_call(client.connect):
        raise RuntimeError("Không kết nối được Outlook. Hãy mở Outlook rồi thử lại.")

    folders = com_call(client.iter_mail_folders)
    targets = [f for f in folders
               if f["is_sent"] or (deep and not f.get("is_inbox"))]
    stores = len({f["store"] for f in targets})
    print(f"Tìm thấy {len(folders)} thư mục thư trên {len({f['store'] for f in folders})} kho.")
    print(f"Sẽ quét {len(targets)} thư mục, tổng {sum(f['count'] for f in targets)} mục.")
    for f in sorted(targets, key=lambda x: -x["count"])[:15]:
        print(f"   {f['count']:>6}  {f['path']}{'  [Sent]' if f['is_sent'] else ''}")
    print("Bắt đầu quét — có thể mất nhiều phút. Dừng bằng Ctrl+C.\n")

    t0 = time.time()
    emails = com_call(client.collect_sent_emails, progress, deep)
    total_raw = len(emails)
    emails = dedupe_emails(emails)
    print(f"\nThu được {total_raw} thư, còn {len(emails)} sau khi lọc trùng "
          f"({time.time() - t0:.0f}s).")
    if not emails:
        raise RuntimeError("Không tìm thấy thư nào do bạn gửi.")

    short, normal = [], []
    for e in emails:
        (short if len((e.get("body") or "").strip()) < SHORT_BODY_CHARS
         else normal).append(e)

    OUT_DIR.mkdir(exist_ok=True)
    for old in OUT_DIR.glob("*.md"):        # xoá kết quả lần xuất trước
        old.unlink()

    stats = analyze_corpus(normal)

    # Thứ tự do group_by_topic quyết định (theo TOPIC_RULES, "Chủ đề khác" cuối cùng),
    # nên đọc 01 → 13 là đi theo một mạch nghiệp vụ chứ không nhảy cóc theo số lượng.
    index = []
    seq = 1
    for key, items in group_by_topic(normal).items():
        note = FALLBACK_NOTE if key == FALLBACK_TOPIC else ""
        chunks = split_chunks(items, MAX_PER_FILE)
        for part, chunk in enumerate(chunks, 1):
            suffix = f"_phan{part}" if len(chunks) > 1 else ""
            name = f"{seq:02d}_{slugify(key)}{suffix}_{len(chunk)}thu.md"
            title = f"Nhóm: {key}" + (f" (phần {part}/{len(chunks)})" if suffix else "")
            write_group(OUT_DIR / name, title, chunk, note)
            index.append((name, key, len(chunk)))
            seq += 1

    if short:
        write_group(OUT_DIR / "99_thu_ngan.md", "Thư trả lời ngắn", short,
                    "Các thư người dùng trả lời rất ngắn. Chúng cho biết **khi nào** "
                    "người này chọn trả lời cụt và dùng cách diễn đạt nào.")
        index.append(("99_thu_ngan.md", "Thư trả lời ngắn", len(short)))

    (OUT_DIR / "00_TONG_QUAN.md").write_text(
        build_overview(stats, index, len(short), total_raw, len(emails),
                       len(targets), stores),
        encoding="utf-8")

    # Ghi hồ sơ văn phong NGAY TẠI ĐÂY để chạy XUAT_THU.bat trực tiếp cũng cập nhật
    # được — hồ sơ này được nạp vào prompt mỗi lần soạn thư. Toàn bộ số liệu do Python
    # đếm, không có lượt gọi AI nào.
    save_style_profile(stats, total_raw, len(emails), len(short))

    return {"total_raw": total_raw, "unique": len(emails), "short": len(short),
            "files": len(index) + 1, "stats": stats, "dir": str(OUT_DIR)}


def save_style_profile(stats: Dict, total_raw: int, unique: int, short: int) -> bool:
    """Lưu thống kê thành hồ sơ văn phong schema v3."""
    if not stats.get("analyzed_count"):
        return False
    from backend.style_analyzer import StyleAnalyzer
    from config import STYLE_PROFILES_DIR

    profile = dict(stats)
    profile["schema_version"] = 3
    profile["stats"] = {
        "source_count": total_raw,
        "unique_count": unique,
        "short_count": short,
        "analyzed_count": stats["analyzed_count"],
        "llm_calls": 0,
    }
    StyleAnalyzer(data_dir=STYLE_PROFILES_DIR).save_profile(profile, "default")
    return True


def _print_progress(phase: str, done: int, total: int, detail: str = "") -> None:
    pct = f"{done * 100 // total:>3}%" if total else "  ?"
    sys.stdout.write(f"\r  [{pct}] {phase}: {detail[:60]:<60}")
    sys.stdout.flush()


def main() -> int:
    logging.basicConfig(level=logging.WARNING,
                        format="%(asctime)s [%(levelname)s] %(message)s")
    deep = "--nhanh" not in sys.argv

    print("=" * 70)
    print("  XUAT THU DA GUI DE PHAN TICH VAN PHONG")
    print("=" * 70)
    print("Che do:", "quet MOI thu muc (ke ca kho luu tru)" if deep
          else "chi quet thu muc kieu Sent (--nhanh)")
    print()

    try:
        res = export(deep=deep, progress=_print_progress)
    except KeyboardInterrupt:
        print("\n\nDa dung theo yeu cau. Chua ghi tep nao.")
        return 1
    except Exception as e:
        print(f"\n[LOI] {e}")
        return 1

    print()
    print("=" * 70)
    print(f"  XONG — {res['files']} tep trong {res['dir']}")
    print("=" * 70)
    print(f"  Tong thu      : {res['total_raw']} (con {res['unique']} sau loc trung)")
    print(f"  Thu ngan      : {res['short']}")
    print(f"  Da thong ke   : {res['stats'].get('analyzed_count', 0)} thu")
    print(f"  Ho so van phong da cap nhat (0 luot goi AI)")
    print()
    print("  CANH BAO: thu muc nay chua thu cong viec THAT.")
    print("            Da nam trong .gitignore. Khong chia se ra ngoai.")
    print()
    print("  BUOC TIEP THEO:")
    print("    1. Mo thu muc xuat_thu/ bang Antigravity hoac Claude Desktop")
    print("    2. Ra lenh: \"doc 00_TONG_QUAN.md roi lam theo de bai\"")
    print("    3. Cong cu se ghi huong dan vao kien_thuc/loai_thu/")
    return 0


if __name__ == "__main__":
    sys.exit(main())
