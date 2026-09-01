"""Thống kê văn phong bằng Python thuần — không AI, không giới hạn số thư.

Vì sao không để AI làm việc này: "bạn hay chào bằng câu gì, bao nhiêu lần" là bài toán
ĐẾM. Python đếm chính xác tuyệt đối trên toàn bộ hộp thư trong tích tắc và miễn phí,
trong khi LLM chỉ đọc được vài thư một lượt rồi ước lượng — và khi hợp nhất nhiều lượt
nó còn trừu tượng hoá mất chi tiết ("Dear CH," thành "Dear [Name],").

Kết quả ở đây dùng làm SỐ LIỆU THAM CHIẾU đặt ở đầu tệp xuất, để công cụ AI bên ngoài
(Antigravity/Claude Desktop) có dữ kiện chính xác thay vì phải đếm bằng mắt.
"""
from __future__ import annotations

import re
import statistics
import unicodedata
from collections import Counter
from typing import Dict, List

# Ngưỡng: một dòng phải xuất hiện ở ít nhất ngần này tỷ lệ thư mới coi là chữ ký.
SIGNATURE_MIN_RATIO = 0.15
MAX_GREETING_LEN = 100
MAX_CLOSING_LEN = 80
MAX_SIGNATURE_LINES = 12
NGRAM_MIN, NGRAM_MAX = 2, 6
NGRAM_BODY_CHARS = 3000       # chặn chi phí n-gram trên thư quá dài
TOP_N = 12

CLOSING_HINTS = (
    "trân trọng", "thân ái", "thân mến", "trân trong", "mong nhận", "mong sớm",
    "cảm ơn", "cám ơn", "regards", "thanks", "thank you", "best", "rgds", "br,",
)

_WORD_RE = re.compile(r"[0-9A-Za-zÀ-ỹ]+", re.UNICODE)
_SENT_SPLIT_RE = re.compile(r"[.!?…]+|\n+")


def strip_accents(text: str) -> str:
    """Bỏ dấu tiếng Việt để so khớp không phụ thuộc cách gõ."""
    nfd = unicodedata.normalize("NFD", text or "")
    out = "".join(c for c in nfd if unicodedata.category(c) != "Mn")
    return out.replace("đ", "d").replace("Đ", "D")


def _lines(body: str) -> List[str]:
    return [ln.strip() for ln in (body or "").split("\n") if ln.strip()]


# ------------------------------------------------------- gom nhóm theo chủ đề

# Cắt tiền tố trả lời/chuyển tiếp để gom theo chủ đề GỐC.
# Lưu ý lịch sử: bản cũ viết r"^(re|fw|fwd):\\s*" — trong raw string "\\s" là dấu \
# cộng chữ s nên không khớp khoảng trắng, khiến mọi thư trả lời dồn vào một nhóm.
_SUBJ_PREFIX_RE = re.compile(
    r"^\s*(re|fw|fwd|trả lời|chuyển tiếp)\s*(\[\d+\])?\s*:\s*", re.I)


def strip_subject_prefix(subject: str) -> str:
    """Bóc hết chuỗi tiền tố lồng nhau kiểu "RE: FW: RE: Báo cáo"."""
    subject = (subject or "").strip()
    while (stripped := _SUBJ_PREFIX_RE.sub("", subject)) != subject:
        subject = stripped
    return subject


# Phân nhóm theo MỤC ĐÍCH THƯ, không theo từ đầu chủ đề.
#
# Bản cũ lấy từ đầu tiên của chủ đề làm khoá, sinh ra hàng chục nhóm vô nghĩa
# ("ra", "check", "kiem", "bao", "giai") và một phần tư số thư rơi vào rổ "nhóm nhỏ".
# Cùng một mục đích viết thư bị xé ra chỉ vì cách đặt chủ đề khác nhau — trong khi
# đó chính là thứ cần gom lại để học văn phong.
#
# Quy tắc: khớp TỪ TRÊN XUỐNG, nhóm đầu tiên trúng từ khoá là nhóm được chọn. Thứ tự
# vì thế là thứ tự ưu tiên — nhóm đặc thù xếp trước nhóm chung chung. Từ khoá viết
# không dấu, chữ thường; so khớp cũng bỏ dấu nên gõ có dấu hay không đều trúng.
#
# Muốn thêm/bớt nhóm: sửa thẳng bảng này, không phải sửa chỗ nào khác.
TOPIC_RULES = [
    # KHÔNG dùng riêng "tham dinh": người dùng còn "thẩm định sự vụ", "thẩm định
    # báo cáo" — chỉ mình từ đó sẽ kéo nhầm thư sang đây.
    ("Thẩm định mặt bằng & mở điểm mới",
     ["mat bang", "mbmm", "tham dinh thuc dia", "tham dinh an", "tham dinh gia",
      "mo moi", "dia diem moi", "mo diem"]),
    ("Xử lý vi phạm & kỷ luật",
     ["vi pham", "xlvp", "bbvp", "bien ban", "ky luat", "giu luong", "canh cao",
      "boi thuong", "boi hoan", "quy trach nhiem", "sai phep", "mo cua muon"]),
    ("Xác minh nghi vấn gian lận",
     ["gian lan", "nghi van", "nghi ngo", "xac minh", "sai pham", "tieu cuc",
      "that thoat", "trom cap", "vu viec", "su vu", "bat thuong", "ton ao",
      "chuyen giao ao"]),
    ("Yêu cầu rà soát & đối chiếu dữ liệu",
     ["ra soat", "check", "doi chieu", "soat xet", "loc du lieu"]),
    # Xếp trên nhóm kiểm kê: "Giải trình về việc KKFF" là thư GIẢI TRÌNH, kiểm kê
    # chỉ là sự việc được nhắc tới. Mục đích viết mới là thứ quyết định văn phong.
    ("Giải trình, phản hồi & khiếu nại",
     ["giai trinh", "phan hoi", "phan anh", "khieu nai", "tra loi", "y kien",
      "lam ro", "ticket"]),
    # "kiem soat ch" chứ không phải "kiem soat": để nguyên sẽ nuốt luôn
    # "Phương án phối hợp kiểm soát…" vốn thuộc nhóm kế hoạch.
    ("Kiểm tra, kiểm kê & giám sát cửa hàng",
     ["kiem ke", "tkk", "kkff", "kknv", "gskk", "kiem tra", "thanh kiem tra",
      "giam sat", "tkmb", "kiem soat tuan thu", "kiem soat ch"]),
    ("Yêu cầu cung cấp thông tin & hồ sơ",
     ["cung cap", "bo sung thong tin", "xac nhan thong tin", "hop dong",
      "danh sach cbnv", "gui lai", "chia se"]),
    ("Đề xuất & trình phê duyệt",
     ["de xuat", "phe duyet", "trinh pd", "xin y kien", "de nghi", "kien nghi",
      "xin duyet", "trinh ky"]),
    ("Báo cáo & tổng hợp kết quả",
     ["bao cao", "b/c", "bc:", "tong hop", "ket qua", "thong ke", "so ket",
      "tong ket", "cap nhat tinh hinh"]),
    ("Kế hoạch, triển khai & chuyên đề",
     ["ke hoach", "trien khai", "phuong an", "chuyen de", "huong dan", "quy trinh",
      "thong bao", "nhac viec", "qua han", "lich "]),
    ("Tồn kho, hàng hoá & chứng từ",
     ["ton kho", "ton cao", "nhap kho", "huy hang", "huy line", "huy dong",
      "tra hang", "no-claim", "noclaim", "no claim", "coupon", "phieu xuat",
      "hoa don", "hang hoa", "sto", "po ", "po kho", "ban hang", "doanh thu"]),
    ("Hệ thống, công cụ & hỗ trợ nội bộ",
     ["request id", "camera", "tool", "phan mem", "loi ", "tai khoan", "phan quyen", "onedrive",
      "he thong", "nhan su", "cham cong", "ovteleport", "dung luong", "cap quyen",
      "ho tro", "nho in", "winx", "sap "]),
]

FALLBACK_TOPIC = "Chủ đề khác"

# Chuẩn hoá sẵn một lần: (tên nhóm, danh sách từ khoá đã bỏ dấu).
_TOPIC_RULES_NORM = [(name, [strip_accents(k).lower() for k in keys])
                     for name, keys in TOPIC_RULES]

# Nhãn hệ thống trong ngoặc vuông ("[KSTT WCM]", "[Request ID :##RE-473581##]") xuất
# hiện ở rất nhiều thư và sẽ kéo mọi thứ về cùng một nhóm. Bỏ trước khi khớp — trừ
# khi bỏ xong không còn gì, lúc đó chính cái nhãn mới là nội dung.
_TAG_RE = re.compile(r"\[[^\]]{0,60}\]")

# Số lượng ký tự đầu thân thư dùng để khớp khi chủ đề không đủ manh mối.
TOPIC_BODY_CHARS = 400


def _match_topic(text: str) -> str:
    for name, keys in _TOPIC_RULES_NORM:
        if any(k in text for k in keys):
            return name
    return ""


def classify_topic(subject: str, body: str = "") -> str:
    """Nhóm mục đích của thư. Ưu tiên chủ đề, thiếu manh mối mới đọc đầu thân thư."""
    cleaned = strip_subject_prefix(subject)
    trimmed = _TAG_RE.sub(" ", cleaned).strip()
    norm = strip_accents(trimmed or cleaned).lower()
    return (_match_topic(norm)
            or _match_topic(strip_accents(body or "")[:TOPIC_BODY_CHARS].lower())
            or FALLBACK_TOPIC)


def dedupe_emails(emails: List[Dict]) -> List[Dict]:
    """Bỏ thư trùng nội dung — thư nằm cả ở Sent Items lẫn bản lưu trữ."""
    import hashlib
    seen, out = set(), []
    for item in emails:
        body = (item.get("body") or "").strip().lower()
        key = hashlib.sha1(body.encode("utf-8")).hexdigest()
        if key in seen:
            continue
        seen.add(key)
        out.append(item)
    return out


def group_by_topic(emails: List[Dict]) -> Dict[str, List[Dict]]:
    """Gom thư theo mục đích. Thứ tự trả về theo TOPIC_RULES, "Chủ đề khác" xếp cuối."""
    groups: Dict[str, List[Dict]] = {}
    for item in emails:
        key = classify_topic(item.get("subject", ""), item.get("body", ""))
        groups.setdefault(key, []).append(item)

    order = {name: i for i, (name, _) in enumerate(TOPIC_RULES)}
    return {k: groups[k]
            for k in sorted(groups, key=lambda k: order.get(k, len(order)))}


def _rank(counter: Counter, total: int, top: int = 8) -> List[Dict]:
    return [{"text": text, "count": n, "ratio": round(n / total, 4) if total else 0.0}
            for text, n in counter.most_common(top)]


# ------------------------------------------------------------------- chữ ký

def is_closing_line(line: str) -> bool:
    low = strip_accents(line).lower()
    return any(strip_accents(h) in low for h in CLOSING_HINTS)


def _split_signature(lines: List[str], doc_freq: Counter, threshold: int) -> tuple:
    """Tách MỘT thư thành (dòng thân, dòng chữ ký).

    Vòng lặp lùi dừng khi chạm câu kết ("Trân trọng,") — chữ ký bắt đầu SAU câu kết,
    không dừng ở đó thì nó nuốt luôn câu kết và cả dòng thân thư hay lặp.
    """
    block, idx = [], len(lines)
    for i in range(len(lines) - 1, -1, -1):
        line = lines[i]
        if is_closing_line(line) or len(block) >= MAX_SIGNATURE_LINES:
            break
        if doc_freq[line] < threshold:
            break
        block.append(line)
        idx = i
    return lines[:idx], list(reversed(block))


def split_corpus(bodies: List[str]) -> tuple:
    """Tách chữ ký khỏi TỪNG thư, trả (danh sách dòng thân, khối chữ ký phổ biến nhất).

    Phải tách theo từng thư chứ không so khớp một chuỗi chữ ký cố định: người dùng
    thực tế có nhiều biến thể chữ ký (bản ngắn, bản đầy đủ của công ty). Chỉ loại
    biến thể phổ biến nhất thì các biến thể còn lại tràn hết vào bảng cụm từ và làm
    sai luôn thống kê độ dài.
    """
    total = len(bodies)
    if not total:
        return [], {"signature": "", "confidence": 0.0}

    # Đếm theo SỐ THƯ chứa dòng đó, không phải số lần xuất hiện, để một thư lặp
    # đi lặp lại một dòng không tự đẩy dòng đó thành chữ ký.
    doc_freq = Counter()
    for body in bodies:
        doc_freq.update(set(_lines(body)))
    threshold = max(2, int(total * SIGNATURE_MIN_RATIO))

    stripped, blocks = [], Counter()
    for body in bodies:
        body_lines, sig = _split_signature(_lines(body), doc_freq, threshold)
        stripped.append(body_lines)
        if sig:
            blocks["\n".join(sig)] += 1

    if not blocks:
        return stripped, {"signature": "", "confidence": 0.0}
    text, count = blocks.most_common(1)[0]
    return stripped, {"signature": text, "confidence": round(count / total, 4)}


# ------------------------------------------------- câu chào và câu kết

def collect_greetings(docs: List[List[str]]) -> Counter:
    counter = Counter()
    for lines in docs:
        if lines and len(lines[0]) <= MAX_GREETING_LEN:
            counter[lines[0]] += 1
    return counter


def collect_closings(docs: List[List[str]]) -> Counter:
    counter = Counter()
    for lines in docs:
        for line in reversed(lines[-3:]):
            if len(line) > MAX_CLOSING_LEN:
                continue
            if is_closing_line(line):
                counter[line] += 1
                break
    return counter


# --------------------------------------------------------------- cụm từ

def collect_phrases(docs: List[List[str]], skip_lines: set) -> List[Dict]:
    """N-gram theo SỐ THƯ chứa cụm đó; giữ cụm dài, bỏ cụm ngắn nằm trọn bên trong.

    skip_lines: câu chào và câu kết — chúng đã được thống kê riêng, để lại chỉ làm
    nhiễu bảng cụm từ.
    """
    total = len(docs)
    doc_freq = Counter()
    surface: Dict[str, Counter] = {}

    for doc in docs:
        lines = [ln for ln in doc
                 if ln not in skip_lines and not is_closing_line(ln)]
        words = _WORD_RE.findall(" ".join(lines)[:NGRAM_BODY_CHARS])
        seen = set()
        for n in range(NGRAM_MIN, NGRAM_MAX + 1):
            for i in range(len(words) - n + 1):
                gram = words[i:i + n]
                key = " ".join(w.lower() for w in gram)
                if key in seen:
                    continue
                seen.add(key)
                surface.setdefault(key, Counter())[" ".join(gram)] += 1
        doc_freq.update(seen)

    threshold = max(3, int(total * 0.01))
    kept = {k: c for k, c in doc_freq.items() if c >= threshold}

    # Hai bước lọc trùng, vì n-gram chồng nhau của cùng một câu sinh ra rất nhiều
    # biến thể gần giống — để nguyên thì bảng tham chiếu đầy nhiễu, khó đọc.
    # Có `k` ở cuối khoá sắp xếp để thứ tự là TOÀN PHẦN. Thiếu nó, các cụm đồng hạng
    # giữ nguyên thứ tự chèn vào dict — vốn đến từ việc duyệt `set`, phụ thuộc hash
    # ngẫu nhiên hoá của Python. Cùng một hộp thư sẽ ra hồ sơ khác nhau mỗi lần chạy.
    ordered = sorted(kept, key=lambda k: (-len(k.split()), -kept[k], k))
    result = []
    for key in ordered:
        # 1) Cụm ngắn nằm trọn trong cụm dài mà tần suất xấp xỉ.
        if any(key in longer and kept[key] <= kept[longer] / 0.8
               for longer in (r["_key"] for r in result)):
            continue
        # 2) Cụm chồng lấn phần lớn từ với một cụm đã giữ, tần suất tương đương.
        words = set(key.split())
        if any(len(words & set(other.split())) >= 0.7 * len(words)
               and kept[key] <= kept[other] * 1.5
               for other in (r["_key"] for r in result)):
            continue
        result.append({
            "_key": key,
            "text": surface[key].most_common(1)[0][0],
            "count": kept[key],
            "ratio": round(kept[key] / total, 4) if total else 0.0,
        })
        if len(result) >= TOP_N:
            break
    for r in result:
        r.pop("_key", None)
    return sorted(result, key=lambda r: -r["count"])


# ---------------------------------------------------------------- độ dài

def length_stats(docs: List[List[str]]) -> Dict:
    """Dùng TRUNG VỊ, không phải trung bình — vài thư dài bất thường sẽ kéo lệch."""
    words, sentences = [], []
    for doc in docs:
        text = "\n".join(doc)
        w = len(_WORD_RE.findall(text))
        if not w:
            continue
        words.append(w)
        sentences.append(max(1, len([s for s in _SENT_SPLIT_RE.split(text) if s.strip()])))
    if not words:
        return {"median_words": 0, "median_sentences": 0, "words_per_sentence": 0}
    return {
        "median_words": int(statistics.median(words)),
        "median_sentences": int(statistics.median(sentences)),
        "words_per_sentence": int(statistics.median(words) /
                                  max(1, statistics.median(sentences))),
    }


def detect_language(bodies: List[str]) -> str:
    joined = " ".join(bodies)[:200_000]
    letters = [c for c in joined if c.isalpha()]
    if not letters:
        return "Unknown"
    accented = sum(1 for c in letters if strip_accents(c) != c)
    ratio = accented / len(letters)
    if ratio > 0.04:
        return "Vietnamese"
    return "English" if ratio < 0.005 else "Bilingual"


# ------------------------------------------------------------------ tổng

def analyze_corpus(emails: List[Dict]) -> Dict:
    """emails: [{subject, body, …}] — body là phần người dùng tự viết, CHƯA cắt."""
    bodies = [(e.get("body") or "").strip() for e in emails]
    bodies = [b for b in bodies if b]
    total = len(bodies)
    if not total:
        return {"analyzed_count": 0}

    docs, sig = split_corpus(bodies)
    greet = collect_greetings(docs)
    close = collect_closings(docs)

    return {
        "analyzed_count": total,
        "greeting_patterns": _rank(greet, total),
        "closing_patterns": _rank(close, total),
        "common_phrases": collect_phrases(docs, set(greet) | set(close)),
        "signature": sig["signature"],
        "signature_confidence": sig["confidence"],
        "length": length_stats(docs),
        "language": detect_language(bodies),
    }
