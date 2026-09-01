from __future__ import annotations

import json
import logging
import os
import re
import urllib.error
import urllib.parse
import urllib.request
from typing import Any, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

def _do_serper_search(query: str, api_key: str, num: int = 10) -> List[Dict[str, Any]]:
    """Gọi API Serper.dev để tìm kiếm kết quả Google Search."""
    if not api_key:
        return []
    url = "https://google.serper.dev/search"
    headers = {
        "X-API-KEY": api_key,
        "Content-Type": "application/json",
        "User-Agent": "EmailAssistant/2.0",
    }
    payload = {
        "q": query,
        "gl": "vn",
        "hl": "vi",
        "num": num,
    }
    try:
        data = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(url, data=data, headers=headers, method="POST")
        with urllib.request.urlopen(req, timeout=8) as resp:
            body = resp.read().decode("utf-8")
            res = json.loads(body)
            organic = res.get("organic", [])
            results = []
            for item in organic:
                link = item.get("link", "")
                title = item.get("title", "")
                snippet = item.get("snippet", "")
                if link and not _is_junk_link(link):
                    results.append({
                        "title": title,
                        "link": link,
                        "snippet": snippet,
                    })
            return results
    except Exception as exc:
        logger.warning("Lỗi tìm kiếm Serper cho '%s': %s", query, exc)
        return []

def _do_google_cse_search(query: str, api_key: str, cse_id: str, num: int = 10) -> List[Dict[str, Any]]:
    """Gọi Google Custom Search JSON API."""
    if not api_key or not cse_id:
        return []
    params = urllib.parse.urlencode({
        "key": api_key,
        "cx": cse_id,
        "q": query,
        "num": min(num, 10),
        "gl": "vn",
        "hl": "vi",
    })
    url = f"https://www.googleapis.com/customsearch/v1?{params}"
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "EmailAssistant/2.0"})
        with urllib.request.urlopen(req, timeout=8) as resp:
            body = resp.read().decode("utf-8")
            res = json.loads(body)
            items = res.get("items", [])
            results = []
            for item in items:
                link = item.get("link", "")
                title = item.get("title", "")
                snippet = item.get("snippet", "")
                if link and not _is_junk_link(link):
                    results.append({
                        "title": title,
                        "link": link,
                        "snippet": snippet,
                    })
            return results
    except Exception as exc:
        logger.warning("Lỗi tìm kiếm Google CSE cho '%s': %s", query, exc)
        return []

def _is_junk_link(url: str) -> bool:
    """Loại bỏ các link mạng xã hội, video, trang rác không có giá trị khảo sát."""
    junk_domains = [
        "facebook.com",
        "fb.com",
        "youtube.com",
        "youtu.be",
        "tiktok.com",
        "instagram.com",
        "twitter.com",
        "x.com",
        "pinterest.com",
        "linkedin.com",
        "shopee.vn",
        "lazada.vn",
        "tiki.vn",
        "sendo.vn",
        "chotot.org",
    ]
    u = url.lower()
    return any(d in u for d in junk_domains)

def _extract_address_components(location: str) -> Tuple[str, str, str, str]:
    """Tách địa chỉ thành (Đường/Phố, Phường/Xã, Quận/Huyện, Tỉnh/TP)."""
    parts = [p.strip() for p in location.split(",") if p.strip()]
    street = ""
    ward = ""
    district = ""
    province = ""
    if len(parts) == 1:
        street = parts[0]
    elif len(parts) == 2:
        street, district = parts[0], parts[1]
    elif len(parts) == 3:
        street, district, province = parts[0], parts[1], parts[2]
    elif len(parts) >= 4:
        street, ward, district, province = parts[0], parts[1], parts[2], parts[3]
    return street, ward, district, province

def _extract_broad_location(location: str) -> str:
    """Lấy phần địa danh rộng hơn (bỏ số nhà/ngõ hẹp)."""
    parts = [p.strip() for p in location.split(",") if p.strip()]
    if len(parts) > 1:
        return ", ".join(parts[1:])
    return location

class WebSearcher:
    """Bộ máy tìm kiếm giá thuê, tin tức quy hoạch và thông tin dự án thực tế."""

    def __init__(self, serper_key: str = "", google_cse_key: str = "", google_cse_id: str = ""):
        self.serper_key = serper_key.strip() if serper_key else os.getenv("SERPER_API_KEY", "").strip()
        self.google_cse_key = google_cse_key.strip() if google_cse_key else os.getenv("GOOGLE_CSE_API_KEY", "").strip()
        self.google_cse_id = google_cse_id.strip() if google_cse_id else os.getenv("GOOGLE_CSE_ID", "").strip()

    @property
    def enabled(self) -> bool:
        return bool(self.serper_key or (self.google_cse_key and self.google_cse_id))

    def _search(self, query: str, num: int = 10) -> List[Dict[str, Any]]:
        if not self.enabled or not query.strip():
            return []
        if self.serper_key:
            results = _do_serper_search(query, self.serper_key, num=num)
            if results:
                return results
        if self.google_cse_key and self.google_cse_id:
            return _do_google_cse_search(query, self.google_cse_key, self.google_cse_id, num=num)
        return []

    def search_general(self, query: str, num: int = 5) -> List[Dict[str, Any]]:
        """Tìm kiếm thông tin tổng quát."""
        return self._search(query, num=num)

    def search_rental_prices(self, location: str, num: int = 6) -> List[Dict[str, Any]]:
        """Tìm kiếm các tin đăng cho thuê nhà/mặt bằng thực tế tại khu vực."""
        if not location.strip():
            return []
        street, ward, district, province = _extract_address_components(location)
        target = f"{street} {ward} {district} {province}".strip()
        queries = [
            f"cho thuê mặt bằng {target}",
            f"cho thuê nhà nguyên căn {target}",
            f"cho thuê nhà kinh doanh {target}",
            f"giá thuê mặt bằng {target}",
            f"cho thuê nhà {street} {district}".strip(),
        ]
        all_results = []
        seen_links = set()
        for q in queries:
            if len(all_results) >= num:
                break
            res = self._search(q, num=num)
            for item in res:
                link = item.get("link", "")
                if link and link not in seen_links:
                    seen_links.add(link)
                    all_results.append(item)
                    if len(all_results) >= num:
                        break
        return all_results

    def search_planning_info(self, location: str, num: int = 6) -> List[Dict[str, Any]]:
        """Tìm các bài báo, văn bản pháp luật, tin tức quy hoạch mở rộng đường, chỉ giới đường đỏ tại khu vực."""
        if not location.strip():
            return []
        street, ward, district, province = _extract_address_components(location)
        target = f"{street} {district} {province}".strip()
        queries = [
            f"quy hoạch mở rộng đường {target}",
            f"chỉ giới đường đỏ {target}",
            f"dự án mở đường {target}",
            f"thông tin quy hoạch {target}",
            f"quy hoạch giao thông {district} {province}".strip(),
            f"bản đồ quy hoạch phân khu {district} {province}".strip(),
            f"quy hoạch hành lang an toàn giao thông {target}",
        ]
        all_results = []
        seen_links = set()
        for q in queries:
            if len(all_results) >= num:
                break
            res = self._search(q, num=num)
            for item in res:
                link = item.get("link", "")
                if link and link not in seen_links:
                    seen_links.add(link)
                    all_results.append(item)
                    if len(all_results) >= num:
                        break
        return all_results

    def search_project_info(self, location: str, num: int = 4) -> List[Dict[str, Any]]:
        """Tìm thông tin dự án, tin tức khu vực bất động sản."""
        if not location.strip():
            return []
        street, ward, district, province = _extract_address_components(location)
        target = f"{street} {district} {province}".strip()
        query = f"thông tin dự án bất động sản {target}"
        return self._search(query, num=num)

    @classmethod
    def _parse_price_area_from_snippet(cls, snippet: str) -> Tuple[Optional[float], Optional[float], Optional[float]]:
        """Trích xuất giá thuê (triệu) và diện tích (m2) từ snippet để tính đơn giá/m²."""
        price_vnd = None
        area_m2 = None
        unit_price = None

        price_m = re.search(r"(\d+[\.,]?\d*)\s*(?:triệu|tr|million|tr/tháng|triệu/tháng)", snippet, re.I)
        if price_m:
            try:
                val = float(price_m.group(1).replace(",", "."))
                price_vnd = val * 1000000
            except Exception:
                pass
        else:
            price_raw = re.search(r"(\d{1,3}(?:[.,]\d{3})+)\s*(?:VND|đồng|vnđ|vnd)?", snippet, re.I)
            if price_raw:
                try:
                    cleaned = re.sub(r"[.,]", "", price_raw.group(1))
                    val = float(cleaned)
                    if val >= 100000:
                        price_vnd = val
                except Exception:
                    pass

        area_m = re.search(r"(\d+[\.,]?\d*)\s*(?:m²|m2|mét vuông)", snippet, re.I)
        if area_m:
            try:
                area_m2 = float(area_m.group(1).replace(",", "."))
            except Exception:
                pass

        if price_vnd and area_m2 and area_m2 > 0:
            unit_price = price_vnd / area_m2

        return price_vnd, area_m2, unit_price

    @classmethod
    def _extract_address_from_snippet(cls, snippet: str, title: str) -> str:
        """Cố gắng trích xuất địa chỉ từ snippet hoặc tiêu đề bài đăng."""
        m = re.search(r"\d+[A-Za-z]?\s+(?:phố|đường|ngõ|ngách|hẻm|pho|duong)\s+[^\,\.;]{3,40}", snippet + " " + title, re.I)
        if m:
            return m.group(0).strip()
        m2 = re.search(r"(?:phường|quận|huyện|xã)\s+[^\,\.;]{2,30}", snippet + " " + title, re.I)
        if m2:
            return m2.group(0).strip()
        return "—"

    @classmethod
    def format_results_for_prompt(cls, results: List[Dict[str, Any]], label: str = "") -> str:
        """Định dạng kết quả tìm kiếm để chèn vào prompt Gemini."""
        if not results:
            return ""
        lines = [f"### {label}" if label else "### Kết quả tìm kiếm thực tế"]
        for i, r in enumerate(results, 1):
            title = r.get("title", "")
            link = r.get("link", "")
            snippet = r.get("snippet", "")
            lines.append(f"{i}. [{title}]({link})")
            if snippet:
                lines.append(f"   Snippet: {snippet[:250]}")
        lines.append("\n*BẮT BUỘC: Bạn phải chèn chính xác các link (URL) ở trên vào từng mục tương ứng trong báo cáo email.*\n")
        return "\n".join(lines)

    @classmethod
    def format_rental_results_for_prompt(cls, results: List[Dict[str, Any]], label: str = "") -> str:
        """Tạo bảng HTML so sánh giá thuê sẵn sàng để AI copy nguyên văn vào email."""
        return cls.build_rental_table_html(results)

    @classmethod
    def build_rental_table_html(cls, results: List[Dict[str, Any]]) -> str:
        """Tạo HTML thuần túy của bảng so sánh giá thuê — dùng để Python inject thẳng vào response.

        Bảng gồm: Thông tin bài đăng | Địa chỉ | Giá thuê | Diện tích | Giá thuê/m²
        """
        if not results:
            return ""
        tbl = 'style="border-collapse:collapse;font-size:10pt;font-family:Calibri,Arial,sans-serif;width:100%;margin:10px 0 6px 0;"'
        th = 'style="background:#2E4057;color:#fff;padding:6px 10px;border:1px solid #888;text-align:left;white-space:nowrap;"'
        td = 'style="padding:5px 8px;border:1px solid #ccc;vertical-align:top;"'
        td_r = 'style="padding:5px 8px;border:1px solid #ccc;text-align:right;white-space:nowrap;"'
        td_hl = 'style="padding:5px 8px;border:1px solid #ccc;text-align:right;white-space:nowrap;color:#C0392B;font-weight:bold;"'

        lines = [
            f'<table {tbl}>',
            '  <thead><tr>',
            f'    <th {th}>Thông tin bài đăng</th>',
            f'    <th {th}>Địa chỉ</th>',
            f'    <th {th}>Giá thuê</th>',
            f'    <th {th}>Diện tích</th>',
            f'    <th {th}>Giá thuê/m²</th>',
            '  </tr></thead>',
            '  <tbody>',
        ]

        for i, r in enumerate(results):
            snippet = r.get("snippet", "")
            title = r.get("title", "")
            link = r.get("link", "")
            price_vnd, area_m2, unit_price = cls._parse_price_area_from_snippet(snippet)
            address = cls._extract_address_from_snippet(snippet, title)

            price_str = f"{price_vnd/1000000:.0f} triệu/tháng" if price_vnd else "—"
            area_str = f"{area_m2:.0f} m²" if area_m2 else "—"
            unit_str = f"{unit_price:,.0f} đ/m²/tháng" if unit_price else "—"
            short_title = title[:75] + "…" if len(title) > 75 else title

            bg = 'style="background:#F8F8F8;"' if i % 2 == 1 else ""
            lines.append(f'  <tr {bg}>')
            lines.append(f'    <td {td}><a href="{link}" style="color:#1155CC;">{short_title}</a></td>')
            lines.append(f'    <td {td}>{address}</td>')
            lines.append(f'    <td {td_r}>{price_str}</td>')
            lines.append(f'    <td {td_r}>{area_str}</td>')
            lines.append(f'    <td {td_hl}>{unit_str}</td>')
            lines.append('  </tr>')

        lines.append('  </tbody>')
        lines.append('</table>\n')
        return "\n".join(lines)

    @classmethod
    def format_rental_context_for_prompt(cls, results: List[Dict[str, Any]], label: str = "") -> str:
        """Tóm tắt ngắn gọn kết quả giá thuê để đưa vào prompt — AI chỉ cần biết khoảng giá."""
        if not results:
            return ""
        header = f"### {label}" if label else "### Giá thuê khu vực (tham khảo)"
        lines = [header]
        unit_prices = []

        for r in results:
            snippet = r.get("snippet", "")
            price_vnd, area_m2, unit_price = cls._parse_price_area_from_snippet(snippet)
            title = r.get("title", "")[:60]
            link = r.get("link", "")
            if unit_price:
                unit_prices.append(unit_price)
                lines.append(f"- [{title}]({link}): ~{price_vnd/1000000:.0f}tr/{area_m2:.0f}m² = **{unit_price:,.0f} đ/m²/tháng**")
            elif price_vnd:
                lines.append(f"- [{title}]({link}): ~{price_vnd/1000000:.0f}tr/tháng (chưa rõ m²)")
            else:
                lines.append(f"- [{title}]({link})")

        if unit_prices:
            lo = min(unit_prices)
            hi = max(unit_prices)
            lines.append(f"\n*Khoảng giá thị trường tính được: {lo:,.0f} – {hi:,.0f} đ/m²/tháng.*")

        lines.append("\n*Python sẽ tự inject bảng so sánh vào email. AI CHỈ CẦN: tính đơn giá đề xuất (đ/m²/tháng) và viết 1 câu so sánh với khoảng giá trên.*")
        return "\n".join(lines)
