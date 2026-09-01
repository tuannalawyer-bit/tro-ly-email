from __future__ import annotations

import logging
import os
import shutil
import subprocess
import tempfile
from typing import Optional

logger = logging.getLogger(__name__)

def _find_browser_binary() -> Optional[str]:
    candidates = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]
    for p in candidates:
        if os.path.isfile(p):
            return p
    return None

def capture_map_screenshot(lat: float, lng: float, timeout_sec: int = 15) -> Optional[bytes]:
    browser_bin = _find_browser_binary()
    if not browser_bin:
        logger.warning("Không tìm thấy Chrome/Edge trên máy để chụp ảnh bản đồ.")
        return None

    temp_dir = tempfile.mkdtemp(prefix="map_capture_")
    out_png = os.path.join(temp_dir, "map_view.png")
    user_data = os.path.join(temp_dir, "user_data")
    target_url = f"https://meeymap.com/?lat={lat}&lng={lng}"

    cmd = [
        browser_bin,
        "--headless=new",
        "--disable-gpu",
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-background-networking",
        f"--user-data-dir={user_data}",
        f"--screenshot={out_png}",
        "--window-size=1280,800",
        "--virtual-time-budget=8000",
        target_url,
    ]

    try:
        logger.info("Đang chụp ảnh bản đồ quy hoạch tại GPS (%.6f, %.6f)...", lat, lng)
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout_sec)
        if os.path.isfile(out_png) and os.path.getsize(out_png) > 1000:
            with open(out_png, "rb") as f:
                img_data = f.read()
            logger.info("Chụp ảnh bản đồ thành công (%d bytes)", len(img_data))
            return img_data
        else:
            logger.warning("Không tạo được ảnh bản đồ: returncode=%d, stderr=%s", res.returncode, res.stderr[:200])
    except subprocess.TimeoutExpired:
        logger.warning("Hết thời gian chờ chụp ảnh bản đồ (> %ds)", timeout_sec)
    except Exception as exc:
        logger.warning("Lỗi chụp ảnh bản đồ: %s", exc)
    finally:
        try:
            shutil.rmtree(temp_dir, ignore_errors=True)
        except Exception:
            pass

    return None

def analyze_map_with_vision(api_key: str, model_name: str, img_bytes: bytes, lat: float, lng: float) -> Optional[str]:
    if not api_key or not img_bytes:
        return None

    try:
        import google.generativeai as genai

        genai.configure(api_key=api_key)
        target_model = model_name if model_name else "gemini-3.1-flash-lite"
        model = genai.GenerativeModel(target_model)

        prompt = (
            f"Dưới đây là ảnh chụp màn hình bản đồ quy hoạch thực tế tại tọa độ GPS ({lat:.6f}, {lng:.6f}).\n"
            "Hãy quan sát kỹ hình ảnh và tóm tắt ngắn gọn các thông tin quy hoạch sau:\n"
            "1. Thông tin quy hoạch/đồ án hoặc số tờ, số thửa, địa chỉ ghi nhận trên bản đồ.\n"
            "2. Tình trạng loại đất theo màu sắc/ký hiệu quy hoạch (Đất ở đô thị ODT, Đất ở nông thôn ONT, Đất giao thông, Đất cây xanh, Đất công cộng, Đất nông nghiệp/vườn/ao...).\n"
            "3. Chỉ giới mở rộng đường, hành lang giao thông hoặc ranh giới quy hoạch có đi qua thửa đất/khu vực không.\n"
            "4. Nhận định rủi ro quy hoạch thực tế và rủi ro mục đích sử dụng đất (ngắn gọn 1-2 câu).\n"
            "Nếu ảnh là trang xác thực hoặc không thấy rõ bản đồ, hãy ghi rõ không thể nhận diện qua ảnh."
        )

        image_part = {
            "mime_type": "image/png",
            "data": img_bytes,
        }

        response = model.generate_content([prompt, image_part])
        text = (response.text or "").strip()
        if text and "xác minh bảo mật" not in text.lower():
            return text
        return None
    except Exception as exc:
        logger.warning("Lỗi phân tích bản đồ bằng Gemini Vision: %s", exc)
        return None
