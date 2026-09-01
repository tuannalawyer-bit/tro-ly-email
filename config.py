import os

from dotenv import load_dotenv

from paths import DATA_ROOT, ENV_FILE, RES_DIR

# Phải truyền đường dẫn tường minh: khi đã đóng gói, find_dotenv của python-dotenv rơi
# về os.getcwd(), mà thư mục làm việc lúc chạy từ lối tắt Startup là System32.
load_dotenv(ENV_FILE if ENV_FILE.is_file() else None)

APP_NAME = "Email Assistant"
VERSION = "2.3.9"

# BASE_DIR giữ lại làm bí danh của DATA_ROOT cho mã cũ; chỗ nào cần tài nguyên chỉ đọc
# (frontend/, addin/, kien_thuc_mau/) thì phải dùng RES_DIR.
BASE_DIR = DATA_ROOT
DATA_DIR = DATA_ROOT / "data"
CACHE_DIR = DATA_DIR / "cache"
STYLE_PROFILES_DIR = DATA_DIR / "style_profiles"

# Kho quy tắc viết tay dạng Markdown, sửa được bằng Claude Desktop / ChatGPT Desktop.
# KHÔNG tạo tự động: tệp mẫu nằm ở kien_thuc_mau/, người dùng tự sao chép sang.
KNOWLEDGE_DIR = DATA_ROOT / "kien_thuc"
KNOWLEDGE_SAMPLE_DIR = RES_DIR / "kien_thuc_mau"

# Create directories if they don't exist
DATA_DIR.mkdir(parents=True, exist_ok=True)
CACHE_DIR.mkdir(exist_ok=True)
STYLE_PROFILES_DIR.mkdir(exist_ok=True)

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
# Hạn mức miễn phí tính RIÊNG cho từng model, và model càng mới hạn mức càng thấp.
# Đo thực tế ngày 05-08-2026 với API key của dự án:
#   gemini-flash-latest -> trỏ tới gemini-3.6-flash, chỉ 20 lượt/NGÀY (~6 thư trả lời).
#   gemini-2.5-flash / 2.5-flash-lite -> 404, Google đã ngừng cấp cho key mới.
#   gemini-3.5-flash -> chạy được nhưng chậm (~9,6s/lượt).
#   gemini-3.1-flash-lite -> chạy được và nhanh (~0,7s/lượt).  <-- đang dùng
# KHÔNG quay lại alias "-latest": nó luôn trỏ tới model mới nhất, tức là model có
# hạn mức miễn phí thấp nhất. Muốn đổi thì sửa GEMINI_MODEL trong .env.
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.5-flash-lite")

# Định dạng cho thư nháp AI soạn. Outlook dựng HTML không khai báo font bằng font
# mặc định của trình duyệt (thường là Times New Roman) trong khi chữ ký và phần trích
# dẫn dùng font soạn thảo của người dùng — hai khối lệch hẳn nhau trông rất xấu.
# Calibri 11pt là font soạn thảo mặc định của Outlook.
DRAFT_FONT_FAMILY = os.getenv("DRAFT_FONT_FAMILY", "Calibri, 'Segoe UI', sans-serif")
DRAFT_FONT_SIZE = os.getenv("DRAFT_FONT_SIZE", "11pt")
DRAFT_TEXT_COLOR = os.getenv("DRAFT_TEXT_COLOR", "#000000")
