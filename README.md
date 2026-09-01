# Trợ lý Email v2.1.0

Trợ lý soạn email chạy hoàn toàn trên máy cá nhân, gồm hai phần dùng chung một lõi Python:

| Thành phần | Vai trò | Khởi chạy |
|---|---|---|
| **Add-in Outlook** | Lớp bề mặt: đọc thư đang mở + tệp đính kèm, soạn thư trả lời, mở form trả lời của Outlook | `CHAY_ADDIN_BACKEND.bat` |
| **Ứng dụng Desktop** | Bảng điều khiển: học văn phong từ hộp thư đã gửi, quét phân loại toàn bộ thư mục, cài API key | `CHAY_UNGDUNG.bat` |
| **Kho tri thức** | Quy tắc viết thư dạng Markdown, sửa bằng Claude Desktop / ChatGPT Desktop | `kien_thuc/` |

Hoặc chạy **`CHAY_NGAM.bat`** để gộp cả hai vào một icon ở khay hệ thống, giống UniKey —
không cửa sổ đen, không chiếm chỗ ở taskbar. Xem mục [Chạy ngầm](#chạy-ngầm-ở-khay-hệ-thống).

Muốn đưa cho người khác dùng thì đóng gói thành **một tệp `.exe`** — xem mục
[Đóng gói](#đóng-gói-thành-một-tệp-exe). Người nhận đọc
[HUONG_DAN_CAI_DAT.md](HUONG_DAN_CAI_DAT.md), không cần đọc tài liệu này.

Không dùng Microsoft Graph, không dùng Entra ID, không gửi dữ liệu đi đâu ngoài lời gọi Gemini API.

---

## Yêu cầu hệ thống

- Windows 10/11
- Microsoft Outlook Desktop (bản classic), tài khoản **Exchange / Microsoft 365** — hộp thư POP/IMAP không nạp được add-in
- Python 3.10+ (dự án đang chạy 3.12)
- Khoá API Google Gemini

---

## Cài đặt

### 1. Môi trường ảo

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

> Nếu máy nằm sau proxy công ty có chặn/giải mã TLS, `pip` hoặc `uv` sẽ báo
> `invalid peer certificate: UnknownIssuer`. Khi đó dùng cờ để lấy chứng chỉ từ
> kho của Windows: `uv pip install --native-tls -r requirements.txt`.

### 2. Khoá API

Tạo tệp `.env` từ `.env.example`:

```env
GEMINI_API_KEY=khoa_api_cua_ban
GEMINI_MODEL=gemini-3.1-flash-lite
```

> **Đừng dùng bí danh `gemini-flash-latest`.** Hạn mức miễn phí tính riêng cho từng
> model, và **model càng mới thì hạn mức càng thấp**. Đo thực tế ngày 05-08-2026:
> `-latest` trỏ tới `gemini-3.6-flash` và chỉ cho **20 lượt/ngày** — khoảng 6 thư
> trả lời là hết. Các model đời 2.5 trở về trước thì đã bị Google gỡ khỏi key mới (404).
>
> Muốn chất lượng cao hơn và chấp nhận chậm hơn (~9,6s/lượt so với ~0,7s) thì đổi
> sang `gemini-3.5-flash`. Khi hết hạn mức, ứng dụng sẽ báo rõ đây là hạn mức
> **theo ngày** kèm gợi ý model thay thế.

Cũng có thể nhập khoá trực tiếp trong màn hình **Cài đặt** của ứng dụng desktop.

### 3. Kho tri thức

```powershell
Copy-Item -Recurse kien_thuc_mau kien_thuc
```

Rồi mở thư mục `kien_thuc/` bằng Claude Desktop (hoặc ChatGPT Desktop / Antigravity)
và sửa lại cho đúng thông tin của bạn. Chi tiết xem [HUONG_DAN_KIEN_THUC.md](HUONG_DAN_KIEN_THUC.md).

Thư mục này **không được đưa lên Git** vì chứa thông tin cá nhân và nội bộ.

### 4. Cài add-in vào Outlook

Xem hướng dẫn đầy đủ ở [HUONG_DAN_ADDIN_LOCAL.md](HUONG_DAN_ADDIN_LOCAL.md). Tóm tắt:

1. Chạy `TAO_CHUNG_CHI_ADDIN.bat` một lần.
2. Chạy `CAI_CHUNG_CHI.bat` — cài chứng chỉ vào **Trusted Root**, không cần quyền Admin.
   **Bỏ qua bước này thì Outlook không tải được icon nên nút ribbon không xuất hiện,
   và task pane hiện trắng trơn không báo lỗi.**
3. Chạy `CHAY_ADDIN_BACKEND.bat` và để cửa sổ đó mở.
4. Chạy `KIEM_TRA_ADDIN.bat` — phải báo *"MOI DIEU KIEN BAT BUOC DEU DAT"* trước khi đi tiếp.
5. Trong Outlook: **Get Add-ins → My add-ins → Add a custom add-in → Add from file** → chọn `addin/manifest.xml`.
6. Mở một email bất kỳ, bấm nút **Soạn trả lời AI** trên ribbon.

> Gặp bất kỳ trục trặc nào với add-in, hãy chạy `KIEM_TRA_ADDIN.bat` **trước tiên**.
> Nó kiểm tra 18 điều kiện và in kèm câu lệnh sửa cho từng mục không đạt.

---

## Tính năng

### Add-in Outlook

- Đọc thư đang mở qua Office.js, tự cắt bỏ phần trích dẫn để AI đọc đúng nội dung mới.
- **Đọc tệp đính kèm**: PDF và ảnh gửi thẳng cho Gemini (đọc được cả PDF scan);
  `.xlsx` và `.docx` được trích thành văn bản tại máy.
- Soạn thư theo **kho tri thức** + **hồ sơ văn phong** + **hướng dẫn riêng cho loại thư**
  đang trả lời.
- **Ô chọn mẫu thư**: để trống thì tự nhận diện loại bằng từ khoá (không tốn lượt gọi AI),
  hoặc chọn tay một mẫu cụ thể — lựa chọn tay luôn thắng kết quả tự khớp. Sau khi soạn
  xong, giao diện cho biết đã dùng mẫu nào và ai chọn.
- Chỉnh sửa bản nháp bằng câu lệnh tiếng Việt, sửa tay trực tiếp trong task pane.
- **Định dạng khớp Outlook**: bản nháp mang sẵn style inline (Calibri 11pt, khoảng cách
  đoạn chuẩn) nên hòa vào chữ ký và phần trích dẫn thay vì lệch font. Đổi được qua
  `DRAFT_FONT_FAMILY` / `DRAFT_FONT_SIZE` trong `.env`.
- **Học văn phong ngay trong task pane**, chạy nền có thanh tiến độ.
- **Chèn vào thư trả lời**: mở cửa sổ trả lời của Outlook đã điền sẵn nội dung, giữ nguyên
  phần trích dẫn gốc. Có nút sao chép HTML dự phòng.
- Phân loại nhanh thư đang mở (mức ưu tiên, có cần trả lời không).

### Ứng dụng Desktop

- Duyệt thư mục Outlook, xem email trong iframe sandbox (không cho chạy script, chống XSS).
- Quét và phân loại toàn bộ thư mục theo lô, lưu đệm ở `data/cache/classifications.json`.
- **Xuất thư để học văn phong**: quét **mọi kho thư** đã gắn vào Outlook (hộp thư chính
  lẫn tệp lưu trữ), xuất ra `xuat_thu/` dưới dạng cặp "thư đến → bạn đã trả lời", nội dung
  không cắt ngắn. Đồng thời đếm câu chào / câu kết / chữ ký / cụm từ / độ dài kèm tần suất
  thật vào `data/style_profiles/default.json`. **Không gọi Gemini lần nào.**
  Sau đó mở `xuat_thu/` bằng Antigravity hoặc Claude Desktop để viết hướng dẫn theo từng
  loại thư — dùng gói thuê bao của công cụ đó, không tiêu hạn mức Gemini.
- Soạn và lưu thư nháp thẳng vào Drafts của Outlook.

---

## Chạy ngầm ở khay hệ thống

```
CHAY_NGAM.bat
```

Một icon ở khay (góc phải thanh tác vụ) thay cho hai cửa sổ đen. Bấm đúp để mở ứng dụng,
bấm **X** là thu về khay chứ không thoát — ứng dụng vẫn sống và backend add-in vẫn phục vụ
Outlook.

| Menu chuột phải | Việc |
|---|---|
| Mở Trợ lý Email | Hiện cửa sổ ứng dụng |
| Backend add-in: … | Trạng thái, kèm nút Khởi động / Dừng / Khởi động lại |
| Khởi động cùng Windows | Tạo hoặc xoá lối tắt trong `shell:startup` — không cần quyền Admin |
| Thoát | Đóng hẳn, dừng luôn backend |

Chi tiết:

- Chạy ngầm thì **không có nút taskbar**; mở cửa sổ lên mới có, đúng như UniKey. Nút thu
  nhỏ vẫn thu xuống taskbar như thường.
- Chỉ chạy được **một bản** cùng lúc. Mở lần hai sẽ báo đang chạy rồi tự thoát.
- Backend add-in chạy như tiến trình con khi chạy từ mã nguồn, và như một thread trong
  bản đóng gói (gọi lại exe onefile sẽ bung nén gói lần thứ hai). Nếu cổng 8765 đã có
  tiến trình khác chiếm, menu sẽ báo và không khởi động chồng lên.
- Vì `pythonw.exe` không có cửa sổ console nên **mọi lỗi ghi vào `data/logs/app.log`**;
  backend ghi riêng ở `data/logs/addin_server.log`. Gặp trục trặc thì mở hai tệp này trước.
- `CHAY_UNGDUNG.bat` và `CHAY_ADDIN_BACKEND.bat` vẫn dùng được như cũ khi cần thấy console.

---

## Đóng gói thành một tệp .exe

```powershell
uv pip install --native-tls -r requirements-build.txt
.\DONG_GOI.bat
```

Kết quả: `dist\CaiTroLyEmail.exe` (~50 MB). Đưa **duy nhất** tệp này cho người dùng; họ
đọc [HUONG_DAN_CAI_DAT.md](HUONG_DAN_CAI_DAT.md).

### Vì sao không dùng thẳng `--onefile`

Đo thực tế trên máy đích:

| Cách | Kích thước | Khởi động |
|---|---|---|
| `--onefile` thuần | 44 MB | **20,5 giây, mỗi lần chạy** |
| `--onedir` | 159 MB (thư mục) | 6,9 giây |
| **Bộ cài tự bung (đang dùng)** | **50 MB, một tệp** | 10 giây lần đầu, **8 giây** những lần sau |

`--onefile` phải bung ~800 tệp ra `%TEMP%` mỗi lần chạy và phần mềm bảo vệ quét lại từng
tệp — 20,5 giây là số đo ổn định qua nhiều lần, không phải cache lạnh. Bộ cài ở đây là
một exe onefile **rất nhỏ** (chỉ thư viện chuẩn) với gói ứng dụng dạng zip **nối vào
đuôi tệp**: `zipfile` đọc được zip nối sau dữ liệu khác, còn bootloader của PyInstaller
vẫn tìm ra archive của nó. Bung một lần vào `%LOCALAPPDATA%\TroLyEmail\app` rồi từ đó
chạy bản nhanh.

### Bản đóng gói khác gì bản chạy từ mã nguồn

| | Mã nguồn | Đóng gói |
|---|---|---|
| Tài nguyên (`frontend/`, `addin/`, `kien_thuc_mau/`) | thư mục dự án | trong gói, chỉ đọc |
| Dữ liệu (`data/`, `certs/`, `kien_thuc/`, `.env`) | thư mục dự án | `%LOCALAPPDATA%\TroLyEmail` |
| Backend add-in | tiến trình con | thread trong cùng tiến trình |
| Lối vào | các tệp `.bat` | tham số: `--cua-so`, `--xuat-thu`, `--kiem-tra`, `--thiet-lap` |

Ranh giới này nằm ở [paths.py](paths.py). Chạy từ mã nguồn thì `RES_DIR == DATA_ROOT` nên
mọi thứ giữ nguyên nếp cũ.

### Lưu ý khi dựng

- **Tắt bản đóng gói đang chạy trước khi dựng lại.** Nó khoá tệp trong `dist\` và
  PyInstaller sẽ báo lỗi khó hiểu.
- Cơ chế tiến trình con của PyInstaller thỉnh thoảng chết bất chợt; chạy lại là qua.
- Đổi phụ thuộc thì kiểm tra lại danh sách `HIDDEN` trong [dong_goi.py](dong_goi.py):
  pywin32 nạp `win32timezone` **động**, thiếu nó thì mọi email đọc từ Outlook đều hỏng.

### Trước khi phát rộng

Tệp exe **chưa được ký số**, trong khi nó đọc toàn bộ hộp thư Outlook, mở máy chủ HTTPS
cục bộ và cài chứng chỉ vào Trusted Root. Đó là hồ sơ mà SmartScreen và phần mềm giám
sát doanh nghiệp phản ứng mạnh — và phản ứng ấy hợp lý. Nên trao đổi với bộ phận CNTT
trước, thay vì chuyền tay tệp exe. Có chứng chỉ ký mã của công ty thì ký sẽ gỡ được phần
lớn ma sát.

---

## Bảo mật

- Backend add-in chỉ lắng nghe `127.0.0.1:8765`, **bắt buộc HTTPS** — thiếu chứng chỉ thì
  chương trình dừng hẳn chứ không tụt xuống HTTP.
- Không phát header CORS nào. Task pane được nạp từ chính origin đó nên fetch là same-origin;
  mọi trang web khác bị trình duyệt chặn đọc phản hồi.
- Mọi endpoint `/api/*` yêu cầu token sinh ngẫu nhiên lúc khởi động, nhúng vào task pane.
  **Khởi động lại backend thì token đổi** — phải đóng và mở lại task pane.
- Manifest chỉ xin quyền `ReadItem`, không xin `ReadWriteMailbox`.
- Nội dung HTML do Gemini sinh được lọc theo danh sách trắng trước khi chèn vào Outlook.
- `.env`, `data/`, `certs/`, `kien_thuc/` đều nằm trong `.gitignore`.

---

## Cấu trúc

```
addin/            Task pane + manifest + icon ribbon
addin_server.py   Backend HTTPS cục bộ cho add-in
backend/          Lõi dùng chung
  api.py              Mặt tiền cho cả desktop lẫn add-in
  outlook_client.py   Truy cập Outlook qua COM, quét mọi kho thư
  com_worker.py       Ép mọi lệnh COM về một thread STA duy nhất
  ai_engine.py        Gọi Gemini, dựng prompt, lọc và định dạng HTML
  knowledge_base.py   Nạp kho tri thức Markdown + nhận diện loại thư
  style_stats.py      Đếm thống kê văn phong bằng Python thuần (không AI)
  attachment_reader.py Đọc PDF/ảnh/xlsx/docx đính kèm
frontend/         Giao diện ứng dụng desktop
paths.py          Tách đường dẫn tài nguyên (chỉ đọc) khỏi dữ liệu (ghi được)
tray.py           Icon khay hệ thống: chạy ngầm, quản lý backend, khởi động cùng Windows
thiet_lap.py      Thiết lập lần đầu cho bản đóng gói
dong_goi.py       Dựng bộ cài một tệp  (DONG_GOI.bat)
cai_dat_sfx.py    Bộ cài tự bung, mang gói ứng dụng ở đuôi tệp exe
tao_icon.py       Sinh addin/assets/app.ico từ PNG, chạy một lần
kien_thuc_mau/    Kho tri thức mẫu (dữ liệu giả) — sao chép thành kien_thuc/
xuat_thu_da_gui.py  Công cụ xuất thư đã gửi để phân tích bằng AI ngoài
main.py           Ứng dụng desktop
```
