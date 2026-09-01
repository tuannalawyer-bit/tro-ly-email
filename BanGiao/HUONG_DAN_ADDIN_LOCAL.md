# Hướng dẫn cài Add-in Outlook (chạy cục bộ)

Add-in này chạy **hoàn toàn trên máy bạn**: không dùng Microsoft Graph, không dùng
Entra ID, không đăng ký ứng dụng đám mây. Backend Python chỉ lắng nghe `127.0.0.1`.

---

## Điều kiện tiên quyết

Kiểm tra hai điều này **trước khi** làm gì khác — nếu một trong hai không đạt thì
add-in không thể cài được và bạn phải dùng ứng dụng desktop thay thế.

1. **Tài khoản phải là Exchange hoặc Microsoft 365.**
   Outlook → **File → Account Settings → Account Settings**, xem cột *Type*.
   Hộp thư POP/IMAP hoặc chỉ có tệp PST **không nạp được add-in**.

2. **Tổ chức phải cho phép cài add-in tùy chỉnh.**
   Outlook → tab **Home → Get Add-ins → My add-ins**, kéo xuống cuối tìm mục
   **Custom Addins**. Nếu không thấy dòng *"Add a custom add-in"* thì quản trị viên
   Exchange đã chặn sideload.

---

## Cài đặt

### Bước 1 — Tạo chứng chỉ (chỉ làm một lần)

```
TAO_CHUNG_CHI_ADDIN.bat
```

Sinh ra `certs/localhost.crt` và `certs/localhost.key`, hạn 825 ngày.

### Bước 2 — Cài chứng chỉ vào Trusted Root

```
CAI_CHUNG_CHI.bat
```

**Đây là bước quan trọng nhất và cũng hay bị bỏ sót nhất.** Task pane chạy trong
WebView2, mà WebView2 tin kho chứng chỉ của Windows. Chứng chỉ chưa được tin cậy sẽ gây
ra **hai** triệu chứng cùng lúc:

- **Nút "Soạn trả lời AI" không xuất hiện trên ribbon** — Outlook không tải nổi icon
  của nút từ `https://localhost:8765`.
- **Task pane hiện trắng trơn**, không báo lỗi, không có nút bỏ qua.

Script tự dọn các chứng chỉ localhost cũ, cài bản mới, rồi **kiểm chứng bằng một kết nối
HTTPS thật** để xác nhận Windows đã thực sự tin. Không cần quyền Administrator. Nếu
Windows hiện hộp thoại cảnh báo bảo mật màu đỏ, bấm **Yes**.

> Sinh lại chứng chỉ (`tao_chung_chi_addin.py --force`) sẽ tạo khoá và thumbprint mới,
> làm mất hiệu lực bản đã cài. **Sinh lại thì phải chạy lại `CAI_CHUNG_CHI.bat`.**

Nếu muốn cài tay: nháy đúp `certs/localhost.crt` → *Install Certificate* → **Current
User** → *Place all certificates in the following store* → **Browse** → **Trusted Root
Certification Authorities** → OK → Next → Finish → Yes. Lỗi hay gặp khi cài tay là chọn
nhầm kho **Personal** thay vì **Trusted Root**.

### Bước 3 — Chạy backend

```
CHAY_ADDIN_BACKEND.bat
```

Để cửa sổ này mở suốt thời gian dùng add-in. Màn hình sẽ hiện:

```
================================================================
  Tro ly Email - backend add-in v1.4.0
  Dia chi     : https://localhost:8765
  Model       : gemini-flash-latest
  API key     : da cau hinh
  Kho tri thuc: 7 tep, 6398 ky tu (co)
  Van phong   : da hoc, 159 mau
----------------------------------------------------------------
  Luu y: moi lan khoi dong lai, token doi -> phai dong va mo lai task pane.
================================================================
```

Nếu báo `[LOI] Thieu certs/...` thì quay lại Bước 1.

### Bước 4 — Kiểm tra trước khi cài

```
KIEM_TRA_ADDIN.bat
```

Kiểm tra 18 điều kiện (chứng chỉ, kho Trusted Root, backend, icon ribbon, cấu trúc
manifest, API key) và in kèm câu lệnh sửa cho từng mục không đạt. **Chỉ đi tiếp khi báo
`MOI DIEU KIEN BAT BUOC DEU DAT`** — cài manifest lúc điều kiện chưa đủ sẽ khiến Outlook
cache trạng thái hỏng và bạn phải làm lại toàn bộ quy trình xoá cache.

Đây cũng là việc **đầu tiên** cần làm mỗi khi add-in trục trặc.

### Bước 5 — Cài manifest vào Outlook

Outlook → **Home → Get Add-ins → My add-ins** → kéo xuống **Custom Addins** →
**Add a custom add-in → Add from file...** → chọn `addin/manifest.xml` → **Install**.

### Bước 6 — Dùng thử

Mở một email bất kỳ. Trên ribbon tab **Home** sẽ có nhóm **Trợ lý Email** với nút
**Soạn trả lời AI**. Bấm vào đó để mở task pane.

---

## Sử dụng

1. Task pane tự đọc thư đang mở, hiện chủ đề, người gửi và danh sách tệp đính kèm.
2. Nhập **chỉ dẫn cho AI** (không bắt buộc), ví dụ *"xác nhận đã nhận, hẹn phản hồi trước thứ Sáu"*.
3. Để nguyên ô **Đọc cả tệp đính kèm** nếu muốn AI đọc nội dung file kèm theo.
   PDF và ảnh được gửi thẳng cho Gemini (đọc được cả PDF scan); `.xlsx` và `.docx`
   được trích thành văn bản ngay trên máy.
4. Bấm **Soạn thư trả lời**. Có tệp đính kèm thì mất khoảng 1–2 phút.
5. Sửa trực tiếp trong ô bản nháp, hoặc nhập yêu cầu vào ô *Yêu cầu chỉnh sửa* rồi
   bấm **Chỉnh sửa lại bản nháp**.
6. Bấm **Chèn vào thư trả lời** (hoặc **Trả lời tất cả**) — Outlook mở cửa sổ trả lời
   đã điền sẵn nội dung, nằm phía trên phần trích dẫn gốc. Kiểm tra rồi bấm **Gửi**.

> Trong Outlook classic, cửa sổ trả lời thường mở ngay trong khung đọc và có thể
> đóng task pane. Đó là bình thường. Nếu cần giữ lại nội dung, dùng nút **Sao chép HTML**.

### Học văn phong ngay trong task pane

Mở mục **"Học văn phong từ thư đã gửi"** ở cuối task pane. Dòng đầu cho biết đã học
chưa và từ bao nhiêu thư mẫu.

Bấm **Bắt đầu học văn phong** để AI quét toàn bộ thư mục Sent Items và học cách bạn
viết. Việc này:

- Chạy **vài phút** — có thanh tiến độ theo từng lô nên bạn biết nó vẫn đang chạy.
- Chạy ở **luồng nền**, không làm treo task pane. Đóng pane rồi mở lại vẫn bắt được
  tiến độ đang chạy dở.
- **Cần Outlook đang mở**, vì nó đọc hộp thư qua Outlook COM.
- Chỉ cần làm lại khi văn phong của bạn thay đổi, không phải làm thường xuyên.

> Đây là thao tác **duy nhất** của add-in chạm tới Outlook COM. Vì backend add-in là
> tiến trình riêng với ứng dụng desktop, lần đầu chạy Outlook có thể hiện hộp thoại
> xin quyền truy cập tự động — chọn cho phép.

Cũng có thể làm việc này ở ứng dụng desktop: **⚙️ Cài đặt → Phân tích văn phong**.

---

## Xử lý sự cố

**Việc đầu tiên luôn là chạy `KIEM_TRA_ADDIN.bat`.** Bảng dưới chỉ dùng khi công cụ đó
đã báo mọi mục đạt mà vẫn có vấn đề.

| Hiện tượng | Nguyên nhân và cách xử lý |
|---|---|
| **Không thấy nút "Soạn trả lời AI" trên ribbon** | Hai nguyên nhân, thường xảy ra cùng lúc. (1) Chứng chỉ chưa được tin cậy nên Outlook không tải nổi icon của nút — chạy `CAI_CHUNG_CHI.bat`. (2) Outlook còn cache manifest cũ — xem mục *Cập nhật manifest* bên dưới. Nếu thấy nút của add-in khác (ví dụ *My Templates* của Microsoft) thì vùng ribbon vẫn tốt, vấn đề nằm ở riêng add-in này. |
| Task pane **trắng trơn** | Chứng chỉ chưa nằm trong Trusted Root — chạy `CAI_CHUNG_CHI.bat`. Đây là nguyên nhân của khoảng 90% ca. |
| Vừa chạy `tao_chung_chi_addin.py --force` xong thì hỏng | Sinh lại chứng chỉ tạo thumbprint mới, tin cậy cũ hết hiệu lực. Chạy `CAI_CHUNG_CHI.bat` rồi khởi động lại `CHAY_ADDIN_BACKEND.bat`. |
| *"Backend đã khởi động lại. Hãy đóng và mở lại task pane."* | Token đổi mỗi lần backend khởi động lại. Đóng task pane rồi bấm lại nút ribbon. |
| *"Không lấy được mã phiên"* | Trang đang mở trực tiếp từ đĩa. Phải mở qua nút ribbon trong Outlook, không mở tệp `taskpane.html` bằng trình duyệt. |
| *"Gọi Gemini quá nhanh (429)"* | Hạn mức gói miễn phí. Chờ khoảng 1 phút rồi thử lại. |
| *"...không có hạn mức miễn phí cho API key này"* | Đặt `GEMINI_MODEL=gemini-flash-latest` trong `.env`. |
| Add-in báo lỗi kết nối | Cửa sổ `CHAY_ADDIN_BACKEND.bat` đã bị đóng. Mở lại rồi mở lại task pane. |
| Tệp đính kèm bị bỏ qua | Task pane liệt kê lý do cụ thể trong dải ghi chú màu vàng (quá lớn, sai định dạng, chạm giới hạn tổng). |

### Cập nhật manifest

Outlook cache manifest rất dai, đặc biệt là nút ribbon — nó **không tự làm mới**
khi bạn cài đè. Mỗi lần sửa `addin/manifest.xml` phải làm đủ bốn bước:

1. **Get Add-ins → My add-ins**, gỡ add-in cũ.
2. **Đóng hẳn Outlook** (kiểm tra Task Manager không còn `OUTLOOK.EXE`).
3. Xoá sạch thư mục `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.
4. Mở lại Outlook và cài lại manifest.

---

## Phân vai giữa add-in và ứng dụng desktop

| Việc | Làm ở đâu |
|---|---|
| Trả lời thư đang đọc | **Add-in** |
| Đọc tệp đính kèm để soạn thư | **Add-in** |
| Học văn phong từ hộp thư đã gửi | **Desktop** (`CHAY_UNGDUNG.bat`) |
| Quét phân loại toàn bộ thư mục | **Desktop** |
| Nhập / đổi API key | **Desktop** → ⚙️ Cài đặt |
| Sửa quy tắc viết thư | Claude Desktop / ChatGPT Desktop, sửa thẳng `kien_thuc/` |

Add-in cố ý **không** đụng tới Outlook COM: mọi dữ liệu nó cần đều lấy qua Office.js.
Nhờ vậy nó không mở kết nối Outlook thứ hai và không phải chờ các tác vụ quét dài
của ứng dụng desktop.

Riêng việc đổi API key chỉ làm ở desktop để tránh hai tiến trình cùng ghi tệp `.env`.
Sau khi đổi khoá, khởi động lại `CHAY_ADDIN_BACKEND.bat` để add-in nhận khoá mới.

---

## Quyền mà add-in yêu cầu

Manifest khai `<Permissions>ReadItem</Permissions>` — mức thấp nhất đủ dùng, cho phép:
đọc nội dung/chủ đề/người gửi thư đang mở, liệt kê và đọc tệp đính kèm, và mở form
trả lời có sẵn nội dung.

Bản v1.3 trước đây khai `ReadWriteMailbox`. Quyền đó mở luôn `makeEwsRequestAsync`,
tức là gửi lệnh EWS tuỳ ý lên **toàn bộ hộp thư** — rộng hơn nhu cầu rất nhiều và
là thứ khiến quản trị viên Exchange có lý do chặn add-in.
