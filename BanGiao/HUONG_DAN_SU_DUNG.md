# HƯỚNG DẪN SỬ DỤNG — TRỢ LÝ EMAIL v2.1.0

Hệ thống gồm ba phần: ứng dụng desktop, add-in Outlook và kho tri thức quy tắc soạn thư.

| Việc | Làm ở đâu |
|---|---|
| Trả lời thư hằng ngày, đọc tệp đính kèm | Add-in Outlook |
| Học văn phong, quét phân loại cả thư mục, cài API key | **Ứng dụng desktop** (tài liệu này) |
| Viết quy tắc soạn thư | Claude Desktop / ChatGPT Desktop / Antigravity → sửa `kien_thuc/` |

Khởi chạy: mở Outlook trước, rồi chạy `CHAY_UNGDUNG.bat`.

Dùng hằng ngày thì nên chạy **`CHAY_NGAM.bat`** — gộp cả ứng dụng lẫn backend add-in vào
một icon ở khay hệ thống giống UniKey, xem mục [7. Chạy ngầm ở khay hệ thống](#7-chạy-ngầm-ở-khay-hệ-thống).

---

## 1. Đọc email và điều hướng thư mục

![Giao diện chính Ứng dụng Desktop Trợ lý Email AI](file:///C:/Users/tuanna13/.gemini/antigravity-ide/brain/b21be490-1aa4-4f61-ab4d-6acb01e42e49/desktop_app_interface_1785919961230.png)

- **Danh sách thư mục**: cột trái hiển thị toàn bộ thư mục trong Outlook (Inbox, Sent, Drafts và thư mục con). Badge màu tím là **số thư chưa đọc**.
- **Đọc email**: bấm một email ở danh sách giữa để xem chi tiết.
- **Iframe an toàn**: thân thư hiển thị trong iframe sandbox không cho chạy script và không cùng origin — email độc hại không chạm được vào ứng dụng.
- **Đánh dấu đã đọc**: bấm xem là email trong Outlook thật chuyển sang trạng thái Read.

---

## 2. Phân loại và tóm tắt hộp thư bằng AI

1. Bấm **✨ Phân tích toàn bộ thư mục** ở phía trên danh sách email.
2. Gemini tóm tắt từng thư thành một câu và gán nhãn:
   - **Cần trả lời** (tím)
   - **Việc cần làm** (cam)
   - **Chỉ để biết** (xám)
   - **Quảng cáo/Rác** (đỏ)
3. Kết quả lưu đệm ở `data/cache/classifications.json` theo EntryID. Chuyển thư mục hay khởi động lại thì nhãn hiện **tức thì**, không tốn lượt gọi AI.
4. Dùng các chip lọc (*Tất cả*, *Cần trả lời*, *Việc cần làm*, *Chỉ để biết*) để lọc nhanh.

---

## 3. Soạn thư trả lời bằng AI

![Bảng điều khiển AI Soạn thư trả lời](file:///C:/Users/tuanna13/.gemini/antigravity-ide/brain/b21be490-1aa4-4f61-ab4d-6acb01e42e49/ai_reply_panel_1785919976556.png)

1. Mở email cần phản hồi, bấm **✨ Trả lời bằng AI** (hoặc **✨ Trả lời tất cả**).
2. Bảng điều khiển AI trượt ra từ bên phải.
3. Chọn **Mẫu thư**. Để nguyên *"AI tự chọn theo nội dung thư"* thì hệ thống tự khớp từ khoá trong chủ đề và thân thư với các tệp trong `kien_thuc/loai_thu/`. Chọn tay khi bạn biết rõ mình đang viết loại gì — lựa chọn của bạn **luôn thắng** kết quả tự khớp.
4. Nhập **chỉ dẫn cho AI**, ví dụ *"Đồng ý họp thứ Ba, hỏi thêm tài liệu cần chuẩn bị"*.
5. Bấm **Tạo nháp**. Nội dung được viết theo ngữ cảnh hội thoại + hồ sơ văn phong + **kho quy tắc trong `kien_thuc/`**. Thông báo sau khi tạo cho biết đã dùng mẫu nào, và mẫu đó do bạn chọn hay AI tự chọn.
6. Muốn sửa thì nhập vào ô *Tinh chỉnh* (ví dụ *"viết trang trọng hơn"*) và bấm **Tinh chỉnh**, hoặc sửa tay trực tiếp trong ô soạn thảo.
7. Bấm **Lưu vào Thư nháp Outlook**. Thư nháp xuất hiện trong Outlook, giữ nguyên chữ ký và trích dẫn email gốc.

---

## 4. Tích hợp Outlook Add-in (Task Pane)

![Giao diện Outlook Add-in Task Pane](file:///C:/Users/tuanna13/.gemini/antigravity-ide/brain/b21be490-1aa4-4f61-ab4d-6acb01e42e49/outlook_addin_taskpane_1785919995205.png)

- Ô **Mẫu thư** ngay đầu task pane: để trống thì AI tự khớp chủ đề email với các quy tắc trong `kien_thuc/loai_thu/`, hoặc chọn tay một mẫu cụ thể.
- Hỗ trợ chèn trực tiếp nội dung soạn bởi AI vào cửa sổ trả lời của Outlook.
- Đọc trực tiếp các tệp đính kèm (`.xlsx`, `.pdf`, `.docx`).

---

## 5. Học văn phong & Cài đặt

![Màn hình Học văn phong & Cài đặt](file:///C:/Users/tuanna13/.gemini/antigravity-ide/brain/b21be490-1aa4-4f61-ab4d-6acb01e42e49/style_analysis_settings_1785920007595.png)

Nút nằm ở góc **dưới cùng bên trái**: **⚙️ Cài đặt** → **Xuất thư đã gửi để phân tích**.
Hoặc chạy thẳng `XUAT_THU.bat`, hoặc dùng mục tương ứng trong task pane của add-in.

1. Công cụ quét **mọi kho thư** đã gắn vào Outlook — hộp thư chính lẫn các tệp lưu trữ.
2. Lọc thư trùng, gom theo chủ đề, rồi ghi ra thư mục `xuat_thu/` dưới dạng **cặp "thư đến → bạn đã trả lời"**.
3. Đồng thời đếm và ghi vào `data/style_profiles/default.json`: câu chào, câu kết, chữ ký, cụm từ đặc trưng, độ dài thư — tất cả **kèm tần suất thật**.
4. Mở `xuat_thu/` bằng Antigravity hoặc Claude Desktop và ra lệnh *"đọc 00_TONG_QUAN.md rồi làm theo đề bài"*. Công cụ sẽ viết hướng dẫn cho từng loại thư vào `kien_thuc/loai_thu/`.

---

## 6. Cài đặt hệ thống

- **API key Gemini**: nhập trong ⚙️ Cài đặt, ứng dụng ghi vào tệp `.env`.
  Đây là **nơi duy nhất** nên đổi khoá. Đổi xong nhớ khởi động lại `CHAY_ADDIN_BACKEND.bat`.
- **Model**: đặt qua biến `GEMINI_MODEL` trong `.env`. Giữ mặc định `gemini-flash-latest`.

---

## 7. Chạy ngầm ở khay hệ thống

Chạy `CHAY_NGAM.bat`. Thay vì hai cửa sổ đen, bạn được **một icon ở góc phải thanh tác vụ**
giống UniKey.

- **Bấm đúp icon** → hiện cửa sổ ứng dụng.
- **Bấm X** → thu về khay, ứng dụng vẫn chạy và backend add-in vẫn phục vụ Outlook.
  Muốn đóng hẳn thì chọn **Thoát** trong menu chuột phải.
- Lúc chạy ngầm **không có nút ở taskbar**; mở cửa sổ lên mới có.

Menu chuột phải trên icon:

| Mục | Việc |
|---|---|
| **Mở Trợ lý Email** | Hiện cửa sổ |
| Backend add-in: … | Cho biết đang chạy hay đã dừng, kèm nút Khởi động / Dừng / Khởi động lại |
| Khởi động cùng Windows | Bật/tắt tự chạy khi đăng nhập. Không cần quyền Admin — chỉ là một lối tắt trong `shell:startup`, bạn xoá tay lúc nào cũng được |
| Thoát | Đóng ứng dụng và dừng luôn backend |

> Khởi động lại backend thì token đổi, nên **phải đóng và mở lại task pane** trong Outlook.

> Chỉ chạy được một bản cùng lúc. Mở lần thứ hai sẽ báo *"đang chạy sẵn ở khay hệ thống"*
> rồi tự thoát.

---

## Xử lý sự cố

| Hiện tượng | Cách xử lý |
|---|---|
| Không kết nối được Outlook | Mở Outlook trước rồi mới chạy ứng dụng. |
| *"Gọi Gemini quá nhanh (429)"* | Hạn mức gói miễn phí. Chờ khoảng 1 phút. |
| *"...không có hạn mức miễn phí cho API key này"* | Đặt `GEMINI_MODEL=gemini-flash-latest` trong `.env`. |
| Cần xem lỗi chi tiết | Chạy `CHAY_KEM_DEBUG.bat` để bật DevTools. |
| Chạy ngầm mà không thấy icon khay | `pythonw.exe` không có console nên lỗi chỉ nằm trong `data/logs/app.log`. Mở tệp đó xem dòng cuối. |
| Menu khay báo cổng 8765 bị chiếm | Có `CHAY_ADDIN_BACKEND.bat` đang mở. Đóng cửa sổ đó rồi bấm Khởi động lại. |
| Backend không lên | Xem `data/logs/addin_server.log`. |
