# HƯỚNG DẪN SỬ DỤNG — TRỢ LÝ EMAIL v2.1.1

> Chưa cài? Xem [HUONG_DAN_CAI_DAT.md](HUONG_DAN_CAI_DAT.md) trước.

Hệ thống gồm ba phần:

| Việc | Làm ở đâu |
|---|---|
| Trả lời thư hằng ngày, đọc tệp đính kèm | **Add-in trong Outlook** |
| Học văn phong, quét phân loại cả thư mục, cài khoá API | **Ứng dụng desktop** (tài liệu này) |
| Viết quy tắc soạn thư | Claude Desktop / ChatGPT Desktop → sửa thư mục `kien_thuc` |

---

## Mở ứng dụng

Ứng dụng chạy ngầm ở khay hệ thống giống UniKey, và đã được đặt **tự chạy khi bạn đăng
nhập Windows**. Bình thường bạn không phải khởi động gì cả.

| Cách | Khi nào dùng |
|---|---|
| **Bấm đúp icon ở khay** (góc phải thanh tác vụ) | Ứng dụng đang chạy sẵn — cách thường dùng |
| **Lối tắt ngoài Desktop** | Bạn đã Thoát hẳn và muốn mở lại |

Lúc chạy ngầm **không có nút ở thanh tác vụ**; mở cửa sổ lên mới có. Bấm **X** là thu về
khay chứ không đóng hẳn.

> **Add-in trong Outlook chỉ chạy khi ứng dụng đang chạy.** Mỗi lần khởi động, ứng dụng tự
> bật sẵn backend phục vụ add-in (mất khoảng 3 giây) nên bạn không phải làm gì thêm.

**Mở Outlook trước rồi mới mở ứng dụng** — ứng dụng đọc thư qua Outlook đang chạy.

---

## 1. Đọc email và điều hướng thư mục

![Giao diện chính của ứng dụng desktop](docs/anh/giao-dien-chinh.png)

- **Danh sách thư mục**: cột trái hiển thị toàn bộ thư mục trong Outlook (Inbox, Sent,
  Drafts và thư mục con). Badge màu tím là **số thư chưa đọc**.
- **Đọc email**: bấm một email ở danh sách giữa để xem chi tiết.
- **Iframe an toàn**: thân thư hiển thị trong iframe sandbox không cho chạy script và
  không cùng origin — email độc hại không chạm được vào ứng dụng.
- **Đánh dấu đã đọc**: bấm xem là email trong Outlook thật chuyển sang trạng thái Read.

---

## 2. Phân loại và tóm tắt hộp thư bằng AI

1. Bấm **✨ Phân tích toàn bộ thư mục** ở phía trên danh sách email.
2. Gemini tóm tắt từng thư thành một câu và gán nhãn:
   - **Cần trả lời** (tím)
   - **Việc cần làm** (cam)
   - **Chỉ để biết** (xám)
   - **Quảng cáo/Rác** (đỏ)
3. Kết quả lưu đệm theo EntryID. Chuyển thư mục hay khởi động lại thì nhãn hiện **tức
   thì**, không tốn lượt gọi AI.
4. Dùng các chip lọc (*Tất cả*, *Cần trả lời*, *Việc cần làm*, *Chỉ để biết*) để lọc nhanh.

---

## 3. Soạn thư trả lời bằng AI

![Bảng điều khiển soạn thư bằng AI](docs/anh/bang-soan-thu-ai.png)

1. Mở email cần phản hồi, bấm **✨ Trả lời bằng AI** (hoặc **✨ Trả lời tất cả**).
2. Bảng điều khiển AI trượt ra từ bên phải.
3. Chọn **Mẫu thư**. Để nguyên *"AI tự chọn theo nội dung thư"* thì hệ thống tự khớp từ
   khoá trong chủ đề và thân thư với các tệp trong `kien_thuc\loai_thu`. Chọn tay khi bạn
   biết rõ mình đang viết loại gì — lựa chọn của bạn **luôn thắng** kết quả tự khớp.
4. Nhập **chỉ dẫn cho AI**, ví dụ *"Đồng ý họp thứ Ba, hỏi thêm tài liệu cần chuẩn bị"*.
5. Bấm **Tạo nháp**. Nội dung được viết theo ngữ cảnh hội thoại + hồ sơ văn phong + kho
   quy tắc trong `kien_thuc`. Thông báo sau khi tạo cho biết đã dùng mẫu nào, và mẫu đó
   do bạn chọn hay AI tự chọn.
6. Muốn sửa thì nhập vào ô *Tinh chỉnh* (ví dụ *"viết trang trọng hơn"*) rồi bấm **Tinh
   chỉnh**, hoặc sửa tay trực tiếp trong ô soạn thảo.
7. Bấm **Lưu vào Thư nháp Outlook**. Thư nháp xuất hiện trong Outlook, giữ nguyên chữ ký
   và trích dẫn email gốc.

---

## 4. Add-in trong Outlook (task pane)

![Task pane của add-in trong Outlook](docs/anh/task-pane-outlook.png)

Đây là chỗ dùng hằng ngày: không cần mở cửa sổ ứng dụng, làm việc thẳng trong Outlook.

- Mở một email rồi bấm **Soạn trả lời AI** trên thanh ribbon.
- Ô **Mẫu thư** ngay đầu task pane: để trống thì AI tự khớp chủ đề email với các quy tắc
  trong `kien_thuc\loai_thu`, hoặc chọn tay một mẫu cụ thể.
- Chèn thẳng nội dung AI soạn vào cửa sổ trả lời của Outlook.
- Đọc được tệp đính kèm (`.xlsx`, `.pdf`, `.docx`).

---

## 5. Học văn phong

![Màn hình học văn phong và cài đặt](docs/anh/hoc-van-phong.png)

Trong ứng dụng: nút **⚙️ Cài đặt** ở góc **dưới cùng bên trái** → **Xuất thư đã gửi để
phân tích**. Hoặc chuột phải icon khay → **Xuất thư đã gửi để phân tích**.

1. Công cụ quét **mọi kho thư** đã gắn vào Outlook — hộp thư chính lẫn các tệp lưu trữ.
   Quét sâu có thể mất **10–30 phút**; cứ để chạy nền, xong sẽ có thông báo.
2. Lọc thư trùng, gom theo chủ đề, rồi ghi ra thư mục `xuat_thu` dưới dạng **cặp "thư đến
   → bạn đã trả lời"**.
3. Đồng thời đếm và ghi hồ sơ văn phong: câu chào, câu kết, chữ ký, cụm từ đặc trưng, độ
   dài thư — tất cả **kèm tần suất thật**.
4. Mở thư mục `xuat_thu` bằng Claude Desktop và ra lệnh *"đọc 00_TONG_QUAN.md rồi làm theo
   đề bài"*. Công cụ sẽ viết hướng dẫn cho từng loại thư vào `kien_thuc\loai_thu`.

> `xuat_thu` chứa **toàn văn thư công việc thật** của bạn. Đừng chia sẻ thư mục này ra
> ngoài.

---

## 6. Cài đặt

**Khoá API Gemini** — nhập trong **⚙️ Cài đặt**. Đây là **nơi duy nhất** nên đổi khoá.

**Model** — đặt qua biến `GEMINI_MODEL` trong tệp `.env` (xem bảng đường dẫn bên dưới).
Mặc định là `gemini-3.1-flash-lite`.

> **Đừng đổi sang các tên kết thúc bằng `-latest`.** Alias đó luôn trỏ tới model mới nhất,
> mà model càng mới hạn mức miễn phí càng thấp — đo thực tế `gemini-flash-latest` chỉ được
> **20 lượt mỗi ngày**, đủ khoảng 6 thư trả lời.

Đổi khoá hoặc model xong thì chuột phải icon khay → **Khởi động lại** để backend nhận
thiết lập mới.

---

## 7. Menu chuột phải trên icon khay

| Mục | Việc |
|---|---|
| **Mở Trợ lý Email** | Hiện cửa sổ ứng dụng (bấm đúp icon cũng vậy) |
| Backend add-in: … | Cho biết đang chạy hay đã dừng |
| Khởi động / Dừng | Bật tắt backend phục vụ add-in |
| Khởi động lại | Dựng lại backend — dùng sau khi đổi khoá API hoặc model |
| Xuất thư đã gửi để phân tích | Chạy phần học văn phong ở mục 5, chạy nền |
| Kiểm tra add-in | Tự chẩn đoán: chứng chỉ, cổng, manifest. Dùng khi Outlook báo lỗi |
| Mở thư mục dữ liệu | Mở Explorer ngay tại `%LOCALAPPDATA%\TroLyEmail` |
| Khởi động cùng Windows | Bật/tắt tự chạy khi đăng nhập. Không cần quyền Admin |
| Tạo lối tắt ngoài Desktop | Dựng lại lối tắt nếu bạn lỡ xoá |
| Thoát | Đóng ứng dụng và dừng backend — **add-in trong Outlook sẽ ngừng hoạt động** |

> Khởi động lại backend thì mã phiên đổi, nên **phải đóng và mở lại task pane** trong
> Outlook.

> Chỉ chạy được một bản cùng lúc. Mở lần thứ hai sẽ báo *"đang chạy sẵn ở khay hệ thống"*
> rồi tự thoát.

---

## Dữ liệu của bạn nằm ở đâu

Tất cả nằm trong `%LOCALAPPDATA%\TroLyEmail` — dán đường dẫn đó vào ô địa chỉ của
Explorer là mở được, hoặc dùng menu khay **Mở thư mục dữ liệu**.

| Thư mục / tệp | Nội dung |
|---|---|
| `.env` | Khoá API và model |
| `kien_thuc\` | Quy tắc soạn thư bạn tự viết |
| `xuat_thu\` | Thư đã gửi xuất ra để phân tích |
| `data\style_profiles\` | Hồ sơ văn phong đã học |
| `data\cache\` | Kết quả phân loại AI lưu đệm |
| `data\logs\` | Nhật ký lỗi |
| `manifest.xml` | Tệp để thêm add-in vào Outlook |
| `certs\` | Chứng chỉ cho backend add-in |

Nâng cấp lên bản mới **không đụng tới** những thứ trên — bộ cài chỉ thay phần chương trình.

---

## Xử lý sự cố

Ứng dụng chạy không có cửa sổ console nên **mọi lỗi ghi vào tệp**:

```
%LOCALAPPDATA%\TroLyEmail\data\logs\app.log            ứng dụng
%LOCALAPPDATA%\TroLyEmail\data\logs\addin_server.log   backend add-in
```

| Hiện tượng | Cách xử lý |
|---|---|
| Không kết nối được Outlook | Mở Outlook trước rồi mở lại ứng dụng |
| *"Sorry, we can't load the add-in… check your network connectivity"* | Ứng dụng chưa chạy, hoặc backend chưa lên. Mở ứng dụng, đợi ~3 giây rồi bấm lại nút trong Outlook. Vẫn lỗi thì menu khay → **Kiểm tra add-in** |
| Không thấy nút trên ribbon Outlook | Chứng chỉ chưa được tin hoặc chưa thêm manifest — xem [HUONG_DAN_CAI_DAT.md](HUONG_DAN_CAI_DAT.md) mục 3 và 4 |
| Task pane trắng trơn | Như trên. Menu khay → **Kiểm tra add-in** để biết thiếu gì |
| *"Phiên không hợp lệ"* trong task pane | Backend vừa khởi động lại nên mã phiên đổi. Đóng và mở lại task pane |
| *"Gọi Gemini quá nhanh (429)"* | Hạn mức gói miễn phí. Chờ khoảng 1 phút |
| *"…không có hạn mức miễn phí cho API key này"* | Model đang dùng đã bị Google ngừng cấp cho khoá mới. Đặt `GEMINI_MODEL=gemini-3.1-flash-lite` trong `.env` |
| Báo hết hạn mức Gemini | Hạn mức tính theo ngày cho từng khoá. Chờ sang hôm sau |
| Không thấy icon khay sau khi đăng nhập | Mở `app.log`, xem dòng cuối |
| Menu khay báo cổng 8765 bị chiếm | Có bản khác đang chạy backend. Đóng nó rồi bấm **Khởi động lại** |

---

## Phụ lục — chạy từ mã nguồn

Chỉ dành cho máy có sẵn Python và mã nguồn dự án; người dùng bản đóng gói bỏ qua mục này.

| Tệp | Việc |
|---|---|
| `CHAY_NGAM.bat` | Chạy ngầm ở khay (giống bản đóng gói) |
| `CHAY_UNGDUNG.bat` | Mở thẳng cửa sổ, không vào khay |
| `CHAY_KEM_DEBUG.bat` | Bật DevTools để xem lỗi giao diện |
| `CHAY_ADDIN_BACKEND.bat` | Chạy riêng backend add-in |
| `XUAT_THU.bat` | Xuất thư đã gửi để phân tích |
| `KIEM_TRA_ADDIN.bat` | Tự chẩn đoán add-in |
| `TAO_CHUNG_CHI_ADDIN.bat` / `CAI_CHUNG_CHI.bat` | Sinh và cài chứng chỉ |
| `DONG_GOI.bat` | Dựng bộ cài `dist\BanGiao\CaiTroLyEmail.exe` |

Chạy từ mã nguồn thì dữ liệu nằm ngay trong thư mục dự án chứ không ở `%LOCALAPPDATA%`.
