# Cài Trợ lý Email — hướng dẫn cho người nhận

Bạn nhận được đúng **một tệp**: `CaiTroLyEmail.exe` (khoảng 50 MB). Không cần cài Python,
không cần quyền Admin.

---

## 1. Chạy tệp cài đặt

Bấm đúp `CaiTroLyEmail.exe`.

> **Windows sẽ cảnh báo.** Màn hình xanh *"Windows protected your PC"* hiện lên là bình
> thường: tệp này chưa được ký số nên Windows không biết ai làm ra nó. Bấm
> **More info → Run anyway**. Nếu máy công ty chặn hẳn, nhờ bộ phận CNTT cho phép.

Lần đầu mất khoảng **10 giây** để bung gói vào `%LOCALAPPDATA%\TroLyEmail`. Những lần
sau ứng dụng mở trong khoảng **8 giây**.

## 2. Nhập khoá API Gemini

Hộp thoại sẽ hỏi khoá. Lấy miễn phí tại **aistudio.google.com/apikey** (bấm luôn đường
dẫn trong hộp thoại), đăng nhập bằng tài khoản Google rồi tạo khoá mới.

> Mỗi người nên dùng **khoá riêng**: hạn mức miễn phí tính theo từng khoá, dùng chung
> thì cả nhóm hết lượt cùng lúc.

Bỏ trống cũng được, sau này nhập trong phần **Cài đặt** của ứng dụng.

## 3. Cho phép cài chứng chỉ

Windows sẽ hỏi có tin một chứng chỉ bảo mật hay không — **bấm Yes**.

Đây là chứng chỉ tự ký cho địa chỉ `localhost`, dùng để Outlook nói chuyện với ứng dụng
qua HTTPS ngay trong máy bạn. Không có nó thì nút add-in không hiện và task pane trắng
trơn. Chứng chỉ chỉ nằm trong hồ sơ người dùng của bạn, gỡ được bất cứ lúc nào bằng
`certmgr.msc`.

## 4. Thêm add-in vào Outlook

Cuối phần thiết lập, một cửa sổ Explorer mở ra ở `%LOCALAPPDATA%\TroLyEmail`. Trong đó
có tệp **`manifest.xml`**.

Mở Outlook (bản desktop), rồi:

1. **Get Add-ins** (hoặc **All Apps → Add-ins**)
2. **My add-ins**
3. **Add a custom add-in → Add from file**
4. Chọn `manifest.xml` vừa nói ở trên
5. Xác nhận **Install**

Mở một email bất kỳ, trên thanh ribbon sẽ có nút **Soạn trả lời AI**.

## 5. Xong

Icon Trợ lý Email nằm ở khay hệ thống (góc phải thanh tác vụ, cạnh đồng hồ). Ứng dụng đã
được đặt tự chạy khi bạn đăng nhập Windows, và có lối tắt ngoài Desktop.

| Thao tác | Kết quả |
|---|---|
| Bấm đúp icon khay | Mở cửa sổ ứng dụng |
| Bấm **X** trên cửa sổ | Thu về khay, ứng dụng vẫn chạy |
| Chuột phải icon khay | Menu: bật/tắt backend, tắt tự khởi động, thoát hẳn |

**Add-in trong Outlook chỉ chạy khi Trợ lý Email đang chạy.** Ứng dụng tự bật backend
sẵn cho add-in mỗi lần khởi động (mất khoảng 3 giây), nên bình thường bạn không phải làm
gì. Nhưng nếu bạn đã **Thoát** hẳn ứng dụng ở menu khay thì nút trong Outlook sẽ báo lỗi
không kết nối được — mở lại ứng dụng là hết.

---

## Yêu cầu

- Windows 10 hoặc 11
- **Outlook bản desktop** (không phải Outlook trên web), tài khoản Exchange /
  Microsoft 365 — hộp thư POP/IMAP không nạp được add-in
- Microsoft Edge WebView2 Runtime: Windows 10/11 có sẵn Edge thì đã có. Nếu thiếu, phần
  thiết lập sẽ báo và cửa sổ ứng dụng sẽ trắng trơn.

## Khi có trục trặc

Ứng dụng chạy không có cửa sổ console nên **mọi lỗi ghi vào tệp**:

```
%LOCALAPPDATA%\TroLyEmail\data\logs\app.log            ứng dụng
%LOCALAPPDATA%\TroLyEmail\data\logs\addin_server.log   backend add-in
```

Dán đường dẫn trên vào ô địa chỉ của Explorer là mở được. Xem vài dòng cuối trước khi
hỏi ai.

| Hiện tượng | Cách xử lý |
|---|---|
| Không thấy icon khay sau khi cài | Mở `app.log`, xem dòng cuối |
| Không thấy nút trên ribbon Outlook | Chưa cài chứng chỉ hoặc chưa thêm manifest — làm lại bước 3 và 4 |
| *"Sorry, we can't load the add-in… check your network connectivity"* | Trợ lý Email chưa chạy, hoặc backend chưa lên. Mở ứng dụng, đợi ~3 giây rồi bấm lại nút trong Outlook. Vẫn lỗi thì chuột phải icon khay → **Khởi động lại** |
| Task pane trắng trơn | Chứng chỉ chưa được tin, hoặc backend chưa chạy. Chuột phải icon khay → Khởi động lại backend |
| *"Phiên không hợp lệ"* trong task pane | Backend vừa khởi động lại nên mã phiên đổi. Đóng và mở lại task pane |
| Báo hết hạn mức Gemini | Hạn mức miễn phí tính theo ngày cho từng khoá. Chờ sang hôm sau |

## Nâng cấp lên bản mới

Chạy tệp `CaiTroLyEmail.exe` của bản mới. Nó nhận ra máy đang có bản nào và hỏi bạn có
cài đè không.

**Khoá API, chứng chỉ, kho tri thức và hồ sơ văn phong của bạn được giữ nguyên** — chỉ
phần chương trình bị thay. Sau khi cài đè, đóng và mở lại Outlook để add-in nhận bản mới.

## Gỡ cài đặt

Không có trình gỡ riêng, xoá tay ba thứ:

1. Xoá thư mục `%LOCALAPPDATA%\TroLyEmail`
2. Xoá lối tắt ngoài Desktop và trong `shell:startup`
3. Trong Outlook: **My add-ins** → gỡ **Trợ lý Email**

Muốn sạch hẳn thì mở `certmgr.msc` → **Trusted Root Certification Authorities** →
**Certificates** → xoá mục `localhost`.
