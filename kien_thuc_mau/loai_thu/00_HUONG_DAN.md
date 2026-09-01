# Hướng dẫn viết theo từng loại thư

Thư mục này chứa hướng dẫn soạn thư **riêng cho từng loại việc**. Khi bạn trả lời một
email, backend tự đọc chủ đề và đoạn đầu thư, khớp với dòng **Từ khoá nhận diện** ở mỗi
tệp, rồi nạp đúng hướng dẫn của loại đó vào prompt.

Toàn bộ việc khớp là so sánh chuỗi tại máy — **không tốn lượt gọi AI nào**.

## Cách thêm một loại thư mới

Thả thêm một tệp `.md` vào thư mục này. Không phải sửa code.

Tên tệp (không dấu, không khoảng trắng) chính là tên loại hiển thị trong log, ví dụ
`xlvp.md`, `bao-cao-thang.md`.

## Khuôn bắt buộc

```markdown
# <Tên loại thư đầy đủ>

**Từ khoá nhận diện:** <các từ/cụm cách nhau bởi dấu phẩy>

## Cấu trúc thư
## Cách lập luận
## Câu mẫu
## Tránh
```

Dòng `**Từ khoá nhận diện:**` là dòng **duy nhất** backend đọc bằng máy. Giữ đúng định
dạng đó — đừng đổi thành chú thích HTML hay YAML frontmatter. Phần còn lại viết tự do.

## Quy tắc khớp

- Bỏ dấu trước khi so, nên `thong ke mat bang` vẫn khớp `thống kê mặt bằng`.
- Loại nào trúng nhiều từ khoá nhất thì thắng; hoà thì từ khoá dài hơn thắng vì cụ thể hơn.
- Không loại nào trúng → dùng `khac.md`.

Vì vậy: **từ khoá càng đặc trưng càng tốt**. Đừng đặt từ khoá quá chung như "báo cáo" ở
nhiều loại, chúng sẽ tranh nhau.

## Cách sinh nội dung tự động

Chạy `XUAT_THU.bat` để xuất toàn bộ thư đã gửi ra `xuat_thu/`, rồi mở thư mục đó bằng
Antigravity hoặc Claude Desktop và ra lệnh *"đọc 00_TONG_QUAN.md rồi làm theo đề bài"*.
Công cụ sẽ tự đọc thư thật và ghi các tệp trong thư mục này.
