# Hướng dẫn sử dụng kho tri thức

Đây là kho quy tắc cho trợ lý soạn email. Mỗi tệp là **Markdown thuần** — không có
cú pháp đặc biệt nào phải nhớ, cứ viết tiếng Việt bình thường theo gạch đầu dòng.

## Dành cho AI đang sửa các tệp này (Claude Desktop / ChatGPT Desktop / Antigravity)

Khi người dùng yêu cầu bổ sung hoặc chỉnh quy tắc:

- Thêm mục dạng `- ` vào **đúng tệp theo chủ đề**, không gộp tất cả vào một tệp.
- Giữ nguyên tiếng Việt, viết ngắn gọn, mỗi ý một dòng.
- **Không** bọc nội dung trong code fence, không thêm YAML frontmatter.
- **Không** đổi tên tệp và không đổi các tiêu đề `##` đã có trong `03_doi_tac_va_xung_ho.md`
  (backend dựa vào tiêu đề đó để lọc theo người gửi).

## Thứ tự ưu tiên

Số ở đầu tên tệp là thứ tự nạp, cũng là thứ tự ưu tiên khi đưa vào prompt.
Nếu tổng nội dung vượt giới hạn độ dài, backend cắt bỏ **từ tệp số lớn nhất ngược lên** —
nên đừng đặt quy tắc sống còn vào `05` hay `06` nếu kho đang rất dài.

Quy tắc trong kho này **được ưu tiên hơn** hồ sơ văn phong do AI tự học từ hộp thư đã gửi.
Khi hai bên mâu thuẫn, AI phải theo kho này.

## Cách áp dụng

Sửa tệp rồi lưu là xong — backend tự nhận biết theo thời gian sửa tệp, **không cần khởi động lại**.
Thư soạn kế tiếp trong Outlook sẽ tuân theo nội dung mới.
