# Hướng dẫn kho tri thức

Kho tri thức là nơi bạn dạy trợ lý **cách viết thư của riêng bạn** — thứ mà AI không thể
tự suy ra từ hộp thư: quy tắc nội bộ, cách xưng hô với từng người, điều cấm tuyệt đối,
và **cách triển khai nội dung cho từng loại thư**.

Kho gồm hai phần:

| Phần | Nội dung | Ai viết |
|---|---|---|
| `kien_thuc/*.md` | Quy tắc chung: hồ sơ cá nhân, giọng văn, xưng hô, thuật ngữ, điều cấm | Bạn, hoặc công cụ AI |
| `kien_thuc/loai_thu/*.md` | Cách viết riêng cho từng loại thư | Antigravity, sau khi đọc thư thật |

Nó bổ sung cho **hồ sơ văn phong** (số liệu đếm được từ thư đã gửi). Khi hai bên
mâu thuẫn, **kho tri thức luôn thắng**.

---

## Quy trình sinh hướng dẫn từ thư thật (khuyến nghị)

Đây là cách hiệu quả nhất, và **không tốn hạn mức Gemini** vì Antigravity/Claude Desktop
chạy bằng gói thuê bao riêng của chúng.

### Bước 1 — Xuất thư đã gửi

```
XUAT_THU.bat
```

Quét **mọi kho thư** đã gắn vào Outlook (hộp thư chính lẫn các tệp lưu trữ) và xuất ra
`xuat_thu/`. Mỗi thư trình bày thành **cặp "thư đến → bạn đã trả lời"** — phần thư đến
lấy từ chính trích dẫn nằm sẵn trong thư bạn gửi, nên không tốn thêm lần đọc nào.

Thêm `--nhanh` nếu chỉ muốn quét các thư mục kiểu Sent Items (nhanh hơn nhiều).

Kết quả gồm:
- `00_TONG_QUAN.md` — **đề bài viết sẵn cho AI** + bảng số liệu đã đếm chính xác
  (câu chào, câu kết, chữ ký, cụm từ, độ dài — tất cả kèm tần suất thật)
- Các tệp nhóm, đánh số theo thứ tự đọc, và `99_thu_ngan.md`

Thư được gom theo **mục đích viết** chứ không theo từ đầu chủ đề: *Yêu cầu rà soát &
đối chiếu dữ liệu*, *Xác minh nghi vấn gian lận*, *Xử lý vi phạm & kỷ luật*, *Báo cáo
& tổng hợp kết quả*, *Thẩm định mặt bằng & mở điểm mới*… Cùng một mục đích thì cách
viết giống nhau — đó mới là thứ cần gom lại để học văn phong.

Muốn thêm, bớt hay đổi tên nhóm: sửa bảng `TOPIC_RULES` ở đầu
[backend/style_stats.py](backend/style_stats.py) rồi chạy lại `XUAT_THU.bat`. Mỗi dòng
là một nhóm kèm danh sách từ khoá; **khớp từ trên xuống, nhóm đầu tiên trúng thì thắng**,
nên nhóm đặc thù phải xếp trước nhóm chung chung. Từ khoá viết không dấu, chữ thường —
so khớp cũng bỏ dấu nên thư gõ có dấu hay không đều nhận ra. Thư không khớp nhóm nào rơi
vào `Chủ đề khác`.

> `xuat_thu/` chứa toàn văn thư công việc thật. Đã nằm trong `.gitignore`.
> Đừng chia sẻ ra ngoài.

### Bước 2 — Để Antigravity đọc và viết hướng dẫn

Mở thư mục `xuat_thu/` bằng Antigravity hoặc Claude Desktop, rồi ra lệnh:

> đọc 00_TONG_QUAN.md rồi làm theo đề bài

Đề bài đã viết sẵn trong tệp đó: yêu cầu công cụ đọc các tệp nhóm, lấy các nhóm mục đích
làm điểm xuất phát cho bộ loại thư (được phép tách hoặc gộp nếu đọc thấy lối viết khác
nhau), rồi ghi ra `kien_thuc/loai_thu/<tên-loại>.md` theo đúng khuôn backend đọc được.

### Bước 3 — Dùng ngay

Không cần khởi động lại gì. Thư kế tiếp bạn soạn trong Outlook sẽ tự khớp loại và nạp
đúng hướng dẫn.

---

## Hướng dẫn theo từng loại thư

Backend đọc chủ đề và đoạn đầu thư đến, so với dòng **Từ khoá nhận diện** ở mỗi tệp
trong `kien_thuc/loai_thu/`, rồi nạp hướng dẫn của loại khớp nhất vào prompt.
Toàn bộ là so khớp chuỗi tại máy — **không tốn lượt gọi AI nào**.

Ô **Mẫu thư** ở bảng soạn thư (cả ứng dụng desktop lẫn task pane) liệt kê mọi tệp trong
thư mục này; nhãn hiển thị lấy từ dòng tiêu đề `# ...` đầu tệp. Để nguyên *"AI tự chọn
theo nội dung thư"* thì chạy đúng cơ chế khớp từ khoá nói trên. Chọn tay một mẫu thì
**lựa chọn của bạn thắng**, kể cả khi từ khoá chỉ sang loại khác — hữu ích khi chủ đề
thư đến không nói lên bạn định viết gì. Sau khi tạo nháp, giao diện cho biết đã dùng mẫu
nào và mẫu đó do bạn chọn hay AI tự chọn.

Thêm hay xoá tệp trong thư mục là danh sách tự cập nhật, không cần khởi động lại backend
— nhưng task pane đang mở phải mở lại vì nó chỉ nạp danh sách một lần.

```markdown
# Báo cáo xử lý vi phạm

**Từ khoá nhận diện:** xử lý vi phạm, không tuân thủ, biên bản vi phạm

## Cấu trúc thư
## Cách lập luận
## Câu mẫu
## Tránh
```

Quy tắc khớp:

- Bỏ dấu trước khi so, nên `xu ly vi pham` vẫn khớp `xử lý vi phạm`.
- Loại nào trúng nhiều từ khoá nhất thì thắng; hoà thì từ khoá dài hơn thắng.
- Không loại nào trúng → dùng `khac.md`.
- Tệp có tên bắt đầu bằng số (`00_HUONG_DAN.md`) là tài liệu, không phải một loại thư.

Thêm loại mới = thả thêm một tệp. Không phải sửa code.

---

---

## Khởi tạo

```powershell
Copy-Item -Recurse kien_thuc_mau kien_thuc
```

Thư mục `kien_thuc/` nằm trong `.gitignore` vì chứa thông tin cá nhân và nội bộ công ty.
Bản mẫu `kien_thuc_mau/` chỉ chứa dữ liệu giả và được commit để làm khuôn.

---

## Bảy tệp

| Tệp | Nội dung |
|---|---|
| `00_HUONG_DAN.md` | Hướng dẫn dành cho chính AI khi nó sửa các tệp này |
| `01_ho_so_ca_nhan.md` | Tên, chức danh, phòng ban, công ty, khối chữ ký, cách tự xưng |
| `02_quy_tac_tra_loi.md` | Độ dài, cấu trúc, giọng văn, khi nào cần thận trọng |
| `03_doi_tac_va_xung_ho.md` | Xưng hô theo từng người / từng tên miền |
| `04_thuat_ngu_va_viet_tat.md` | Từ viết tắt nội bộ để AI không bịa nghĩa |
| `05_mau_cau_thuong_dung.md` | Mẫu câu theo tình huống |
| `06_khong_duoc_lam.md` | Cấm tuyệt đối |

Số ở đầu tên tệp là **thứ tự ưu tiên**. Nếu tổng nội dung vượt 12 000 ký tự, backend
cắt bỏ từ tệp số lớn nhất ngược lên — nên đừng đặt quy tắc sống còn vào `05` hay `06`
nếu kho của bạn đang rất dài.

---

## Sửa bằng Claude Desktop / ChatGPT Desktop / Antigravity

Đây là cách dùng chính. Mở thư mục `kien_thuc/` bằng công cụ AI trên máy rồi ra lệnh
bằng tiếng Việt, ví dụ:

> Đọc thư mục kien_thuc/ rồi bổ sung quy tắc: với anh Hùng bên phòng Kế toán,
> luôn CC chị Lan và luôn nhắc số chứng từ trong thư.

Tệp `00_HUONG_DAN.md` chứa sẵn chỉ dẫn cho AI về cách sửa (thêm vào đúng tệp theo chủ đề,
giữ tiếng Việt, không bọc code fence, không đổi tiêu đề `##`), nên chỉ cần bảo nó đọc
thư mục trước là được.

**Sửa xong lưu là xong.** Backend theo dõi thời gian sửa tệp và tự nạp lại —
**không cần khởi động lại** `CHAY_ADDIN_BACKEND.bat`. Thư soạn kế tiếp trong Outlook
đã tuân theo nội dung mới.

---

## Quy ước riêng của `03_doi_tac_va_xung_ho.md`

Tệp này phình to theo số đối tác, nên backend chỉ nạp phần liên quan:

- Mục `## Chung` — **luôn** được nạp.
- Mục `## <chuỗi>` — chỉ nạp khi chuỗi đó nằm trong địa chỉ email người gửi.

```markdown
## Chung
- Mặc định xưng "em", gọi "anh/chị".

## @doitac.example
- Luôn viết tiếng Anh với nhóm này.

## sep@congty-mau.example
- Xưng "em", gọi "anh". Luôn nêu mốc thời gian cụ thể.
```

Nhờ vậy bạn có thể liệt kê hàng trăm đối tác mà prompt gửi cho Gemini vẫn gọn.
Đây là lý do **không được đổi các tiêu đề `##` trong tệp này** thành dạng khác.

---

## Kiểm tra kho có được nạp không

Trong cửa sổ `CHAY_ADDIN_BACKEND.bat`, dòng khởi động in ra số tệp và số ký tự:

```
Kho tri thuc: 7 tep, 6398 ky tu (co)
```

Nếu thấy `THIEU THU MUC kien_thuc/` nghĩa là bạn chưa sao chép từ `kien_thuc_mau/`.

---

## Cách kiểm chứng quy tắc thật sự có tác dụng

1. Thêm vào `06_khong_duoc_lam.md` một dòng không thể nhầm lẫn, ví dụ
   `- Tuyệt đối không dùng từ "trân trọng".`
2. Soạn một thư trả lời bất kỳ — kết quả không được chứa từ đó.
3. Xoá dòng vừa thêm, lưu lại, soạn tiếp **mà không khởi động lại gì**.
   Từ "trân trọng" quay lại là đúng.
