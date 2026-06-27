# 📄 Word Template Pro

**Word Template Pro** là ứng dụng giao diện (GUI) hiện đại giúp bạn tạo hàng loạt các tài liệu Word (DOCX) dựa trên biểu mẫu (template) một cách dễ dàng và hoàn toàn tự động. 

Ứng dụng hỗ trợ các loại dữ liệu nhập liệu đa dạng, khả năng nhúng script tính toán tự động bằng Python, và tạo file hàng loạt (batch generate) từ Excel.

---

## 🚀 Tính năng nổi bật
- **Giao diện trực quan (Dark Mode):** Thiết kế tối giản, hiện đại và tập trung vào trải nghiệm người dùng.
- **Tự động nhận diện trường dữ liệu:** Chỉ cần tạo file Word chứa cú pháp `{{ten_truong|type}}`, ứng dụng sẽ tự động sinh form nhập liệu tương ứng.
- **Tính toán tự động (Scripting):** Tích hợp mạnh mẽ với Python, cho phép bạn viết code để tự động tính toán các trường phức tạp (ví dụ: tính thuế, thành tiền, viết số thành chữ).
- **Tạo hàng loạt từ Excel (Batch Generate):** Trích xuất dữ liệu từ hàng ngàn dòng Excel và tạo ra hàng ngàn file Word tương ứng chỉ trong 1 click.
- **Quản lý thông minh:** Quản lý kho mẫu tài liệu (template), ghi nhớ cài đặt thư mục xuất file, và quy tắc đặt tên tự động.

---

## 📦 Cài đặt và Chạy ứng dụng

### Yêu cầu hệ thống
- Python 3.10+
- `uv` (hoặc `pip`) để quản lý thư viện.

### Các bước cài đặt
1. Cài đặt các thư viện cần thiết:
   ```cmd
   uv sync
   ```
2. Chạy ứng dụng:
   ```cmd
   run_app.cmd
   ```
   Hoặc bằng lệnh trực tiếp:
   ```cmd
   uv run python -m word_template_pro
   ```

---

## 📖 Hướng dẫn sử dụng cơ bản

### 1. Chuẩn bị file Word (Template)
Mở file Word (ví dụ: `hop_dong.docx`), tại bất kỳ vị trí nào bạn muốn chèn dữ liệu tự động, hãy gõ theo cú pháp:
```
{{ten_truong|kieu_du_lieu}}
```
Hoặc ngắn gọn hơn (mặc định là kiểu văn bản thuần):
```
{{ten_truong}}
```

**Các kiểu dữ liệu (Type) hỗ trợ:**
| Type | Ý nghĩa | Ví dụ form nhập liệu |
|---|---|---|
| `text` | Đoạn văn bản ngắn 1 dòng | `{{ho_va_ten\|text}}` |
| `m_text` | Đoạn văn bản dài (nhiều dòng) | `{{ghi_chu\|m_text}}` |
| `text_upper` | Chữ in hoa toàn bộ | `{{ho_va_ten\|text_upper}}` |
| `number`, `float`, `integer` | Chỉ cho phép nhập số | `{{don_gia\|number}}` |
| `date` | Ngày tháng năm (có nút chọn lịch) | `{{ngay_sinh\|date}}` |
| `script` | Trường chạy tính toán tự động qua code Python | `{{tong_tien\|script}}` |

### 2. Sử dụng Script (Tính toán tự động)
Nếu bạn có một trường tên là `{{tong_tien|script}}`, hãy tạo một file `.py` có cùng tên với file Word của bạn (ví dụ file Word là `hop_dong.docx`, hãy tạo `hop_dong.py` để cạnh đó).

Bên trong `hop_dong.py`, định nghĩa một hàm cùng tên với tên trường:
```python
def tong_tien(so_luong: int, don_gia: float):
    return str(so_luong * don_gia)
```
Ứng dụng sẽ tự động phân tích và tạo form yêu cầu bạn nhập `so_luong` và `don_gia`, sau đó tự động tính ra `tong_tien`. **(Lưu ý: Bạn hoàn toàn có thể build file `.exe`, tính năng chạy file script rời này vẫn sẽ hoạt động bình thường nhờ cơ chế import linh hoạt).**

### 3. Đặt quy tắc tên file đầu ra
Khi bạn bấm nút "Tạo file", tên file sẽ được tạo ra dựa trên cấu hình trong phần **Quản lý Template**.
Các từ khóa có thể dùng để ghép tên file:
- `{template_name}`: Tên template
- `{date}`: Ngày hiện tại (20260627)
- `{datetime}`: Ngày giờ hiện tại
- `{ten_truong}`: Giá trị của bất kỳ trường nào có trong file Word (ví dụ: `{ho_va_ten}`)

*Ví dụ:* Cấu hình `{template_name}_{ho_va_ten}_{date}` sẽ sinh ra file tên `hop_dong_Nguyen_Van_A_20260627.docx`. (Nếu file đã tồn tại, ứng dụng sẽ tự động thêm đuôi `_1`, `_2` để tránh ghi đè).

### 4. Tạo hàng loạt bằng Excel (Batch Generate)
- Mở template, bấm vào nút **"📊 Batch"**.
- Chọn file Excel của bạn. **Lưu ý quan trọng**: Hàng đầu tiên (Header) của file Excel phải là tên các cột trùng với tên các trường `{{ten_truong}}` trong file Word.
- Nhấn Bắt đầu, và ứng dụng sẽ tạo ra toàn bộ file Word trong chớp mắt!

---
*Phát triển bởi đội ngũ Word Template Pro.*
