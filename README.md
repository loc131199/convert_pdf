# Excel & Word to PDF Converter with Background Watermark

Chương trình tự động hóa quy trình chuyển đổi các tệp **Microsoft Excel** và **Microsoft Word** sang định dạng **PDF**, hỗ trợ chèn ảnh nền (background/watermark) theo trang (trang đầu và các trang tiếp theo) một cách chuyên nghiệp.

---

## 📌 Tính năng chính

- **Tự động quét thư mục**: Tự động phát hiện tất cả các file Excel (`.xlsx`, `.xls`, `.xlsm`) và Word (`.doc`, `.docx`) trong cùng thư mục chạy chương trình (hoặc thư mục chứa file `.exe`).
- **Xử lý Excel thông minh**:
  - Export **Sheet 1** thành file PDF tên dạng: `K1 GCN <tên_file_gốc>.pdf` (có hỗ trợ chèn background).
  - Export **Sheet 2** (nếu có) thành file PDF tên dạng: `K1 biên bản <tên_file_gốc>.pdf`.
- **Xử lý Word**: Export tài liệu Word thành file PDF tên dạng `K1 GCN <tên_file_gốc>.pdf` (có hỗ trợ chèn background).
- **Tự động chèn Background/Khung viền**:
  - Trang 1 sử dụng hình nền: `background_page1.jpg`
  - Các trang sau (trang 2 trở đi) sử dụng hình nền: `background_other_pages.jpg`
  - Hình nền được chèn chìm phía dưới nội dung (`overlay=False`), đảm bảo không đè lên văn bản/dữ liệu.
- **Tự động hóa an toàn & mượt mà**:
  - Chạy ẩn ứng dụng MS Office (Excel & Word) trong quá trình xử lý.
  - Tự động tắt cảnh báo, không cập nhật link ngoài, vô hiệu hóa Macro để tránh đơ/treo chương trình.
- **Tương thích linh hoạt**: Chạy trực tiếp dưới dạng script Python (`.py`) hoặc file thực thi (`.exe`).

---

## 🛠️ Yêu cầu hệ thống & Môi trường

1. **Hệ điều hành**: Windows (do chương trình sử dụng Microsoft Office COM Automation).
2. **Phần mềm bắt buộc**: Microsoft Office (Excel và Word) đã được cài đặt trên máy.
3. **Python**: Phiên bản 3.8 trở lên (nếu chạy bằng source code Python).

---

## 📦 Cài đặt thư viện (Dành cho nhà phát triển / Chạy file `.py`)

Cài đặt các thư viện Python cần thiết bằng câu lệnh sau trong Command Prompt / Terminal:

```bash
pip install pywin32 PyMuPDF
```

*Chi tiết các thư viện:*
- `pywin32`: Điều khiển Microsoft Excel và Word thông qua COM API.
- `PyMuPDF` (`fitz`): Thao tác với file PDF (chèn hình nền background vào từng trang).

---

## 📁 Cấu trúc thư mục

Để chương trình hoạt động đúng, hãy đặt các file liên quan chung một thư mục:

```text
📂 Thu_Muc_Chieu_Xy_Ly/
 ├── 📄 main.py (hoặc main.exe)
 ├── 🖼️ background_page1.jpg (Ảnh nền cho trang 1 - Tùy chọn)
 ├── 🖼️ background_other_pages.jpg (Ảnh nền cho các trang từ trang 2 trở đi - Tùy chọn)
 ├── 📊 Bang_ke_01.xlsx
 └── 📝 Van_ban_01.docx
```

---

## 🖼️ Quy tắc sử dụng Background

Chương trình có cơ chế kiểm tra file background thông minh:

| Trường hợp | Trạng thái các file background trong thư mục | Hành vi của chương trình |
| :--- | :--- | :--- |
| **Trường hợp 1** | Có đủ 2 file `background_page1.jpg` và `background_other_pages.jpg` | ✅ Tạo PDF và **chèn background** tương ứng vào từng trang. |
| **Trường hợp 2** | Không có cả 2 file background | ✅ Tạo PDF bình thường **không chèn background**. |
| **Trường hợp 3** | Chỉ có 1 trong 2 file background | ❌ **Báo lỗi** và dừng chương trình (yêu cầu cung cấp đủ cả 2 file hoặc xóa hết). |

---

## 🚀 Hướng dẫn sử dụng

### Cách 1: Chạy bằng file `.exe` (Đã đóng gói)
1. Copy file `.exe` vào thư mục chứa các file Excel/Word cần chuyển đổi.
2. (Tùy chọn) Copy 2 file hình nền `background_page1.jpg` và `background_other_pages.jpg` vào cùng thư mục nếu muốn chèn background.
3. Nhấp đôi chuột vào file `.exe` để khởi chạy.
4. Chờ chương trình xử lý xong và kiểm tra các file PDF được tạo ra trong thư mục.

### Cách 2: Chạy bằng lệnh Python
1. Mở Terminal / Command Prompt tại thư mục chứa dự án.
2. Chạy lệnh:
   ```bash
   python main.py
   ```

---

## 📋 Kết quả đầu ra (Output)

Với file đầu vào là `HopDong_01.xlsx` (chứa 2 sheets) và `CongVan_02.docx`:

- **Từ Excel (`HopDong_01.xlsx`)**:
  - Sheet 1 ➔ `K1 GCN HopDong_01.pdf` (có background nếu kích hoạt)
  - Sheet 2 ➔ `K1 biên bản HopDong_01.pdf`
- **Từ Word (`CongVan_02.docx`)**:
  - Tài liệu ➔ `K1 GCN CongVan_02.pdf` (có background nếu kích hoạt)

---

## ⚠️ Lưu ý quan trọng

1. **Đóng các file đang mở**: Trước khi chạy chương trình, hãy đóng tất cả các file Excel/Word đang mở để tránh xung đột quyền truy cập tệp.
2. **Microsoft Office**: Máy tính phải cài sẵn MS Office. Chương trình không hoạt động trên các bộ ứng dụng thay thế như WPS Office hay LibreOffice nếu không có giao diện COM của MS Office.
3. **Định dạng ảnh nền**: Nên dùng ảnh có tỉ lệ vừa vặn với khổ giấy A4 để background khi chèn không bị méo hoặc hở viền.

---

## 📄 Giấy phép (License)

Dự án này được phát hành dưới giấy phép **MIT License**. Bạn có thể tự do sử dụng, chỉnh sửa và phân phối.
