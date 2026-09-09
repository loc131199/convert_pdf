# Certification - Excel Certificate Automation

## 1. Giới thiệu

**Certification** là chương trình Python dùng để tự động tạo các file Excel giấy chứng nhận/biên bản hiệu chuẩn từ một file danh mục Excel và các file mẫu Excel có sẵn.

Chương trình được xây dựng nhằm thay thế quá trình xử lý thủ công:

- Đọc danh sách giấy cần tạo từ file `danh_muc.xlsm`.
- Tự động xác định file mẫu tương ứng.
- Truyền dữ liệu từ danh mục vào file mẫu.
- Xử lý các giá trị đặc biệt về ngày tháng.
- Xử lý công thức Random trong Sheet2.
- Giữ nguyên định dạng file Excel của file mẫu.
- Xuất ra các file Excel hoàn chỉnh.
- Tạo báo cáo PDF đối với những dòng không tìm thấy file mẫu.

---

# 2. Nguyên lý hoạt động

Quy trình tổng quát của chương trình:

```text
                    danh_muc.xlsm
                          │
                          ▼
                 Đọc danh sách dữ liệu
                          │
                          ▼
              Lấy [Type_paper] + [Type]
                          │
                          ▼
                Xác định tên file mẫu
                          │
                          ▼
                    Folder /data
                          │
              ┌───────────┴───────────┐
              │                       │
        Tìm thấy file mẫu        Không tìm thấy
              │                       │
              ▼                       ▼
        Mở file mẫu             Ghi nhận lỗi
              │                       │
              ▼                       │
       Thay thế dữ liệu               │
              │                       │
              ▼                       │
       Xử lý Sheet1                   │
              │                       │
              ▼                       │
       Xử lý Sheet2                   │
              │                       │
              ▼                       │
    Random → Value ở Sheet2           │
              │                       │
              ▼                       │
       Lưu file Excel                 │
              │                       │
              └───────────┬───────────┘
                          │
                          ▼
                 Kết quả trong folder
                 Certification
                          │
                          ▼
                 BaoCao_Loi.pdf
             nếu có mẫu bị thiếu
