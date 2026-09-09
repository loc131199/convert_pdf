import os
import fitz
import win32com.client
import sys


if getattr(sys, 'frozen', False):
    # chạy dưới dạng exe
    folder = os.path.dirname(sys.executable)
else:
    # chạy bằng file .py
    folder = os.path.dirname(os.path.abspath(__file__))

print("Folder đang quét:", folder)

# KIỂM TRA BACKGROUND

background_page1 = os.path.join(folder, "background_page1.jpg")
background_other_pages = os.path.join(folder, "background_other_pages.jpg")

# KIỂM TRA BACKGROUND

has_background_page1 = os.path.exists(background_page1)
has_background_other = os.path.exists(background_other_pages)

# TRƯỜNG HỢP 1:
# Chỉ có 1 trong 2 file
# → BÁO LỖI VÀ DỪNG

if has_background_page1 != has_background_other:

    print()
    print("=" * 60)
    print("LỖI BACKGROUND")
    print("=" * 60)
    print("Phải có ĐỦ CẢ 2 file:")
    print("  - background_page1.jpg")
    print("  - background_other_pages.jpg")
    print()
    print("Hoặc không có cả 2 file.")
    print("=" * 60)

    input("Nhấn Enter để thoát...")
    sys.exit(1)

# TRƯỜNG HỢP 2:
# Có ĐỦ cả 2 file
# → Sử dụng background

if has_background_page1 and has_background_other:

    use_background = True

    print()
    print("Đã tìm thấy đầy đủ 2 file background.")
    print("Sẽ chèn background vào PDF.")


# TRƯỜNG HỢP 3:
# Không có cả 2 file
# → Không sử dụng background

else:

    use_background = False

    print()
    print("Không tìm thấy background.")
    print("PDF sẽ được tạo mà không chèn background.")

excel_extensions = (".xls", ".xlsx", ".xlsm")
word_extensions = (".doc", ".docx")


# Khởi động Excel

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
# Không hỏi Update Links
excel.AskToUpdateLinks = False

# Không chạy macro
excel.AutomationSecurity = 3

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = 0
try:

    for filename in os.listdir(folder):

        if filename.lower().endswith(excel_extensions) and not filename.startswith("~$"):

            try:

                print(f"Đang xử lý: {filename}")

                file_path = os.path.join(folder, filename)
                base_name = os.path.splitext(filename)[0]

                workbook = excel.Workbooks.Open(
                    file_path,
                    UpdateLinks=0,
                    ReadOnly=True,
                    IgnoreReadOnlyRecommended=True
                )

                total_sheets = workbook.Worksheets.Count

                # ====================================
                # SHEET 1
                # ====================================
                sheet1 = workbook.Worksheets(1)

                pdf_path = os.path.join(folder,f"K1 GCN {base_name}.pdf")

                sheet1.ExportAsFixedFormat(
                    Type=0,
                    Filename=pdf_path
                )


                # Chèn background
                if use_background:

                    doc = fitz.open(pdf_path)

                    for i in range(len(doc)):

                        page = doc[i]
                        rect = page.rect

                        if i == 0:
                            img_path = background_page1
                        else:
                            img_path = background_other_pages

                        # Chèn ảnh nền phía dưới nội dung
                        page.insert_image(
                            rect,
                            filename=img_path,
                            overlay=False
                        )

                    temp_pdf = pdf_path.replace(".pdf", "_tmp.pdf")

                    doc.save(
                        temp_pdf,
                        garbage=4,
                        deflate=True
                    )

                    doc.close()

                    os.remove(pdf_path)
                    os.replace(temp_pdf, pdf_path)

                    print(f"Đã tạo: K1 GCN {base_name}.pdf")


                # SHEET 2
                
                if total_sheets >= 2:

                    sheet2 = workbook.Worksheets(2)

                    pdf_path2 = os.path.join(
                        folder,
                        f"K1 biên bản {base_name}.pdf"
                    )

                    sheet2.ExportAsFixedFormat(
                        Type=0,
                        Filename=pdf_path2
                    )

                    print(f"Đã tạo: K1 biên bản {base_name}.pdf")

                workbook.Close(False)

            except Exception as e:

                print(f"Lỗi với file {filename}: {e}")

                try:
                    workbook.Close(False)
                except:
                    pass

        elif filename.lower().endswith(word_extensions) and not filename.startswith("~$"):

            try:

                print(f"Đang xử lý Word: {filename}")

                file_path = os.path.join(folder, filename)
                base_name = os.path.splitext(filename)[0]

                doc = word.Documents.Open(
                    file_path,
                    ReadOnly=True
                )

                pdf_path = os.path.join(
                    folder,
                    f"K1 GCN {base_name}.pdf"
                )

                # 17 = wdExportFormatPDF
                doc.ExportAsFixedFormat(
                    OutputFileName=pdf_path,
                    ExportFormat=17
                )

                doc.Close(False)

                # Add background
               
                if use_background:

                    pdf = fitz.open(pdf_path)

                    for i in range(len(pdf)):

                        page = pdf[i]

                        r = page.rect

                        # tránh hở viền
                        r.x0 -= 1
                        r.y0 -= 1
                        r.x1 += 1
                        r.y1 += 1

                        if i == 0:
                            img_path = background_page1
                        else:
                            img_path = background_other_pages

                        page.insert_image(
                            r,
                            filename=img_path,
                            overlay=False
                        )

                    temp_pdf = pdf_path.replace(".pdf", "_tmp.pdf")

                    pdf.save(
                        temp_pdf,
                        garbage=4,
                        deflate=True
                    )

                    pdf.close()

                    os.remove(pdf_path)
                    os.replace(temp_pdf, pdf_path)                

                print(f"Đã tạo: {base_name}.pdf")

            except Exception as e:

                print(f"Lỗi với file Word {filename}: {e}")

finally:

    excel.Quit()

print("Hoàn thành!")