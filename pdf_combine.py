import os
from PyPDF2 import PdfMerger

# مسیر پوشه‌ای که فایل‌های PDF در آن قرار دارند
folder_path = 'word'

# فهرست فایل‌های PDF در پوشه، بر اساس نام فایل‌ها
input_files = sorted([f for f in os.listdir(folder_path) if f.startswith('certificate') and f.endswith('.pdf')])

# ایجاد یک شیء PdfMerger برای ترکیب PDF‌ها
pdf_merger = PdfMerger()

# افزودن تمام فایل‌های PDF به شیء PdfMerger
for file in input_files:
    file_path = os.path.join(folder_path, file)
    pdf_merger.append(file_path)

# ذخیره کردن فایل ترکیب‌شده به عنوان یک PDF جدید
output_pdf = os.path.join(folder_path, 'combined_certificates.pdf')
pdf_merger.write(output_pdf)

# بستن شیء PdfMerger
pdf_merger.close()


print(f"تمام فایل‌های PDF به یک فایل واحد تبدیل شدند: {output_pdf}")


for paragraph in doc.paragraphs:
    for key, value in data.items():
        if f"{{{{{key}}}}}" in paragraph.text:
            paragraph.text = paragraph.text.replace(f"{{{{{key}}}}}", value)
