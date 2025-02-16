import os
import comtypes.client
from PyPDF2 import PdfMerger

# مسیر پوشه‌ای که فایل‌های ورد در آن قرار دارند
folder_path = 'E:\\code\\word'

# فهرست فایل‌های ورد در پوشه، بر اساس نام فایل‌ها
input_files = sorted([f for f in os.listdir(folder_path) if f.startswith('certificate') and f.endswith('.docx')])

# تبدیل فایل ورد به PDF
def convert_to_pdf(docx_path, pdf_path):
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(docx_path)
    doc.SaveAs(pdf_path, FileFormat=17)  # فرمت PDF
    doc.Close()
    word.Quit()
    os.remove(docx_path)  # حذف فایل ورد پس از تبدیل

# تبدیل همه فایل‌های ورد به PDF
for file in input_files:
    file_path = os.path.join(folder_path, file)
    pdf_output = os.path.join(folder_path, file.replace('.docx', '.pdf'))
    convert_to_pdf(file_path, pdf_output)

# فهرست فایل‌های PDF در پوشه
input_pdf_files = sorted([f for f in os.listdir(folder_path) if f.startswith('certificate') and f.endswith('.pdf')])

# ترکیب فایل‌های PDF
pdf_merger = PdfMerger()
for file in input_pdf_files:
    file_path = os.path.join(folder_path, file)
    pdf_merger.append(file_path)

# ذخیره کردن فایل ترکیب‌شده به عنوان یک PDF جدید
output_pdf = os.path.join(folder_path, 'combined_certificates.pdf')
pdf_merger.write(output_pdf)

# بستن شیء PdfMerger
pdf_merger.close()

# حذف فایل‌های PDF پس از ترکیب
for file in input_pdf_files:
    file_path = os.path.join(folder_path, file)
    os.remove(file_path)  # حذف فایل PDF

print(f"تمام فایل‌های ورد به PDF تبدیل شدند و فایل‌های PDF حذف شدند.")
