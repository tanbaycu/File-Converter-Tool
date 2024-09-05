import os
from pdf2docx import Converter
import tabula
import html2text
from docx import Document
from bs4 import BeautifulSoup
from openpyxl import load_workbook, Workbook
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import markdown
import csv
from pptx import Presentation
""" hãy chuyển đổi định dạng của đường dẫn thành đường dẫn tuyệt đối có \\"""
# PDF sang DOCX
def convert_pdf_to_docx(pdf_path):
    base_name = os.path.splitext(pdf_path)[0]
    docx_path = base_name + '.docx'
    
    cv = Converter(pdf_path)
    cv.convert(docx_path)
    cv.close()
    
    print(f"Converted '{pdf_path}' to '{docx_path}'")

# PDF sang XLSX
def convert_pdf_to_xlsx(pdf_path):
    base_name = os.path.splitext(pdf_path)[0]
    xlsx_path = base_name + '.xlsx'
    
    tabula.convert_into(pdf_path, xlsx_path, output_format="xlsx")
    
    print(f"Converted '{pdf_path}' to '{xlsx_path}'")

#  DOCX sang PDF
def convert_docx_to_pdf(docx_path):
    base_name = os.path.splitext(docx_path)[0]
    pdf_path = base_name + '.pdf'
    
    pdf = canvas.Canvas(pdf_path, pagesize=letter)
    doc = Document(docx_path)
    
    for para in doc.paragraphs:
        pdf.drawString(100, 750, para.text)
    
    pdf.save()
    
    print(f"Converted '{docx_path}' to '{pdf_path}'")
    
def convert_xlsx_to_docx(xlsx_path):
    base_name = os.path.splitext(xlsx_path)[0]
    docx_path = base_name + '.docx'
    
    workbook = load_workbook(xlsx_path)
    sheet = workbook.active
    
    doc = Document()
    
    for row in sheet.iter_rows(values_only=True):
        text = " ".join([str(cell) for cell in row if cell is not None])
        doc.add_paragraph(text)
    
    doc.save(docx_path)
    
    print(f"Converted '{xlsx_path}' to '{docx_path}'")

#  XLSX sang PDF
def convert_xlsx_to_pdf(xlsx_path):
    base_name = os.path.splitext(xlsx_path)[0]
    pdf_path = base_name + '.pdf'
    
    pdf = canvas.Canvas(pdf_path, pagesize=letter)
    workbook = load_workbook(xlsx_path)
    sheet = workbook.active
    
    row_y = 750
    for row in sheet.iter_rows(values_only=True):
        text = " ".join([str(cell) for cell in row if cell is not None])
        pdf.drawString(100, row_y, text)
        row_y -= 20
    
    pdf.save()
    
    print(f"Converted '{xlsx_path}' to '{pdf_path}'")

def convert_txt_to_docx(txt_path):
    base_name = os.path.splitext(txt_path)[0]
    docx_path = base_name + '.docx'
    
    with open(txt_path, 'r', encoding='utf-8') as file:
        text = file.read()
    
    doc = Document()
    doc.add_paragraph(text)
    doc.save(docx_path)
    
    print(f"Converted '{txt_path}' to '{docx_path}'")

#  TXT sang PDF
def convert_txt_to_pdf(txt_path):
    base_name = os.path.splitext(txt_path)[0]
    pdf_path = base_name + '.pdf'
    
    pdf = canvas.Canvas(pdf_path, pagesize=letter)
    with open(txt_path, 'r') as file:
        text = file.read()
        pdf.drawString(100, 750, text)
    
    pdf.save()
    
    print(f"Converted '{txt_path}' to '{pdf_path}'")

#  PDF sang TXT
def convert_pdf_to_txt(pdf_path):
    base_name = os.path.splitext(pdf_path)[0]
    txt_path = base_name + '.txt'
    
    from PyPDF2 import PdfReader
    pdf = PdfReader(pdf_path)
    with open(txt_path, 'w') as file:
        for page in pdf.pages:
            file.write(page.extract_text())
    
    print(f"Converted '{pdf_path}' to '{txt_path}'")

# CSV sang XLSX
def convert_csv_to_xlsx(csv_path):
    base_name = os.path.splitext(csv_path)[0]
    xlsx_path = base_name + '.xlsx'
    
    wb = Workbook()
    ws = wb.active
    
    with open(csv_path, 'r') as file:
        reader = csv.reader(file)
        for row in reader:
            ws.append(row)
    
    wb.save(xlsx_path)
    
    print(f"Converted '{csv_path}' to '{xlsx_path}'")

#  XLSX sang CSV
def convert_xlsx_to_csv(xlsx_path):
    base_name = os.path.splitext(xlsx_path)[0]
    csv_path = base_name + '.csv'
    
    wb = load_workbook(xlsx_path)
    ws = wb.active
    
    with open(csv_path, 'w', newline='') as file:
        writer = csv.writer(file)
        for row in ws.iter_rows(values_only=True):
            writer.writerow(row)
    
    print(f"Converted '{xlsx_path}' to '{csv_path}'")

# PPTX sang PDF
def convert_pptx_to_pdf(pptx_path):
    base_name = os.path.splitext(pptx_path)[0]
    pdf_path = base_name + '.pdf'
    
    presentation = Presentation(pptx_path)
    pdf = canvas.Canvas(pdf_path, pagesize=letter)
    
    for slide in presentation.slides:
        pdf.drawString(100, 750, f"Slide title: {slide.shapes.title.text}")
        pdf.showPage()
    
    pdf.save()
    
    print(f"Converted '{pptx_path}' to '{pdf_path}'")

#  PPTX sang DOCX
def convert_pptx_to_docx(pptx_path):
    base_name = os.path.splitext(pptx_path)[0]
    docx_path = base_name + '.docx'
    
    presentation = Presentation(pptx_path)
    doc = Document()
    
    for slide in presentation.slides:
        doc.add_heading(slide.shapes.title.text, level=1)
        for shape in slide.shapes:
            if shape.has_text_frame:
                doc.add_paragraph(shape.text)
    
    doc.save(docx_path)
    
    print(f"Converted '{pptx_path}' to '{docx_path}'")

# Markdown (MD) sang TXT
def convert_md_to_txt(md_path):
    base_name = os.path.splitext(md_path)[0]
    txt_path = base_name + '.txt'
    
    with open(md_path, 'r') as file:
        md_content = file.read()
    
    html_content = markdown.markdown(md_content)
    
    with open(txt_path, 'w') as file:
        file.write(html_content)
    
    print(f"Converted '{md_path}' to '{txt_path}'")

#  TXT sang Markdown (MD)
def convert_txt_to_md(txt_path):
    base_name = os.path.splitext(txt_path)[0]
    md_path = base_name + '.md'
    
    with open(txt_path, 'r') as file:
        txt_content = file.read()
    
    with open(md_path, 'w') as file:
        file.write(txt_content)
    
    print(f"Converted '{txt_path}' to '{md_path}'")

#  Markdown (MD) sang HTML
def convert_md_to_html(md_path):
    base_name = os.path.splitext(md_path)[0]
    html_path = base_name + '.html'
    
    with open(md_path, 'r', encoding='utf-8') as file:
        md_content = file.read()
    
    html_content = markdown.markdown(md_content)
    
    with open(html_path, 'w', encoding='utf-8') as file:
        file.write(html_content)
    
    print(f"Converted '{md_path}' to '{html_path}'")


def handle_conversion(file_path):
    ext = os.path.splitext(file_path)[1].lower()

    if ext == ".pdf":
        print("Chọn phương thức chuyển đổi:")
        print("1. PDF sang DOCX")
        print("2. PDF sang XLSX")
        print("3. PDF sang TXT")
        choice = input("Nhập số lựa chọn: ")
        
        if choice == "1":
            convert_pdf_to_docx(file_path)
        elif choice == "2":
            convert_pdf_to_xlsx(file_path)
        elif choice == "3":
            convert_pdf_to_txt(file_path)
        else:
            print("Lựa chọn không hợp lệ.")
    
    elif ext == ".docx":
        print("Chọn phương thức chuyển đổi:")
        print("1. DOCX sang PDF")
        print("2. DOCX sang PPTX")
        choice = input("Nhập số lựa chọn: ")
        
        if choice == "1":
            convert_docx_to_pdf(file_path)
        elif choice == "2":
            convert_docx_to_pptx(file_path)
        else:
            print("Lựa chọn không hợp lệ.")
    
    elif ext == ".xlsx":
        print("Chọn phương thức chuyển đổi:")
        print("1. XLSX sang DOCX")
        print("2. XLSX sang PDF")
        print("3. XLSX sang CSV")
        choice = input("Nhập số lựa chọn: ")
        
        if choice == "1":
            convert_xlsx_to_docx(file_path)
        elif choice == "2":
            convert_xlsx_to_pdf(file_path)
        elif choice == "3":
            convert_xlsx_to_csv(file_path)
        else:
            print("Lựa chọn không hợp lệ.")
    
    elif ext == ".txt":
        print("Chọn phương thức chuyển đổi:")
        print("1. TXT sang DOCX")
        print("2. TXT sang PDF")
        print("3. TXT sang Markdown")
        choice = input("Nhập số lựa chọn: ")
        
        if choice == "1":
            convert_txt_to_docx(file_path)
        elif choice == "2":
            convert_txt_to_pdf(file_path)
        elif choice == "3":
            convert_txt_to_md(file_path)
        else:
            print("Lựa chọn không hợp lệ.")
    
    elif ext == ".csv":
        print("Chọn phương thức chuyển đổi:")
        print("1. CSV sang XLSX")
        choice = input("Nhập số lựa chọn: ")
        
        if choice == "1":
            convert_csv_to_xlsx(file_path)
        else:
            print("Lựa chọn không hợp lệ.")
    
    elif ext == ".pptx":
        print("Chọn phương thức chuyển đổi:")
        print("1. PPTX sang PDF")
        print("2. PPTX sang DOCX")
        choice = input("Nhập số lựa chọn: ")
        
        if choice == "1":
            convert_pptx_to_pdf(file_path)
        elif choice == "2":
            convert_pptx_to_docx(file_path)
        else:
            print("Lựa chọn không hợp lệ.")
    
    elif ext == ".md":
        print("Chọn phương thức chuyển đổi:")
        print("1. Markdown sang TXT")
        print("2. Markdown sang HTML")
        choice = input("Nhập số lựa chọn: ")
        
        if choice == "1":
            convert_md_to_txt(file_path)
        elif choice == "2":
            convert_md_to_html(file_path)
        else:
            print("Lựa chọn không hợp lệ.")
    
    else:
        print("Định dạng tệp không hỗ trợ.")


print("Vui lòng định dạng đường dẫn với dạng \\\ và bỏ dấu nháy")
file_path = input("Nhập đường dẫn tệp: ")
if os.path.isfile(file_path):
    handle_conversion(file_path)
else:
    print(f"'{file_path}' không phải là tệp hợp lệ hoặc tệp không hề tồn tại.")

