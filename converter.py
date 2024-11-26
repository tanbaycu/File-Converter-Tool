import os #8.24.22 - pdf -> xlsx , docx - pdf
import logging
from pathlib import Path
from typing import Callable, Dict
import pandas as pd
from pdf2docx import Converter
from docx import Document
from openpyxl import load_workbook, Workbook
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import markdown
import csv
from PyPDF2 import PdfReader
from pptx import Presentation
from docx2pdf import convert
# Thiết lập logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def safe_convert(func: Callable) -> Callable:
    """Decorator để xử lý ngoại lệ trong các hàm chuyển đổi."""
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            logging.error(f"Lỗi trong {func.__name__}: {str(e)}")
            return None
    return wrapper

@safe_convert
def convert_pdf_to_docx(pdf_path: str) -> str:
    docx_path = Path(pdf_path).with_suffix('.docx')
    try:
        with Converter(pdf_path) as cv:
            cv.convert(str(docx_path))
        return str(docx_path)
    except Exception as e:
        logging.error(f"Lỗi khi sử dụng pdf2docx: {str(e)}")
        logging.info("Đang thử phương pháp thay thế...")
        
        try:
            from PyPDF2 import PdfReader
            from docx import Document
            
            pdf = PdfReader(pdf_path)
            doc = Document()
            
            for page in pdf.pages:
                text = page.extract_text()
                doc.add_paragraph(text)
            
            doc.save(docx_path)
            return str(docx_path)
        except Exception as e:
            logging.error(f"Lỗi khi sử dụng phương pháp thay thế: {str(e)}")
            return None

@safe_convert
def convert_pdf_to_xlsx(pdf_path: str) -> str:
    xlsx_path = Path(pdf_path).with_suffix('.xlsx')
    pdf = PdfReader(pdf_path)
    data = []
    for page in pdf.pages:
        text = page.extract_text()
        data.extend([line.split() for line in text.split('\n') if line.strip()])
    
    df = pd.DataFrame(data)
    df.to_excel(xlsx_path, index=False, header=False)
    return str(xlsx_path)

@safe_convert
def convert_docx_to_pdf(docx_path: str) -> str:
    pdf_path = Path(docx_path).with_suffix('.pdf')
    convert(docx_path, str(pdf_path))
    return str(pdf_path)

@safe_convert
def convert_xlsx_to_docx(xlsx_path: str) -> str:
    docx_path = Path(xlsx_path).with_suffix('.docx')
    workbook = load_workbook(xlsx_path)
    sheet = workbook.active
    
    doc = Document()
    for row in sheet.iter_rows(values_only=True):
        doc.add_paragraph(" ".join(str(cell) for cell in row if cell is not None))
    
    doc.save(docx_path)
    return str(docx_path)

@safe_convert
def convert_xlsx_to_pdf(xlsx_path: str) -> str:
    pdf_path = Path(xlsx_path).with_suffix('.pdf')
    workbook = load_workbook(xlsx_path)
    sheet = workbook.active
    
    pdf = canvas.Canvas(str(pdf_path), pagesize=letter)
    for row in sheet.iter_rows(values_only=True):
        pdf.drawString(100, 750, " ".join(str(cell) for cell in row if cell is not None))
        pdf.showPage()
    pdf.save()
    
    return str(pdf_path)

@safe_convert
def convert_txt_to_docx(txt_path: str) -> str:
    docx_path = Path(txt_path).with_suffix('.docx')
    with open(txt_path, 'r', encoding='utf-8') as file:
        text = file.read()
    
    doc = Document()
    doc.add_paragraph(text)
    doc.save(docx_path)
    
    return str(docx_path)

@safe_convert
def convert_txt_to_pdf(txt_path: str) -> str:
    pdf_path = Path(txt_path).with_suffix('.pdf')
    with open(txt_path, 'r', encoding='utf-8') as file:
        text = file.read()
    
    pdf = canvas.Canvas(str(pdf_path), pagesize=letter)
    pdf.drawString(100, 750, text)
    pdf.save()
    
    return str(pdf_path)

@safe_convert
def convert_pdf_to_txt(pdf_path: str) -> str:
    txt_path = Path(pdf_path).with_suffix('.txt')
    pdf = PdfReader(pdf_path)
    
    with open(txt_path, 'w', encoding='utf-8') as file:
        for page in pdf.pages:
            file.write(page.extract_text())
    
    return str(txt_path)

@safe_convert
def convert_csv_to_xlsx(csv_path: str) -> str:
    xlsx_path = Path(csv_path).with_suffix('.xlsx')
    
    wb = Workbook()
    ws = wb.active
    
    with open(csv_path, 'r', encoding='utf-8') as file:
        reader = csv.reader(file)
        for row in reader:
            ws.append(row)
    
    wb.save(xlsx_path)
    return str(xlsx_path)

@safe_convert
def convert_xlsx_to_csv(xlsx_path: str) -> str:
    csv_path = Path(xlsx_path).with_suffix('.csv')
    
    wb = load_workbook(xlsx_path)
    ws = wb.active
    
    with open(csv_path, 'w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        for row in ws.iter_rows(values_only=True):
            writer.writerow(row)
    
    return str(csv_path)

@safe_convert
def convert_pptx_to_pdf(pptx_path: str) -> str:
    pdf_path = Path(pptx_path).with_suffix('.pdf')
    
    presentation = Presentation(pptx_path)
    pdf = canvas.Canvas(str(pdf_path), pagesize=letter)
    
    for slide in presentation.slides:
        if slide.shapes.title:
            pdf.drawString(100, 750, f"Slide title: {slide.shapes.title.text}")
        pdf.showPage()
    
    pdf.save()
    return str(pdf_path)

@safe_convert
def convert_pptx_to_docx(pptx_path: str) -> str:
    docx_path = Path(pptx_path).with_suffix('.docx')
    
    presentation = Presentation(pptx_path)
    doc = Document()
    
    for slide in presentation.slides:
        if slide.shapes.title:
            doc.add_heading(slide.shapes.title.text, level=1)
        for shape in slide.shapes:
            if hasattr(shape, 'text'):
                doc.add_paragraph(shape.text)
    
    doc.save(docx_path)
    return str(docx_path)

@safe_convert
def convert_md_to_txt(md_path: str) -> str:
    txt_path = Path(md_path).with_suffix('.txt')
    
    with open(md_path, 'r', encoding='utf-8') as file:
        md_content = file.read()
    
    html_content = markdown.markdown(md_content)
    
    with open(txt_path, 'w', encoding='utf-8') as file:
        file.write(html_content)
    
    return str(txt_path)

@safe_convert
def convert_txt_to_md(txt_path: str) -> str:
    md_path = Path(txt_path).with_suffix('.md')
    
    with open(txt_path, 'r', encoding='utf-8') as file:
        txt_content = file.read()
    
    with open(md_path, 'w', encoding='utf-8') as file:
        file.write(txt_content)
    
    return str(md_path)

@safe_convert
def convert_md_to_html(md_path: str) -> str:
    html_path = Path(md_path).with_suffix('.html')
    
    with open(md_path, 'r', encoding='utf-8') as file:
        md_content = file.read()
    
    html_content = markdown.markdown(md_content)
    
    with open(html_path, 'w', encoding='utf-8') as file:
        file.write(html_content)
    
    return str(html_path)

CONVERSION_MAP: Dict[str, Dict[str, Callable]] = {
    '.pdf': {
        '.docx': convert_pdf_to_docx,
        '.xlsx': convert_pdf_to_xlsx,
        '.txt': convert_pdf_to_txt,
    },
    '.docx': {
        '.pdf': convert_docx_to_pdf,
    },
    '.xlsx': {
        '.docx': convert_xlsx_to_docx,
        '.pdf': convert_xlsx_to_pdf,
        '.csv': convert_xlsx_to_csv,
    },
    '.txt': {
        '.docx': convert_txt_to_docx,
        '.pdf': convert_txt_to_pdf,
        '.md': convert_txt_to_md,
    },
    '.csv': {
        '.xlsx': convert_csv_to_xlsx,
    },
    '.pptx': {
        '.pdf': convert_pptx_to_pdf,
        '.docx': convert_pptx_to_docx,
    },
    '.md': {
        '.txt': convert_md_to_txt,
        '.html': convert_md_to_html,
    },
}

def handle_conversion(file_path: str) -> None:
    file_path = Path(file_path)
    if not file_path.is_file():
        logging.error(f"'{file_path}' không phải là tệp hợp lệ hoặc không tồn tại.")
        return

    src_ext = file_path.suffix.lower()
    if src_ext not in CONVERSION_MAP:
        logging.error(f"Định dạng tệp nguồn không được hỗ trợ: {src_ext}")
        return

    print(f"Các tùy chọn chuyển đổi có sẵn cho {src_ext}:")
    for i, target_ext in enumerate(CONVERSION_MAP[src_ext].keys(), 1):
        print(f"{i}. {src_ext[1:].upper()} sang {target_ext[1:].upper()}")

    choice = input("Nhập số lựa chọn của bạn: ")
    try:
        choice = int(choice)
        target_ext = list(CONVERSION_MAP[src_ext].keys())[choice - 1]
    except (ValueError, IndexError):
        logging.error("Lựa chọn không hợp lệ.")
        return

    conversion_func = CONVERSION_MAP[src_ext][target_ext]
    result = conversion_func(str(file_path))
    
    if result:
        logging.info(f"Đã chuyển đổi thành công '{file_path}' thành '{result}'")
    else:
        logging.error(f"Chuyển đổi thất bại cho '{file_path}'")

if __name__ == "__main__":
    print("Vui lòng cung cấp đường dẫn tệp sử dụng \\ cho dấu phân cách thư mục.")
    file_path = input("Nhập đường dẫn tệp: ")
    handle_conversion(file_path)


