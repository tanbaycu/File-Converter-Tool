import os #ver12.13.58 docx - xlsx, thêm try 
import logging
from pathlib import Path
from typing import Callable, Dict
import pandas as pd
import numpy as np
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
import mammoth
import html2text
import pdfplumber

# Thiết lập logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def chuyen_doi_an_toan(func: Callable) -> Callable:
    """Decorator để xử lý ngoại lệ trong các hàm chuyển đổi."""
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            logging.error(f"Lỗi trong {func.__name__}: {str(e)}")
            return None
    return wrapper

@chuyen_doi_an_toan
def chuyen_doi_pdf_sang_docx(duong_dan_pdf: str) -> str:
    duong_dan_docx = Path(duong_dan_pdf).with_suffix('.docx')
    try:
        with Converter(duong_dan_pdf) as cv:
            cv.convert(str(duong_dan_docx))
        return str(duong_dan_docx)
    except Exception as e:
        logging.error(f"Lỗi khi sử dụng pdf2docx: {str(e)}")
        logging.info("Đang thử phương pháp thay thế...")
        
        try:
            pdf = PdfReader(duong_dan_pdf)
            doc = Document()
            
            for page in pdf.pages:
                text = page.extract_text()
                doc.add_paragraph(text)
            
            doc.save(duong_dan_docx)
            return str(duong_dan_docx)
        except Exception as e:
            logging.error(f"Lỗi khi sử dụng phương pháp thay thế: {str(e)}")
            return None

@chuyen_doi_an_toan
def chuyen_doi_pdf_sang_xlsx(duong_dan_pdf: str) -> str:
    """
    Chuyển đổi PDF sang XLSX với khả năng phát hiện và trích xuất bảng.
    
    Args:
        duong_dan_pdf: Đường dẫn đến tệp PDF
        
    Returns:
        Đường dẫn đến tệp XLSX đã tạo
    """
    duong_dan_xlsx = Path(duong_dan_pdf).with_suffix('.xlsx')
    
    try:
        with pdfplumber.open(duong_dan_pdf) as pdf:
            
            with pd.ExcelWriter(duong_dan_xlsx, engine='openpyxl') as writer:
                for i, page in enumerate(pdf.pages):
                    
                    tables = page.extract_tables()
                    
                    if tables:
                        
                        for j, table in enumerate(tables):
                            df = pd.DataFrame(table[1:], columns=table[0] if table else None)
                            sheet_name = f"Trang_{i+1}_Bảng_{j+1}"
                            df.to_excel(writer, sheet_name=sheet_name, index=False)
                            logging.info(f"Đã trích xuất bảng {j+1} từ trang {i+1}")
                    else:
                        
                        text = page.extract_text()
                        if text:
                            lines = [line.split() for line in text.split('\n') if line.strip()]
                            df = pd.DataFrame(lines)
                            sheet_name = f"Trang_{i+1}_Văn_bản"
                            df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
        
        logging.info(f"Đã chuyển đổi thành công '{duong_dan_pdf}' sang '{duong_dan_xlsx}'")
        return str(duong_dan_xlsx)
    
    except Exception as e:
        logging.error(f"Lỗi khi chuyển đổi PDF sang XLSX: {str(e)}")
        return None

@chuyen_doi_an_toan
def chuyen_doi_docx_sang_pdf(duong_dan_docx: str) -> str:
    duong_dan_pdf = Path(duong_dan_docx).with_suffix('.pdf')
    convert(duong_dan_docx, str(duong_dan_pdf))
    return str(duong_dan_pdf)

@chuyen_doi_an_toan
def chuyen_doi_xlsx_sang_docx(duong_dan_xlsx: str) -> str:
    duong_dan_docx = Path(duong_dan_xlsx).with_suffix('.docx')
    df = pd.read_excel(duong_dan_xlsx)
    
    doc = Document()
    for column in df.columns:
        doc.add_heading(column, level=1)
        for value in df[column]:
            doc.add_paragraph(str(value))
        doc.add_paragraph()  # Thêm một dòng trống giữa các cột
    
    doc.save(duong_dan_docx)
    return str(duong_dan_docx)

@chuyen_doi_an_toan
def chuyen_doi_docx_sang_xlsx(duong_dan_docx: str) -> str:
    duong_dan_xlsx = Path(duong_dan_docx).with_suffix('.xlsx')
    doc = Document(duong_dan_docx)
    
    data = []
    for paragraph in doc.paragraphs:
        data.append([paragraph.text])
    
    df = pd.DataFrame(data)
    df.to_excel(duong_dan_xlsx, index=False, header=False)
    return str(duong_dan_xlsx)

@chuyen_doi_an_toan
def chuyen_doi_xlsx_sang_pdf(duong_dan_xlsx: str) -> str:
    duong_dan_pdf = Path(duong_dan_xlsx).with_suffix('.pdf')
    df = pd.read_excel(duong_dan_xlsx)
    
    pdf = canvas.Canvas(str(duong_dan_pdf), pagesize=letter)
    y = 750  # Tọa độ y bắt đầu
    for column in df.columns:
        pdf.drawString(100, y, column)
        y -= 20
        for value in df[column]:
            pdf.drawString(120, y, str(value))
            y -= 15
            if y < 50:  # Bắt đầu trang mới nếu gần đến cuối trang
                pdf.showPage()
                y = 750
    pdf.save()
    
    return str(duong_dan_pdf)

@chuyen_doi_an_toan
def chuyen_doi_xlsx_sang_csv(duong_dan_xlsx: str) -> str:
    duong_dan_csv = Path(duong_dan_xlsx).with_suffix('.csv')
    df = pd.read_excel(duong_dan_xlsx)
    df.to_csv(duong_dan_csv, index=False)
    return str(duong_dan_csv)

@chuyen_doi_an_toan
def chuyen_doi_txt_sang_docx(duong_dan_txt: str) -> str:
    duong_dan_docx = Path(duong_dan_txt).with_suffix('.docx')
    with open(duong_dan_txt, 'r', encoding='utf-8') as file:
        text = file.read()
    
    doc = Document()
    for paragraph in text.split('\n'):
        doc.add_paragraph(paragraph)
    doc.save(duong_dan_docx)
    
    return str(duong_dan_docx)

@chuyen_doi_an_toan
def chuyen_doi_txt_sang_pdf(duong_dan_txt: str) -> str:
    duong_dan_pdf = Path(duong_dan_txt).with_suffix('.pdf')
    with open(duong_dan_txt, 'r', encoding='utf-8') as file:
        text = file.read()
    
    pdf = canvas.Canvas(str(duong_dan_pdf), pagesize=letter)
    y = 750  # Tọa độ y bắt đầu
    for line in text.split('\n'):
        pdf.drawString(100, y, line)
        y -= 15
        if y < 50:  # Bắt đầu trang mới nếu gần đến cuối trang
            pdf.showPage()
            y = 750
    pdf.save()
    
    return str(duong_dan_pdf)

@chuyen_doi_an_toan
def chuyen_doi_txt_sang_md(duong_dan_txt: str) -> str:
    duong_dan_md = Path(duong_dan_txt).with_suffix('.md')
    
    with open(duong_dan_txt, 'r', encoding='utf-8') as file:
        noi_dung_txt = file.read()
    
    # Chuyển đổi đơn giản: giả định mỗi dòng là một đoạn văn
    noi_dung_md = '\n\n'.join(noi_dung_txt.split('\n'))
    
    with open(duong_dan_md, 'w', encoding='utf-8') as file:
        file.write(noi_dung_md)
    
    return str(duong_dan_md)

@chuyen_doi_an_toan
def chuyen_doi_csv_sang_xlsx(duong_dan_csv: str) -> str:
    duong_dan_xlsx = Path(duong_dan_csv).with_suffix('.xlsx')
    df = pd.read_csv(duong_dan_csv)
    df.to_excel(duong_dan_xlsx, index=False)
    return str(duong_dan_xlsx)

@chuyen_doi_an_toan
def chuyen_doi_pptx_sang_pdf(duong_dan_pptx: str) -> str:
    from win32com import client
    duong_dan_pdf = Path(duong_dan_pptx).with_suffix('.pdf')
    
    powerpoint = client.Dispatch("Powerpoint.Application")
    deck = powerpoint.Presentations.Open(duong_dan_pptx)
    deck.SaveAs(duong_dan_pdf, 32)  # 32 là mã định dạng PDF
    deck.Close()
    powerpoint.Quit()
    
    return str(duong_dan_pdf)

@chuyen_doi_an_toan
def chuyen_doi_pptx_sang_docx(duong_dan_pptx: str) -> str:
    duong_dan_docx = Path(duong_dan_pptx).with_suffix('.docx')
    
    presentation = Presentation(duong_dan_pptx)
    doc = Document()
    
    for slide in presentation.slides:
        if slide.shapes.title:
            doc.add_heading(slide.shapes.title.text, level=1)
        for shape in slide.shapes:
            if hasattr(shape, 'text'):
                doc.add_paragraph(shape.text)
        doc.add_page_break()
    
    doc.save(duong_dan_docx)
    return str(duong_dan_docx)

@chuyen_doi_an_toan
def chuyen_doi_md_sang_txt(duong_dan_md: str) -> str:
    duong_dan_txt = Path(duong_dan_md).with_suffix('.txt')
    
    with open(duong_dan_md, 'r', encoding='utf-8') as file:
        noi_dung_md = file.read()
    
    h = html2text.HTML2Text()
    h.ignore_links = True
    noi_dung_txt = h.handle(markdown.markdown(noi_dung_md))
    
    with open(duong_dan_txt, 'w', encoding='utf-8') as file:
        file.write(noi_dung_txt)
    
    return str(duong_dan_txt)

@chuyen_doi_an_toan
def chuyen_doi_md_sang_html(duong_dan_md: str) -> str:
    duong_dan_html = Path(duong_dan_md).with_suffix('.html')
    
    with open(duong_dan_md, 'r', encoding='utf-8') as file:
        noi_dung_md = file.read()
    
    noi_dung_html = markdown.markdown(noi_dung_md, extensions=['extra'])
    
    with open(duong_dan_html, 'w', encoding='utf-8') as file:
        file.write(f"<html><body>{noi_dung_html}</body></html>")
    
    return str(duong_dan_html)

@chuyen_doi_an_toan
def chuyen_doi_pdf_sang_txt(duong_dan_pdf: str) -> str:
    """
    Chuyển đổi PDF sang TXT với hỗ trợ OCR cho PDF quét và cải thiện định dạng.
    
    Args:
        duong_dan_pdf: Đường dẫn đến tệp PDF
        
    Returns:
        Đường dẫn đến tệp TXT đã tạo
    """
    duong_dan_txt = Path(duong_dan_pdf).with_suffix('.txt')
    
    # Phương pháp 1: Sử dụng pdfplumber (tốt cho PDF có văn bản)
    try:
        logging.info(f"Đang chuyển đổi '{duong_dan_pdf}' sang TXT bằng pdfplumber...")
        
        with pdfplumber.open(duong_dan_pdf) as pdf:
            with open(duong_dan_txt, 'w', encoding='utf-8') as txt_file:
                for i, page in enumerate(pdf.pages):
                    text = page.extract_text()
                    if text:
                        txt_file.write(f"--- Trang {i+1} ---\n\n")
                        txt_file.write(text)
                        txt_file.write("\n\n")
        
        # Kiểm tra kết quả
        with open(duong_dan_txt, 'r', encoding='utf-8') as f:
            content = f.read().strip()
        
        if len(content) > 100:  # Nếu có đủ văn bản
            logging.info(f"Đã chuyển đổi thành công bằng pdfplumber: '{duong_dan_txt}'")
            return str(duong_dan_txt)
        else:
            logging.warning("pdfplumber trích xuất ít văn bản. Có thể là PDF quét. Thử OCR...")
    except Exception as e:
        logging.warning(f"pdfplumber không thành công: {str(e)}")
    
    # Phương pháp 2: Sử dụng PyMuPDF + OCR cho PDF quét
    try:
        logging.info("Đang thử chuyển đổi bằng PyMuPDF + OCR...")
        
        # Tạo thư mục tạm cho hình ảnh
        with tempfile.TemporaryDirectory() as temp_dir:
            # Mở PDF
            pdf_doc = fitz.open(duong_dan_pdf)
            
            with open(duong_dan_txt, 'w', encoding='utf-8') as txt_file:
                for page_num, page in enumerate(pdf_doc):
                    # Thử trích xuất văn bản trực tiếp
                    text = page.get_text("text")
                    
                    # Nếu ít văn bản, thử OCR
                    if len(text.strip()) < 100:
                        # Chuyển trang thành hình ảnh
                        pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72))
                        img_path = os.path.join(temp_dir, f"page_{page_num+1}.png")
                        pix.save(img_path)
                        
                        # OCR hình ảnh
                        try:
                            img = Image.open(img_path)
                            ocr_text = pytesseract.image_to_string(img, lang='vie+eng')
                            text = ocr_text
                        except Exception as ocr_err:
                            logging.warning(f"OCR không thành công cho trang {page_num+1}: {str(ocr_err)}")
                    
                    # Ghi văn bản vào tệp
                    if text.strip():
                        txt_file.write(f"--- Trang {page_num+1} ---\n\n")
                        txt_file.write(text)
                        txt_file.write("\n\n")
        
        # Kiểm tra kết quả
        with open(duong_dan_txt, 'r', encoding='utf-8') as f:
            content = f.read().strip()
        
        if content:
            logging.info(f"Đã chuyển đổi thành công bằng PyMuPDF + OCR: '{duong_dan_txt}'")
            return str(duong_dan_txt)
    except Exception as e:
        logging.warning(f"PyMuPDF + OCR không thành công: {str(e)}")
    
    # Phương pháp 3: Phương pháp dự phòng với PyPDF2
    try:
        logging.info("Đang thử phương pháp dự phòng với PyPDF2...")
        
        pdf = PdfReader(duong_dan_pdf)
        with open(duong_dan_txt, 'w', encoding='utf-8') as txt_file:
            for i, page in enumerate(pdf.pages):
                text = page.extract_text()
                if text:
                    txt_file.write(f"--- Trang {i+1} ---\n\n")
                    txt_file.write(text)
                    txt_file.write("\n\n")
        
        logging.info(f"Đã chuyển đổi với phương pháp dự phòng: '{duong_dan_txt}'")
        return str(duong_dan_txt)
    except Exception as e:
        logging.error(f"Tất cả các phương pháp chuyển đổi đều thất bại: {str(e)}")
        
        # Tạo tệp trống nếu tất cả đều thất bại
        with open(duong_dan_txt, 'w', encoding='utf-8') as f:
            f.write(f"Không thể trích xuất văn bản từ {Path(duong_dan_pdf).name}.\n")
            f.write(f"Lỗi: {str(e)}")
        
        return str(duong_dan_txt)

# Cập nhật BANG_CHUYEN_DOI với các hàm mới và cải tiến
BANG_CHUYEN_DOI: Dict[str, Dict[str, Callable]] = {
    '.pdf': {
        '.docx': chuyen_doi_pdf_sang_docx,
        '.xlsx': chuyen_doi_pdf_sang_xlsx,
        '.txt': chuyen_doi_pdf_sang_txt,
    },
    '.xlsx': {
        '.docx': chuyen_doi_xlsx_sang_docx,
        '.pdf': chuyen_doi_xlsx_sang_pdf,
        '.csv': chuyen_doi_xlsx_sang_csv,
    },
    '.docx': {
        '.pdf': chuyen_doi_docx_sang_pdf,
        '.xlsx': chuyen_doi_docx_sang_xlsx,
    },
    '.txt': {
        '.docx': chuyen_doi_txt_sang_docx,
        '.pdf': chuyen_doi_txt_sang_pdf,
        '.md': chuyen_doi_txt_sang_md,
    },
    '.csv': {
        '.xlsx': chuyen_doi_csv_sang_xlsx,
    },
    '.pptx': {
        '.pdf': chuyen_doi_pptx_sang_pdf,
        '.docx': chuyen_doi_pptx_sang_docx,
    },
    '.md': {
        '.txt': chuyen_doi_md_sang_txt,
        '.html': chuyen_doi_md_sang_html,
    },
}

def xu_ly_chuyen_doi(duong_dan_tep: str) -> None:
    duong_dan_tep = Path(duong_dan_tep)
    if not duong_dan_tep.is_file():
        logging.error(f"'{duong_dan_tep}' không phải là tệp hợp lệ hoặc không tồn tại.")
        return

    dinh_dang_nguon = duong_dan_tep.suffix.lower()
    if dinh_dang_nguon not in BANG_CHUYEN_DOI:
        logging.error(f"Định dạng tệp nguồn không được hỗ trợ: {dinh_dang_nguon}")
        return

    print(f"Các tùy chọn chuyển đổi có sẵn cho {dinh_dang_nguon}:")
    for i, dinh_dang_dich in enumerate(BANG_CHUYEN_DOI[dinh_dang_nguon].keys(), 1):
        print(f"{i}. {dinh_dang_nguon[1:].upper()} sang {dinh_dang_dich[1:].upper()}")

    lua_chon = input("Nhập số lựa chọn của bạn: ")
    try:
        lua_chon = int(lua_chon)
        dinh_dang_dich = list(BANG_CHUYEN_DOI[dinh_dang_nguon].keys())[lua_chon - 1]
    except (ValueError, IndexError):
        logging.error("Lựa chọn không hợp lệ.")
        return

    ham_chuyen_doi = BANG_CHUYEN_DOI[dinh_dang_nguon][dinh_dang_dich]
    ket_qua = ham_chuyen_doi(str(duong_dan_tep))
    
    if ket_qua:
        logging.info(f"Đã chuyển đổi thành công '{duong_dan_tep}' sang '{ket_qua}'")
    else:
        logging.error(f"Chuyển đổi thất bại cho '{duong_dan_tep}'")

def main():
    while True:
        print("\nChương trình chuyển đổi định dạng tệp")
        print("\nAuthor: tanbaycu")
        print("Vui lòng cung cấp đường dẫn tệp sử dụng \\ cho dấu phân cách thư mục.")
        duong_dan_tep = input("Nhập đường dẫn tệp (hoặc 'q' để thoát): ")
        
        if duong_dan_tep.lower() == 'q':
            print("Cảm ơn bạn đã sử dụng chương trình. Tạm biệt!")
            break
        
        xu_ly_chuyen_doi(duong_dan_tep)
        
        lua_chon = input("\nBạn có muốn chuyển đổi tệp khác không? (y/n): ")
        if lua_chon.lower() != 'y':
            print("Cảm ơn bạn đã sử dụng chương trình. Tạm biệt!")
            break

if __name__ == "__main__":
    main()
