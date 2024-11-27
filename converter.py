import os #ver12.13.58 docx - xlsx, add try 
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

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def safe_convert(func: Callable) -> Callable:
    """Decorator to handle exceptions in conversion functions."""
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            logging.error(f"Error in {func.__name__}: {str(e)}")
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
        logging.error(f"Error when using pdf2docx: {str(e)}")
        logging.info("Trying alternative method...")
        
        try:
            pdf = PdfReader(pdf_path)
            doc = Document()
            
            for page in pdf.pages:
                text = page.extract_text()
                doc.add_paragraph(text)
            
            doc.save(docx_path)
            return str(docx_path)
        except Exception as e:
            logging.error(f"Error when using alternative method: {str(e)}")
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
    df = pd.read_excel(xlsx_path)
    
    doc = Document()
    for column in df.columns:
        doc.add_heading(column, level=1)
        for value in df[column]:
            doc.add_paragraph(str(value))
        doc.add_paragraph()  # Add a blank line between columns
    
    doc.save(docx_path)
    return str(docx_path)

@safe_convert
def convert_docx_to_xlsx(docx_path: str) -> str:
    xlsx_path = Path(docx_path).with_suffix('.xlsx')
    doc = Document(docx_path)
    
    data = []
    for paragraph in doc.paragraphs:
        data.append([paragraph.text])
    
    df = pd.DataFrame(data)
    df.to_excel(xlsx_path, index=False, header=False)
    return str(xlsx_path)

@safe_convert
def convert_xlsx_to_pdf(xlsx_path: str) -> str:
    pdf_path = Path(xlsx_path).with_suffix('.pdf')
    df = pd.read_excel(xlsx_path)
    
    pdf = canvas.Canvas(str(pdf_path), pagesize=letter)
    y = 750  # Starting y-coordinate
    for column in df.columns:
        pdf.drawString(100, y, column)
        y -= 20
        for value in df[column]:
            pdf.drawString(120, y, str(value))
            y -= 15
            if y < 50:  # Start a new page if we're near the bottom
                pdf.showPage()
                y = 750
    pdf.save()
    
    return str(pdf_path)

@safe_convert
def convert_xlsx_to_csv(xlsx_path: str) -> str:
    csv_path = Path(xlsx_path).with_suffix('.csv')
    df = pd.read_excel(xlsx_path)
    df.to_csv(csv_path, index=False)
    return str(csv_path)

@safe_convert
def convert_txt_to_docx(txt_path: str) -> str:
    docx_path = Path(txt_path).with_suffix('.docx')
    with open(txt_path, 'r', encoding='utf-8') as file:
        text = file.read()
    
    doc = Document()
    for paragraph in text.split('\n'):
        doc.add_paragraph(paragraph)
    doc.save(docx_path)
    
    return str(docx_path)

@safe_convert
def convert_txt_to_pdf(txt_path: str) -> str:
    pdf_path = Path(txt_path).with_suffix('.pdf')
    with open(txt_path, 'r', encoding='utf-8') as file:
        text = file.read()
    
    pdf = canvas.Canvas(str(pdf_path), pagesize=letter)
    y = 750  # Starting y-coordinate
    for line in text.split('\n'):
        pdf.drawString(100, y, line)
        y -= 15
        if y < 50:  # Start a new page if we're near the bottom
            pdf.showPage()
            y = 750
    pdf.save()
    
    return str(pdf_path)

@safe_convert
def convert_txt_to_md(txt_path: str) -> str:
    md_path = Path(txt_path).with_suffix('.md')
    
    with open(txt_path, 'r', encoding='utf-8') as file:
        txt_content = file.read()
    
    # Simple conversion: assume each line is a paragraph
    md_content = '\n\n'.join(txt_content.split('\n'))
    
    with open(md_path, 'w', encoding='utf-8') as file:
        file.write(md_content)
    
    return str(md_path)

@safe_convert
def convert_csv_to_xlsx(csv_path: str) -> str:
    xlsx_path = Path(csv_path).with_suffix('.xlsx')
    df = pd.read_csv(csv_path)
    df.to_excel(xlsx_path, index=False)
    return str(xlsx_path)

@safe_convert
def convert_pptx_to_pdf(pptx_path: str) -> str:
    from win32com import client
    pdf_path = Path(pptx_path).with_suffix('.pdf')
    
    powerpoint = client.Dispatch("Powerpoint.Application")
    deck = powerpoint.Presentations.Open(pptx_path)
    deck.SaveAs(pdf_path, 32)  # 32 is the PDF format code
    deck.Close()
    powerpoint.Quit()
    
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
        doc.add_page_break()
    
    doc.save(docx_path)
    return str(docx_path)

@safe_convert
def convert_md_to_txt(md_path: str) -> str:
    txt_path = Path(md_path).with_suffix('.txt')
    
    with open(md_path, 'r', encoding='utf-8') as file:
        md_content = file.read()
    
    h = html2text.HTML2Text()
    h.ignore_links = True
    txt_content = h.handle(markdown.markdown(md_content))
    
    with open(txt_path, 'w', encoding='utf-8') as file:
        file.write(txt_content)
    
    return str(txt_path)

@safe_convert
def convert_md_to_html(md_path: str) -> str:
    html_path = Path(md_path).with_suffix('.html')
    
    with open(md_path, 'r', encoding='utf-8') as file:
        md_content = file.read()
    
    html_content = markdown.markdown(md_content, extensions=['extra'])
    
    with open(html_path, 'w', encoding='utf-8') as file:
        file.write(f"<html><body>{html_content}</body></html>")
    
    return str(html_path)
@safe_convert
def convert_pdf_to_txt(pdf_path: str) -> str:
    txt_path = Path(pdf_path).with_suffix('.txt')
    pdf = PdfReader(pdf_path)
    with open(txt_path, 'w', encoding='utf-8') as file:
        for page in pdf.pages:
            file.write(page.extract_text())
    return str(txt_path)
# Update the CONVERSION_MAP with the new and improved functions
CONVERSION_MAP: Dict[str, Dict[str, Callable]] = {
    '.pdf': {
        '.docx': convert_pdf_to_docx,
        '.xlsx': convert_pdf_to_xlsx,
        '.txt': convert_pdf_to_txt,
    },
    '.xlsx': {
        '.docx': convert_xlsx_to_docx,
        '.pdf': convert_xlsx_to_pdf,
        '.csv': convert_xlsx_to_csv,
    },
    '.docx': {
        '.pdf': convert_docx_to_pdf,
        '.xlsx': convert_docx_to_xlsx,
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
        logging.error(f"'{file_path}' is not a valid file or does not exist.")
        return

    src_ext = file_path.suffix.lower()
    if src_ext not in CONVERSION_MAP:
        logging.error(f"Source file format not supported: {src_ext}")
        return

    print(f"Available conversion options for {src_ext}:")
    for i, target_ext in enumerate(CONVERSION_MAP[src_ext].keys(), 1):
        print(f"{i}. {src_ext[1:].upper()} to {target_ext[1:].upper()}")

    choice = input("Enter your choice number: ")
    try:
        choice = int(choice)
        target_ext = list(CONVERSION_MAP[src_ext].keys())[choice - 1]
    except (ValueError, IndexError):
        logging.error("Invalid choice.")
        return

    conversion_func = CONVERSION_MAP[src_ext][target_ext]
    result = conversion_func(str(file_path))
    
    if result:
        logging.info(f"Successfully converted '{file_path}' to '{result}'")
    else:
        logging.error(f"Conversion failed for '{file_path}'")

def main():
    while True:
        print("\nFile Format Conversion Program")
        print("Please provide the file path using \\ for directory separators.")
        file_path = input("Enter the file path (or 'q' to quit): ")
        
        if file_path.lower() == 'q':
            print("Thank you for using the program. Goodbye!")
            break
        
        handle_conversion(file_path)
        
        choice = input("\nDo you want to convert another file? (y/n): ")
        if choice.lower() != 'y':
            print("Thank you for using the program. Goodbye!")
            break

if __name__ == "__main__":
    main()




