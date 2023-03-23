import os
import glob
from PyPDF3 import PdfFileReader
import pandas as pd
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.layout import LAParams
from pdfminer.converter import TextConverter
from io import StringIO
import PyPDF2


# função para extrair texto do pdf usando PyPDF2
def extract_text_PyPDF2(pdf_path):
    with open(pdf_path, 'rb') as f:
        pdf_reader = PyPDF2.PdfFileReader(f)
        num_pages = pdf_reader.getNumPages()
        text = ''
        for page_num in range(num_pages):
            page = pdf_reader.getPage(page_num)
            text += page.extractText()
        return text

# função para extrair texto do pdf usando pdfminer
def extract_text_pdfminer(pdf_path):
    resource_manager = PDFResourceManager()
    string_buffer = StringIO()
    layout_params = LAParams()
    device = TextConverter(resource_manager, string_buffer, laparams=layout_params)
    with open(pdf_path, 'rb') as f:
        interpreter = PDFPageInterpreter(resource_manager, device)
        for page in PDFPage.get_pages(f):
            interpreter.process_page(page)
    text = string_buffer.getvalue()
    device.close()
    string_buffer.close()
    return text

# função para extrair metadata do pdf usando PyPDF2
def extract_metadata_PyPDF2(pdf_path):
    with open(pdf_path, 'rb') as f:
        pdf_reader = PyPDF2.PdfFileReader(f)
        metadata = pdf_reader.getDocumentInfo()
        return metadata

# get list of PDF files in directory
pdf_files = glob.glob(os.path.join(pdf_path, '*.pdf'))

# loop through PDF files and extract metadata
for pdf_file in pdf_files:
    with open(pdf_file, 'rb') as f:
        pdf_reader = PdfFileReader(f)
        metadata = pdf_reader.getDocumentInfo()
        author = metadata.author
        title = metadata.title
        year = metadata.year
        keywords = metadata.keywords
        abstract = metadata.abstract



   # diretório com os arquivos pdf
pdf_dir = '/Users/flavi/OneDrive/Documentos/TESE/WOKPROJECT'

#lista com o nome dos arquivos pdf
pdf_files = os.listdir(pdf_dir)

#lista com as informações dos pdfs
pdf_info = []

for pdf_file in pdf_files:
    pdf_path = os.path.join(pdf_dir, pdf_file)
author, title, date, topic, keywords, abstract = extract_pdf_info(pdf_path)
pdf_info.append([pdf_file, author, title, date, topic, keywords, abstract])

#criar tabela com as informações
df = pd.DataFrame(pdf_info, columns=['File Name', 'Author', 'Title', 'Date', 'Topic', 'Keywords', 'Abstract'])

#salvar tabela em um arquivo excel
excel_file = 'pdf_info.xlsx'
df.to_excel(excel_file, index=False)

#criar novo diretório com cópias dos arquivos pdf
pdf_new_dir = '/Users/flavi/OneDrive/Documentos/TESE/WOKPROJECT/teste'
os.makedirs(pdf_new_dir, exist_ok=True)
for i, pdf_file in enumerate(pdf_files):
    pdf_path = os.path.join(pdf_dir, pdf_file)
pdf_new_path = os.path.join(pdf_new_dir, f'{i+1} - {pdf_file}')
os.copy(pdf_path, pdf_new_path)


