import os
import io
import shutil
import glob
import pandas as pd
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.layout import LAParams, LTTextBoxHorizontal
import spacy
import PyPDF4

from maisimples import extract_metadata

nlp = spacy.load("en_core_web_sm")

# Define o caminho para o diretório que contém os arquivos PDF
pdf_dir = "/Users/flavi/OneDrive/Documentos/TESE/WOKPROJECT"

# Define o caminho para o diretório de saída
output_dir = "/Users/flavi/OneDrive/Documentos/TESE/WOKPROJECT/output"

# Cria o diretório de saída, se ele não existir
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Cria um dataframe vazio para armazenar as informações extraídas
df = pd.DataFrame(columns=['Author', 'Title', 'Year', 'Topic', 'Keywords', 'Abstract'])

# Inicializa uma lista vazia para armazenar os nomes dos arquivos PDF de saída
output_filenames = []

# Loop pelos arquivos PDF no diretório de entrada
for filename in os.listdir(pdf_dir):
    if filename.endswith('.pdf'):
        print(f'A processar... :-)')
        # Open the PDF file and create a PDF reader object
        with open(os.path.join(pdf_dir, filename), 'rb') as file:
            pdf = PyPDF4.PdfFileReader(file)

           

            # function to extract text from a PDF file using PDFMiner
            def extract_text(filepath):
                fp = open(filepath, 'rb')
                rsrcmgr = PDFResourceManager()
                laparams = LAParams()
                fpout = io.StringIO()
                device = LTTextBoxHorizontal()
                interpreter = PDFPageInterpreter(rsrcmgr, device)
                for page in PDFPage.get_pages(fp):
                    interpreter.process_page(page)
                    layout = device.get_result()
                    for lt_obj in layout:
                        fpout.write(lt_obj.get_text())
                text = fpout.getvalue()
                fp.close()
                return text

            # function to analyze a PDF file to extract metadata and text
            def analyze_pdf(filepath):
                metadata = extract_metadata(filepath)
                if not metadata:
                    text = extract_text(filepath)
                    # TODO: analyze text to extract metadata
                else:
                    text = None
                return metadata, text

            metadata, text = analyze_pdf(os.path.join(pdf_dir, filename))

            # Extrai as informações de autor, título, ano, tópico principal, palavras-chave e resumo (abstract)
            doc = nlp(pdf_dir)
            author = metadata.get('/Author') or next((ent.text for ent in doc.ents if ent.label_ == 'Autor'), None)
            title = metadata.get('/Title') or next((ent.text for ent in doc.ents if ent.label_ == 'Titulo'), None)
            year = metadata.get('/CreationDate')[:4] if metadata.get('/CreationDate') else None
            topic = next((ent.text for ent in doc.ents if ent.label_ == 'ORG' or ent.label_ == 'GPE' or ent.label_ == 'NORP'), None)
            keywords = metadata.get('/Keywords')
            abstract = next((ent.text for ent in doc.ents if ent.label_ == 'ABSTRACT'), None)
     
        
        # Adiciona as informações extraídas ao dataframe
        df = df.append({'Author': author, 'Title': title, 'Year': year, 'Topic': topic, 'Keywords': keywords, 'Abstract': abstract}, ignore_index=True)

        # Cria uma cópia do arquivo PDF no diretório de saída com o nome original precedido por sequência numérica
        i = 1
        while True:
            output_filename = str(i) + '_' + filename
            if os.path.exists(os.path.join(output_dir, output_filename)):
                i += 1
            else:
                break
            output_filenames.append(output_filename)
        shutil.copyfile(os.path.join(pdf_dir, filename), os.path.join(output_dir, output_filename))

        # Salva o dataframe como um arquivo Excel no diretório de saída
        df.to_excel(os.path.join(output_dir, 'Analise.xlsx'), index=False)
        print(f'já está! :-)')