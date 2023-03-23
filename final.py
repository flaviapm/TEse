import os
import shutil
import PyPDF3
import pandas as pd
import nltk
import re

from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
from collections import Counter
from PyPDF3 import PdfFileReader
from tika import parser
import matplotlib.pyplot as plt




# Diretórios de entrada e saída
input_dir = '/Users/flavi/OneDrive/Documentos/TESE/WOKPROJECT'
output_dir = '/Users/flavi/OneDrive/Documentos/TESE/WOKPROJECT/teste'

# Criação do diretório de saída, caso ele não exista
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Criação do dataframe para armazenar as informações
df = pd.DataFrame(columns=['Autor', 'Título', 'Ano', 'Tópico Principal', 'Keywords', 'Resumo'])

# Contador para numeração dos arquivos
file_num = 1

# Loop pelos arquivos no diretório de entrada
for filename in os.listdir(input_dir):
    if filename.endswith(".pdf"):
        filepath = os.path.join(input_dir, filename)
        try:
            # código para abrir e ler o arquivo PDF    
            # Leitura do arquivo PDF
            with open(filepath, 'rb') as f:
                pdf_reader = PyPDF3.PdfFileReader(f)
                # faça o que quiser com o objeto pdf_reader
                pdf_info = pdf_reader.getDocumentInfo()
                print(pdf_info)

                # Extração das informações de metadata do arquivo
                author = pdf_info.author if pdf_info.author else ''
                title = pdf_info.title if pdf_info.title else ''
                year = pdf_info.get('/CreationDate').getObject().split('-')[0][2:] if '/CreationDate' in pdf_info else ''
                topic = pdf_info.subject if pdf_info.subject else ''
                keywords = pdf_info.get('/Keywords') if '/Keywords' in pdf_info else ''
                abstract = pdf_info.get('/Abstract') if '/Abstract' in pdf_info else ''

        except FileNotFoundError:
            print(f"O arquivo PDF em {filepath} não foi encontrado")
            continue
        except PyPDF3.utils.PdfReadError:
            print(f"Skipping encrypted file {filepath}")
            continue
        except Exception as e:
            print(f"Erro ao abrir o arquivo PDF em {filepath}: {e}")
            continue

        # Se a metadata não estiver disponível, utilizar a NLTK para extrair informações do texto
        if not all([author, title, year, topic, keywords, abstract]):
            raw_text = parser.from_file(filepath)['content']

            # Download dos recursos necessários da NLTK
            nltk.download('punkt')
            nltk.download('stopwords')

            # Processamento do texto
            stop_words = set(stopwords.words('english') + stopwords.words('portuguese'))
            word_tokens = word_tokenize(raw_text)
            filtered_tokens = [word for word in word_tokens if word.lower() not in stop_words]
            processed_text = ' '.join(filtered_tokens)

            # Extração das informações
            author = ''
            title = ''
            year = ''
            topic = ''
            keywords = ''
            abstract = ''
            sentences = nltk.sent_tokenize(processed_text)
            for sent in sentences:
                if 'author' in sent.lower():
                    parts = sent.split(':')
                    if len(parts) > 1:
                        author = parts[1].strip()
                elif 'title' in sent.lower():
                    parts = sent.split(':')
                    if len(parts) > 1:
                        title = parts[1].strip()
                elif 'year' in sent.lower():
                    parts = sent.split(':')
                    if len(parts) > 1:
                        year = parts[1].strip()
                elif 'topic' in sent.lower():
                    parts = sent.split(':')
                    if len(parts) > 1:
                        topic = parts[1].strip()
        
        # Adição das informações ao dataframe
        df = pd.DataFrame(pdf_info, columns=['File Name', 'Author', 'Title', 'Date', 'Topic', 'Keywords', 'Abstract'])
        df = df.append({
            'Autor': author,
            'Título': title,
            'Ano': year,
            'Tópico Principal': topic,
            'Keywords': keywords,
            'Resumo': abstract
        }, ignore_index=True)
        
    # Definição das funções para extração de informações
    def extract_author(text):
        author_regex = r"Author[:\s]*(.*)"
        match = re.search(author_regex, text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        return ""

    def extract_title(text):
        title_regex = r"Title[:\s]*(.*)"
        match = re.search(title_regex, text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        return ""

    def extract_date(text):
        date_regex = r"Date[:\s]*(.*)"
        match = re.search(date_regex, text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        return ""

    def extract_topic(text):
        topic_regex = r"Topic[:\s]*(.*)"
        match = re.search(topic_regex, text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        return ""

    def extract_keywords(text):
        keywords_regex = r"Keywords[:\s]*(.*)"
        match = re.search(keywords_regex, text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        return ""

    def extract_abstract(text):
        abstract_regex = r"Abstract[:\s]*(.*)"
        match = re.search(abstract_regex, text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        return ""

        #criar tabela com as informações
        df = pd.DataFrame(pdf_info, columns=['File Name', 'Author', 'Title', 'Date', 'Topic', 'Keywords', 'Abstract'])

       
        # Definição das funções para extração de informações
    def extract_author(text):
        author_regex = r"Author[:\s]*(.*)"
        if match := re.search(author_regex, text, re.IGNORECASE):
            return match[1].strip()
        return ""

    def extract_title(text):
        title_regex = r"Title[:\s]*(.*)"
        match = re.search(title_regex, text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        return ""

    def extract_date(text):
        date_regex = r"Date[:\s]*(.*)"
        match = re.search(date_regex, text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        return ""

    def extract_topic(text):
        topic_regex = r"Topic[:\s]*(.*)"
        match = re.search(topic_regex, text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        return ""

    def extract_keywords(text):
        keywords_regex = r"Keywords[:\s]*(.*)"
        match = re.search(keywords_regex, text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        return ""

    def extract_abstract(text):
        abstract_regex = r"Abstract[:\s]*(.*)"
        match = re.search(abstract_regex, text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        return ""

       # Salvar dataframe em um arquivo Excel
    excel_file_name = 'Analise_Sistematica_Literatura.xlsx'
    writer = pd.ExcelWriter(excel_file_name, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Artigos', index=False)

    # Adição de uma nova sheet com as palavras chaves mais utilizadas
    keywords = ' '.join(df['Keywords'].tolist())
    keywords_list = keywords.split()
    keywords_count = Counter(keywords_list)
    keywords_dict = {'Palavras-Chaves': list(keywords_count.keys()), 'Frequência': list(keywords_count.values())}
    keywords_df = pd.DataFrame.from_dict(keywords_dict)
    keywords_df = keywords_df.sort_values(by='Frequência', ascending=False)
    keywords_df.to_excel(writer, sheet_name='Palavras-Chaves', index=False)

    # Salvar o arquivo Excel
    writer.save()

    # Fechar o arquivo Excel
    writer.close()

    # Criação da cópia do arquivo no diretório de saída
    new_filename = f"{file_num}_{filename}"
    new_filepath = os.path.join(output_dir, new_filename)
    shutil.copyfile(filepath, new_filepath)
     
    file_num += 1