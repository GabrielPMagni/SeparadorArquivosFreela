from os import path, listdir as ls, path, system, mkdir
import re
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfpage import PDFTextExtractionNotAllowed
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfdevice import PDFDevice
from pdfminer.layout import LAParams
from pdfminer.converter import PDFPageAggregator
import pandas as pd
import openpyxl

class OriginTypeInterface:
    def handle_file(files, doNext):
        pass

class OriginType2(OriginTypeInterface):
    def handle_file(files, doNext):
        nf_locator = 'Número da\n\nNFS-e\n\n'
        nf_locator1 = 'Nº '
        city_locator = 'Dados do Tomador de Serviços'
        city_locator1 = 'MUNÍCIPIO\n\n'
        for index, file in enumerate(files):
            print(f'{index+1}/{len(files)} - {(index/len(files)) * 100}%')
            file_binary = open(file, 'rb')
            parser = PDFParser(file_binary)
            doc = PDFDocument(parser)
            laparams = LAParams()
            if not doc.is_extractable:
                print('PDFTextExtractionNotAllowed')
            rsrcmgr = PDFResourceManager()
            device = PDFPageAggregator(rsrcmgr, laparams=laparams)
            interpreter = PDFPageInterpreter(rsrcmgr, device)
            for index, page in enumerate(PDFPage.create_pages(doc)):
                if index > 0: break
                interpreter.process_page(page)
                layout = device.get_result()
                layout.analyze(laparams)
                pdf_text: str = layout.groups[0].get_text()

                try:
                    if pdf_text.find(nf_locator) >= 0:
                        nf_position = pdf_text.find(nf_locator) + len(nf_locator)
                    else:
                        nf_position = pdf_text.find(nf_locator1) + len(nf_locator1)
                    nf_content = pdf_text[nf_position:pdf_text.find('\n', nf_position)].strip()
                    
                    if pdf_text.find(city_locator) >= 0:
                        city_position = (pdf_text.find(city_locator) + len(city_locator))
                        city_position = re.search(r'Município\n\n', pdf_text[city_position:-1]).end() + city_position
                    else:
                        city_position = (pdf_text.find(city_locator1) + len(city_locator1))
                except AttributeError:
                    print('DEBUG')
                    print(file)
                    print(city_content)

                    print('--------CONTEÚDO--------')
                    print(pdf_text)

                    print('FIM DEBUG')
                    exit()

                city_content = pdf_text[city_position:pdf_text.find('\n', city_position + 2)].strip()
                try:
                    try:
                        city_content = re.search(r'^(.+?[\-])', city_content).group(0)
                    except AttributeError:
                        pass
                except AttributeError:
                    print('DEBUG')
                    print(file)
                    print(city_content)

                    print('--------CONTEÚDO--------')
                    print(pdf_text)

                    print('FIM DEBUG')
                    exit()

                city_content = re.sub(r'[^\w ]', '', city_content).capitalize()
                city_content = city_content.strip()
                doNext(nf_content, city_content, file)


class OriginType1(OriginTypeInterface):
    def handle_file(files, doNext):
        nf_locator = 'Número:\n\n'
        city_locator = 'Local da Prestação do Serviço:'
        city_locator_1 = 'Endereço Obra:'
        city_locator_2 = 'Natureza da Operação:'
        for index, file in enumerate(files):
            print(f'{index+1}/{len(files)} - {(index/len(files)) * 100}%')
            file_binary = open(file, 'rb')
            parser = PDFParser(file_binary)
            doc = PDFDocument(parser)
            laparams = LAParams()
            if not doc.is_extractable:
                print('PDFTextExtractionNotAllowed')
            rsrcmgr = PDFResourceManager()
            device = PDFPageAggregator(rsrcmgr, laparams=laparams)
            interpreter = PDFPageInterpreter(rsrcmgr, device)
            for index, page in enumerate(PDFPage.create_pages(doc)):
                if index > 0: break
                interpreter.process_page(page)
                layout = device.get_result()
                layout.analyze(laparams)
                pdf_text: str = layout.groups[0].get_text()

                nf_position = pdf_text.find(nf_locator) + len(nf_locator)
                nf_content = pdf_text[nf_position:pdf_text.find('\n', nf_position)].strip()

                city_position = (pdf_text.find(city_locator) + len(city_locator))
                city_position = re.search(r'\d?\d/\d\d\d\d', pdf_text[city_position:-1]).end() + city_position
                city_content = pdf_text[city_position:pdf_text.find('\n', city_position + 2)].strip()
                if city_content == 'A.R.T:':
                    city_position = (pdf_text.find(city_locator_1) + len(city_locator_1))              
                    city_content = pdf_text[city_position:pdf_text.find('\n', city_position + 2)].strip()
                try:
                    try:
                        city_content = re.search(r'^(.+?[\-])', city_content).group(0)
                    except AttributeError:
                        city_position = (pdf_text.find(city_locator_2) + len(city_locator_2))              
                        city_content = pdf_text[city_position:pdf_text.find('\n', city_position + 2)].strip()
                except AttributeError:
                    print('DEBUG')
                    print(file)
                    print(city_content)

                    print('--------CONTEÚDO--------')
                    print(pdf_text)

                    print('FIM DEBUG')
                    break

                city_content = re.sub(r'[^\w ]', '', city_content).capitalize()
                doNext(nf_content, city_content, file)

class PDFManager:
    def __init__(self, folder, origin_type: OriginTypeInterface):
        self.set_origin_type(origin_type)
        self.files = []
        self.nome_pasta_onde_salvar = 'final'
        self.folder = folder
        self.table_file = self.get_table_file()
        self.list_folder_files(self.folder)
        self.get_file_data()
        self.generate_table()


    def set_origin_type(self, origin_type: OriginTypeInterface):
        self.origin_type = origin_type

    def generate_table(self):
        self.table_file.to_excel('Relação de Notas x Cidades.xlsx')


    def get_table_file(self):
        return pd.DataFrame({'Número da nota': [], 'Cidade': []})
    
    
    def get_file_data(self):
        self.origin_type.handle_file(self.files, self.organize_files)


    def organize_files(self, nf_number: str, nf_city: str, file):
            pasta_completa_para_salvar = f'{self.nome_pasta_onde_salvar}{path.sep}{nf_city}'
            new_row = pd.DataFrame({'Número da nota': [nf_number], 'Cidade': [nf_city]}, columns=self.table_file.columns)
            self.table_file = pd.concat([self.table_file, new_row], ignore_index=True)
            try:
                mkdir(pasta_completa_para_salvar)
            except FileExistsError:
                pass

            try:
                system(f'cp "{file}" "{pasta_completa_para_salvar}{path.sep}{nf_number}.pdf"')
            except FileNotFoundError:
                print('Erro, arquivo não encontrado')
    

    def list_folder_files(self, dir, debug=False):
        if debug:
            print('Listando diretórios...')
        try:
            for item in ls(dir):
                d = path.join(dir, item)
                if path.isdir(d):
                    if debug:
                        print('Pasta encontrada')
                    self.list_folder_files(d, debug)
                else:
                    if debug:
                        print('Arquivo Encontrado')
                    self.files.append(str(d))
            else:
                if len(self.files) == 0:
                    print('Não encontrados arquivos')
                    exit(3)
                    
        except PermissionError:
            if debug:
                print('Permissão Negada à Pasta')
        except Exception as e:
            if debug:
                print('Erro não tratado list_folder_files: '+str(e))
            

PDFManager('arquivos', OriginType2)