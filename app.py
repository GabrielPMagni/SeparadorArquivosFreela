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

class PDFManager:
    def __init__(self, folder):
        self.files = []
        self.nome_pasta_onde_salvar = 'final'
        self.folder = folder
        self.list_folder_files()
        self.get_file_data()

    
    def get_file_data(self):
        nf_locator = 'Número:\n\n'
        city_locator = 'Local da Prestação do Serviço:'
        for file in self.files:
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
                pdf_text = layout.groups[0].get_text()

                nf_position = pdf_text.find(nf_locator) + len(nf_locator)
                nf_content = pdf_text[nf_position:pdf_text.find('\n', nf_position)].strip()

                city_position = (pdf_text.find(city_locator) + len(city_locator))              
                city_position = re.search(r'\d?\d/\d\d\d\d', pdf_text[city_position:-1]).end() + city_position
                city_content = pdf_text[city_position:pdf_text.find('\n', city_position+2)].strip()
                city_content = re.sub(r'[^\w\- ]', '', city_content).capitalize()
                self.organize_files(nf_content, city_content, file)


    def organize_files(self, nf_number, nf_city, file):
            pasta_completa_para_salvar = f'{self.nome_pasta_onde_salvar}{path.sep}{nf_city}'
            try:
                mkdir(pasta_completa_para_salvar)
            except FileExistsError:
                pass

            try:
                system(f'copy {file} "{pasta_completa_para_salvar}{path.sep}{nf_number}.pdf"')
            except FileNotFoundError:
                print('Erro, arquivo não encontrado')
    

    def list_folder_files(self, debug=False):
        if debug:
            print('Listando diretórios...')
        try:
            for item in ls(self.folder):
                d = path.join(self.folder, item)
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
            

app = PDFManager('arquivos')