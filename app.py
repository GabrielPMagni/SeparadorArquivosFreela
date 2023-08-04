from os import listdir as ls
from os import mkdir, path, system
from pathlib import Path
from platform import system as currentOS
import openpyxl
from logging import getLogger, FileHandler, Formatter, DEBUG
import pandas as pd

from extractMethods import (ErrorOnPDFHandle, ExtractMethodList,
                            IncorrectMimeType, execute_pipeline)


fileLogger = getLogger('file_logger')
hd = FileHandler('app.log', 'a+')
hd.setFormatter(Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s'))
fileLogger.addHandler(hd)
fileLogger.setLevel(DEBUG)


def _debug(on=False):
        def _inner_debug(msg: any):
            print(msg)
        
        return _inner_debug if on else lambda _: None


class PDFManager:
    def __init__(self, folder: str,  debug_on=False) -> None:
        self.files: list[str] = []
        self.nome_pasta_onde_salvar: str = 'final'
        self.errorLogFile = open('error.log', '+a', encoding='utf8')
        self.folder: str = folder
        self.debug = self._debug(debug_on)
        self.set_table_file()
        self.list_folder_files(self.folder)
        self.get_file_data()
        self.generate_table()


    def logProgress(self, actual, total):
        print(f'{actual+1}/{total} - {(actual/total) * 100:.3f}%')


    def generate_table(self):
        self.table_file.to_excel('Relação de Notas x Cidades.xlsx')


    def set_table_file(self):
        self.table_file = pd.DataFrame({'Número da nota': [], 'Cidade': []})
    
    
    def get_file_data(self):
        for index, file in enumerate(self.files):
            self.logProgress(index, len(self.files))

            for results in list(execute_pipeline(file, fileLogger)):
                print(results)
                # self.organize_files(opResult., nf_city, file)


    def organize_files(self, nf_number: str, nf_city: str, file):
            if nf_number == 'cancelada' and nf_city == 'cancelada':
                nf_city = 'Notas_Canceladas'
            pasta_completa_para_salvar = f'{self.nome_pasta_onde_salvar}{path.sep}{nf_city}'
            new_row = pd.DataFrame({'Número da nota': [nf_number], 'Cidade': [nf_city]}, columns=self.table_file.columns)
            self.table_file = pd.concat([self.table_file, new_row], ignore_index=True)
            try:
                mkdir(pasta_completa_para_salvar)
            except FileExistsError:
                pass

            try:
                if currentOS() == 'Windows':
                    system(f'copy "{file}" "{pasta_completa_para_salvar}{path.sep}{nf_number}.pdf" 1>NULL')
                else:
                    system(f'cp "{file}" "{pasta_completa_para_salvar}{path.sep}{nf_number}.pdf"')
            except FileNotFoundError:
                print('Erro, arquivo não encontrado')
    
    def list_folder_files(self, dir):
        self.debug('Listando diretórios...')

        try:
            for item in ls(dir):
                d = path.join(dir, item)
                if path.isdir(d):
                    self.debug('Pasta encontrada')
                    self.list_folder_files(d)
                else:
                    self.debug('Arquivo Encontrado')
                    self.files.append(str(d))
            else:
                if len(self.files) == 0:
                    print('Não encontrados arquivos')
                    exit(3)
                    
        except PermissionError:
            self.debug('Permissão Negada à Pasta')
        except Exception as e:
            self.debug('Erro não tratado list_folder_files: ' + e.args)


from argparse import ArgumentParser


parser = ArgumentParser(
                    prog='ProgramName',
                    description='What the program does',
                    epilog='Text at the bottom of help')


parser.add_argument('-f', '--foldername')
parser.add_argument('-D', '--debug', default=False, type=str)


if __name__ == '__main__':
    args = parser.parse_args()

    PDFManager(args.foldername, debug_on=args.debug)
