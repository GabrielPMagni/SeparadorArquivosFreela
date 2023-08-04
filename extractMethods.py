import mimetypes
import re

from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams
from pdfminer.pdfdevice import PDFDevice
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFPageInterpreter, PDFResourceManager
from pdfminer.pdfpage import PDFPage, PDFTextExtractionNotAllowed
from pdfminer.pdfparser import PDFParser

from pdfExceptions import ErrorOnPDFHandle, IncorrectMimeType




from dataclasses import dataclass
from typing import Callable
from pathlib import Path


@dataclass(slots=True)
class OperationResult:
    nf_content: str = str()
    nf_city: str = str()
    origin_file_path: str = str()


@dataclass(slots=True)
class FileAttributes:
    file_name: str = str()
    expectedMimeType: str = str()
    nf_locator: str = str()
    city_locator: str = str()
    situation: str = str()


OperationGuard = Callable[[[str], FileAttributes], bool]
Operation = Callable[[[str], FileAttributes], OperationResult]


@dataclass(slots=True)
class FileOperation:
    fileAttr: FileAttributes    
    operationGuard: OperationGuard
    operation: Operation
    debug: bool



def isValidMimeTypeOrError(file: str,  expectedMimeType: str = '*'):
    if expectedMimeType == '*':
        return True
    guessedMimeType = mimetypes.guess_type(file)[0]
    if guessedMimeType != None and expectedMimeType == guessedMimeType:
        return True
    raise IncorrectMimeType(expectedMimeType, guessedMimeType)


def getPDFText(file, numberOfPages: str = 1) -> list[str]:
    pagesText: list[str] = []
    file_binary = open(file, 'rb')
    parser = PDFParser(file_binary)
    doc = PDFDocument(parser)
    laparams = LAParams()
    if not doc.is_extractable:
        print('PDFTextExtractionNotAllowed')
    rsrcmgr = PDFResourceManager()
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    for pageNumber, page in enumerate(PDFPage.create_pages(doc)):
        if pageNumber >= numberOfPages: break
        interpreter.process_page(page)
        layout = device.get_result()
        layout.analyze(laparams)
        pdf_text: str = layout.groups[0].get_text()
        pagesText.append(pdf_text)
    return pagesText



def find_city_position(content: str, city_locator: str, find_seq: list[int]) -> int:
    last_position = 0

    FINDER_PATTERN: dict[int, Callable[[], int] ] = {
        0: lambda: content.find(city_locator) + len(city_locator),
        1: lambda: re.search(r'(\d+\n\n)?', content[last_position:-1]).end() + last_position,
        2: lambda: re.search(r'\d?\d/\d\d\d\d', content[last_position:-1]).end() + last_position,        
    }

    for f in find_seq:
        last_position = FINDER_PATTERN[f]()
        
    return last_position


def find_city_content(content: str, city_position: int, find_seq: list[int]) -> str:
    last_content = ""

    FINDER_PATTERN: dict[int, Callable[[], str] ] = {
        0: lambda: content[city_position:content.find(' -', city_position + 2)].strip(),
        1: lambda: re.sub(r'[^\w ]', '', last_content).capitalize(),
        2: lambda: content[city_position:content.find('\n', city_position + 2)].strip(),
        3: lambda: re.search(r'^(.+?[\-])', last_content).group(0),
    }
    
    for f in find_seq:
        last_content = FINDER_PATTERN[f]()
        
    return last_content.strip()


def nf_content_by_locator(content: str, nf_locator: str):
    nf_position = content.find(nf_locator) + len(nf_locator)
    return content[nf_position:content.find('\n', nf_position)].strip()


def build_file_operation1() -> FileOperation:    
    def operation(content: str, fAttr: FileAttributes) -> OperationResult:
        nf_content = nf_content_by_locator(content, fAttr.nf_locator)
        
        city_position = find_city_position(content, fAttr.city_locator, [0, 1])
        city_content = find_city_content(content, city_position, [0, 1])

        return OperationResult(nf_city=city_content, nf_content=nf_content, origin_file_path=fAttr.file_name)

    def operation_guard(content, fAttr, ): 
        return content.find(fAttr.nf_locator) < 0 or content.find(fAttr.city_locator) < 0
    
    return FileOperation(
        fileAttr=FileAttributes(
            expectedMimeType = 'application/pdf',
            nf_locator = 'Número da\n\nNFS-e\n\n',
            city_locator = 'Local da Prestação\n\n',
        ),
        operation=operation,
        operationGuard=operation_guard,
        debug=False,
    )



def build_file_operation2() -> FileOperation:    
    def operation(content: str, fAttr: FileAttributes) -> OperationResult:
        nf_content = nf_content_by_locator(fAttr.nf_locator)
                        
        city_position = find_city_position(content, fAttr.city_locator, [0, 2])
        city_content = find_city_content(content, city_position, [2, 3, 1])

        return OperationResult(nf_city=city_content, nf_content=nf_content, origin_file_path=fAttr.file_name)

    def operation_guard(content, fAttr): 
        return content.find(fAttr.nf_locator) < 0 or content.find(fAttr.city_locator) < 0
    
    return FileOperation(
        fileAttr=FileAttributes(
            expectedMimeType = 'application/pdf',
            nf_locator = 'Número:\n\n',
            city_locator = 'Local da Prestação do Serviço:',
        ),
        operation=operation,
        operationGuard=operation_guard,
        debug=False,
    )



def build_file_operation3() -> FileOperation:    
    def operation(content: str, fAttr: FileAttributes) -> OperationResult:
        nf_content = nf_content_by_locator(fAttr.nf_locator)

        city_position = find_city_position(content, fAttr.city_locator, [0])
        city_content = find_city_content(content, city_position, [2, 3, 1])

        return OperationResult(nf_city=city_content, nf_content=nf_content, origin_file_path=fAttr.file_name)

    def operation_guard(content, fAttr): 
        return content.find(fAttr.nf_locator) < 0 or content.find(fAttr.city_locator) < 0
    
    return FileOperation(
        fileAttr=FileAttributes(
            expectedMimeType = 'application/pdf',
            nf_locator = 'Número:\n\n',
            city_locator = 'Endereço Obra:',
        ),
        operation=operation,
        operationGuard=operation_guard,
        debug=False,
    )

# class extractMethod3(extractMethodInterface):
#     def __init__(self) -> None:
#         self.expectedMimeType = 'application/pdf'
#         self.nf_locator = 'Número:\n\n'
#         self.city_locator = 'Natureza da Operação:'

#     def execute(self, file: str) -> list[str]:
#         super().isValidMimeTypeOrError(file, self.expectedMimeType)
#         pdfContentArray = getPDFText(file, 1)
#         for content in pdfContentArray:
#             try:
#                 if content.find(self.nf_locator) < 0 or content.find(self.city_locator) < 0:
#                     raise AttributeError
#                 nf_position = content.find(self.nf_locator) + len(self.nf_locator)
#                 nf_content = content[nf_position:content.find('\n', nf_position)].strip()
                
#                 city_position = (content.find(self.city_locator) + len(self.city_locator))
#                 city_position = re.search(r'\d?\d/\d\d\d\d', content[city_position:-1]).end() + city_position
#                 city_content = content[city_position:content.find('\n', city_position + 2)].strip()
#                 city_content = re.search(r'^(.+?[\-])', city_content).group(0)
#                 city_content = re.sub(r'[^\w ]', '', city_content).capitalize()
#                 city_content = city_content.strip()
#             except AttributeError:
#                 raise ErrorOnPDFHandle(content, [f'file {file}'])

#         return [nf_content, city_content, file]
    

# class extractMethod4(extractMethodInterface):
#     def __init__(self) -> None:
#         self.expectedMimeType = 'application/pdf'
#         self.nf_locator = 'Número:\n\n'
#         self.city_locator = 'Natureza da Operação:'

#     def execute(self, file: str) -> list[str]:
#         super().isValidMimeTypeOrError(file, self.expectedMimeType)
#         pdfContentArray = getPDFText(file, 1)
#         for content in pdfContentArray:
#             try:
#                 if content.find(self.nf_locator) < 0 or content.find(self.city_locator) < 0:
#                     raise AttributeError
#                 nf_position = content.find(self.nf_locator) + len(self.nf_locator)
#                 nf_content = content[nf_position:content.find('\n', nf_position)].strip()
                
#                 city_position = (content.find(self.city_locator) + len(self.city_locator))
#                 city_position = re.search(r'\d?\d/\d\d\d\d', content[city_position:-1]).end() + city_position
#                 city_content = content[city_position:content.find('\n', city_position + 2)].strip()
#                 city_content = re.search(r'^(.+?[\-])', city_content).group(0)
#                 city_content = re.sub(r'[^\w ]', '', city_content).capitalize()
#                 city_content = city_content.strip()
#             except AttributeError:
#                 raise ErrorOnPDFHandle(content, [f'file {file}'])

#         return [nf_content, city_content, file]
    

# class extractMethod5(extractMethodInterface):
#     def __init__(self) -> None:
#         self.expectedMimeType = 'application/pdf'
#         self.nf_locator = 'Nº '
#         self.city_locator = 'MUNÍCIPIO\n\n'

#     def execute(self, file: str) -> list[str]:
#         super().isValidMimeTypeOrError(file, self.expectedMimeType)
#         pdfContentArray = getPDFText(file, 1)
#         for content in pdfContentArray:
#             try:
#                 if content.find(self.nf_locator) < 0 or content.find(self.city_locator) < 0:
#                     raise AttributeError
#                 nf_position = content.find(self.nf_locator) + len(self.nf_locator)
#                 nf_content = content[nf_position:content.find('\n', nf_position)].strip()
                
#                 city_position = (content.find(self.city_locator) + len(self.city_locator))
#                 city_content = content[city_position:content.find('\n', city_position + 2)].strip()
#                 city_content = re.sub(r'[^\w ]', '', city_content).capitalize()
#                 city_content = city_content.strip()
#             except AttributeError:
#                 raise ErrorOnPDFHandle(content, [f'file {file}'])

#         return [nf_content, city_content, file]
    

# class extractMethod6(extractMethodInterface):
#     def __init__(self) -> None:
#         self.expectedMimeType = 'application/pdf'
#         self.nf_locator = 'Número:\n\n'
#         self.city_locator = 'Endereço Obra:\n\n'

#     def execute(self, file: str) -> list[str]:
#         super().isValidMimeTypeOrError(file, self.expectedMimeType)
#         pdfContentArray = getPDFText(file, 1)
#         for content in pdfContentArray:
#             try:
#                 if content.find(self.nf_locator) < 0 or content.find(self.city_locator) < 0:
#                     raise AttributeError
#                 nf_position = content.find(self.nf_locator) + len(self.nf_locator)
#                 nf_content = content[nf_position:content.find('\n', nf_position)].strip()
                
#                 city_position = (content.find(self.city_locator) + len(self.city_locator))
#                 city_content = content[city_position:content.find('-', city_position + 2)].strip()
#                 city_content = re.sub(r'[^\w ]', '', city_content).capitalize()
#                 city_content = city_content.strip()
#             except AttributeError:
#                 raise ErrorOnPDFHandle(content, [f'file {file}'])

#         return [nf_content, city_content, file]
    

# class extractMethod7(extractMethodInterface):
#     def __init__(self) -> None:
#         self.expectedMimeType = 'application/pdf'
#         self.nf_locator = 'Número:\n\n'
#         self.city_locator = 'Natureza da Operação:\n\n'

#     def execute(self, file: str) -> list[str]:
#         super().isValidMimeTypeOrError(file, self.expectedMimeType)
#         pdfContentArray = getPDFText(file, 1)
#         for content in pdfContentArray:
#             try:
#                 if content.find(self.nf_locator) < 0 or content.find(self.city_locator) < 0:
#                     raise AttributeError
#                 nf_position = content.find(self.nf_locator) + len(self.nf_locator)
#                 nf_content = content[nf_position:content.find('\n', nf_position)].strip()
                
#                 city_position = (content.find(self.city_locator) + len(self.city_locator))
#                 city_content = content[city_position:content.find('-', city_position + 2)].strip()
#                 city_content = re.sub(r'[^\w ]', '', city_content).capitalize()
#                 city_content = city_content.strip()
#             except AttributeError:
#                 raise ErrorOnPDFHandle(content, [f'file {file}'])

#         return [nf_content, city_content, file]


# class extractMethod8(extractMethodInterface):
#     def __init__(self) -> None:
#         self.expectedMimeType = 'application/pdf'
#         self.nf_locator = 'Número da NFS-e\n\n'
#         self.city_locator = 'Cidade - Estado\n\n'

#     def execute(self, file: str) -> list[str]:
#         super().isValidMimeTypeOrError(file, self.expectedMimeType)
#         pdfContentArray = getPDFText(file, 1)
#         for content in pdfContentArray:
#             try:
#                 if content.find(self.nf_locator) < 0 or content.find(self.city_locator) < 0:
#                     raise AttributeError
#                 nf_position = content.find(self.nf_locator) + len(self.nf_locator)
#                 nf_content = content[nf_position:content.find('\n', nf_position)].strip()
                
#                 city_position = (content.find(self.city_locator) + len(self.city_locator))
#                 city_content = content[city_position:content.find('-', city_position + 2)].strip()
#                 city_content = re.sub(r'[^\w ]', '', city_content).capitalize()
#                 city_content = city_content.strip()
#             except AttributeError:
#                 raise ErrorOnPDFHandle(content, [f'file {file}'])

#         return [nf_content, city_content, file]


# class extractMethod9(extractMethodInterface):
#     def __init__(self) -> None:
#         self.expectedMimeType = 'application/pdf'
#         self.situation = 'Situação\n\n'

#     def execute(self, file: str) -> list[str]:
#         super().isValidMimeTypeOrError(file, self.expectedMimeType)
#         pdfContentArray = getPDFText(file, 1)
#         for content in pdfContentArray:
#             try:
#                 if content.find(self.situation) < 0:
#                     raise AttributeError
#                 situation_position = content.find(self.situation) + len(self.situation)
#                 situation_content = content[situation_position:content.find('\n', situation_position)].strip()
#                 if situation_content != 'Cancelada':
#                     raise AttributeError
#             except AttributeError:
#                 raise ErrorOnPDFHandle(content, [f'file {file}'])

#         return ['cancelada', 'cancelada', file]



# class ExtractMethodList:
#     def __init__(self) -> None:
#         self.methods = [extractMethod1, extractMethod2, extractMethod3, extractMethod4, extractMethod5, extractMethod6, extractMethod7, extractMethod8, extractMethod9]
    
#     def getList(self) -> list[extractMethodInterface]:
#         return self.methods
    

OPERATION_PIPELINE = [
    build_file_operation1(),
    build_file_operation2(),
    build_file_operation3(),
]


def _operation_error_handler(operation: Operation, fAttr: FileAttributes, log_handler: Callable[[str], None]) -> OperationResult:
    try:
        isValidMimeTypeOrError(fAttr.file_name, fAttr.expectedMimeType)
        pdfContentArray = getPDFText(fAttr.file_name, 1)

        for content in pdfContentArray:
            yield operation(content, fAttr)
    except ErrorOnPDFHandle as err:
        log_handler(err.message)
        exit(1)
    except IncorrectMimeType as err:        
        return log_handler(err.message)
    except AttributeError:
        return log_handler(f'{content} {fAttr.file_name}')
        

def execute_pipeline(file, logger):
    for op_n, op in enumerate(OPERATION_PIPELINE):
        op.fileAttr.file_name = file

        def log_handler(msg):
            if op.debug:
                logger.debug(f'operation {op_n} fail with {msg}')

        yield _operation_error_handler(op.operation, op.fileAttr, log_handler)
