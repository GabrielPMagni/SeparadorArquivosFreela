import mimetypes
import re

from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams
from pdfminer.pdfdevice import PDFDevice
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFPageInterpreter, PDFResourceManager
from pdfminer.pdfpage import PDFPage, PDFTextExtractionNotAllowed
from pdfminer.pdfparser import PDFParser
import pytesseract
from pdf2image import convert_from_path

from pdfExceptions import ErrorOnPDFHandle, IncorrectMimeType


class extractMethodInterface:
    def __init__(self) -> None:
        pass

    def execute(self, file: str) -> list[int]:
        pass

    @staticmethod
    def isValidMimeTypeOrError(file: str,  expectedMimeType: str = '*'):
        if expectedMimeType == '*':
            return True
        guessedMimeType = mimetypes.guess_type(file)[0]
        if guessedMimeType != None and expectedMimeType == guessedMimeType:
            return True
        raise IncorrectMimeType(expectedMimeType, guessedMimeType)


def getPDFText(file: str, numberOfPages: str = 1) -> list[str]:
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
        pdf_text: str = ''
        try:
            pdf_text = layout.groups[0].get_text()
        except TypeError as e:
            pdf_text = getPDFTextAsImage(file)
        pagesText.append(pdf_text)
    return pagesText


def getPDFTextAsImage(file):
    images = convert_from_path(file)
    for image in images:
        text = pytesseract.image_to_string(image, config='--psm 6 -l por')
        return text

class extractMethod1(extractMethodInterface):
    def __init__(self) -> None:
        self.expectedMimeType = 'application/pdf'
        self.nf_locator = 'Número da\n\nNFS-e\n\n'
        self.city_locator = 'Local da Prestação\n\n'

    def execute(self, file: str) -> list[str]:
        super().isValidMimeTypeOrError(file, self.expectedMimeType)
        pdfContentArray = getPDFText(file, 1)
        for content in pdfContentArray:
            try:
                if content.find(self.nf_locator) < 0 or content.find(self.city_locator) < 0:
                    raise AttributeError
                nf_position = content.find(self.nf_locator) + len(self.nf_locator)
                nf_content = content[nf_position:content.find('\n', nf_position)].strip()
                
                city_position = (content.find(self.city_locator) + len(self.city_locator))
                city_position = re.search(r'(\d+\n\n)?', content[city_position:-1]).end() + city_position
                city_content = content[city_position:content.find(' -', city_position + 2)].strip()
                city_content = re.sub(r'[^\w ]', '', city_content).capitalize()
                city_content = city_content.strip()
            except AttributeError:
                raise ErrorOnPDFHandle(content, [f'file {file}'])

        return [nf_content, city_content, file]


class extractMethod2(extractMethodInterface):
    def __init__(self) -> None:
        self.expectedMimeType = 'application/pdf'
        self.nf_locator = 'Número:\n\n'
        self.city_locator = 'Local da Prestação do Serviço:'

    def execute(self, file: str) -> list[str]:
        super().isValidMimeTypeOrError(file, self.expectedMimeType)
        pdfContentArray = getPDFText(file, 1)
        for content in pdfContentArray:
            try:
                if content.find(self.nf_locator) < 0 or content.find(self.city_locator) < 0:
                    raise AttributeError
                nf_position = content.find(self.nf_locator) + len(self.nf_locator)
                nf_content = content[nf_position:content.find('\n', nf_position)].strip()
                
                city_position = (content.find(self.city_locator) + len(self.city_locator))
                city_position = re.search(r'\d?\d/\d\d\d\d', content[city_position:-1]).end() + city_position
                city_content = content[city_position:content.find('\n', city_position + 2)].strip()
                city_content = re.search(r'^(.+?[\-])', city_content).group(0)
                city_content = re.sub(r'[^\w ]', '', city_content).capitalize()
                city_content = city_content.strip()
            except AttributeError:
                raise ErrorOnPDFHandle(content, [f'file {file}'])

        return [nf_content, city_content, file]


class extractMethod3(extractMethodInterface):
    def __init__(self) -> None:
        self.expectedMimeType = 'application/pdf'
        self.nf_locator = 'Número:\n\n'
        self.city_locator = 'Endereço Obra:'

    def execute(self, file: str) -> list[str]:
        super().isValidMimeTypeOrError(file, self.expectedMimeType)
        pdfContentArray = getPDFText(file, 1)
        for content in pdfContentArray:
            try:
                if content.find(self.nf_locator) < 0 or content.find(self.city_locator) < 0:
                    raise AttributeError
                nf_position = content.find(self.nf_locator) + len(self.nf_locator)
                nf_content = content[nf_position:content.find('\n', nf_position)].strip()
                
                city_position = (content.find(self.city_locator) + len(self.city_locator))
                city_content = content[city_position:content.find('\n', city_position + 2)].strip()
                city_content = re.search(r'^(.+?[\-])', city_content).group(0)
                city_content = re.sub(r'[^\w ]', '', city_content).capitalize()
                city_content = city_content.strip()
            except AttributeError:
                raise ErrorOnPDFHandle(content, [f'file {file}'])

        return [nf_content, city_content, file]


class extractMethod3(extractMethodInterface):
    def __init__(self) -> None:
        self.expectedMimeType = 'application/pdf'
        self.nf_locator = 'Número:\n\n'
        self.city_locator = 'Natureza da Operação:'

    def execute(self, file: str) -> list[str]:
        super().isValidMimeTypeOrError(file, self.expectedMimeType)
        pdfContentArray = getPDFText(file, 1)
        for content in pdfContentArray:
            try:
                if content.find(self.nf_locator) < 0 or content.find(self.city_locator) < 0:
                    raise AttributeError
                nf_position = content.find(self.nf_locator) + len(self.nf_locator)
                nf_content = content[nf_position:content.find('\n', nf_position)].strip()
                
                city_position = (content.find(self.city_locator) + len(self.city_locator))
                city_position = re.search(r'\d?\d/\d\d\d\d', content[city_position:-1]).end() + city_position
                city_content = content[city_position:content.find('\n', city_position + 2)].strip()
                city_content = re.search(r'^(.+?[\-])', city_content).group(0)
                city_content = re.sub(r'[^\w ]', '', city_content).capitalize()
                city_content = city_content.strip()
            except AttributeError:
                raise ErrorOnPDFHandle(content, [f'file {file}'])

        return [nf_content, city_content, file]
    

class extractMethod4(extractMethodInterface):
    def __init__(self) -> None:
        self.expectedMimeType = 'application/pdf'
        self.nf_locator = 'Número:\n\n'
        self.city_locator = 'Natureza da Operação:'

    def execute(self, file: str) -> list[str]:
        super().isValidMimeTypeOrError(file, self.expectedMimeType)
        pdfContentArray = getPDFText(file, 1)
        for content in pdfContentArray:
            try:
                if content.find(self.nf_locator) < 0 or content.find(self.city_locator) < 0:
                    raise AttributeError
                nf_position = content.find(self.nf_locator) + len(self.nf_locator)
                nf_content = content[nf_position:content.find('\n', nf_position)].strip()
                
                city_position = (content.find(self.city_locator) + len(self.city_locator))
                city_position = re.search(r'\d?\d/\d\d\d\d', content[city_position:-1]).end() + city_position
                city_content = content[city_position:content.find('\n', city_position + 2)].strip()
                city_content = re.search(r'^(.+?[\-])', city_content).group(0)
                city_content = re.sub(r'[^\w ]', '', city_content).capitalize()
                city_content = city_content.strip()
            except AttributeError:
                raise ErrorOnPDFHandle(content, [f'file {file}'])

        return [nf_content, city_content, file]
    

class extractMethod5(extractMethodInterface):
    def __init__(self) -> None:
        self.expectedMimeType = 'application/pdf'
        self.nf_locator = 'Nº '
        self.city_locator = 'MUNÍCIPIO\n\n'

    def execute(self, file: str) -> list[str]:
        super().isValidMimeTypeOrError(file, self.expectedMimeType)
        pdfContentArray = getPDFText(file, 1)
        for content in pdfContentArray:
            try:
                if content.find(self.nf_locator) < 0 or content.find(self.city_locator) < 0:
                    raise AttributeError
                nf_position = content.find(self.nf_locator) + len(self.nf_locator)
                nf_content = content[nf_position:content.find('\n', nf_position)].strip()
                
                city_position = (content.find(self.city_locator) + len(self.city_locator))
                city_content = content[city_position:content.find('\n', city_position + 2)].strip()
                city_content = re.sub(r'[^\w ]', '', city_content).capitalize()
                city_content = city_content.strip()
            except AttributeError:
                raise ErrorOnPDFHandle(content, [f'file {file}'])

        return [nf_content, city_content, file]
    

class extractMethod6(extractMethodInterface):
    def __init__(self) -> None:
        self.expectedMimeType = 'application/pdf'
        self.nf_locator = 'Número:\n\n'
        self.city_locator = 'Endereço Obra:\n\n'

    def execute(self, file: str) -> list[str]:
        super().isValidMimeTypeOrError(file, self.expectedMimeType)
        pdfContentArray = getPDFText(file, 1)
        for content in pdfContentArray:
            try:
                if content.find(self.nf_locator) < 0 or content.find(self.city_locator) < 0:
                    raise AttributeError
                nf_position = content.find(self.nf_locator) + len(self.nf_locator)
                nf_content = content[nf_position:content.find('\n', nf_position)].strip()
                
                city_position = (content.find(self.city_locator) + len(self.city_locator))
                city_content = content[city_position:content.find('-', city_position + 2)].strip()
                city_content = re.sub(r'[^\w ]', '', city_content).capitalize()
                city_content = city_content.strip()
            except AttributeError:
                raise ErrorOnPDFHandle(content, [f'file {file}'])

        return [nf_content, city_content, file]
    

class extractMethod7(extractMethodInterface):
    def __init__(self) -> None:
        self.expectedMimeType = 'application/pdf'
        self.nf_locator = 'Número:\n\n'
        self.city_locator = 'Natureza da Operação:\n\n'

    def execute(self, file: str) -> list[str]:
        super().isValidMimeTypeOrError(file, self.expectedMimeType)
        pdfContentArray = getPDFText(file, 1)
        for content in pdfContentArray:
            try:
                if content.find(self.nf_locator) < 0 or content.find(self.city_locator) < 0:
                    raise AttributeError
                nf_position = content.find(self.nf_locator) + len(self.nf_locator)
                nf_content = content[nf_position:content.find('\n', nf_position)].strip()
                
                city_position = (content.find(self.city_locator) + len(self.city_locator))
                city_content = content[city_position:content.find('-', city_position + 2)].strip()
                city_content = re.sub(r'[^\w ]', '', city_content).capitalize()
                city_content = city_content.strip()
            except AttributeError:
                raise ErrorOnPDFHandle(content, [f'file {file}'])

        return [nf_content, city_content, file]


class extractMethod8(extractMethodInterface):
    def __init__(self) -> None:
        self.expectedMimeType = 'application/pdf'
        self.nf_locator = 'Número da NFS-e\n\n'
        self.city_locator = 'Cidade - Estado\n\n'

    def execute(self, file: str) -> list[str]:
        super().isValidMimeTypeOrError(file, self.expectedMimeType)
        pdfContentArray = getPDFText(file, 1)
        for content in pdfContentArray:
            try:
                if content.find(self.nf_locator) < 0 or content.find(self.city_locator) < 0:
                    raise AttributeError
                nf_position = content.find(self.nf_locator) + len(self.nf_locator)
                nf_content = content[nf_position:content.find('\n', nf_position)].strip()
                
                city_position = (content.find(self.city_locator) + len(self.city_locator))
                city_content = content[city_position:content.find('-', city_position + 2)].strip()
                city_content = re.sub(r'[^\w ]', '', city_content).capitalize()
                city_content = city_content.strip()
            except AttributeError:
                raise ErrorOnPDFHandle(content, [f'file {file}'])

        return [nf_content, city_content, file]


class extractMethod9(extractMethodInterface):
    def __init__(self) -> None:
        self.expectedMimeType = 'application/pdf'
        self.situation = 'Situação\n\n'

    def execute(self, file: str) -> list[str]:
        super().isValidMimeTypeOrError(file, self.expectedMimeType)
        pdfContentArray = getPDFText(file, 1)
        for content in pdfContentArray:
            try:
                if content.find(self.situation) < 0:
                    raise AttributeError
                situation_position = content.find(self.situation) + len(self.situation)
                situation_content = content[situation_position:content.find('\n', situation_position)].strip()
                if situation_content != 'Cancelada':
                    raise AttributeError
            except AttributeError:
                raise ErrorOnPDFHandle(content, [f'file {file}'])

        return ['cancelada', 'cancelada', file]



class ExtractMethodList:
    def __init__(self) -> None:
        self.methods = [extractMethod1, extractMethod2, extractMethod3, extractMethod4, extractMethod5, extractMethod6, extractMethod7, extractMethod8, extractMethod9]
    
    def getList(self) -> list[extractMethodInterface]:
        return self.methods