class ErrorOnPDFHandle(AttributeError):
    def __init__(self, pdfText: str, args: list[str]) -> None:
        self.items = args
        self.pdfText = pdfText
        self.groupArgs = '\n'.join(self.items)
        self.message = f'''
Falha ao tratar arquivo do tipo PDF

DEBUG

{self.items}

--------CONTEÚDO--------
{self.pdfText}

FIM DEBUG
        '''

class IncorrectMimeType(ValueError):
    def __init__(self, extpectedType: str, foundType: str) -> None:
        self.expectedType = extpectedType
        self.foundType = foundType
        self.message = f'Tipo de arquivo incorreto para operação. Esperado {self.expectedType}, encontrado {self.foundType}'        
