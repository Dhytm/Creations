
from docx import Document

from datetime import datetime

contrato_ = Document('contrato.docx')

nome = 'DDhytm :|'
text_01 = 'TOMAR CAFÉ'
text_02 = 'Programando em VBA'
text_03 ='Estudando Python'

dictionary_text = {
    'XXXX': nome,
    'YYYY': text_01,
    'ZZZZ': text_02,
    'WWWW': text_03,
    'DD': str(datetime.now().day),
    'MM': str(datetime.now().month),
    'AAAA': str(datetime.now().year),
}

for paragrafo in contrato_.paragraphs:
    for placeholder in dictionary_text:
        if placeholder in paragrafo.text:
            paragrafo.text = paragrafo.text.replace(placeholder, dictionary_text[placeholder])


contrato_.save(f'contrato_novo {nome}.docx')