# coding: utf-8

from format.excel import XLSX
from format.powerpoint import PPTX, PPT
from format.pdf import PDF
from format.word import DOCX
from format.dummy import File

with open('sample/sample.pdf', 'rb') as f:
    data = File(f.read())

pptx = PDF()
text = pptx.read(data)

with open('text.txt', 'w') as f:
    f.write('\n'.join(text).encode('utf-8'))