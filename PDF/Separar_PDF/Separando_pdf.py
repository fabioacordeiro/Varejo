# https://pypdf2.readthedocs.io/en/3.0.0/_modules/PyPDF2/_reader.html#DocumentInformation

# pip install PyPDF2
# pip install icecream

# Importando as bibliotecas necessárias
from PyPDF2 import PdfReader, PdfWriter
from icecream import ic
from math import trunc, ceil
import os
from pathlib import Path
from time import sleep

PDF_Raiz = Path(__file__).parent
# Nome do arquivo que deseja separar as páginas
Apostila = PDF_Raiz / 'Cavalo.pdf'
print(PDF_Raiz)
print(Apostila)

# sleep(50)

os.system('cls')

newManual = PdfWriter()
# Abrindo o pdf
manual = PdfReader(open(Apostila, 'rb'))
pages = len(manual.pages)
size = 3
total_sections = pages/size
sections = trunc(total_sections)
rest = ceil((total_sections-sections)*size)
lastpage = pages-rest
lenIndex = len(str(sections))
ic(pages, size, total_sections, sections, rest, lastpage)
cont = 0

for pg in manual.pages:
    # Pegando os dados da segunda página
    pagina = manual.pages[cont]
    # Criando o escritor
    writer = PdfWriter()
    # Adicionando a página
    writer.add_page(pagina)
    # Criando o arquivo
    writer.write(f"{PDF_Raiz}\\{cont}.pdf")
    cont = cont + 1
