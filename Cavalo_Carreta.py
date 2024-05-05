# Bibliotecas de instalação na primeira execução
#!pip install PyPDF2
#!pip install tabula-pyplot
from PyPDF2 import PdfReader
import pandas as pd
# import tabula

reader = PdfReader("Cavalo.pdf")
page = reader.pages[0]
arquivo = page.extract_text()
print(page.extract_text())
arquivo.to_csv("Cavalo1.txt", index=False, encoding="utf-8")

with open("Cavalo.txt", "w") as arquivo:
    arquivo.write(f'{page}')
