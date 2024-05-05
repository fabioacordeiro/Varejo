import time
#Pegando o tempo de execução do código
tempo_inicial = time.time()
seconds = 0
hour = 0
minutes = 0

import tabula
import pandas as pd
from tabula.io import read_pdf

def pdf_to_excel(pdf_file_path, excel_file_path):
    # Read PDF file
    tables = tabula.read_pdf(pdf_file_path, pages='all')

    # Write each table to a separate sheet in the Excel file
    with pd.ExcelWriter(excel_file_path) as writer:
        for i, table in enumerate(tables):
            table.to_excel(writer, sheet_name=f'Sheet{i+1}')


pdf_to_excel('c:\\Desenvolvimento\\Varejo\\Sample.pdf', 'c:\\Desenvolvimento\\Varejo\\Sample.xlsx')

print("--- %s segundos ---" % (time.time() - tempo_inicial))
seconds = (time.time() - tempo_inicial)
hour = seconds // 3600
minutes = seconds // 60
seconds %= 60

for i in range(0,1000000):
    i = i+1
    print (i)

print('Fim')

print("--- %s segundos ---" % (time.time() - tempo_inicial))
print("--- %s segundos ---" % ("%d:%02d:%02d" % (hour, minutes, seconds) ))
print("Conversion process completed")