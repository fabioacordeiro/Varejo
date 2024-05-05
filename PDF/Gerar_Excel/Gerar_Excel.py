#pip install pypdf2 #necessário instalar
import os # Leitura arquivos, pastas e diretórios
import time #Tempo de execução
import pandas as pd
import tabula #Manipulação de Tabelas 
import PyPDF2 # import PdfMerger, PdfReader, PdfWriter #Manipular pdf
from pathlib import Path #Manipular pastas, diretórios e arquivos
from openpyxl import Workbook, worksheet
from openpyxl.utils.dataframe import dataframe_to_rows

# PyPDF2 para manipular arquivos PDF (PdfMerger)
# PyPDF2 é uma biblioteca de manipulação de arquivos PDF feita em Python puro,
# gratuita e de código aberto. Ela é capaz de ler, manipular, escrever e unir
# dados de arquivos PDF, assim como adicionar anotações, transformar páginas.
#------------------------------------------
#Pegando o tempo de execução do código
tempo_inicial = time.time()
seconds = 0
hour = 0
minutes = 0
print("--- %s segundos ---" % (time.time() - tempo_inicial))
nomeprograma = 'Gerador de Excel'
print(50*'_')
print(f'{nomeprograma:^50}')
print(50*'_')
#-------------Fim Variação tempo -------------------------------
#Incluindo o Try e except que força a execução do código e detalhe
# de erros de execução
try:
    PASTA_RAIZ = Path(__file__).parent
    PASTA_ORIGINAIS = PASTA_RAIZ / 'Arquivo_pdf'
    PASTA_ALTERADOS = PASTA_RAIZ / 'Arquivo_Excel'
    PASTA_FINAL = PASTA_ORIGINAIS/'Teste.pdf'
    print(f'Caminho PDFS_ORIGINAIS:',(PASTA_ORIGINAIS))
    print(f'Caminho Raiz:',(PASTA_RAIZ))
    print(f'Caminho PASTA_FINAL:',(PASTA_FINAL))
    print('Passei ponto 0')
    with open('Teste.pdf', 'rb') as texto1:
        texto2 = PyPDF2.PdfReader(texto1)
        texto_acumulado = ""
        #print('Passei ponto 1')
        for tex in texto2.pages:
            texto_acumulado += tex.extract_text()
            #print('Passei ponto 2')
            #print(texto_acumulado)
        print(f'Qtde páginas:{len(texto2.pages)}')
        print('Passei ponto 3')
        #df = pd.DataFrame(texto_acumulado)
        df = texto_acumulado
        print('-'*65)
        print(df)
        print('-'*65)
        wb  = Workbook()
        ws = wb.active()
        dd = dataframe_to_rows(df, index=True, header=True)
        for r in dataframe_to_rows(df, index=True, header=True):
            ws.append(r)
        wb.save("Novo_Excel.xlsx")
        #df.to_excel("Novo_Excel.xlsx")
        print('Passei ponto 4')
    print('Passei ponto 4.1')
    print('Fim')
except ZeroDivisionError:
    print('Dividiu por zero')
except NameError:
    print('Algum nome não está definido')
except(TypeError, IndexError) as error:
    print('TypeError + IndexError')
    print('MSG:', error)
    print('MSG:', error.__class__.__name__)
except Exception:
    print('ERRO DESCONHECIDO')
else:
    print('Não deu erro')
finally:
    print('Fechar arquivo')
print('Fim')
#----------------------------------------------------
#Imprimindo o tempo inicial menos o final, ou seja, tempo total
seconds = (time.time() - tempo_inicial)
hour = seconds // 3600
minutes = seconds // 60
seconds %= 60
print("--- %s segundos ---" % (time.time() - tempo_inicial))
print("--- %s segundos ---" % ("%d:%02d:%02d" % (hour, minutes, seconds) ))
