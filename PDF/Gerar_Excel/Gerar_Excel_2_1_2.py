#pip install pypdf2 #necessário instalar
import os # Leitura arquivos, pastas e diretórios
from pathlib import Path
import chardet #pip install chardet #para ver qual o tipo de dados do csv
import csv
from openpyxl import Workbook, load_workbook
from openpyxl.cell import Cell  # Aqui eu corrijo a tipagem de Cell
from openpyxl.worksheet.worksheet import Worksheet
import time #Tempo de execução
import pandas as pd
import tabula #Manipulação de Tabelas 
import PyPDF2 # import PdfMerger, PdfReader, PdfWriter #Manipular pdf
import PyPDF2.pagerange
from pathlib import Path #Manipular pastas, diretórios e arquivos
#from openpyxl import Workbook, worksheet
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
lista_pdf = []
lista_pdf2 = []
try:
    PASTA_RAIZ = Path(__file__).parent
    PASTA_ORIGINAIS = PASTA_RAIZ / 'Arquivo_pdf'
    CRIADO1 = PASTA_RAIZ / 'Arquivo_csv'
    CRIADO2 = PASTA_RAIZ / 'Arquivo_Excel'
    PASTA_FINAL = PASTA_ORIGINAIS/'Teste.pdf'
    print(' ----- Início ---- ') 
    print('Caminhos onde serão lidos e gravados os arquivos')    
    print(f'Caminho PDFS_ORIGINAIS:\n',(PASTA_ORIGINAIS))
    print(f'Caminho Raiz:\n',(PASTA_RAIZ))
    print(f'Caminho arquivo csv criado:\n',(CRIADO1))
    print(f'Caminho arquivo XlS criado:\n',(CRIADO2))
    print(f'Caminho PASTA_FINAL:\n',(PASTA_FINAL))
    print('Passo 1')
    reader = PyPDF2.PdfReader(PASTA_ORIGINAIS/'Teste.pdf')
    page = len(reader.pages)
    print(f"Quantidade de páginas do pdf:{page}")
    for p in reader.pages:
        #print(p.extract_text())
        lista_pdf.append(p.extract_text())
        
    print('Passo 2')
    print(f"------- Criando arquivo CSV -------")
    with open(CRIADO1/'Teste_1.csv', 'w', newline='\n') as d:
            writer = csv.writer(d)
            writer.writerow(lista_pdf)
    print('Passo 3')  
    # Step 2: Read CSV File in Binary Mode
    with open(CRIADO1/'Teste_1.csv','rb') as f:
        data = f.read()
    
    # Step 3: Detect Encoding using chardet Library
    encoding_result = chardet.detect(data)
    # Step 4: Retrieve Encoding Information
    encoding = encoding_result['encoding']
    # Step 5: Print Detected Encoding Information
    #print("Detected Encoding:", encoding)
    print(f"------- Transformando CSV em Dataframe -------")
    N3 = pd.read_csv(CRIADO1/"Teste_1.csv",encoding='latin-1')
    N4 = []
    for linha in N3:
        N4.append(linha)
    DF0 =pd.DataFrame(N4) 
    print('Passo 4')
    print(f'Transformando Dataframe em Excel (páginas em linhas)') 
    DF0.to_excel(CRIADO2/'Teste_1.xlsx')    
    print('Passo 5')
    print(f'---- Ler Excel e manipular extração ----')
    #N2 = pd.read_csv(CRIADO1/"Teste_1.csv",skiprows=[683,684,685,686,687,
    #        688,689,690,691,692,693,694,695,696,697,698,699,700,701,
    #        702,703,704,705,706,707,708,710,711,712,713,714,715, 716,
    #        717,718,719,720,721,722,723,724,725,726,727,728,729,
    #        730,731,732,733,734,735,736,737,738,739,740, 1705], encoding='ISO-8859-1', delimiter='str')
    #encoding='utf8' 
    #encoding_errors='utf-8'
    #ASCII, UTF-7, UTF-8, UTF-16 e UTF-32, latin-1, cp1252 , ISO-8859-1      
     # abrindo o arquivo CSV e mostrando 5 linhas.
    # skiprows => informar quais as linhas que serão ignoradas
    # nrows => qual a quantidade de linhas que serão percorridas
    # usecols => qual as colunas que serão utilizadas
    # BD1 = pd.read_csv(filepath_or_buffer='C:/Fabio/Python/Carrefour/CTE.csv',
#                      sep='|', nrows=5, index_col=None, skiprows=[1], low_memory=False, usecols=[
#                          1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16,
#                          17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30,
#                          31, 32])
    #encoding='utf8' 
    #encoding_errors='utf-8'
    #ASCII, UTF-7, UTF-8, UTF-16 e UTF-32, latin-1, cp1252
    #df0 = pd.read_csv(filepath_or_buffer='C:\\Desenvolvimento\\Varejo\\PDF\\Gerar_Excel\\Arquivo_csv\\Teste_1.csv', 
     #               encoding='ISO-8859-1', sep='str', skiprows=[68,682,683,684,685,
     #               686,687,688,691,692,693,694,695,696,697,698,699,700,701,
     #               702,703,704,705,706,707,708,709,710,711,712,713,714,715], nrows=None, index_col=None,  
     #               usecols=None)
    print('-'*65)
    
    print('Passei ponto 5')
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
