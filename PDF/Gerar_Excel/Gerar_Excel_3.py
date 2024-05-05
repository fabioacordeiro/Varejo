#pip install pypdf2 #necessário instalar
#import os # Leitura arquivos, pastas e diretórios
import time #Tempo de execução
import pandas as pd
from pathlib import Path #Manipular pastas, diretórios e arquivos
#from openpyxl import Workbook, load_workbook
#from openpyxl.cell import Cell  # Aqui eu corrijo a tipagem de Cell
#from openpyxl.worksheet.worksheet import Worksheet
#import tabula #Manipulação de Tabelas 
#import PyPDF2 # import PdfMerger, PdfReader, PdfWriter #Manipular pdf
#from openpyxl import Workbook, worksheet
#from openpyxl.utils.dataframe import dataframe_to_rows
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
    print('Passei ponto inicio')    
    print(f'Caminho PDFS_ORIGINAIS:',(PASTA_ORIGINAIS))
    print(f'Caminho Raiz:',(PASTA_RAIZ))
    print(f'Caminho PASTA_FINAL:',(PASTA_FINAL))
    print('Passei ponto 1')
    #N2 = pd.read_csv("Teste.csv",skiprows=[41, 1705], encoding='latin-1', delimiter=';')
    #encoding='utf8' 
    #encoding_errors='utf-8'
    #ASCII, UTF-7, UTF-8, UTF-16 e UTF-32, latin-1, cp1252
    N2 = pd.read_csv("Teste.csv",skiprows=[1,2,3,4,5,6,7, 1705], encoding='latin-1', delimiter='str')
    qtde=0
    print(f'----------- Imprimindo Dataframe --------')
    print(N2)
    N2.to_excel('Teste_1.xlsx')
    # Criamos um dicionário de substituições onde
    #  , vai ser substituido por ''
    #    substituir = {',':''}
    #N3 = N2.replace(substituir, inplace=True) #Não funcionou neste código
    print(f'----- Imprimindo o nome das colunas ----')
    for col in N2.columns: 
        print(col)
    print(f'------- Fim do nome das colunas ----- ')
    #Axis: ‘0’ ou ‘index’ para linhas, e ‘1’ ou ‘columns’ para colunas.
    #Inplace: True, altera o dataframe original sem precisar atribuir
    # ao próprio dataframe.
    #DF = N2.columns.str.replace(',','')
    D = N2
    DF = D['Placa'].str.replace(',', '')
    
    
    
    #pd.concat([df1, df2, ...], axis=0/1)# 0 para combinar linhas e 1 para colunas
    #df1 = pd.DataFrame({'A': ['A0', 'A1'],'B': ['B0', 'B1']})
    #df2 = pd.DataFrame({'A': ['A2', 'A3'],'B': ['B2', 'B3']})
    #result = pd.concat([df1, df2])
    DF0 =pd.DataFrame({'A':['Fabio_Alves_Cordeiro']})
    DF1 =pd.DataFrame({'A':['Extração_Relatorio_Temperatura']})
    Final = pd.concat([DF0, DF1, DF])
    print(f'Novo Dataframe')
    print(Final)
    print(f'----- Imprimindo Novo nome das colunas ----')
    for col in Final.columns: 
        print(col)
    print(f'------- Fim do nome das colunas ----- ')
    #DF = DF['Placa'].drop([',', 'l'], axis=1)
    #DF = DF['P,l,a,c,a,:,'''].drop([',', 'l'], axis=0)
    #df = df.drop(['coluna 1', 'coluna 2' ...], axis=1)
    #df.columns.str.replace('_', '-')
    print(f'Imprimindo data frame Final:{Final}')
    #Vamos supor uma coluna “Filhos” com valores “True” ou “False”, e queremos substitui-los por “sim” e “não”:
    #df["filhos"] = df["filhos"].map({False:"Não", True:"Sim"})
    qtde = len(Final)   
    print(f'------ Tamanho de DF: {qtde} ------')
    Final.to_excel('Teste4.xlsx')
    
    print('Passei ponto 2')
    print('-'*65)
    print('Passei ponto 3')
    #BD = pd.read_csv('Teste.csv')
    #print(f'lendo o BD')
    print('-'*65)
    print(Final)
    print('-'*65)
    print('Passei ponto 4')
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
print("--- %s HORAS:MINUTOS:SEGUNDOS ---" % ("%d:%02d:%02d" % (hour, minutes, seconds) ))
