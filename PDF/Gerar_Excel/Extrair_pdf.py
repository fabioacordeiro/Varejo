#pip install pdferli
#pip install PrettyColorPrinter
#from PrettyColorPrinter import add_printer
#add_printer(1)
#pip install tabula-py
#pip install tabula-py[jpype]
#pip install pandas
#Precisa verificar qual a pasta do java no Windows
#Minha_pasta_Java_no_meu_pc =>C:\Program Files (x86)\Java\jre-1.8\bin
#Salvar este caminho como JAVA_HOME em 'variáveis de ambiente' no Windows
#JAVA_HOME
#C:\Program Files (x86)\Java\jre-1.8\bin
from pdferli import get_pdfdf
from PrettyColorPrinter import add_printer
from pathlib import Path
import pandas as pd
#import tabula #Manipulação de Tabelas 
import time #Tempo de execução
import numpy as np

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
    dd = PASTA_ORIGINAIS/"Teste.pdf"
    print(' ----- Início ---- ') 
    print('Caminhos onde serão lidos e gravados os arquivos')    
    print(f'Caminho PDFS_ORIGINAIS:\n',(PASTA_ORIGINAIS))
    print(f'Caminho Raiz:\n',(PASTA_RAIZ))
    print(f'Caminho arquivo csv criado:\n',(CRIADO1))
    print(f'Caminho arquivo XlS criado:\n',(CRIADO2))
    print(f'Caminho PASTA_FINAL:\n',(PASTA_FINAL))
    print('Passo 1')
    add_printer(1)
    path = "C:\\Desenvolvimento\\Varejo\\PDF\\Gerar_Excel\\Arquivo_pdf\\Teste.pdf"
    df = get_pdfdf(path, normalize_content=False)
    print(df)
    df.ds_color_print_all()
    #DF =  tabula.read_pdf("C:\\Desenvolvimento\\Varejo\\PDF\\Gerar_Excel\\Arquivo_pdf\\Teste.pdf", pages='all') # type: ignore
    #tabula.convert_into(dd, "C:\\Desenvolvimento\\Varejo\\PDF\\Gerar_Excel\\Arquivo_csv\\output1.csv", output_format="csv", pages='all')# type: ignore
    #Filtrando o elemento 'LTAnno'
    #df.loc[df.aa_element_type == 'LTAnno']
    np.split(df,df.loc[df.aa_element_type == 'LTAnno'].index)
    #looping retirando as N/A
    for r in np.split(df,df.loc[df.aa_element_type == 'LTAnno'].index):
        df2 = r.dropna(subset="aa_size")
    #criando uma lista e colocando os valores diferentes de vazio dentro dela
    # do campo 'aa_size', somente o primeiro elemento
    togi=[]
    for r in np.split(df,df.loc[df.aa_element_type == 'LTAnno'].index):
        df2 = r.dropna(subset="aa_size")
        if not df2.empty:
            df3 = df2.sort_values(by='aa_x0')
            togi.append(df3.iloc[:1].copy())
            
    print(f'--------- Imprimindo a lista togi --------')
    print(f'Lista togi: {togi}')
    print(f'--------- Fim da lista togi --------')
    #print(len(df))
    #print(df)
    
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
