#pip install pypdf2 #necessário instalar
import os # Leitura arquivos, pastas e diretórios
import time #Tempo de execução
from PyPDF2 import PdfMerger, PdfReader, PdfWriter #Manipular pdf
from pathlib import Path #Manipular pastas, diretórios e arquivos
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
nomeprograma = 'Gerador de CPF'
print(50*'_')
print(f'{nomeprograma:^50}')
print(50*'_')
#-------------Fim Variação tempo -------------------------------
#Incluindo o Try e except que força a execução do código e detalhe
# de erros de execução
try:
    PASTA_RAIZ = Path(__file__).parent
    PASTA_ORIGINAIS = PASTA_RAIZ / 'PDFS_ORIGINAIS'
    PASTA_ALTERADOS = PASTA_RAIZ / 'PDFS_ALTERADOS'
    print(f'Caminho PDFS_ORIGINAIS:',(PASTA_ORIGINAIS))
    print(f'Caminho Raiz:',(PASTA_RAIZ))
    print('Passei ponto 0')
    merger = PdfMerger()
    
    for pdf in (PASTA_ORIGINAIS/"0.pdf", PASTA_ORIGINAIS/"1.pdf"):
        print(f'nome:', pdf)
        merger.append(pdf)
    with open(PASTA_ALTERADOS/'Juntar.pdf', 'wb') as m:
        merger.write(m)
    merger.close()
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
