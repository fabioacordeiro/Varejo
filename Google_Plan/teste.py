
from pathlib import Path #Biblioteca para manipular caminhos
import pandas as pd #Biblioteca do pandas para manipular o banco de dados
import csv #Bpm o arquivo CSV
import time #Bpm a pausa
import datetime #Bpm a o tempo
import pyautogui # Biblioteca para manipular o mouse
import webbrowser # Bpm Web

agora = datetime.datetime.now()
agora_string = agora.strftime("%A %d %B %y %I:%M")
agora_datetime = datetime.datetime.strptime(agora_string, "%A %d %B %y %I:%M")
hora = datetime.date.today()
inicio = time.time()
print(f'Time Início: {agora_datetime}')
# time.sleep(50)

try:
    print(f'{"Step 1: Reading file CSV":.^60}')
    # abrindo o arquivo CSV e mostrando 5 linhas.
    # skiprows => informar quais as linhas que serão ignoradas
    # nrows => qual a quantidade de linhas que serão percorridas
    # usecols => qual as colunas que serão utilizadas
    BD1 = pd.read_csv(filepath_or_buffer='C:\\Fabio\\Desenvolvimento\\Varejo\\Google_Plan\\Rel_Teste.csv',
                      sep='|', index_col=None, skiprows=[1, 2, 3], low_memory=False, usecols=[
                          1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16,
                          17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30,
                          31, 32])

    coluna = {
        "CPF/CNPJ": "CNPJ",
        "Data de Término de Descarregamento": "DT Descarga"}
    
    BD1.rename(columns=coluna)

    BD2 = pd.DataFrame(BD1)
    print(f'{"Step 2: Dataframe created":.^60}')
    
    # print('Readying file completed')
    print(f'{"Step 3: Show size Dataframe":.^60}')
    print(f'Linhas do BD2:{BD2[BD2.columns[0]].count()}')
    # print(BD2.head())
    # print(f'{"Visualização com describe()":.^60}')
    # print(BD2.describe())
    # print(f'{"Visualização com info()":.^60}')
    # print(BD2.info())

except Exception as error:
    print(error.__class__.__name__)
    print(error.args)
    print(error)

finally:
    print('Finished reading process ....')


# Copiar os dados do DataFrame para a área de transferência
BD2.to_clipboard(index=False, sep='\t')  # Usa tabulação para compatibilidade com a Google Sheet

# Abrir a Google Sheet no navegador
sheet_url = "https://docs.google.com/spreadsheets/d/1gnmYocybaHksZGLqq-4jRM-McbaaZW-NZ4IBgIi8vyw/edit?gid=0#gid=0"
webbrowser.open(sheet_url)

# Aguarde o carregamento do navegador
time.sleep(10)  # Ajuste o tempo se necessário

# Simular o comando de colar na Google Sheet
# Certifique-se de que a célula inicial está selecionada antes de rodar o script
pyautogui.hotkey('ctrl', 'v')

#Finalizando e cronometrando o tempo do processo
print(f'{"Step 4: Process finished ":.^60}')
fim = time.time()
print(
    f'Tempo de processamento: {int(fim-inicio)} segundos {int((fim-inicio)/60)} Minutos')
