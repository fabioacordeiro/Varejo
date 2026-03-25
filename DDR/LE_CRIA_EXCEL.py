# pip install pandas
# pip install openpyxl

import os
import pandas as pd

# Caminho da pasta onde estão os arquivos PDF
pasta = r'C:\\Fabio\Desenvolvimento\\Varejo\\DDR'  # <-- Altere se necessário

# Lista para armazenar os nomes dos arquivos PDF
nomes_arquivos = []

# Percorre os arquivos da pasta
for arquivo in os.listdir(pasta):
    if arquivo.lower().endswith('.pdf'):
        nomes_arquivos.append({'NOME DO ARQUIVO': arquivo})

# Cria o DataFrame
df = pd.DataFrame(nomes_arquivos)

# Caminho do arquivo Excel de saída
caminho_saida = os.path.join(pasta, 'Lista_Arquivos_PDF.xlsx')

# Salva em Excel
df.to_excel(caminho_saida, index=False)

print(f'Planilha criada com sucesso: {caminho_saida}')