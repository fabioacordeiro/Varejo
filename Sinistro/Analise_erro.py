import os
import pandas as pd
import tabula
from openpyxl import load_workbook
from pptx import Presentation

 # 1. Leitura dos dados do arquivo CSV
df_carga = pd.read_csv("C:\\Fabio\\Desenvolvimento\\Varejo\\Sinistro\\ppt\\Carga.csv")
print(df_carga.columns)
print("Fim")

df_temperatura = tabula.read_pdf("C:\\Fabio\\Desenvolvimento\\Varejo\\Sinistro\\ppt\\Temperatura.pdf", pages=1)[0]
print(df_temperatura.columns)
print("Fim")

df = df_carga


sheets_validas = [df for df in sheets if not df.empty]
if sheets_validas:
    df_xlsb = pd.concat(sheets_validas)
else:
    df_xlsb = pd.DataFrame()

if len(df_carga) == len(df_xlsb):
    # Inserir dados do df_carga no DataFrame df_xlsb
    for col_num, col_name in enumerate(df_carga.columns):
        df_xlsb.insert(col_num, col_name, df_carga[col_name])
else:
    print("Erro: Os DataFrames têm tamanhos diferentes.")

# ... (seu código para salvar o arquivo XLSX) ...