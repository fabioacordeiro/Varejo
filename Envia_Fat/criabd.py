# pip install pandas
# pip install openpyxl

import os
import pandas as pd

# Caminho para o arquivo e pasta de saída (igual ao seu script)
input_path = r"C:\\Fabio\\Desenvolvimento\\Varejo\\Envia_Fat\\TRATAMENTO_BD\\Base_Pagamento.xlsx"
output_dir = r"C:\\Fabio\\Desenvolvimento\\Varejo\\Envia_Fat\\TRATAMENTO_BD"
os.makedirs(output_dir, exist_ok=True)

# Lê a planilha
df = pd.read_excel(input_path)
prestadores = df.DataFrame()

# Valida colunas usadas
required_cols = ["Documento do Favorecido", "Nome do Favorecido"]
missing = [c for c in required_cols if c not in df.columns]
if missing:
    raise RuntimeError(f"Colunas ausentes no Excel: {missing}")
    prestadores = df["Nome do Favorecido"].dropna().unique()

