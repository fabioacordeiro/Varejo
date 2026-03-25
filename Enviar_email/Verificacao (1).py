import pandas as pd

# Defina o caminho do arquivo Excel
excel_file = 'C:\\Fabio\\CARREFOUR\\BRK\\PRONTA_RESPOSTA\\Pronta_resposta.xlsx'

# Leia os dados
df_dados = pd.read_excel(excel_file, sheet_name="Dados").fillna('')

# Exibir as colunas disponíveis para garantir que o nome está correto
print("Colunas disponíveis:", df_dados.columns.tolist())

# Exibir as primeiras linhas para ver os dados da coluna "Email_transp"
print(df_dados[["TRANSPORTADORA", "Email_Transp"]].head(10))

# Verificar se há valores NaN ou vazios
print("Valores únicos na coluna 'Email_transp':", df_dados["Email_Transp"].unique())