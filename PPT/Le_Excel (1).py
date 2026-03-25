import pandas as pd

# Caminho do arquivo
file_path = "SINISTROS1.xlsx"

# Ler a planilha, pulando as 2 primeiras linhas de cabeçalho
#df = pd.read_excel(file_path, skiprows=2)
df = pd.read_excel(file_path, sheet_name="Dados")

# Conferir nomes das colunas
print("Colunas encontradas:")
print(df.columns.tolist())

print("Exibir algumas linhas")
print(df)

# Definir a coluna de filtro
coluna_filtro = "Nº Reguladora"

# Exemplo: filtrar somente registros onde há número de reguladora válido
df_filtrado = df[df[coluna_filtro].notna()]
# Filtrar para o sinistro especificado (na coluna Status_Cordeiro)
sinistro_df = df[df['Nº Reguladora'] == 65346]

# Exibir algumas linhas
print(df)
print("\nDados filtrados:")
print(df_filtrado.head())

# Se quiser salvar em novo Excel com os dados filtrados
sinistro_df.to_excel("SINISTROS_FILTRADO.xlsx", index=False)