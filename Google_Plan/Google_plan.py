import pandas as pd #Biblioteca do pandas para manipular o banco de dados
import pyautogui # Biblioteca para manipular o mouse
import webbrowser # Bpm Web
import time #Bpm a pausa

# Caminho do arquivo Excel
file_path = 'C:\\Fabio\\Desenvolvimento\\Varejo\\Google_Plan\\Rel_Teste.xlsx'

# Nome da planilha específica que você deseja importar
sheet_name = 'Base'

# Carregar a planilha específica em um DataFrame
df = pd.read_csv(file_path, sheet_name=sheet_name)
# Ordenar a planilha de forma ascendente
df1 = df.sort_values(by=['DATA INVOICE', 'VALOR TOTAL'], ascending=[True, True])

# Criar novas colunas com valores deslocados
df1['VALOR FINAL'] = df1['VALOR TOTAL'].shift(1)
df1['COD_FORN'] = df1['FORNECEDOR'].shift(1)
# procure no google "google developers console"
# selecione o primeiro site "Google Developers Console"
# 1 - Aceitar os termos 
# Acessar Console
# Primeiros passos do google Cloud 
# API serviços
# Criar Projeto

# Salvar o DataFrame atualizado em um novo arquivo Excel
#df1.to_excel('Rel_Teste2.xlsx', index=False)

# Copiar os dados do DataFrame para a área de transferência
df1.to_clipboard(index=False, sep='\t')  # Usa tabulação para compatibilidade com a Google Sheet

# Abrir a Google Sheet no navegador
sheet_url = "https://docs.google.com/spreadsheets/d/1gnmYocybaHksZGLqq-4jRM-McbaaZW-NZ4IBgIi8vyw/edit?gid=0#gid=0"
webbrowser.open(sheet_url)

# Aguarde o carregamento do navegador
time.sleep(10)  # Ajuste o tempo se necessário

# Simular o comando de colar na Google Sheet
# Certifique-se de que a célula inicial está selecionada antes de rodar o script
pyautogui.hotkey('ctrl', 'v')