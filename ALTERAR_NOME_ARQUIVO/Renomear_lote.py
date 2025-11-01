import os

# Caminho da pasta com os arquivos PDF
pasta = r'C:\\Fabio\\Desenvolvimento\\Varejo\\ALTERAR_NOME_ARQUIVO'  # <-- Altere para o caminho correto

# Prefixos
prefixo_antigo = "Carta_Conforto_Carrefour_-_"
prefixo_novo = "DDR-TOTAL-2024-2025 - "

# Itera sobre todos os arquivos na pasta
for nome_arquivo in os.listdir(pasta):
    if nome_arquivo.lower().endswith(".pdf") and nome_arquivo.startswith(prefixo_antigo):
        novo_nome = nome_arquivo.replace(prefixo_antigo, prefixo_novo, 1)
        caminho_antigo = os.path.join(pasta, nome_arquivo)
        caminho_novo = os.path.join(pasta, novo_nome)

        # Renomeia o arquivo
        os.rename(caminho_antigo, caminho_novo)
        print(f"Renomeado para: {novo_nome}")

print("Processo concluído.")