# pip install PyPDF2
# pip install pandas

import os
import re
from PyPDF2 import PdfReader
import pandas as pd

# Caminho onde estão os PDFs
pasta = r'C:\\Fabio\\Desenvolvimento\\Varejo\\DDR'

# Caracteres inválidos no Windows
caracteres_invalidos = r'[\\/:*?"<>|\r\n]'

# Lista para armazenar resultados
dados_transportadoras = []

def extrair_dados(pdf_path):
    try:
        reader = PdfReader(pdf_path)
        texto = ''
        for page in reader.pages:
            texto += page.extract_text() or ''

        # Extrair nome da transportadora
        nome = None
        matches = re.findall(r'([A-ZÁÉÍÓÚÂÊÔÃÕÇ]{2,}(?:\s+[A-ZÁÉÍÓÚÂÊÔÃÕÇ]{2,}){1,})\s+CNPJ', texto)
        for match in matches:
            if "CARREFOUR" not in match and "FILIAIS" not in match and len(match) > 5:
                nome = re.sub(caracteres_invalidos, '', match.strip())
                nome = re.sub(r'\s+', ' ', nome)
                break

        # Extrair data após "Revisão ... DATA"
        data_match = re.search(r'Revisão\s+\d+\s*-\s*DATA\s+(\d{2}/\d{2}/\d{4})', texto)
        data_pgr = data_match.group(1) if data_match else None

        if nome:
            dados_transportadoras.append({'TRANSPORTADORA': nome, 'DT PGR': data_pgr})

            # Renomear arquivo
            novo_nome_base = f"Carta_Conforto_Carrefour_24_a_25_-_{nome}.pdf"
            caminho_novo = os.path.join(pasta, novo_nome_base)

            # Evitar sobrescrita
            contador = 1
            while os.path.exists(caminho_novo):
                novo_nome_base = f"Carta_Conforto_Carrefour_24_a_25_-_{nome}_{contador}.pdf"
                caminho_novo = os.path.join(pasta, novo_nome_base)
                contador += 1

            try:
                os.rename(pdf_path, caminho_novo)
                print(f"Renomeado para: {novo_nome_base}")
            except Exception as e:
                print(f"[ERRO] Falha ao renomear {os.path.basename(pdf_path)}: {e}")
        else:
            print(f"[!] Nome não encontrado no arquivo: {os.path.basename(pdf_path)}")

    except Exception as e:
        print(f"Erro ao ler {os.path.basename(pdf_path)}: {e}")

# Processar PDFs
for arquivo in os.listdir(pasta):
    if arquivo.lower().endswith(".pdf"):
        caminho_completo = os.path.join(pasta, arquivo)
        extrair_dados(caminho_completo)

# Gerar planilha com resultados
df = pd.DataFrame(dados_transportadoras)
excel_saida = os.path.join(pasta, "Resumo_Transportadoras_DT_PGR.xlsx")
df.to_excel(excel_saida, index=False)
print(f"Planilha gerada: {excel_saida}")
