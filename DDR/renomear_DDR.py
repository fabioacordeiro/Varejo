# pip install PyPDF2

import os
import re
from PyPDF2 import PdfReader

# Caminho onde estão os PDFs
pasta = r'C:\Fabio\Desenvolvimento\Varejo\DDR'

# Caracteres inválidos no Windows
caracteres_invalidos = r'[\\/:*?"<>|\r\n]'

def extrair_nome(pdf_path):
    try:
        reader = PdfReader(pdf_path)
        texto = ''
        for page in reader.pages:
            texto += page.extract_text() or ''

        # Capturar nome antes de "CNPJ"
        matches = re.findall(r'([A-ZÁÉÍÓÚÂÊÔÃÕÇ]{2,}(?:\s+[A-ZÁÉÍÓÚÂÊÔÃÕÇ]{2,}){1,})\s+CNPJ', texto)
        for nome in matches:
            if "CARREFOUR" not in nome and "FILIAIS" not in nome and len(nome) > 5:
                nome_limpo = re.sub(caracteres_invalidos, '', nome.strip())
                nome_limpo = re.sub(r'\s+', ' ', nome_limpo)  # remover espaços extras
                return nome_limpo
    except Exception as e:
        print(f"Erro ao ler {pdf_path}: {e}")
    return None

# Processar PDFs
for arquivo in os.listdir(pasta):
    if arquivo.lower().endswith(".pdf"):
        caminho_antigo = os.path.join(pasta, arquivo)
        nome = extrair_nome(caminho_antigo)
        if nome:
            novo_nome_base = f"Carta_Conforto_Carrefour_-_{nome}.pdf"
            caminho_novo = os.path.join(pasta, novo_nome_base)

            # Evitar sobrescrita
            contador = 1
            while os.path.exists(caminho_novo):
                novo_nome_base = f"Carta_Conforto_Carrefour_-_{nome}_{contador}.pdf"
                caminho_novo = os.path.join(pasta, novo_nome_base)
                contador += 1

            try:
                os.rename(caminho_antigo, caminho_novo)
                print(f"Renomeado para: {novo_nome_base}")
            except Exception as e:
                print(f"[ERRO] Falha ao renomear {arquivo}: {e}")
        else:
            print(f"[!] Nome não encontrado no arquivo: {arquivo}")
