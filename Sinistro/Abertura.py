# Desenvolvido por Fábio A Cordeiro (ajustado)
# Objetivo: Ler CSV, imprimir, gerar TXT, enviar e-mail e gerar Excel no layout solicitado

# pip install pandas python-dotenv openpyxl

from __future__ import annotations

import os
import smtplib
from datetime import datetime
from pathlib import Path
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import pandas as pd
from dotenv import load_dotenv


# =========================
# CONFIGURAÇÕES
# =========================
load_dotenv()

FROM_EMAIL = os.getenv("FROM_EMAIL")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

CSV_PATH = Path(r"C:\\Fabio\\Desenvolvimento\\Varejo\\Sinistro\\Carga.csv")
OUTPUT_DIR = Path(r"C:\\Fabio\\Desenvolvimento\\Varejo\\Sinistro")
OUTPUT_TXT_PATH = OUTPUT_DIR / "Carga_org_gerado.txt"
OUTPUT_EXCEL_PATH = OUTPUT_DIR / "Carga_org_gerado_layout_sinistro.xlsx"

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

DESTINATARIO_EMAIL = "fabio_cordeiro@carrefour.com"
EMAIL_TRANSP = "fabioacordeiro@yahoo.com.br"  # mantido como você tinha (não usado no sendmail hoje)


# =========================
# COLUNAS DO EXCEL (ORDEM EXATA)
# =========================
COLUNAS_EXCEL = [
    "Tipo Carga", "Cliente", "Revisão Causa", "Causa Final", "Causa do Sinistro",
    "Nº Corretora", "Nº Reguladora", "Reguladora", "Seguradora", "Apólice",
    "Data do Sinistro", "Hora do Sinistro", "Data do Aviso Regulador",
    "Mercadoria", "Transportador", "Nota Fiscal", "N_Carga", "Origem Ajustado",
    "UF - Origem", "Cidade Origem", "UF - Destino", "Cidade - Destino",
    "Complemento_info", "Motorista", "CPF", "QTDE_CARGAS_MOT", "Placa",
    "Qtde Viagens Mot", "Local do Sinistro", "Cidade do Sinistro", "UF - Sinistro",
    "Valor do Embarque", "Imp. Segurada", "Estimativa Prejuizo", "Prejuizo Apurado",
    "Franquia", "Valor Indenizavel", "Observação", "Valor Descarte",
    "Solic Doc Transp", "Ret DOC Transp", "Solic DOC BRK", "Ret DOC BRK",
    "Solic DOC CD", "Ret DOC CD", "MÉTODO", "MATERIAL", "MAQUINA",
    "MEIO AMBIENTE", "MÃO DE OBRA", "M MEIO AMBIENTE", "RESPONSAVEL_SINI",
    "DET_RESPONSABILIDADE", "ANO", "Status_Cordeiro", "Credito Seguro",
    "Saldo", "USAR CRÉDITO", "Responsabilidade Pagamento", "Ação_1",
    "Cont", "Sequência", "Nome_Ano_e_Mes", "Ano", "Mês", "CONT_MES",
    "Ano_e_Mes", "Nome_Mês", "TRANSP", "UF_Destino_BR", "Região_Destino_BR",
    "UF_Local_Sinistro_BR", "Local_Sinistro_Coordenadas", "Região_Sinistro_BR",
    "ENCOSTA_EM_DOCA", "INICIO_CARREGAMENTO", "FIM_CARREGAMENTO", "EMISSAO_NF",
    "INICIO_VIAGEM", "CHEGADA_EM_LOJA", "Ação", "CNPJ Transp"
]


# =========================
# MAPEAMENTO (EXCEL -> CSV)
# =========================
# Observação importante: seu CSV tem "\ufeffNúmero da Carga". Vamos normalizar isso ao ler.
MAPEAMENTO_EXCEL_PARA_CSV = {
    "Tipo Carga": "Tipo de Carga",
    "Cliente": "Filial",
    "Transportador": "Transportador",
    "Nota Fiscal": "Nº NF-e",
    "N_Carga": "Número da Carga",  # após normalização do BOM
    "Origem Ajustado": "Filial",
    "UF - Origem": "UF Origem",
    "Cidade Origem": "Origem",
    "UF - Destino": "UF Destino",
    "Cidade - Destino": "Destino",
    "Motorista": "Motoristas",
    "Placa": "Veículo",
    "Valor do Embarque": "Valor NF",
    "Imp. Segurada": "Valor NF",
    "Estimativa Prejuizo": "Valor NF",
    "ENCOSTA_EM_DOCA": "Data Carregamento",
    "INICIO_VIAGEM": "Data Início Viagem",
    "CHEGADA_EM_LOJA": "Data Fim Viagem",
    "CNPJ Transp": "CNPJ do Transportador",
}


# =========================
# FUNÇÕES AUXILIARES
# =========================
def normalizar_colunas(df: pd.DataFrame) -> pd.DataFrame:
    """
    - Remove espaços extras
    - Remove BOM do começo do nome de coluna, quando existir
    - Ajusta especificamente 'Número da Carga' vindo com BOM
    """
    df = df.copy()
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
    df.columns = [c.replace("\ufeff", "") if isinstance(c, str) else c for c in df.columns]
    return df


def agora_data_br() -> str:
    return datetime.now().strftime("%d/%m/%Y")


def agora_hora() -> str:
    return datetime.now().strftime("%H:%M")


def pegar_valor(row: pd.Series, col_csv: str):
    if col_csv not in row.index:
        return ""
    val = row[col_csv]
    if pd.isna(val):
        return ""
    return val


def primeiro_valor_se_lista(val) -> str:
    """
    'Somente 1 nota fiscal': se vier '123, 456', pega a primeira.
    """
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    if not s:
        return ""
    return s.split(",")[0].strip()


def enviar_email(destinatario: str, assunto: str, corpo_html: str) -> None:
    msg = MIMEMultipart()
    msg["From"] = FROM_EMAIL
    msg["To"] = destinatario
    msg["Subject"] = assunto
    msg.attach(MIMEText(corpo_html, "html"))

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(FROM_EMAIL, EMAIL_PASSWORD)
        server.sendmail(FROM_EMAIL, destinatario, msg.as_string())


# =========================
# MAIN
# =========================
def main():
    # 1) Ler CSV uma única vez (utf-8-sig remove BOM do arquivo)
    if not CSV_PATH.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {CSV_PATH}")

    df = pd.read_csv(CSV_PATH, sep=None, engine="python", encoding="utf-8-sig")
    df = normalizar_colunas(df)

    # Debug: mostrar colunas encontradas
    print(df.columns.tolist())
    print("-" * 65)

    # 2) Mostrar primeiras linhas (corrigido)
    print(df.head())
    print("-" * 65)

    # 3) Trabalhar com a primeira linha 
    row = df.iloc[0]

    cliente = pegar_valor(row, "Filial")
    carga = pegar_valor(row, "Número da Carga")
    data_carga = pegar_valor(row, "Data da Carga")

    motivo = "VARIAÇÃO DE TEMPERATURA / ROUBO / ACIDENTE"
    faixa_temperatura = pegar_valor(row, "Temperatura")

    origem = f"{pegar_valor(row, 'Origem')} - {pegar_valor(row, 'UF Origem')}"
    destino = f"{pegar_valor(row, 'Código Destinatário')} - {pegar_valor(row, 'Destino')} - {pegar_valor(row, 'UF Destino')} ({pegar_valor(row, 'Entregas')} EntregaS)"

    local_sinistro = "Verificar"

    valor_nf = pegar_valor(row, "Valor NF")
    valor_embarcado = f"R$ {valor_nf}"

    transportador_nome = pegar_valor(row, "Transportador")
    cnpj_transp = pegar_valor(row, "CNPJ do Transportador")
    transportadora = f"{transportador_nome} - {cnpj_transp}"

    nota_fiscal = primeiro_valor_se_lista(pegar_valor(row, "Nº NF-e"))

    mdfe = pegar_valor(row, "MDF-es")
    cte = pegar_valor(row, "Número CTe")
    motorista = pegar_valor(row, "Motoristas")

    veiculo = str(pegar_valor(row, "Veículo"))
    placas = veiculo.replace(",", " / CARRETA:") if veiculo else ""

    # 4) Montar texto formatado e salvar TXT
    texto_formatado = f"""cliente: {cliente}

Carga: {carga}

Data: {data_carga}

MOTIVO: {motivo}

FAIXA DE TEMPERATURA: {faixa_temperatura}

Origem: {origem}

Destino: {destino}

LOCAL DO SINISTRO: {local_sinistro}

Valor Embarcado: {valor_embarcado}

Transportadora: {transportadora}

Nota fiscal: {nota_fiscal}

cte: {cte}

MDFE: {mdfe}

Motorista: {motorista}

Placas: {placas}
"""

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_TXT_PATH.write_text(texto_formatado, encoding="utf-8")

    print("Concluído (TXT)")
    print(texto_formatado)

    # 5) Enviar e-mail (mantendo sua estrutura)
    assunto = f"Dados de Sinistro - {transportador_nome}"
    corpo_email = f"""
    <html>
        <body style="font-family: Arial, sans-serif; color: #333;">
            <p>Prezados saudações !</p>
            <p>Favor seguir com o sinistro abaixo:</p>

            <p>Sinistro: </p>
            <p>Cliente: {cliente}</p>
            <p>Reguladora: GLOBAL COMISSARIA</p>
            <p>Seguradora: Ezze</p>
            <p>Corretora: Wiz</p>
            <p>Carga: {carga}</p>
            <p>Data: {data_carga}</p>
            <p>Motivo: {motivo}</p>
            <p>FAIXA DE TEMPERATURA: {faixa_temperatura}</p>
            <p>Origem: {origem}</p>
            <p>Destino: {destino}</p>
            <p>LOCAL DO SINISTRO: {local_sinistro}</p>
            <p>Valor Embarcado: {valor_embarcado}</p>
            <p>Transportadora: {transportadora}</p>
            <p>Nota fiscal: {nota_fiscal}</p>
            <p>CT-e: {cte}</p>
            <p>MDF-e: {mdfe}</p>
            <p>Motorista: {motorista}</p>
            <p>Placas: {placas}</p>
        </body>
    </html>
    """

    #enviar_email(DESTINATARIO_EMAIL, assunto, corpo_email)
    #print("Enviado (e-mail)")

    # 6) Gerar Excel 
    # Estratégia: cria 1 linha com TODAS as colunas; preenche o que foi mapeado; o resto fica vazio.
    linha_excel = {col: "" for col in COLUNAS_EXCEL}

    # Datas/hora (Agora())
    linha_excel["Data do Sinistro"] = agora_data_br()
    linha_excel["Data do Aviso Regulador"] = agora_data_br()
    linha_excel["Hora do Sinistro"] = agora_hora()

    # Campos mapeados
    for col_excel, col_csv in MAPEAMENTO_EXCEL_PARA_CSV.items():
        if col_excel == "Nota Fiscal":
            linha_excel[col_excel] = primeiro_valor_se_lista(pegar_valor(row, col_csv))
        else:
            linha_excel[col_excel] = pegar_valor(row, col_csv)

    # Garantir tipos simples (evita alguns problemas de escrita)
    df_excel = pd.DataFrame([linha_excel], columns=COLUNAS_EXCEL)

    # Salvar Excel
    df_excel.to_excel(OUTPUT_EXCEL_PATH, index=False, engine="openpyxl")
    print(f"Excel gerado: {OUTPUT_EXCEL_PATH}")

    enviar_email(DESTINATARIO_EMAIL, assunto, corpo_email)
    print("Enviado (e-mail)")

    print("Fim")


if __name__ == "__main__":
    main()