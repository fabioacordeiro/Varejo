# pip install pandas
# pip install python-dotenv
# pip install openpyxl

import pandas as pd
import os
import smtplib
import time
import socket
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from dotenv import load_dotenv

# =========================
# CONFIGURAÇÕES / TIMEOUTS
# =========================
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 465

SMTP_TIMEOUT = 60        # tempo máximo (segundos) aguardando resposta do servidor SMTP
RETRY_MAX = 3            # tentativas por prestador
WAIT_BEFORE_SEND = 5    # espera (segundos) antes de tentar enviar (para estabilizar 4G/5G)
WAIT_RETRY = 10          # espera (segundos) entre tentativas quando falhar

# =========================
# VARIÁVEIS DE AMBIENTE
# =========================
load_dotenv()
FROM_EMAIL = os.getenv("FROM_EMAIL")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

if not FROM_EMAIL or not EMAIL_PASSWORD:
    raise ValueError("❌ Variáveis de ambiente FROM_EMAIL e/ou EMAIL_PASSWORD não foram encontradas no .env")

# =========================
# CAMINHOS
# =========================
input_path = r"C:\\Fabio\\Desenvolvimento\\Varejo\\Envia_Fat\\BD_PAGAMENTOS.xlsx"
output_dir = r"C:\\Fabio\\Desenvolvimento\\Varejo\\Envia_Fat\\relatorios_pgts"
os.makedirs(output_dir, exist_ok=True)

# =========================
# LER PLANILHA
# =========================
df = pd.read_excel(input_path)

# Valida colunas essenciais
colunas_necessarias = ["PRESTADOR", "Email_transp", "Valor Transação"]
faltando = [c for c in colunas_necessarias if c not in df.columns]
if faltando:
    raise ValueError(f"❌ A planilha não contém as colunas necessárias: {faltando}")

# Agrupar por prestador
prestadores = df["PRESTADOR"].dropna().unique()

# =========================
# LOOP PRINCIPAL
# =========================
for prestador in prestadores:
    try:
        df_prestador = df[df["PRESTADOR"] == prestador].copy()

        # Trata múltiplos e-mails separados por ponto e vírgula
        email_raw = str(df_prestador["Email_transp"].iloc[0]).strip()
        emails_to = [e.strip() for e in email_raw.split(";") if e.strip() and "@" in e]

        if not emails_to:
            print(f"⚠️ Nenhum e-mail válido para: {prestador}")
            continue

        print(f"📧 Enviando para: {prestador} | {', '.join(emails_to)}")

        # Criar arquivo Excel com dados do prestador
        safe_prestador = "".join(ch for ch in str(prestador) if ch not in r'\/:*?"<>|').strip()
        file_name = os.path.join(output_dir, f"{safe_prestador}.xlsx")
        df_prestador.to_excel(file_name, index=False)

        # Total de pagamentos
        total_pagamento = df_prestador["Valor Transação"].sum()

        # Assunto e corpo do e-mail
        assunto = f"{prestador} - BD Pagto - CRFLOG(BOMPREÇO) - DEZEMBRO/2025"
        corpo = f"""Prezado(a) {prestador},

Segue em anexo a base de pagamentos referente ao período informado.

Total de pagamentos: R$ {total_pagamento:,.2f}

Atenciosamente,
Equipe Financeira
"""

        # Lista de cópias (Cc)
        emails_cc = "fabio_cordeiro@carrefour.com; br_financeiro_crflog@carrefour.com"
        lista_cc = [e.strip() for e in emails_cc.split(";") if e.strip() and "@" in e]

        # Monta o e-mail
        msg = MIMEMultipart()
        msg["From"] = FROM_EMAIL
        msg["To"] = ", ".join(emails_to)
        msg["Cc"] = ", ".join(lista_cc)
        msg["Subject"] = assunto
        msg.attach(MIMEText(corpo, "plain"))

        # Anexa o relatório
        with open(file_name, "rb") as f:
            attachment = MIMEApplication(f.read(), _subtype="xlsx")
            attachment.add_header(
                "Content-Disposition",
                "attachment",
                filename=os.path.basename(file_name),
            )
            msg.attach(attachment)

        # Lista final de destinatários
        destinatarios_finais = emails_to + lista_cc

        # =========================
        # ENVIO COM RETRY + TIMEOUT
        # =========================
        tentativa = 1
        enviado = False

        while tentativa <= RETRY_MAX and not enviado:
            try:
                print(f"⏳ Aguardando conexão estabilizar ({WAIT_BEFORE_SEND}s)...")
                #time.sleep(WAIT_BEFORE_SEND)

                with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=SMTP_TIMEOUT) as server:
                    server.login(FROM_EMAIL, EMAIL_PASSWORD)
                    server.sendmail(FROM_EMAIL, destinatarios_finais, msg.as_string())

                print(f"✅ E-mail enviado com sucesso para: {prestador}")
                enviado = True

            except (smtplib.SMTPException, socket.timeout, OSError) as e:
                print(f"⚠️ Tentativa {tentativa}/{RETRY_MAX} falhou para {prestador}: {e}")
                tentativa += 1

                if tentativa <= RETRY_MAX:
                    print(f"🔄 Aguardando {WAIT_RETRY}s para tentar novamente...")
                    #time.sleep(WAIT_RETRY)
                else:
                    raise e

        # (Opcional) Pausa pequena entre prestadores para reduzir risco de bloqueio/instabilidade
        time.sleep(3)

    except Exception as e:
        print(f"❌ Erro ao enviar e-mail para: {prestador} | Erro: {e}")






