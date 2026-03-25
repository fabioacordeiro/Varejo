# pip install pandas
# pip install dotenv
# pip install openpyxl
import pandas as pd
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from dotenv import load_dotenv

# Carrega variáveis de ambiente
load_dotenv()
FROM_EMAIL = os.getenv("FROM_EMAIL")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

# Caminho para o arquivo e pasta de saída
input_path = "C:\\Fabio\\Desenvolvimento\\Varejo\\Envia_Fat\\BD_PAGAMENTOS.xlsx"
output_dir = "C:\\Fabio\\Desenvolvimento\\Varejo\\Envia_Fat\\relatorios_pgts"
os.makedirs(output_dir, exist_ok=True)

# Lê a planilha
df = pd.read_excel(input_path)

# Agrupar por prestador
prestadores = df['PRESTADOR'].unique()

# Enviar e-mail por prestador
for prestador in prestadores:
    try:
        df_prestador = df[df['PRESTADOR'] == prestador]

        # Trata múltiplos e-mails separados por ponto e vírgula
        email_raw = str(df_prestador['Email_transp'].iloc[0]).strip()
        emails_to = [e.strip() for e in email_raw.split(';') if e.strip() and '@' in e]

        if not emails_to:
            print(f"⚠️ Nenhum e-mail válido para: {prestador}")
            continue

        print(f"📧 Enviando para: {prestador} | {', '.join(emails_to)}")

        # Criar arquivo Excel com dados do prestador
        file_name = os.path.join(output_dir, f'{prestador}.xlsx')
        df_prestador.to_excel(file_name, index=False)

        # Total de pagamentos
        total_pagamento = df_prestador['Valor Transação'].sum()

        # Assunto e corpo do e-mail
        assunto = f'{prestador} - BD Pagto - CRFLOG(BOMPREÇO) - Dezembro/2025'
        corpo = f"""
        Prezado(a) {prestador},

        Segue em anexo a base de pagamentos referente ao período informado.

        Total de pagamentos: R$ {total_pagamento:,.2f}

        Atenciosamente,
        Equipe Financeira
        """

        # Lista de cópias (Cc)
        emails_cc = "fabio_cordeiro@carrefour.com; br_financeiro_crflog@carrefour.com"
        lista_cc = [e.strip() for e in emails_cc.split(';') if e.strip() and '@' in e]

        # Monta o e-mail
        msg = MIMEMultipart()
        msg['From'] = FROM_EMAIL
        msg['To'] = ', '.join(emails_to)
        msg['Cc'] = ', '.join(lista_cc)
        msg['Subject'] = assunto
        msg.attach(MIMEText(corpo, 'plain'))

        # Anexa o relatório
        with open(file_name, 'rb') as f:
            attachment = MIMEApplication(f.read(), _subtype='xlsx')
            attachment.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file_name))
            msg.attach(attachment)

        # Lista final de destinatários
        destinatarios_finais = emails_to + lista_cc

        # Envia o e-mail
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(FROM_EMAIL, EMAIL_PASSWORD)
            server.sendmail(FROM_EMAIL, destinatarios_finais, msg.as_string())

        print(f"✅ E-mail enviado com sucesso para: {prestador}")

    except Exception as e:
        print(f"❌ Erro ao enviar e-mail para: {prestador} | Erro: {e}")