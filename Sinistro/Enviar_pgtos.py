import pandas as pd
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from dotenv import load_dotenv

# Carrega as variáveis de ambiente (.env) com e-mail e senha
load_dotenv()
FROM_EMAIL = os.getenv("FROM_EMAIL")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

# Caminho para o arquivo e pasta de saída
input_path = "C:\\Fabio\\Desenvolvimento\\Varejo\\Sinistro\\BD_PAGAMENTOS.xlsx"
output_dir = "C:\\Fabio\\Desenvolvimento\\Varejo\\Sinistro\\relatorios_pgts"
os.makedirs(output_dir, exist_ok=True)

# Lê a planilha
df = pd.read_excel(input_path)

# Agrupar por prestador
prestadores = df['PRESTADOR'].unique()

# Enviar e-mail por prestador
for prestador in prestadores:
    df_prestador = df[df['PRESTADOR'] == prestador]
    email = df_prestador['Email_transp'].iloc[0]

    # Criar arquivo Excel para o prestador
    file_name = f'{output_dir}/{prestador}.xlsx'
    df_prestador.to_excel(file_name, index=False)

    # Calcular total de pagamentos
    total_pagamento = df_prestador['Valor Transação'].sum()

    # Corpo e assunto do e-mail
    assunto = f'{prestador} - Base de Pagamentos'
    corpo = f"""
    Prezado(a) {prestador},

    Segue em anexo a base de pagamentos referente ao período informado.

    Total de pagamentos: R$ {total_pagamento:,.2f}

    Atenciosamente,
    Equipe Financeira
    """

    # Montar e enviar o e-mail
    msg = MIMEMultipart()
    msg['From'] = FROM_EMAIL
    msg['To'] = email
    msg['Subject'] = assunto
    msg['Cc'] = 'fabioacordeiro@yahoo.com.br'
    msg.attach(MIMEText(corpo, 'plain'))

    # Anexo
    with open(file_name, 'rb') as f:
        attachment = MIMEApplication(f.read(), _subtype='xlsx')
        attachment.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file_name))
        msg.attach(attachment)

    # Lista de destinatários
    destinatarios = [email, 'fabioacordeiro@yahoo.com.br']

    # Enviar via SMTP
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(FROM_EMAIL, EMAIL_PASSWORD)
        server.sendmail(FROM_EMAIL, destinatarios, msg.as_string())

print("E-mails enviados com sucesso.")