#Instale as bibliotecas necessárias:

#pandas para manipulação de dados.
#openpyxl para leitura de arquivos Excel.
#smtplib para envio de e-mails.

import os
import pathlib
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv
# Defina o caminho do arquivo Excel
excel_file = 'Cadastro_Fornecedores.xlsx'

# Leia a planilha
df = pd.read_excel(excel_file)

# Carrega variáveis de ambiente
load_dotenv()
FROM_EMAIL = os.getenv("FROM_EMAIL")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

# Defina o assunto e o corpo do e-mail
assunto = 'Prospecção de Veículos - Carrefour'
corpo = '''Boa tarde !!!
AGUARDANDO CONTRATAÇÃO - Data-16/07/04
Origem CD OSASCO (CARRETA)
UF-CIDADE	   QTDE
RS-SERTORIO      1
MG-JUIZ DE FORA  1
TOTAL            2

Tem veículo disponível para atender alguma rota ?
'''
# Configurações do servidor SMTP
servidor = 'smtp.gmail.com'
porta = 587

# Função para enviar o e-mail
def enviar_email(destinatario):
    # Crie a mensagem do e-mail
    msg = MIMEMultipart()
    msg['From'] = seu_email
    msg['To'] = destinatario
    msg['Subject'] = assunto

    # Adicione o corpo do e-mail
    msg.attach(MIMEText(corpo, 'plain'))

    # Envie o e-mail
    with smtplib.SMTP(servidor, porta) as server:
        server.starttls()
        server.login(seu_email, sua_senha)
        text = msg.as_string()
        server.sendmail(seu_email, destinatario, text)

# Envie e-mails para cada destinatário na planilha
for email in df['Email']:
    try:
        enviar_email(email)
        print(f'E-mail enviado para {email}')
    except Exception as e:
        print(f'Erro ao enviar e-mail para {email}: {e}')