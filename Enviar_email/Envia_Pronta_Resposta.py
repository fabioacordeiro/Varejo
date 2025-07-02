# Desenvolvido por Fábio A Cordeiro 
# Em 22/02/2025 

import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
from dotenv import load_dotenv #type:igore
load_dotenv()
FROM_EMAIL= os.getenv('FROM_EMAIL')
BD_PASSWORD = os.getenv('EMAIL_PASSWORD')

# Defina o seu e-mail e senha (use um e-mail e senha de aplicação)
seu_email = FROM_EMAIL
sua_senha = BD_PASSWORD

# Defina o caminho do arquivo Excel
excel_file = 'C:\\Fabio\\CARREFOUR\\BRK\\PRONTA_RESPOSTA\\Pronta_resposta.xlsx'

# Leia a planilha
df = pd.read_excel(excel_file).fillna('')  # Substitui valores NaN por string vazia

# Configurações do servidor SMTP
servidor = 'smtp.gmail.com'
porta = 587

# Função para enviar o e-mail
def enviar_email(destinatario, transportadora, email_transp, cnpj):
    # Define o assunto personalizado
    assunto = f'DESCONTO - ESCOLTA - BRK/Carrefour - {transportadora}'

    # Define o corpo personalizado
    corpo = f"""{email_transp}
samantha_campos@carrefour.com; ana_cristina_moraes@carrefour.com;
br_torre_controle_crf@carrefour.com;leticia_andrade_reis@carrefour.com;
clovis_da_silva@carrefour.com;


Prezados saudações!

Favor seguir com o desconto abaixo da transportadora, referente a escolta enviada pela nossa gerenciadora BRK.
Conforme PGR assinado pelo transportador especificado no item 11.
Item 11. PROCEDIMENTOS DE PRONTA RESPOSTA
11.1.6. Para proteção de veículos quebrados e/ou com problemas mecânicos ou não cumprimento das normas contidas no PGR custo será revertido ao transportador.

"NR ORDEM", "DATA_HORA DO ACIONAMENTO", "PLACAS", "STATUS", "CLIENTE", "TRANSPORTADORA", "VALOR"
{NR ORDEM}, {DATA_HORA DO ACIONAMENTO}, {PLACAS}, {STATUS}, {CLIENTE}, {TRANSPORTADORA},{VALOR}

{Soma de VALOR}

Obrigado
"""

    # Cria a mensagem do e-mail
    msg = MIMEMultipart()
    msg['From'] = seu_email
    msg['To'] = destinatario
    msg['Subject'] = assunto

    # Adiciona o corpo do e-mail
    msg.attach(MIMEText(corpo, 'plain'))

    # Envia o e-mail
    with smtplib.SMTP(servidor, porta) as server:
        server.starttls()
        server.login(seu_email, sua_senha)
        server.sendmail(seu_email, destinatario, msg.as_string())

# Envie e-mails para cada destinatário na planilha
for _, row in df.iterrows():
    email = row['Email']
    transportadora = row['TRANSP']
    email_transp = row['Email_transp']
    cnpj = row['CNPJ']
    
    try:
        enviar_email(email, transportadora, email_transp, cnpj)
        print(f'E-mail enviado para {email} - Transportadora: {transportadora}')
    except Exception as e:
        print(f'Erro ao enviar e-mail para {email}: {e}')