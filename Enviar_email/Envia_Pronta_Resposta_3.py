# Desenvolvido por Fábio A Cordeiro
# Em 22/02/2025
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

# Leia as abas da planilha
df_dados = pd.read_excel(excel_file, sheet_name="Dados").fillna('')
df_resumo = pd.read_excel(excel_file, sheet_name="Resumo").fillna('')
print('-'*30)
print(df.columns)
print('-'*30)

# Configurações do servidor SMTP
servidor = 'smtp.gmail.com'
porta = 587

# Função para enviar o e-mail
def enviar_email(destinatario, transportadora, email_transp, corpo_email):
    assunto = f'DESCONTO - ESCOLTA - BRK/Carrefour - {transportadora}'

    msg = MIMEMultipart()
    msg['From'] = seu_email
    msg['To'] = destinatario
    msg['Subject'] = assunto
    msg.attach(MIMEText(corpo_email, 'plain'))

    # Envia o e-mail
    try:
        with smtplib.SMTP(servidor, porta) as server:
            server.starttls()
            server.login(seu_email, sua_senha)
            server.sendmail(seu_email, destinatario, msg.as_string())
        print(f'E-mail enviado para {destinatario} - Transportadora: {transportadora}')
    except Exception as e:
        print(f'Erro ao enviar e-mail para {destinatario}: {e}')

# Processar os dados e enviar os e-mails
for transportadora in df_dados["TRANSPORTADORA"].unique():
    # Filtrar os dados da transportadora
    df_filtro = df_dados[df_dados["TRANSPORTADORA"] == transportadora]

    # Obter os e-mails da transportadora
    email_transp = ", ".join(df_filtro["Email_Transp"].dropna().unique())

    # Criar corpo do e-mail
    corpo_email = f"""{email_transp}
samantha_campos@carrefour.com; ana_cristina_moraes@carrefour.com;
br_torre_controle_crf@carrefour.com;leticia_andrade_reis@carrefour.com;
clovis_da_silva@carrefour.com;

Prezados saudações!

Favor seguir com o desconto abaixo da transportadora, referente à escolta enviada pela nossa gerenciadora BRK.

"NR ORDEM", "DATA_HORA DO ACIONAMENTO", "PLACAS", "STATUS", "CLIENTE", "TRANSPORTADORA", "VALOR"
"""

    for _, row in df_filtro.iterrows():
        corpo_email += f'{row["NR ORDEM"]}, {row["DATA_HORA DO ACIONAMENTO"]}, {row["PLACAS"]}, {row["STATUS"]}, {row["CLIENTE"]}, {row["TRANSPORTADORA"]}, {row["VALOR"]}\n'

    # Obter o total do "Resumo" correspondente à transportadora
    total_resumo = df_resumo[df_resumo.iloc[:, 0] == transportadora].iloc[:, 1].sum()

    corpo_email += f'\nTotal: R$ {total_resumo:.2f}\n\nObrigado!'

    # Obter os destinatários únicos
    emails = df_filtro["Email"].dropna().unique()

    # Enviar e-mails
    for email in emails:
        enviar_email(email, transportadora, email_transp, corpo_email)