# Desenvolvido por Fábio A Cordeiro 
# Em 22/02/2025 

import pandas as pd
import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv #type:igore
load_dotenv()

FROM_EMAIL= os.getenv('FROM_EMAIL')
BD_PASSWORD = os.getenv('EMAIL_PASSWORD')

# Defina o seu e-mail e senha (use um e-mail e senha de aplicação)
seu_email = FROM_EMAIL
sua_senha = BD_PASSWORD
# Defina o caminho do arquivo Excel
excel_file = 'licenca_sanitaria_vencida.xlsx'

# Leia a planilha
df = pd.read_excel(excel_file).fillna('')  # Substitui valores NaN por string vazia

# Configurações do servidor SMTP
servidor = 'smtp.gmail.com'
porta = 587
copias = "samantha_campos@carrefour.com; br_torre_controle_crf@carrefour.com; leticia_andrade_reis@carrefour.com "
# Função para enviar o e-mail
def enviar_email(destinatario, copias, transportadora, cnpj):
    # Define o assunto personalizado
    assunto = f'Licença da Vigilância Sanitária - Carrefour - {transportadora}'

    # Define o corpo personalizado
    corpo = f"""
Prezados saudações !!!

Ultima cobrança, favor nos encaminhar a Licença da Vigilância Sanitária até o dia 30/06/25.
Reforço que passaremos por auditoria de documentação e todas as transportadoras
devem estar com a documentação em dia.
Qualquer dificuldade na apresentação do documento válido, favor enviar o protocolo.
Favor confirmar o recebimento deste e-mail.

Documento requerido do(s) CNPJ(s): {cnpj}

Obrigado !!!
"""

    # Cria a mensagem do e-mail
    msg = MIMEMultipart()
    msg['From'] = seu_email
    msg['To'] = destinatario
    msg['Cc'] = copias
    msg['Subject'] = assunto

    # Adiciona o corpo do e-mail
    msg.attach(MIMEText(corpo, 'plain'))
    
    # Lista final de destinatários (To + Cc)
    destinatarios_finais = [destinatario] + [e.strip() for e in copias.split(';') if e.strip()]
    
    # Envia o e-mail
    with smtplib.SMTP(servidor, porta) as server:
        server.starttls()
        server.login(seu_email, sua_senha)
        server.sendmail(seu_email, destinatarios_finais, msg.as_string())

# Envie e-mails para cada destinatário na planilha
for _, row in df.iterrows():
    destinatario = row['Email_transp']
    copias = row.get('copias', '')  # campo copias na planilha
    transportadora = row['TRANSP']
    cnpj = row['CNPJ']
    
    try:
        enviar_email(destinatario, copias, transportadora, cnpj)
        print(f'E-mail enviado para {destinatario} - Transportadora: {transportadora}')
    except Exception as e:
        print(f'Erro ao enviar e-mail para {destinatario}: {e}')