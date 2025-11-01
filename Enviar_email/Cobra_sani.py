#Desenvolvido por Fábio A Cordeiro 
#Em 22/02/2025 
#Instale as bibliotecas necessárias:

#pandas para manipulação de dados.
#openpyxl para leitura de arquivos Excel.
#smtplib para envio de e-mails.

import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv

# Defina o caminho do arquivo Excel
excel_file = 'licenca_sanitaria_vencida.xlsx'

# Leia a planilha
df = pd.read_excel(excel_file).fillna('')  # Substitui valores NaN por string vazia

# Carrega variáveis de ambiente
load_dotenv()
FROM_EMAIL = os.getenv("FROM_EMAIL")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

# Configurações do servidor SMTP
servidor = 'smtp.gmail.com'
porta = 587

# Função para enviar o e-mail
def enviar_email(destinatario, transportadora, email_transp, cnpj):
    # Defina o assunto personalizado
    assunto = f'Licença da Vigilância Sanitária - Carrefour - 
    {transportadora}'

    # Defina o corpo personalizado
    corpo = f'''{email_transp}

Prezados, saudações!

Favor nos encaminhar a Licença da Vigilância Sanitária até o dia 20/06/25.
Reforço que passaremos por auditoria de documentação e todas as transportadoras
devem estar com a documentação em dia.
Qualquer dificuldade na apresentação do documento válido, favor enviar o protocolo.
Favor confirmar o recebimento deste e-mail.

{cnpj}

Obrigado
'''

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
        server.sendmail(seu_email, destinatario, msg.as_string())

# Envie e-mails para cada destinatário na planilha
for _, row in df.iterrows():
    email = row['Email']
    transportadora = row['TRANSP']
    email_transp = row['email_transp']
    cnpj = row['CNPJ']
    
    try:
        enviar_email(email, transportadora, email_transp, cnpj)
        print(f'E-mail enviado para {email}')
    except Exception as e:
        print(f'Erro ao enviar e-mail para {email}: {e}')