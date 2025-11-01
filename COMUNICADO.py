# pip install python-dotenv
# pip install pandas
# pip install openpyxl
# Desenvolvido por Fábio A Cordeiro
# Atualizado em 05/08/2025 - Corpo do e-mail alterado para comunicado CRFLOG

import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
from dotenv import load_dotenv

load_dotenv()
FROM_EMAIL = os.getenv('FROM_EMAIL')
BD_PASSWORD = os.getenv('EMAIL_PASSWORD')
BD_PORT = os.getenv('BD_PORT')

# Defina o caminho do arquivo Excel
excel_file = 'C:\\Fabio\\CARREFOUR\\BRK\\PRONTA_RESPOSTA\\Rota_MG.xlsx'

# Leia as abas da planilha
df_dados = pd.read_excel(excel_file, sheet_name="Dados_Final").fillna('')
df_postos = pd.read_excel(excel_file, sheet_name="Postos").fillna('')

# Configurações do servidor SMTP
servidor = 'smtp.gmail.com'
porta = int(BD_PORT) if BD_PORT else 587

# Função para enviar o e-mail
def enviar_email(destinatarios, copia, transportadora, corpo_email):
    assunto = f'COMUNICADO - CRFLOG - {transportadora}'
    msg = MIMEMultipart()
    msg['From'] = FROM_EMAIL
    msg['To'] = ", ".join(destinatarios)
    msg['Cc'] = ", ".join(copia) if copia else ""
    msg['Subject'] = assunto
    msg.attach(MIMEText(corpo_email, 'html'))

    destinatarios_finais = destinatarios + copia

    try:
        with smtplib.SMTP(servidor, porta) as server:
            server.starttls()
            server.login(FROM_EMAIL, BD_PASSWORD)
            server.sendmail(FROM_EMAIL, destinatarios_finais, msg.as_string())
        print(f'E-mail enviado para {", ".join(destinatarios)} com cópia para {", ".join(copia)} - Transportadora: {transportadora}')
    except Exception as e:
        print(f'Erro ao enviar e-mail para {", ".join(destinatarios)}: {e}')

# Corpo do e-mail conforme comunicado CRFLOG
corpo_email_modelo = """
<html>
<body style="font-family: Arial, sans-serif; color: #333;">

    <h2 style="color: #004aad; text-align: center;">
        COMUNICADO AOS TERCEIROS CRFLOG
    </h2>
    <h3 style="color: red; text-align: center;">
        Atenção Srs. Transportadores Subcontratados CRFLOG
    </h3>

    <p>
        Pensando em melhoria nos processos e controles dos Transportadores, informamos que a partir do mês 
        <strong>09/2025</strong>, os pagamentos da CRFLOG aos Subcontratados Pessoa Jurídica serão feitos nas
        dezenas, dias <strong>10, 20</strong> e <strong>30</strong> de cada mês, assim como é feito no Grupo Carrefour 
    </p>
    <p>
        O prazo de pagamento continua 30 dias, porém respeita a janela de pagamento conforme demonstrado abaixo:
    </p>

    <p>
       <strong>#Exemplo:</strong>
    </p>
    <p>

        Emissões 01 a 09/08 - Serão pagos no dia 10/09
    </p>
    <p>
        Emissões 10 a 19/08 - Serão pagos no dia 22/09
    </p>
    <p>
        Emissões 20 a 31/08 - Serão pagos no dia 30/09
    </p>

    <p>
        Sobre os pagamentos dos Subcontratados Autônomos, não terão nenhuma alteração.

        Em caso de dúvidas sobre esse assunto, favor enviar por e-mail para:
    </p>
    <p style="text-align: center; font-size: 16px;">
        <a href="mailto:br_financeiro_tb@carrefour.com" style="color: #004aad; font-weight: bold;">
            br_financeiro_tb@carrefour.com
        </a>
    </p>
    <p style="margin-top: 40px;">
        Atenciosamente,<br>

        <strong>Time Financeiro CRFLOG</strong>
        <p>
            <strong>Osasco,05 de Agosto de 2025</strong>
        </p>
    </p>

</body>
</html>
"""

# Processar e enviar
for transportadora in df_dados["TRANSPORTADORA"].unique():
    df_filtro = df_dados[df_dados["TRANSPORTADORA"] == transportadora]

    emails_transp = df_filtro["Email_transp"].astype(str).str.strip().replace('', None).dropna().unique()
    destinatarios = emails_transp.tolist() if len(emails_transp) > 0 else []

    emails_carrefour = df_filtro["Email_Carrefour"].astype(str).str.strip().replace('', None).dropna().unique()
    copia = emails_carrefour.tolist() if len(emails_carrefour) > 0 else []

    if destinatarios:
        enviar_email(destinatarios, copia, transportadora, corpo_email_modelo)
