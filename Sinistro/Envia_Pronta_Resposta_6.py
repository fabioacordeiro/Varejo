# Desenvolvido por Fábio A Cordeiro
# Em 22/02/2025
# Desenvolvido por Fábio A Cordeiro
# Em 22/02/2025

import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import locale
import os
from dotenv import load_dotenv #type:igore
load_dotenv()
FROM_EMAIL= os.getenv('FROM_EMAIL')
BD_PASSWORD = os.getenv('EMAIL_PASSWORD')

# Defina o seu e-mail e senha (use um e-mail e senha de aplicação)
seu_email = FROM_EMAIL
sua_senha = BD_PASSWORD

# Definir local para formato de moeda brasileira
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# Defina o caminho do arquivo Excel
excel_file = 'C:\\Fabio\\CARREFOUR\\BRK\\PRONTA_RESPOSTA\\Pronta_resposta.xlsx'

# Leia as abas da planilha
df_dados = pd.read_excel(excel_file, sheet_name="Dados").fillna('')
df_resumo = pd.read_excel(excel_file, sheet_name="Resumo").fillna('')

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
    msg.attach(MIMEText(corpo_email, 'html'))  # Define o corpo como HTML

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
    emails_transp = df_filtro["Email_transp"].astype(str).str.strip().replace('', None).dropna().unique()
    email_transp = ", ".join(emails_transp) if len(emails_transp) > 0 else "Nenhum e-mail disponível"

    # Criar corpo do e-mail em HTML
    corpo_email = f"""
    <html>
    <body style="font-family: Arial, sans-serif; color: #333;">
        <p><strong>{email_transp}</strong></p>
        <p> samantha_campos@carrefour.com; ana_cristina_moraes@carrefour.com;
br_torre_controle_crf@carrefour.com;leticia_andrade_reis@carrefour.com;
clovis_da_silva@carrefour.com; </p>
        <p>Prezados saudações!</p>
        <p>Favor seguir com o desconto abaixo da transportadora, referente à escolta enviada pela nossa gerenciadora de risco BRK.</p>
        <p>Conforme PGR assinado pelo transportador especificado no item 11.</p>
        <p>Item 11. PROCEDIMENTOS DE PRONTA RESPOSTA</p>
        <p>11.1.6. Para proteção de veículos quebrados e/ou com problemas mecânicos ou não cumprimento das normas contidas no PGR custo será revertido ao transportador.</p>
        <table style="width: 100%; border-collapse: collapse; background-color: #f0f0f0;">
            <tr style="background-color: #d3d3d3; text-align: left;">
                <th style="padding: 8px; border: 1px solid #999;">NR ORDEM</th>
                <th style="padding: 8px; border: 1px solid #999;">DATA_HORA DO ACIONAMENTO</th>
                <th style="padding: 8px; border: 1px solid #999;">PLACAS</th>
                <th style="padding: 8px; border: 1px solid #999;">STATUS</th>
                <th style="padding: 8px; border: 1px solid #999;">CLIENTE</th>
                <th style="padding: 8px; border: 1px solid #999;">TRANSPORTADORA</th>
                <th style="padding: 8px; border: 1px solid #999;">VALOR</th>
            </tr>
    """

    # Adicionar as linhas da tabela
    for _, row in df_filtro.iterrows():
        valor_formatado = locale.currency(row["VALOR"], grouping=True, symbol=True)
        corpo_email += f"""
            <tr>
                <td style="padding: 8px; border: 1px solid #999;">{row["NR ORDEM"]}</td>
                <td style="padding: 8px; border: 1px solid #999;">{row["DATA_HORA DO ACIONAMENTO"]}</td>
                <td style="padding: 8px; border: 1px solid #999;">{row["PLACAS"]}</td>
                <td style="padding: 8px; border: 1px solid #999;">{row["STATUS"]}</td>
                <td style="padding: 8px; border: 1px solid #999;">{row["CLIENTE"]}</td>
                <td style="padding: 8px; border: 1px solid #999;">{row["TRANSPORTADORA"]}</td>
                <td style="padding: 8px; border: 1px solid #999;">{valor_formatado}</td>
            </tr>
        """

    # Obter o total do "Resumo" correspondente à transportadora
    total_resumo = df_resumo[df_resumo.iloc[:, 0] == transportadora].iloc[:, 1].sum()
    total_formatado = locale.currency(total_resumo, grouping=True, symbol=True)

    # Fechar a tabela e incluir o total
    corpo_email += f"""
        </table>
        <p><strong>Total: {total_formatado}</strong></p>
        <p>Obrigado!</p>
    </body>
    </html>
    """

    # Obter os destinatários únicos
    emails = df_filtro["Email"].dropna().unique()

    # Enviar e-mails
    for email in emails:
        enviar_email(email, transportadora, email_transp, corpo_email)