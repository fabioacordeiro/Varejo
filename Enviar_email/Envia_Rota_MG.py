# Desenvolvido por Fábio A Cordeiro
# Em 28/02/2025

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
excel_file = 'C:\\Fabio\\CARREFOUR\\BRK\\PRONTA_RESPOSTA\\Rota_MG.xlsx'

# Leia as abas da planilha
df_dados = pd.read_excel(excel_file, sheet_name="Dados_Final").fillna('')
df_resumo = pd.read_excel(excel_file, sheet_name="Postos").fillna('')

# Configurações do servidor SMTP
servidor = 'smtp.gmail.com'
porta = 587

# Função para enviar o e-mail
def enviar_email(destinatario, transportadora, email_transp, corpo_email):
    assunto = f'Informativo - ROTA SP x MG - CARREFOUR - {transportadora}'

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
        <p>Prezados saudações!</p>
        <p>Segue abaixo as diretrizes do PGR para as cargas efetuadas no Carrefour e Sam's em São Paulo para Minas Gerais.</p>
        <p>Seguir pela Fernão Dias sem desvios;</p>
        <p>Ao sair carregado com destino a qualquer loja de MG é expressamente proibido a parada em qualquer local antes de rodar o mínimo de 150km da origem;</p>
        <p>Postos Homologados para Paradas e Pernoite:</p>
        <table style="width: 100%; border-collapse: collapse; background-color: #f0f0f0;">
            <tr style="background-color: #d3d3d3; text-align: left;">
                <th style="padding: 8px; border: 1px solid #999;">NOME DO POSTO</th>
                <th style="padding: 8px; border: 1px solid #999;">CIDADE</th>
                <th style="padding: 8px; border: 1px solid #999;">UF</th>
                <th style="padding: 8px; border: 1px solid #999;">DISTÂNCIA DE OSASCO-SP</th>
                <th style="padding: 8px; border: 1px solid #999;">LATITUDE</th>
                <th style="padding: 8px; border: 1px solid #999;">LONGITUDE</th>
                <th style="padding: 8px; border: 1px solid #999;">HOMOLOGADO</th>
            </tr>
    """

    # Adicionar as linhas da tabela
    for _, row in df_filtro.iterrows():
        #valor_formatado = locale.currency(row["VALOR"], grouping=True, symbol=True)
        corpo_email += f"""
            <tr>
                <td style="padding: 8px; border: 1px solid #999;">{row["NOME DO POSTO"]}</td>
                <td style="padding: 8px; border: 1px solid #999;">{row["CIDADE"]}</td>
                <td style="padding: 8px; border: 1px solid #999;">{row["UF"]}</td>
                <td style="padding: 8px; border: 1px solid #999;">{row["DISTANCIA DO CDA"]}</td>
                <td style="padding: 8px; border: 1px solid #999;">{row["LATITUDE"]}</td>
                <td style="padding: 8px; border: 1px solid #999;">{row["LONGITUDE"]}</td>
                <td style="padding: 8px; border: 1px solid #999;">{row["HOMOLOGADO"]}</td>
            </tr>
        """

    # Obter o total do "Resumo" correspondente à transportadora
    total_resumo = df_resumo[df_resumo.iloc[:, 0] == transportadora].iloc[:, 1].sum()
    total_formatado = locale.currency(total_resumo, grouping=True, symbol=True)

    # Fechar a tabela e incluir o total
    corpo_email += f"""
        </table>
        <p>O restante dos postos são de alto risco e proibidos para parada/pernoite</p>
        <p>A loja Carrefour BHA - em Belo Horizonte o recebimento é das 08:00 até 12:00</p>
        <p>As outras lojas recebem entre 06:30 e 10:00 da manhã, caso a loja não receba por qualquer motivo se deslocar para o posto Homologado Estacionamento São Rafael Carretas End: Rua Maria de Araújo,790 - Santa Cruz - Betim - MG - CEP:32667-464.</p>
        <p>Em caso dúvidas ou necessidad de suporte entre em contato com a BRK</p>
        <p>Contatos BRK: 3028-1600 OU 0800 600 1499</p>
        <p>Obrigado!</p>
    </body>
    </html>
    """

    # Obter os destinatários únicos
    emails = df_filtro["Email"].dropna().unique()

    # Enviar e-mails
    for email in emails:
        enviar_email(email, transportadora, email_transp, corpo_email)