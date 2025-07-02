# pip install python-dotenv
# pip install pandas
# pip install MIMEText
# pip install smtplib
# Desenvolvido por Fábio A Cordeiro
# Em 28/02/2025

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
excel_file = 'C:\\Fabio\\CARREFOUR\\BRK\\PRONTA_RESPOSTA\\Rota_MG.xlsx'

# Leia as abas da planilha
df_dados = pd.read_excel(excel_file, sheet_name="Dados_Final").fillna('')
df_postos = pd.read_excel(excel_file, sheet_name="Postos").fillna('')

# Configurações do servidor SMTP
servidor = 'smtp.gmail.com'
porta = 587

# Função para enviar o e-mail
def enviar_email(destinatarios, copia, transportadora, corpo_email):
    assunto = f'Informativo - ROTA-SP x BH-MG - CARREFOUR - {transportadora}'

    msg = MIMEMultipart()
    msg['From'] = seu_email
    msg['To'] = ", ".join(destinatarios)
    msg['Cc'] = ", ".join(copia) if copia else ""
    msg['Subject'] = assunto
    msg.attach(MIMEText(corpo_email, 'html'))  # Define o corpo como HTML

    destinatarios_finais = destinatarios + copia

    # Envia o e-mail
    try:
        with smtplib.SMTP(servidor, porta) as server:
            server.starttls()
            server.login(seu_email, sua_senha)
            server.sendmail(seu_email, destinatarios_finais, msg.as_string())
        print(f'E-mail enviado para {", ".join(destinatarios)} com cópia para {", ".join(copia)} - Transportadora: {transportadora}')
    except Exception as e:
        print(f'Erro ao enviar e-mail para {", ".join(destinatarios)}: {e}')

# Processar os dados e enviar os e-mails
for transportadora in df_dados["TRANSPORTADORA"].unique():
    # Filtrar os dados da transportadora
    df_filtro = df_dados[df_dados["TRANSPORTADORA"] == transportadora]

    # Obter os destinatários principais (Email_Transp)
    emails_transp = df_filtro["Email_transp"].astype(str).str.strip().replace('', None).dropna().unique()
    destinatarios = emails_transp.tolist() if len(emails_transp) > 0 else []

    # Obter os e-mails para cópia (Email_Carrefour)
    emails_carrefour = df_filtro["Email_Carrefour"].astype(str).str.strip().replace('', None).dropna().unique()
    copia = emails_carrefour.tolist() if len(emails_carrefour) > 0 else []

    # Criar corpo do e-mail em HTML
    corpo_email = f"""
    <html>
    <body style="font-family: Arial, sans-serif; color: #333;">
        <p>Prezados saudações!</p>
        <p>Segue abaixo as diretrizes do PGR para as cargas efetuadas no Carrefour e Sam's em São Paulo para Belo Horizonte - Minas Gerais.</p>
        <p><strong>Regras principais:</strong></p>
        <ul>
            <li>Seguir pela Fernão Dias sem desvios.</li>
            <li>Após sair carregado, é proibido parar antes de rodar no mínimo 150km.</li>
        </ul>
        <p><strong>Postos Homologados para Paradas e Pernoite:</strong></p>
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

    # Adicionar os dados dos postos homologados
    for _, row in df_postos.iterrows():
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

    corpo_email += """
        </table>
        <p><strong>Observações:</strong></p>
        <ul>
            <li>Os demais postos são considerados de alto risco e não devem ser utilizados.</li>
            <li>A loja Carrefour BHA em Belo Horizonte recebe entre 08:00 e 12:00hs.</li>
            <li>As outras lojas recebem entre 06:30 e 10:00hs.</li>
            <li>Se houver problemas no recebimento, deslocar-se para o posto homologado São Rafael Carretas em Betim - MG.</li>
            <li>Endereço: Rua Maria de Araújo, 790 - Santa Cruz - Betim - MG.</li>
        </ul>
        <p><strong>Contato BRK:</strong> 3028-1600 ou 0800 600 1499</p>
        <p>Atenciosamente,</p>
        <p>Equipe Carrefour / BRK</p>
    </body>
    </html>
    """

    # Enviar e-mails se houver destinatários
    if destinatarios:
        enviar_email(destinatarios, copia, transportadora, corpo_email)