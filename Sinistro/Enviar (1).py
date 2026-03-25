# Desenvolvido por Fábio A Cordeiro
# Em 22/06/2025
# pip install python-dotenv

import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import locale
import os

# Importando os dados do arquivo .env e-mail e senha
from dotenv import load_dotenv
load_dotenv()
FROM_EMAIL = os.getenv('FROM_EMAIL')
BD_PASSWORD = os.getenv('EMAIL_PASSWORD')

# Definir local para formato de moeda brasileira
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# Função para envio de e-mail
def enviar_email(destinatario, transportadora, email_transp, corpo_email):
    assunto = f'Sinistro {num_sinistro} - {transportadora} - {motivo}'
    msg = MIMEMultipart()
    msg['From'] = FROM_EMAIL
    msg['To'] = destinatario
    msg['Subject'] = assunto
    msg.attach(MIMEText(corpo_email, 'html'))

    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(FROM_EMAIL, BD_PASSWORD)
            server.sendmail(FROM_EMAIL, destinatario, msg.as_string())
        print(f'E-mail enviado para {destinatario} - Transportadora: {transportadora}')
    except Exception as e:
        print(f'Erro ao enviar e-mail para {destinatario}: {e}')

# Carregar base de dados
xls_path = "C:\\Fabio\\Desenvolvimento\\Varejo\\Sinistro\\SINISTROS1.xlsx"
df_sinistro = pd.read_excel(xls_path, sheet_name='Dados')

# Processar e-mails
for index, row in df_sinistro.iterrows():
    if str(row.get("Email_enviado", "")).strip().upper() != "NÃO":
        continue
    
    num_sinistro = row["Nº Reguladora"]
    dt_sinistro= row["Data do Sinistro"]
    causa = str(row["Revisão Causa"]).strip().upper()
    cliente = row["Cliente"]
    carga = row["N_Carga"]
    data = row["Data do Sinistro"]
    motivo = row["Causa do Sinistro"]
    origem = f"{row['Cidade Origem']} - {row['UF - Origem']}"
    destino = f"{row['Cidade - Destino']} - {row['UF - Destino']}"
    local_sinistro = row["Local do Sinistro"]
    valor_embarcado = f"R$ {row['Valor do Embarque']}"
    transportadora = row["Transportador"]
    nota_fiscal = row["Nota Fiscal"]
    motorista = row["Motorista"]
    placas = row["Placa"]
    cte = row.get("CT-e", "Não informado")
    mdfe = row.get("MDF-e", "Não informado")

    if "TEMPERATURA" in causa:
        corpo_adicional = """<b>Documentos Pendentes - BRK - Rastreamento</b><ul>
<li>Autorização de Embarque / Solicitação de Monitoramento</li>
<li>Histórico de Posições</li>
<li>Histórico de Temperatura</li>
<li>Relatório de mensagens recebidas e enviadas</li>
<li>Relatório de comandos enviados</li>
<li>Relatório de Alertas recebidos</li>
<li>Histórico de alertas e comandos registrando o último teste ocorrido</li>
<li>Liberação do Motorista, Veículo e Ajudante</li>
<li>Histórico de Posições da Isca</li>
<li>Relatório da Gerenciadora de Riscos sobre acionamento de segurança</li>
<li>Cópia da Liberação do Motorista e Veículo</li></ul>
<b>Documentos da Transportadora:</b><ul>
<li>CNH</li><li>CRLV</li><li>ANTT</li><li>Ficha Cadastral</li><li>Discos de Tacógrafo</li></ul>"""
    elif "ACIDENTE" in causa:
        corpo_adicional = """<b>Documentos Pendentes - BRK - Rastreamento</b><ul>
<li>Autorização de Embarque / Solicitação de Monitoramento</li>
<li>Histórico de Temperatura - em caso de cargas Refrigeradas/Congeladas</li>
<li>Relatório de mensagens recebidas e enviadas</li>
<li>Relatório de comandos enviados</li>
<li>Relatório de Alertas recebidos</li>
<li>Liberação do Motorista, Veículo e Ajudante</li>
<li>Cópia da Liberação do Motorista e Veículo</li></ul>
<b>Documentos da Transportadora:</b><ul>
<li>CNH</li><li>CRLV</li><li>ANTT</li><li>Ficha Cadastral</li><li>Declaração do Motorista</li><li>Discos de Tacógrafo</li></ul>"""
    elif "ROUBO" in causa:
        corpo_adicional = """<b>Documentos Pendentes - BRK - Rastreamento</b><ul>
<li>Autorização de Embarque / Solicitação de Monitoramento</li>
<li>Histórico de Posições</li>
<li>Relatório de mensagens recebidas e enviadas</li>
<li>Relatório de comandos enviados</li>
<li>Relatório de Alertas recebidos</li>
<li>Histórico de alertas e comandos registrando o último teste ocorrido</li>
<li>Liberação do Motorista, Veículo e Ajudante</li>
<li>Histórico de Posições da Isca</li>
<li>Relatório da Gerenciadora de Riscos sobre acionamento de segurança</li>
<li>Cópia da Liberação do Motorista e Veículo</li></ul>
<b>Documentos da Transportadora:</b><ul>
<li>CNH</li><li>CRLV</li><li>ANTT</li><li>Ficha Cadastral</li><li>Discos de Tacógrafo</li></ul>"""
    else:
        corpo_adicional = "<p><b>Sem documentos adicionais definidos para esta causa.</b></p>"

    corpo_email = f"""
    <html>
        <body style="font-family: Arial, sans-serif; color: #333;">
            <p>Prezados,</p>
            <p>Identificamos o sinistro {num_sinistro} abaixo:</p>
            <p><b>Cliente:</b> {cliente}<br>
            <b>Reguladora: GLOBAL COMISSARIA</b><br>
            <b>Seguradora Akad</b><br>
            <b>Corretora: AON</b><br>
            <b>Carga:</b> {carga}<br>
            <b>Data:</b> {dt_sinistro}<br>
            <b>Motivo:</b> {motivo}<br>
            <b>Origem:</b> {origem}<br>
            <b>Destino:</b> {destino}<br>
            <b>Local do Sinistro:</b> {local_sinistro}<br>
            <b>Valor Embarcado:</b> {valor_embarcado}<br>
            <b>Nota Fiscal:</b> {nota_fiscal}<br>
            <b>Transportadora:</b> {transportadora}<br>
            <b>Motorista:</b> {motorista}<br>
            <b>Placas:</b> {placas}</p>
            {corpo_adicional}
        </body>
    </html>
    """

    destinatario = "fabio_cordeiro@carrefour.com"
    enviar_email(destinatario, transportadora, destinatario, corpo_email)
    print(enviar_email(destinatario, transportadora, destinatario, corpo_email))

print("Fim")
