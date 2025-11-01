# pip install python-dotenv
# pip install pandas
# pip install openpyxl

import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
from dotenv import load_dotenv

# === CONFIGURAÇÃO ===
EMAIL_TESTE = "seuemail@dominio.com"  # <- coloque aqui seu e-mail de teste

# Carregar variáveis de ambiente
load_dotenv()
FROM_EMAIL = os.getenv('FROM_EMAIL')
BD_PASSWORD = os.getenv('EMAIL_PASSWORD')
BD_PORT = os.getenv('BD_PORT') or 587

# Caminho do arquivo Excel
excel_file = r'C:\\Fabio\\CARREFOUR\BRK\\PRONTA_RESPOSTA\\Rota_MG.xlsx'

# Ler planilha
df_dados = pd.read_excel(excel_file, sheet_name="Dados_Final").fillna('')
print(f"📊 Registros lidos da planilha: {len(df_dados)}")

# Configurações SMTP
servidor = 'smtp.gmail.com'
porta = int(BD_PORT) if BD_PORT else 587

# Função para limpar lista de e-mails
def limpar_emails(coluna):
    emails_limpos = []
    for e in coluna:
        if pd.notna(e):
            for email in str(e).replace(',', ';').split(';'):
                email = email.strip()
                if email:
                    emails_limpos.append(email)
    return list(set(emails_limpos))

# Função para enviar e-mail
def enviar_email(destinatarios, copia, transportadora, corpo_email):
    assunto = f'COMUNICADO - CRFLOG - {transportadora}'
    msg = MIMEMultipart()
    msg['From'] = FROM_EMAIL
    msg['To'] = ", ".join(destinatarios)
    msg['Cc'] = ", ".join(copia) if copia else ""
    msg['Subject'] = assunto
    msg.attach(MIMEText(corpo_email, 'html', 'utf-8'))

    destinatarios_finais = destinatarios + (copia if copia else [])

    try:
        with smtplib.SMTP(servidor, porta) as server:
            server.starttls()
            server.login(FROM_EMAIL, BD_PASSWORD)
            server.sendmail(FROM_EMAIL, destinatarios_finais, msg.as_string())
        print(f"✅ E-mail enviado para {', '.join(destinatarios_finais)} - Transportadora: {transportadora}")
    except Exception as e:
        print(f"❌ Erro ao enviar e-mail para {', '.join(destinatarios_finais)}: {e}")

# Corpo do e-mail
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

# Processamento e debug
for transportadora in df_dados["TRANSPORTADORA"].unique():
    df_filtro = df_dados[df_dados["TRANSPORTADORA"] == transportadora]

    print(f"\n📦 Transportadora: {transportadora}")
    print("   Email_transp original:", list(df_filtro["Email_transp"]))
    print("   Email_Carrefour original:", list(df_filtro["Email_Carrefour"]))

    destinatarios = limpar_emails(df_filtro["Email_transp"])
    copia = limpar_emails(df_filtro["Email_Carrefour"])

    print("   Destinatários limpos:", destinatarios)
    print("   Cópia limpa:", copia)

    if not destinatarios and not copia:
        print(f"⚠️ Nenhum e-mail válido encontrado. Usando e-mail de teste: {EMAIL_TESTE}")
        destinatarios = [EMAIL_TESTE]

    enviar_email(destinatarios, copia, transportadora, corpo_email_modelo)