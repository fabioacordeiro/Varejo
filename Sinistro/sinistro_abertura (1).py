# Desenvolvido por Fábio A Cordeiro
# Em 12/04/2025
# pip install pandas
# pip install python-dotenv

import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import locale
import os
# Importando os dados do arquivo .env e-mail e senha
from dotenv import load_dotenv #type:igore
load_dotenv()
FROM_EMAIL= os.getenv('FROM_EMAIL')
BD_PASSWORD = os.getenv('EMAIL_PASSWORD')

# Definir local para formato de moeda brasileira
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# Carregar o arquivo CSV para visualizar os dados
csv_path = "C:\\Fabio\\Desenvolvimento\\Varejo\\Sinistro\\Carga.csv"
df_carga = pd.read_csv(csv_path, sep=None, engine='python')  # Detectar automaticamente o separador

# Mostrar o nome das colunas se tiver algum erro no nome \ufeff no começo
print(df_carga.columns.tolist())
print("-" * 65)
bom_char = "\ufeff"
try:
    count = 0
    for col in df_carga.columns:
        if isinstance(col, str) and col.startswith(bom_char):
            print(col)  # 1 por linha
            count += 1

        print(f"Encontradas {count} colunas com início BOM (\ufeff).")

except FileNotFoundError:
    print(f"Arquivo não encontrado: {csv_path}")
except UnicodeDecodeError as e:
    print(f"Erro de decodificação ao ler o CSV: {e}")
    print("Dica: tente ler com encoding='utf-8-sig' para remover o BOM automaticamente.")

print("-" * 65)
# Mostrar as primeiras linhas do DataFrame
print(df_carga.head())

# Recarregar o CSV
df = pd.read_csv("C:\\Fabio\\Desenvolvimento\\Varejo\\Sinistro\\Carga.csv", sep=None, engine='python')
row = df.iloc[0]
print(df.head)
# Extração de dados equivalentes (com alguns campos preenchidos para exemplo)
cliente =  row["Filial"]
carga = row["\ufeffNúmero da Carga"]
data = row["Data da Carga"]
motivo = "VARIAÇÃO DE TEMPERATURA / ROUBO / ACIDENTE"
tipo_carga = row["Tipo de Carga"]
faixa_temperatura = row["Temperatura"]
origem = f"{row['Origem']} - {row['UF Origem']}"
destino = f"{row['Código Destinatário']} - {row['Destino']} - {row['UF Destino']} ({row['Entregas']} EntregaS)"
local_sinistro = "Verificar"
valor_embarcado = f"R$ {row['Valor NF']}"
transportadora = f"{row['Transportador']} - {row['CNPJ do Transportador']}"
nota_fiscal = f"{row['Nº NF-e']}" 
mdfe = f"{row['MDF-es']}"
cte = f"{row['Número CTe']}"          
motorista = f"{row['Motoristas']}" 
placas = row["Veículo"].replace(",", " / CARRETA:")

# Montar o texto formatado
texto_formatado = f"""cliente: {cliente}
Carga: {carga}
Data: {data}
MOTIVO: {motivo}
FAIXA DE TEMPERATURA: {faixa_temperatura}
Origem: {origem}
Destino: {destino}
LOCAL DO SINISTRO: {local_sinistro}
Valor Embarcado: {valor_embarcado}
Transportadora: {transportadora}
Nota fiscal: {nota_fiscal}
cte: {cte}
MDFE: {mdfe}
Motorista: {motorista}
Placas: {placas}
"""

# Salvar no novo arquivo TXT
output_txt_path = "C:\\Fabio\\Desenvolvimento\\Varejo\\Sinistro\\Carga_org_gerado.txt"
with open(output_txt_path, "w", encoding="utf-8") as f:
    f.write(texto_formatado)
    

output_txt_path

print("Concluído")
print(texto_formatado)



# Defina o seu e-mail e senha (use um e-mail e senha de aplicação)
seu_email = FROM_EMAIL
sua_senha = BD_PASSWORD

# Configurações do servidor SMTP
servidor = 'smtp.gmail.com'
porta = 587

# Função para enviar o e-mail
def enviar_email(destinatario, transportadora, email_transp, corpo_email):
    assunto = f'Dados de Sinistro - {transportadora}'

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

destinatario = 'fabio_cordeiro@carrefour.com'
email_transp = 'fabioacordeiro@yahoo.com.br'
corpo_email = f"""
<html>
    <body style="font-family: Arial, sans-serif; color: #333;">
        <p>Prezados saudações !</p>
        <p>Favor seguir com o sinistro abaixo:</p>
        <p></p>
        <p>Sinistro: </p>
        <p>Cliente: {cliente}</p>
        <p>Reguladora: GLOBAL COMISSARIA</p>
        <p>Seguradora: Akad</p>
        <p>Corretora: AON</p>
        <p>Carga: {carga}</p>
        <p>Data: {data}</p>
        <p>Motivo: {motivo}</p>
        <p>FAIXA DE TEMPERATURA: {faixa_temperatura}</p>
        <p>Origem: {origem}</p>
        <p>Destino: {destino}</p>
        <p>LOCAL DO SINISTRO: {local_sinistro}</p>
        <p>Valor Embarcado: {valor_embarcado}</p>
        <p>Transportadora: {transportadora}</p>
        <p>Nota fiscal: {nota_fiscal}</p>
        <p>CT-e: {cte}</p>
        <p>MDF-e: {mdfe}</p>
        <p>Motorista: {motorista}</p>
        <p>Placas: {placas}</p>
               

"""
#email = "fabio_cordeiro@carrefour.com; fabioacordeiro@yahoo.com.br"
email = "fabio_cordeiro@carrefour.com"
enviar_email(email, transportadora, email_transp, corpo_email)
        

print("Enviado")

# Criar DataFrame com os dados extraídos do ARQUIVO CSV
df_saida = pd.DataFrame([{
    "Cliente": cliente,
    "Carga": carga,
    "Data": data,
    "Motivo": motivo,
    "Faixa Temperatura": faixa_temperatura,
    "Origem": origem,
    "Destino": destino,
    "Local Sinistro": local_sinistro,
    "Valor Embarcado": valor_embarcado,
    "Transportadora": transportadora,
    "Nota Fiscal": nota_fiscal,
    "CT-e": cte,
    "MDF-e": mdfe,
    "Motorista": motorista,
    "Placas": placas
}])

# Salvar em planilha Excel 
output_excel_path = "C:\\Fabio\\Desenvolvimento\\Varejo\\Sinistro\\Carga_org_gerado.xlsx"
df_saida.to_excel(output_excel_path, index=False)

print ("Fim")