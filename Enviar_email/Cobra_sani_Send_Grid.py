# pip install pandas
# pip install python-dotenv
# pip install openpyxl
# pip install sendgrid


import os
import pandas as pd
from dotenv import load_dotenv
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail, Personalization, To, Cc, Email

# Tenta importar exceções específicas do cliente HTTP do SendGrid
try:
    from python_http_client.exceptions import HTTPError
except Exception:
    HTTPError = Exception

# Carrega variáveis de ambiente
load_dotenv()

# Configurações iniciais
SENDGRID_API_KEY = os.getenv("SENDGRID_API_KEY")
FROM_EMAIL = os.getenv("FROM_EMAIL", "br_backoffice_transporte_cda@carrefour.com")

if not SENDGRID_API_KEY:
    raise RuntimeError("Variável SENDGRID_API_KEY não encontrada. Configure no .env.")
if not FROM_EMAIL:
    raise RuntimeError("Variável FROM_EMAIL não encontrada. Configure no .env.")

def parse_emails_semicolon(raw: str) -> list[str]:
    """
    Extrai e-mails separados por ';', filtra '@' e remove duplicados 
    preservando a ordem. Útil para o campo 'copias' e destinatários múltiplos.
    """
    raw = str(raw or "").strip()
    # Divide por ; ou , para garantir
    parts = raw.replace(',', ';').split(";")
    emails = [e.strip() for e in parts if e.strip() and "@" in e]
    
    seen = set()
    out = []
    for e in emails:
        k = e.lower()
        if k not in seen:
            out.append(e)
            seen.add(k)
    return out

def build_message(
    from_email: str,
    to_emails: list[str],
    cc_emails: list[str],
    subject: str,
    body_plain: str
) -> Mail:
    """
    Constrói o objeto de e-mail do SendGrid sem anexo (texto simples),
    conforme a lógica original do script de cobrança.
    """
    message = Mail(
        from_email=from_email,
        subject=subject,
        plain_text_content=body_plain
    )

    personalization = Personalization()
    
    # Adiciona Destinatários (To)
    if not to_emails:
        raise ValueError("Lista de destinatários vazia.")
        
    for e in to_emails:
        personalization.add_to(To(e))
        
    # Adiciona Cópias (Cc)
    for e in cc_emails:
        personalization.add_cc(Cc(e))
        
    message.add_personalization(personalization)
    
    return message

def main():
    # Defina o caminho do arquivo Excel
    excel_file = 'licenca_sanitaria_vencida.xlsx'

    # Verifica se arquivo existe
    if not os.path.exists(excel_file):
        raise FileNotFoundError(f"Arquivo não encontrado: {excel_file}")

    # Leia a planilha
    print(f"Lendo arquivo: {excel_file}...")
    df = pd.read_excel(excel_file).fillna('')

    # Validação básica de colunas
    colunas_necessarias = ['Email_transp', 'TRANSP', 'CNPJ']
    for col in colunas_necessarias:
        if col not in df.columns:
            raise RuntimeError(f"Coluna obrigatória '{col}' não encontrada na planilha.")

    # Inicializa cliente SendGrid
    sg = SendGridAPIClient(SENDGRID_API_KEY)

    # Cópias fixas (se houver, baseado no script original havia uma variável 'copias' fora do loop)
    # Se quiser adicionar um CC fixo para todos os emails, adicione aqui:
    fixed_cc_list = parse_emails_semicolon("leticia_andrade_reis@carrefour.com") 

    # Itera sobre as linhas da planilha
    for index, row in df.iterrows():
        transportadora = row['TRANSP']
        cnpj = row['CNPJ']
        
        # Tratamento dos e-mails da linha
        email_destinatario_raw = row['Email_transp']
        email_copias_raw = row.get('copias', '') # Pega da coluna 'copias' se existir
        
        # Processa as listas de e-mail
        emails_to = parse_emails_semicolon(email_destinatario_raw)
        emails_cc_row = parse_emails_semicolon(email_copias_raw)
        
        # Junta CC fixo com CC da linha
        final_cc_list = list(set(fixed_cc_list + emails_cc_row))

        if not emails_to:
            print(f"⚠️ Pular linha {index}: E-mail inválido para {transportadora}")
            continue

        # Define o assunto
        assunto = f'Licença da Vigilância Sanitária - Carrefour - {transportadora}'

        # Define o corpo (Mantido exatamente como no original)
        corpo = f"""
Prezados saudações !!!

Idenficamos o(s) CNPJ(s) abaixo com a Licença da Vigilância Sanitária vencida(s).
Reforço que passaremos por auditoria de documentação e todas as transportadoras
devem estar com a documentação em dia.
Qualquer dificuldade na apresentação do documento válido, favor enviar o protocolo.
Favor confirmar o recebimento deste e-mail.

Documento requerido do(s) CNPJ(s): {cnpj}



Obs: As transportadoras que não enviarem a documentação válida até o dia 20/02/2026 e 
não se justificar, serão bloqueadas para carregamento.

Conto com o apoio de todos.
Obrigado !!!
"""
        
        try:
            # Constrói a mensagem
            message = build_message(
                from_email=FROM_EMAIL,
                to_emails=emails_to,
                cc_emails=final_cc_list,
                subject=assunto,
                body_plain=corpo
            )

            # Envia
            print(f"📤 Enviando para: {transportadora} ({emails_to[0]})...")
            response = sg.send(message)

            if response.status_code in [200, 201, 202]:
                print(f"✅ Sucesso! Status: {response.status_code}")
            else:
                print(f"❌ Falha no envio. Status: {response.status_code}")

        except HTTPError as e:
            print(f"❌ Erro HTTP ao enviar para {transportadora}:")
            print(f"   Status: {e.status_code}")
            if e.body:
                print(f"   Detalhe: {e.body}")
                
        except Exception as e:
            print(f"❌ Erro genérico ao enviar para {transportadora}: {e}")

if __name__ == "__main__":
    main()