# pip install pandas
# pip install python-dotenv
# pip install openpyxl
# pip install sendgrid


import os
import pandas as pd
from dotenv import load_dotenv
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail, Personalization, To, Cc

# Tratamento de exceção do SendGrid
try:
    from python_http_client.exceptions import HTTPError
except Exception:
    HTTPError = Exception

# Carrega variáveis
load_dotenv()

SENDGRID_API_KEY = os.getenv("SENDGRID_API_KEY")
FROM_EMAIL = os.getenv("FROM_EMAIL")

if not SENDGRID_API_KEY or not FROM_EMAIL:
    raise RuntimeError("Verifique as variáveis SENDGRID_API_KEY e FROM_EMAIL no arquivo .env")

def parse_emails_semicolon(raw: str) -> list[str]:
    """Limpa, separa por ponto e vírgula e remove duplicatas."""
    raw = str(raw or "").strip()
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

def build_message(from_email, to_emails, cc_emails, subject, body_plain) -> Mail:
    message = Mail(
        from_email=from_email,
        subject=subject,
        plain_text_content=body_plain
    )
    personalization = Personalization()
    for e in to_emails:
        personalization.add_to(To(e))
    for e in cc_emails:
        personalization.add_cc(Cc(e))
    message.add_personalization(personalization)
    return message

def main():
    excel_file = 'licenca_sanitaria_vencida.xlsx'

    if not os.path.exists(excel_file):
        raise FileNotFoundError(f"Arquivo não encontrado: {excel_file}")

    # Lê a planilha convertendo tudo para string para evitar erros de formatação
    print(f"Lendo e agrupando dados de: {excel_file}...")
    df = pd.read_excel(excel_file, dtype=str).fillna('')

    colunas_necessarias = ['Email_transp', 'TRANSP', 'CNPJ']
    for col in colunas_necessarias:
        if col not in df.columns:
            raise RuntimeError(f"Coluna '{col}' não encontrada.")

    sg = SendGridAPIClient(SENDGRID_API_KEY)
    
    # CC Fixo (Solicitado no original)
    fixed_cc_raw = "leticia_andrade_reis@carrefour.com"
    fixed_cc_list = parse_emails_semicolon(fixed_cc_raw)

    # --- NOVA LÓGICA: AGRUPAR POR TRANSPORTADORA ---
    # Isso cria grupos onde a chave é o nome da transportadora
    grupos = df.groupby('TRANSP')

    total_grupos = len(grupos)
    print(f"Total de transportadoras únicas encontradas: {total_grupos}\n")

    for transportadora, dados_transp in grupos:
        # 'dados_transp' é um DataFrame contendo apenas as linhas dessa transportadora
        
        # 1. Obter lista de CNPJs únicos para esta transportadora
        lista_cnpjs = dados_transp['CNPJ'].unique()
        # Formata a lista com um traço e quebra de linha para ficar bonito no e-mail
        cnpjs_formatados = "\n".join([f"- {cnpj.strip()}" for cnpj in lista_cnpjs if cnpj.strip()])

        # 2. Obter E-mails
        # Pega o email principal da primeira linha (assume-se que é o mesmo para a transportadora)
        email_to_raw = dados_transp['Email_transp'].iloc[0]
        lista_to = parse_emails_semicolon(email_to_raw)

        if not lista_to:
            print(f"⚠️ Pular: {transportadora} (Sem e-mail válido cadastrado)")
            continue

        # 3. Obter Cópias (Junta as cópias de TODAS as linhas dessa transportadora, caso variem)
        copias_raw_list = dados_transp.get('copias', pd.Series()).tolist()
        lista_cc_variavel = []
        for c in copias_raw_list:
            lista_cc_variavel.extend(parse_emails_semicolon(c))
        
        # Junta CC fixo + CC variável e remove duplicatas
        final_cc_list = list(set(fixed_cc_list + lista_cc_variavel))

        # 4. Montar o Corpo do E-mail
        assunto = f'Licença da Vigilância Sanitária - Carrefour - {transportadora}'

        corpo = f"""
Prezados saudações !!!

Identificamos o(s) CNPJ(s) abaixo com a Licença da Vigilância Sanitária vencida(s).
Reforço que passaremos por auditoria de documentação e todas as transportadoras
devem estar com a documentação em dia.
Qualquer dificuldade na apresentação do documento válido, favor enviar o protocolo.
Favor confirmar o recebimento deste e-mail.

Documento requerido do(s) CNPJ(s):
{cnpjs_formatados}



Obs: As transportadoras que não enviarem a documentação válida até o dia 20/02/2026 e 
não se justificar, serão bloqueadas para carregamento.

Conto com o apoio de todos.
Obrigado !!!
"""

        # 5. Enviar
        try:
            message = build_message(
                from_email=FROM_EMAIL,
                to_emails=lista_to,
                cc_emails=final_cc_list,
                subject=assunto,
                body_plain=corpo
            )
            
            print(f"📤 Enviando para: {transportadora} | CNPJs: {len(lista_cnpjs)} | Dest: {lista_to[0]}")
            response = sg.send(message)
            
            if response.status_code in [200, 201, 202]:
                print(f"   ✅ Sucesso! (Status: {response.status_code})")
            else:
                print(f"   ❌ Falha (Status: {response.status_code})")

        except HTTPError as e:
            print(f"   ❌ Erro API SendGrid: {e.status_code}")
            if e.body:
                try:
                    print(f"   Detalhe: {e.body.decode('utf-8')}")
                except:
                    print(e.body)
        except Exception as e:
            print(f"   ❌ Erro genérico: {e}")

    print("\nProcesso finalizado.")

if __name__ == "__main__":
    main()