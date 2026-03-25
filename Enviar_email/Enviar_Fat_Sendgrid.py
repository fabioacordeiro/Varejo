# pip install pandas
# pip install python-dotenv
# pip install openpyxl
# pip install sendgrid

import os
import re
import base64
import pandas as pd
from dotenv import load_dotenv
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import (
    Mail, Personalization, To, Cc, Email,
    Attachment, FileContent, FileName, FileType, Disposition
)

try:
    from python_http_client.exceptions import HTTPError
except Exception:
    HTTPError = Exception

load_dotenv()
SENDGRID_API_KEY = os.getenv("SENDGRID_API_KEY")
FROM_EMAIL = os.getenv("FROM_EMAIL", "br_backoffice_transporte_cda@carrefour.com")

print("DEBUG_FROM_EMAIL =", FROM_EMAIL)

def sanitize_filename(name: str) -> str:
    """Evita caracteres inválidos no Windows e nomes vazios."""
    name = str(name or "").strip()
    name = re.sub(r'[<>:"/\|?*]', "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name if name else "prestador_sem_nome"


def parse_emails_semicolon(raw: str) -> list[str]:
    """Extrai e-mails separados por ';', filtra '@' e remove duplicados preservando ordem."""
    raw = str(raw or "").strip()
    emails = [e.strip() for e in raw.split(";") if e.strip() and "@" in e]
    seen = set()
    out = []
    for e in emails:
        k = e.lower()
        if k not in seen:
            out.append(e)
            seen.add(k)
    return out


def format_brl_currency(valor: float) -> str:
    """Formata 1234.5 -> 1.234,50"""
    s = f"{valor:,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def build_message(
    from_email: str,
    reply_to: str | None,
    to_emails: list[str],
    cc_emails: list[str],
    subject: str,
    body_plain: str,
    attachment_path: str
) -> Mail:
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

    if reply_to:
        message.reply_to = Email(reply_to)

    with open(attachment_path, "rb") as f:
        data = f.read()

    encoded = base64.b64encode(data).decode("utf-8")

    attachment = Attachment()
    attachment.file_content = FileContent(encoded)
    attachment.file_name = FileName(os.path.basename(attachment_path))
    attachment.file_type = FileType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    attachment.disposition = Disposition("attachment")

    message.add_attachment(attachment)
    return message


def main():
    # Carrega variáveis do .env (no mesmo diretório do script) ou do ambiente do sistema
    load_dotenv()

    SENDGRID_API_KEY = os.getenv("SENDGRID_API_KEY")
    FROM_EMAIL = os.getenv("FROM_EMAIL", "br_backoffice_transporte_cda@carrefour.com")
    REPLY_TO = os.getenv("REPLY_TO")  # opcional

    if not SENDGRID_API_KEY:
        raise RuntimeError("Variável SENDGRID_API_KEY não encontrada. Configure no .env.")
    if not FROM_EMAIL:
        raise RuntimeError("Variável FROM_EMAIL não encontrada. Configure no .env.")

    # Caminho para o arquivo e pasta de saída (igual ao seu script)
    input_path = r"C:\\Fabio\Desenvolvimento\\Varejo\\Envia_Fat\\BD_PAGAMENTOS.xlsx"
    output_dir = r"C:\\Fabio\Desenvolvimento\\Varejo\\Envia_Fat\\relatorios_pgts"
    os.makedirs(output_dir, exist_ok=True)

    # Lê a planilha
    df = pd.read_excel(input_path)

    # Valida colunas usadas
    required_cols = ["PRESTADOR", "Email_transp", "Valor Transação"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise RuntimeError(f"Colunas ausentes no Excel: {missing}")

    prestadores = df["PRESTADOR"].dropna().unique()

    # Cliente SendGrid fora do loop
    sg = SendGridAPIClient(SENDGRID_API_KEY)

    for prestador in prestadores:
        try:
            df_prestador = df[df["PRESTADOR"] == prestador]

            email_raw = str(df_prestador["Email_transp"].iloc[0]).strip()
            emails_to = parse_emails_semicolon(email_raw)

            if not emails_to:
                print(f"⚠️ Nenhum e-mail válido para: {prestador}")
                continue

            print(f"📧 Enviando para: {prestador} | {', '.join(emails_to)}")

            # Criar arquivo Excel com dados do prestador
            safe_name = sanitize_filename(prestador)
            file_name = os.path.join(output_dir, f"{safe_name}.xlsx")
            df_prestador.to_excel(file_name, index=False)

            # Total de pagamentos
            total_pagamento = float(df_prestador["Valor Transação"].sum())
            total_fmt = format_brl_currency(total_pagamento)

            # Assunto e corpo do e-mail (igual ao original)
            assunto = f"{prestador} - BD Pagto - CRFLOG(BOMPREÇO) - Dezembro/2025"
            corpo = (
                f"Prezado(a) {prestador},\n\n"
                f"Segue em anexo a base de pagamentos referente ao período informado.\n\n"
                f"Total de pagamentos: R$ {total_fmt}\n\n"
                f"Atenciosamente,\n"
                f"Equipe Financeira\n"
            )

            # CC fixo (igual ao original)
            emails_cc = "fabio_cordeiro@carrefour.com; br_financeiro_crflog@carrefour.com"
            lista_cc = parse_emails_semicolon(emails_cc)

            message = build_message(
                from_email=FROM_EMAIL,
                reply_to=REPLY_TO,
                to_emails=emails_to,
                cc_emails=lista_cc,
                subject=assunto,
                body_plain=corpo,
                attachment_path=file_name
            )

            response = sg.send(message)

            if response.status_code == 202:
                print(f"✅ E-mail enviado com sucesso para: {prestador} | Status: {response.status_code}")
            else:
                print(f"❌ Falha no envio para: {prestador} | Status: {response.status_code}")
                body = getattr(response, "body", b"")
                if body:
                    try:
                        print(body.decode("utf-8", errors="replace"))
                    except Exception:
                        print(body)

        except HTTPError as e:
            print(f"❌ Erro HTTP ao enviar e-mail para: {prestador}")
            status = getattr(e, "status_code", None)
            if status:
                print(f"Status: {status}")

            body = getattr(e, "body", None)
            if body:
                if isinstance(body, (bytes, bytearray)):
                    print(body.decode("utf-8", errors="replace"))
                else:
                    print(body)
            else:
                print(str(e))

        except Exception as e:
            print(f"❌ Erro ao enviar e-mail para: {prestador} | Erro: {e}")


if __name__ == "__main__":
    main()