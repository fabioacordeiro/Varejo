from pathlib import Path
import json
import re
from collections import defaultdict

# ============================================================
# GERADOR DE HTML INTERATIVO - FLUXO DE SINISTRO
# Baseado no fluxo de comunicação e na matriz de documentos,
# responsáveis e prazos.
# Saída: fluxo_sinistro_interativo.html
# ============================================================

# ---------------------------
# DADOS BASE
# ---------------------------
# Fluxo macro consolidado a partir do PDF do processo.
FLOW_STEPS = [
    {
        "id": "inicio",
        "lane": "LOJAS / CDS",
        "title": "Início / Receber documentação da carga",
        "detail": "Recebimento da documentação da carga e início da conferência.",
        "type": "start",
        "next": ["conferencia_lacre"],
    },
    {
        "id": "conferencia_lacre",
        "lane": "LOJAS / CDS",
        "title": "Processo de conferência e quebra de lacre",
        "detail": "Conferência inicial da carga e verificação do lacre.",
        "type": "process",
        "next": ["aferir_temperatura"],
    },
    {
        "id": "aferir_temperatura",
        "lane": "LOJAS / CDS",
        "title": "Aferir temperatura com termômetro espeto",
        "detail": "Etapa crítica para sinistros de variação de temperatura.",
        "type": "process",
        "next": ["temperatura_ok"],
    },
    {
        "id": "temperatura_ok",
        "lane": "LOJAS / CDS",
        "title": "Temperatura dentro do limite aceitável?",
        "detail": "Se sim, descarga normal. Se não, segue apuração do sinistro.",
        "type": "decision",
        "next": ["descarga", "verificar_divisoria"],
    },
    {
        "id": "descarga",
        "lane": "LOJAS / CDS",
        "title": "Efetuar a descarga",
        "detail": "Fluxo normal quando a temperatura está aceitável.",
        "type": "success",
        "next": ["fim"],
    },
    {
        "id": "verificar_divisoria",
        "lane": "LOJAS / CDS",
        "title": "Tem divisória?",
        "detail": "Se houver divisória, solicitar retirada pela transportadora.",
        "type": "decision",
        "next": ["retirada_divisoria", "liberar_motorista"],
    },
    {
        "id": "retirada_divisoria",
        "lane": "TRANSPORTADORA",
        "title": "Processo de desmontagem da divisória",
        "detail": "A transportadora realiza a retirada/desmontagem da divisória.",
        "type": "process",
        "next": ["abrir_sinistro"],
    },
    {
        "id": "liberar_motorista",
        "lane": "LOJAS / CDS",
        "title": "Processo de conferência e liberação do motorista",
        "detail": "Fluxo alternativo sem divisória.",
        "type": "process",
        "next": ["fim"],
    },
    {
        "id": "abrir_sinistro",
        "lane": "TRANSPORTE",
        "title": "Abertura de sinistro",
        "detail": "Início formal do processo de comunicação e tratativas.",
        "type": "alert",
        "next": ["coletar_evidencias"],
    },
    {
        "id": "coletar_evidencias",
        "lane": "TRANSPORTE",
        "title": "Coletar evidências e documentos",
        "detail": "Solicitar relatório de temperatura BRK, evidências da reguladora, documentos da transportadora e do CD expedidor.",
        "type": "process",
        "next": ["analisar_responsavel"],
    },
    {
        "id": "analisar_responsavel",
        "lane": "TRANSPORTE",
        "title": "Analisar / apurar responsável",
        "detail": "Classificação do responsável: transportadora, CD, outra loja ou indefinido.",
        "type": "decision",
        "next": ["pericia", "classificar_responsavel"],
    },
    {
        "id": "pericia",
        "lane": "REGULADORA",
        "title": "Informar número do sinistro e chegada do perito",
        "detail": "A reguladora informa o número do sinistro, previsão do perito e solicita/recebe evidências.",
        "type": "process",
        "next": ["classificar_responsavel"],
    },
    {
        "id": "classificar_responsavel",
        "lane": "TRANSPORTE",
        "title": "Classificação do responsável",
        "detail": "Responsável transportadora, CD, outra loja ou indefinido.",
        "type": "decision",
        "next": ["tratativa_transportadora", "tratativa_cd", "tratativa_outra_loja", "tratativa_indefinido"],
    },
    {
        "id": "tratativa_transportadora",
        "lane": "TRANSPORTE",
        "title": "Responsável: Transportadora",
        "detail": "Encerrar sinistro com tratativa financeira e informar desconto do transportador.",
        "type": "alert",
        "next": ["seguradora_docs"],
    },
    {
        "id": "tratativa_cd",
        "lane": "TRANSPORTE",
        "title": "Responsável: CD",
        "detail": "Solicitar desconto ao financeiro e reclassificação em loja.",
        "type": "process",
        "next": ["seguradora_docs"],
    },
    {
        "id": "tratativa_outra_loja",
        "lane": "TRANSPORTE",
        "title": "Responsável: Outra loja",
        "detail": "Informar conta e seguir tratativa interna.",
        "type": "process",
        "next": ["seguradora_docs"],
    },
    {
        "id": "tratativa_indefinido",
        "lane": "TRANSPORTE",
        "title": "Responsável: Indefinido / Quebra da CIA",
        "detail": "Fluxo de tratativa quando não é possível definir responsabilidade imediata.",
        "type": "warning",
        "next": ["seguradora_docs"],
    },
    {
        "id": "seguradora_docs",
        "lane": "SEGURADORA",
        "title": "Aguardar documentos do processo",
        "detail": "Validação documental e decisão sobre ressarcimento.",
        "type": "process",
        "next": ["docs_recebidos"],
    },
    {
        "id": "docs_recebidos",
        "lane": "SEGURADORA",
        "title": "Documentos recebidos?",
        "detail": "Se não, informar motivo. Se sim, avaliar ressarcimento.",
        "type": "decision",
        "next": ["ressarcimento_aprovado"],
    },
    {
        "id": "ressarcimento_aprovado",
        "lane": "SEGURADORA",
        "title": "Ressarcimento aprovado?",
        "detail": "Se não, informar motivo. Se sim, encerrar o sinistro.",
        "type": "decision",
        "next": ["encerramento", "informar_motivo"],
    },
    {
        "id": "informar_motivo",
        "lane": "SEGURADORA",
        "title": "Informar motivo",
        "detail": "Motivo de recusa ou quebra da companhia.",
        "type": "warning",
        "next": ["fim"],
    },
    {
        "id": "encerramento",
        "lane": "TRANSPORTE",
        "title": "Encerrar o sinistro",
        "detail": "Finalização do processo após aprovação do ressarcimento ou conclusão da tratativa.",
        "type": "success",
        "next": ["fim"],
    },
    {
        "id": "fim",
        "lane": "LOJAS / CDS",
        "title": "Fim",
        "detail": "Encerramento do fluxo.",
        "type": "end",
        "next": [],
    },
]

# Matriz documental consolidada.
# Ajuste livremente para incluir/retirar documentos futuramente.
DOCUMENT_ROWS = [
    # VARIAÇÃO DE TEMPERATURA
    ("VARIAÇÃO DE TEMPERATURA", "MDF-e", "TRANSPORTE", "1 DIA"),
    ("VARIAÇÃO DE TEMPERATURA", "CT-e", "TRANSPORTE", "1 DIA"),
    ("VARIAÇÃO DE TEMPERATURA", "Romaneio", "TRANSPORTE", "1 DIA"),
    ("VARIAÇÃO DE TEMPERATURA", "Boletim de Viagem", "TRANSPORTE", "1 DIA"),
    ("VARIAÇÃO DE TEMPERATURA", "Gráfico de análise de temperatura", "TRANSPORTE", "4 DIAS"),
    ("VARIAÇÃO DE TEMPERATURA", "PPT - Análise do Sinistro", "TRANSPORTE", "10 DIAS"),
    ("VARIAÇÃO DE TEMPERATURA", "Atualização Google Sheet", "TRANSPORTE", "DIÁRIO"),
    ("VARIAÇÃO DE TEMPERATURA", "KPI's", "TRANSPORTE", "DIÁRIO"),
    ("VARIAÇÃO DE TEMPERATURA", "Autorização de Embarque / Solicitação de Monitoramento", "BRK", "3 DIAS"),
    ("VARIAÇÃO DE TEMPERATURA", "Histórico de Posições", "BRK", "3 DIAS"),
    ("VARIAÇÃO DE TEMPERATURA", "Histórico de Temperatura", "BRK", "3 DIAS"),
    ("VARIAÇÃO DE TEMPERATURA", "Relatório de mensagens recebidas e enviadas", "BRK", "3 DIAS"),
    ("VARIAÇÃO DE TEMPERATURA", "Relatório de comandos enviados", "BRK", "3 DIAS"),
    ("VARIAÇÃO DE TEMPERATURA", "Relatório de Alertas recebidos", "BRK", "3 DIAS"),
    ("VARIAÇÃO DE TEMPERATURA", "Histórico de alertas e comandos registrando o último teste ocorrido", "BRK", "3 DIAS"),
    ("VARIAÇÃO DE TEMPERATURA", "Liberação do Motorista, Veículo e Ajudante", "BRK", "3 DIAS"),
    ("VARIAÇÃO DE TEMPERATURA", "Cópia da Liberação do Motorista e Veículo", "BRK", "3 DIAS"),
    ("VARIAÇÃO DE TEMPERATURA", "CNH", "TRANSPORTADORA", "3 DIAS"),
    ("VARIAÇÃO DE TEMPERATURA", "CRLV", "TRANSPORTADORA", "3 DIAS"),
    ("VARIAÇÃO DE TEMPERATURA", "ANTT", "TRANSPORTADORA", "3 DIAS"),
    ("VARIAÇÃO DE TEMPERATURA", "Ficha Cadastral", "TRANSPORTADORA", "3 DIAS"),
    ("VARIAÇÃO DE TEMPERATURA", "Discos de Tacógrafo", "TRANSPORTADORA", "3 DIAS"),
    ("VARIAÇÃO DE TEMPERATURA", "Laudo de Expedição", "CD EXPEDIDOR", "3 DIAS"),
    ("VARIAÇÃO DE TEMPERATURA", "Planilha de Expedição / Capa de Embarque", "CD EXPEDIDOR", "3 DIAS"),
    ("VARIAÇÃO DE TEMPERATURA", "Lista detalhada do prejuízo", "LOJA/CD", "3 DIAS"),
    ("VARIAÇÃO DE TEMPERATURA", "Laudo do recebimento da Qualidade", "LOJA/CD", "3 DIAS"),
    ("VARIAÇÃO DE TEMPERATURA", "Informar data e hora do descarte", "LOJA/CD", "3 DIAS"),
    ("VARIAÇÃO DE TEMPERATURA", "Laudo da qualidade do descarte com evidências", "LOJA/CD", "3 DIAS"),
    ("VARIAÇÃO DE TEMPERATURA", "Nota fiscal do descarte", "MTR", "30 DIAS"),

    # ROUBO
    ("ROUBO", "MDF-e", "TRANSPORTE", "1 DIA"),
    ("ROUBO", "CT-e", "TRANSPORTE", "1 DIA"),
    ("ROUBO", "Romaneio", "TRANSPORTE", "1 DIA"),
    ("ROUBO", "Boletim de Viagem", "TRANSPORTE", "1 DIA"),
    ("ROUBO", "PPT - Análise do Sinistro", "TRANSPORTE", "10 DIAS"),
    ("ROUBO", "Atualização Google Sheet", "TRANSPORTE", "DIÁRIO"),
    ("ROUBO", "KPI's", "TRANSPORTE", "DIÁRIO"),
    ("ROUBO", "NF's", "FINANCEIRO", "10 DIAS"),
    ("ROUBO", "Autorização de Embarque / Solicitação de Monitoramento", "BRK", "3 DIAS"),
    ("ROUBO", "Histórico de Posições", "BRK", "3 DIAS"),
    ("ROUBO", "Histórico de Temperatura", "BRK", "3 DIAS"),
    ("ROUBO", "Relatório de mensagens recebidas e enviadas", "BRK", "3 DIAS"),
    ("ROUBO", "Relatório de comandos enviados", "BRK", "3 DIAS"),
    ("ROUBO", "Relatório de Alertas recebidos", "BRK", "3 DIAS"),
    ("ROUBO", "Histórico de alertas e comandos registrando o último teste ocorrido", "BRK", "3 DIAS"),
    ("ROUBO", "Liberação do Motorista, Veículo e Ajudante", "BRK", "3 DIAS"),
    ("ROUBO", "Histórico de Posições da Isca", "BRK", "3 DIAS"),
    ("ROUBO", "Relatório da Gerenciadora de Riscos sobre acionamento de segurança", "BRK", "3 DIAS"),
    ("ROUBO", "Cópia da Liberação do Motorista e Veículo", "BRK", "3 DIAS"),
    ("ROUBO", "CNH", "TRANSPORTADORA", "3 DIAS"),
    ("ROUBO", "CRLV", "TRANSPORTADORA", "3 DIAS"),
    ("ROUBO", "ANTT", "TRANSPORTADORA", "3 DIAS"),
    ("ROUBO", "Ficha Cadastral", "TRANSPORTADORA", "3 DIAS"),
    ("ROUBO", "Discos de Tacógrafo", "TRANSPORTADORA", "3 DIAS"),
    ("ROUBO", "Laudo de Expedição", "CD EXPEDIDOR", "3 DIAS"),
    ("ROUBO", "Planilha de Expedição / Capa de Embarque", "CD EXPEDIDOR", "3 DIAS"),

    # ROUBO PARCIAL
    ("ROUBO PARCIAL", "MDF-e", "TRANSPORTE", "1 DIA"),
    ("ROUBO PARCIAL", "CT-e", "TRANSPORTE", "1 DIA"),
    ("ROUBO PARCIAL", "Romaneio", "TRANSPORTE", "1 DIA"),
    ("ROUBO PARCIAL", "Boletim de Viagem", "TRANSPORTE", "1 DIA"),
    ("ROUBO PARCIAL", "PPT - Análise do Sinistro", "TRANSPORTE", "10 DIAS"),
    ("ROUBO PARCIAL", "Atualização Google Sheet", "TRANSPORTE", "DIÁRIO"),
    ("ROUBO PARCIAL", "KPI's", "TRANSPORTE", "DIÁRIO"),
    ("ROUBO PARCIAL", "NF's", "FINANCEIRO", "10 DIAS"),
    ("ROUBO PARCIAL", "Autorização de Embarque / Solicitação de Monitoramento", "BRK", "3 DIAS"),
    ("ROUBO PARCIAL", "Histórico de Posições", "BRK", "3 DIAS"),
    ("ROUBO PARCIAL", "Histórico de Temperatura", "BRK", "3 DIAS"),
    ("ROUBO PARCIAL", "Relatório de mensagens recebidas e enviadas", "BRK", "3 DIAS"),
    ("ROUBO PARCIAL", "Relatório de comandos enviados", "BRK", "3 DIAS"),
    ("ROUBO PARCIAL", "Relatório de Alertas recebidos", "BRK", "3 DIAS"),
    ("ROUBO PARCIAL", "Histórico de alertas e comandos registrando o último teste ocorrido", "BRK", "3 DIAS"),
    ("ROUBO PARCIAL", "Liberação do Motorista, Veículo e Ajudante", "BRK", "3 DIAS"),
    ("ROUBO PARCIAL", "Histórico de Posições da Isca", "BRK", "3 DIAS"),
    ("ROUBO PARCIAL", "Relatório da Gerenciadora de Riscos sobre acionamento de segurança", "BRK", "3 DIAS"),
    ("ROUBO PARCIAL", "Cópia da Liberação do Motorista e Veículo", "BRK", "3 DIAS"),
    ("ROUBO PARCIAL", "CNH", "TRANSPORTADORA", "3 DIAS"),
    ("ROUBO PARCIAL", "CRLV", "TRANSPORTADORA", "3 DIAS"),
    ("ROUBO PARCIAL", "ANTT", "TRANSPORTADORA", "3 DIAS"),
    ("ROUBO PARCIAL", "Ficha Cadastral", "TRANSPORTADORA", "3 DIAS"),
    ("ROUBO PARCIAL", "Discos de Tacógrafo", "TRANSPORTADORA", "3 DIAS"),
    ("ROUBO PARCIAL", "Laudo de Expedição", "CD EXPEDIDOR", "3 DIAS"),
    ("ROUBO PARCIAL", "Planilha de Expedição / Capa de Embarque", "CD EXPEDIDOR", "3 DIAS"),
    ("ROUBO PARCIAL", "Lista detalhada do prejuízo", "LOJA/CD", "3 DIAS"),
    ("ROUBO PARCIAL", "Laudo do recebimento da Qualidade", "LOJA/CD", "3 DIAS"),
    ("ROUBO PARCIAL", "Informar data e hora do descarte", "LOJA/CD", "3 DIAS"),
    ("ROUBO PARCIAL", "Laudo da qualidade do descarte com evidências", "LOJA/CD", "3 DIAS"),
    ("ROUBO PARCIAL", "Nota fiscal do descarte", "MTR", "30 DIAS"),

    # FURTO
    ("FURTO", "MDF-e", "TRANSPORTE", "1 DIA"),
    ("FURTO", "CT-e", "TRANSPORTE", "1 DIA"),
    ("FURTO", "Romaneio", "TRANSPORTE", "1 DIA"),
    ("FURTO", "Boletim de Viagem", "TRANSPORTE", "1 DIA"),
    ("FURTO", "PPT - Análise do Sinistro", "TRANSPORTE", "10 DIAS"),
    ("FURTO", "Atualização Google Sheet", "TRANSPORTE", "DIÁRIO"),
    ("FURTO", "KPI's", "TRANSPORTE", "DIÁRIO"),
    ("FURTO", "NF's", "FINANCEIRO", "10 DIAS"),
    ("FURTO", "Autorização de Embarque / Solicitação de Monitoramento", "BRK", "3 DIAS"),
    ("FURTO", "Histórico de Posições", "BRK", "3 DIAS"),
    ("FURTO", "Histórico de Temperatura", "BRK", "3 DIAS"),
    ("FURTO", "Relatório de mensagens recebidas e enviadas", "BRK", "3 DIAS"),
    ("FURTO", "Relatório de comandos enviados", "BRK", "3 DIAS"),
    ("FURTO", "Relatório de Alertas recebidos", "BRK", "3 DIAS"),
    ("FURTO", "Histórico de alertas e comandos registrando o último teste ocorrido", "BRK", "3 DIAS"),
    ("FURTO", "Liberação do Motorista, Veículo e Ajudante", "BRK", "3 DIAS"),
    ("FURTO", "Histórico de Posições da Isca", "BRK", "3 DIAS"),
    ("FURTO", "Relatório da Gerenciadora de Riscos sobre acionamento de segurança", "BRK", "3 DIAS"),
    ("FURTO", "Cópia da Liberação do Motorista e Veículo", "BRK", "3 DIAS"),
    ("FURTO", "CNH", "TRANSPORTADORA", "3 DIAS"),
    ("FURTO", "CRLV", "TRANSPORTADORA", "3 DIAS"),
    ("FURTO", "ANTT", "TRANSPORTADORA", "3 DIAS"),
    ("FURTO", "Ficha Cadastral", "TRANSPORTADORA", "3 DIAS"),
    ("FURTO", "Discos de Tacógrafo", "TRANSPORTADORA", "3 DIAS"),
    ("FURTO", "Laudo de Expedição", "CD EXPEDIDOR", "3 DIAS"),
    ("FURTO", "Planilha de Expedição / Capa de Embarque", "CD EXPEDIDOR", "3 DIAS"),
    ("FURTO", "Lista detalhada do prejuízo", "LOJA/CD", "3 DIAS"),
    ("FURTO", "Laudo do recebimento da Qualidade", "LOJA/CD", "3 DIAS"),
    ("FURTO", "Informar data e hora do descarte", "LOJA/CD", "3 DIAS"),
    ("FURTO", "Laudo da qualidade do descarte com evidências", "LOJA/CD", "3 DIAS"),
    ("FURTO", "Nota fiscal do descarte", "MTR", "30 DIAS"),

    # AVARIAS
    ("AVARIAS", "MDF-e", "TRANSPORTE", "1 DIA"),
    ("AVARIAS", "CT-e", "TRANSPORTE", "1 DIA"),
    ("AVARIAS", "Romaneio", "TRANSPORTE", "1 DIA"),
    ("AVARIAS", "Boletim de Viagem", "TRANSPORTE", "1 DIA"),
    ("AVARIAS", "PPT - Análise do Sinistro", "TRANSPORTE", "10 DIAS"),
    ("AVARIAS", "Atualização Google Sheet", "TRANSPORTE", "DIÁRIO"),
    ("AVARIAS", "KPI's", "TRANSPORTE", "DIÁRIO"),
    ("AVARIAS", "Autorização de Embarque / Solicitação de Monitoramento", "BRK", "3 DIAS"),
    ("AVARIAS", "Histórico de Posições", "BRK", "3 DIAS"),
    ("AVARIAS", "Histórico de Temperatura", "BRK", "3 DIAS"),
    ("AVARIAS", "Relatório de mensagens recebidas e enviadas", "BRK", "3 DIAS"),
    ("AVARIAS", "Relatório de comandos enviados", "BRK", "3 DIAS"),
    ("AVARIAS", "Relatório de Alertas recebidos", "BRK", "3 DIAS"),
    ("AVARIAS", "Histórico de alertas e comandos registrando o último teste ocorrido", "BRK", "3 DIAS"),
    ("AVARIAS", "Liberação do Motorista, Veículo e Ajudante", "BRK", "3 DIAS"),
    ("AVARIAS", "Histórico de Posições da Isca", "BRK", "3 DIAS"),
    ("AVARIAS", "Relatório da Gerenciadora de Riscos sobre acionamento de segurança", "BRK", "3 DIAS"),
    ("AVARIAS", "Cópia da Liberação do Motorista e Veículo", "BRK", "3 DIAS"),
    ("AVARIAS", "CNH", "TRANSPORTADORA", "3 DIAS"),
    ("AVARIAS", "CRLV", "TRANSPORTADORA", "3 DIAS"),
    ("AVARIAS", "ANTT", "TRANSPORTADORA", "3 DIAS"),
    ("AVARIAS", "Ficha Cadastral", "TRANSPORTADORA", "3 DIAS"),
    ("AVARIAS", "Discos de Tacógrafo", "TRANSPORTADORA", "3 DIAS"),
    ("AVARIAS", "Laudo de Expedição", "CD EXPEDIDOR", "3 DIAS"),
    ("AVARIAS", "Planilha de Expedição / Capa de Embarque", "CD EXPEDIDOR", "3 DIAS"),
    ("AVARIAS", "Lista detalhada do prejuízo", "LOJA/CD", "3 DIAS"),
    ("AVARIAS", "Laudo do recebimento da Qualidade", "LOJA/CD", "3 DIAS"),
    ("AVARIAS", "Informar data e hora do descarte", "LOJA/CD", "3 DIAS"),
    ("AVARIAS", "Laudo da qualidade do descarte com evidências", "LOJA/CD", "3 DIAS"),
    ("AVARIAS", "Nota fiscal do descarte", "MTR", "30 DIAS"),
]


def prazo_ordem(prazo: str) -> int:
    prazo = prazo.upper().strip()
    if "DIÁRIO" in prazo or "DIARIO" in prazo:
        return 0
    match = re.search(r"(\d+)", prazo)
    if match:
        return int(match.group(1))
    return 999


def resumo_por_tipo(rows):
    agrupado = defaultdict(list)
    for tipo, documento, responsavel, prazo in rows:
        agrupado[tipo].append({
            "tipo": tipo,
            "documento": documento,
            "responsavel": responsavel,
            "prazo": prazo,
            "prazo_ordem": prazo_ordem(prazo),
        })

    resumo = {}
    for tipo, itens in agrupado.items():
        prazos = sorted({item["prazo"] for item in itens}, key=prazo_ordem)
        responsaveis = sorted({item["responsavel"] for item in itens})
        resumo[tipo] = {
            "total_documentos": len(itens),
            "prazos": prazos,
            "responsaveis": responsaveis,
            "criticos_1_dia": sum(1 for item in itens if prazo_ordem(item["prazo"]) == 1),
            "criticos_3_dias": sum(1 for item in itens if prazo_ordem(item["prazo"]) == 3),
            "longos_10_30": sum(1 for item in itens if prazo_ordem(item["prazo"]) >= 10),
        }
    return resumo


def estatisticas(rows):
    tipos = sorted({r[0] for r in rows})
    responsaveis = sorted({r[2] for r in rows})
    prazos = sorted({r[3] for r in rows}, key=prazo_ordem)
    return {
        "total_registros": len(rows),
        "tipos": tipos,
        "responsaveis": responsaveis,
        "prazos": prazos,
    }


SUMMARY = resumo_por_tipo(DOCUMENT_ROWS)
STATS = estatisticas(DOCUMENT_ROWS)

HTML_TEMPLATE = r'''<!DOCTYPE html>
<html lang="pt-br">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Fluxo de Sinistro Interativo</title>
  <style>
    :root {
      --bg: #f4f7fb;
      --card: #ffffff;
      --text: #1f2a37;
      --muted: #6b7280;
      --line: #d9e3f0;
      --primary: #0f4c81;
      --secondary: #3b82f6;
      --danger: #dc2626;
      --warning: #f59e0b;
      --success: #16a34a;
      --shadow: 0 12px 30px rgba(15, 76, 129, 0.10);
      --radius: 18px;
    }

    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: Arial, Helvetica, sans-serif;
      background: linear-gradient(180deg, #eef4fb 0%, #f7f9fc 100%);
      color: var(--text);
    }

    .container {
      width: min(1400px, calc(100% - 32px));
      margin: 20px auto 40px;
    }

    .hero {
      background: linear-gradient(135deg, #0f4c81 0%, #1f6db1 100%);
      color: #fff;
      border-radius: 24px;
      padding: 28px;
      box-shadow: var(--shadow);
      margin-bottom: 20px;
    }

    .hero h1 {
      margin: 0 0 8px;
      font-size: 32px;
    }

    .hero p {
      margin: 0;
      opacity: 0.95;
      line-height: 1.5;
      max-width: 980px;
    }

    .grid {
      display: grid;
      grid-template-columns: repeat(12, 1fr);
      gap: 16px;
      margin-top: 18px;
    }

    .panel {
      grid-column: span 12;
      background: var(--card);
      border: 1px solid #e6edf5;
      border-radius: var(--radius);
      box-shadow: var(--shadow);
      padding: 18px;
    }

    .panel h2 {
      margin: 0 0 14px;
      font-size: 22px;
      color: var(--primary);
    }

    .stats {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
      gap: 12px;
    }

    .stat {
      border: 1px solid #e5edf6;
      border-radius: 16px;
      padding: 14px;
      background: #fbfdff;
    }

    .stat .label {
      font-size: 12px;
      color: var(--muted);
      text-transform: uppercase;
      letter-spacing: .08em;
      margin-bottom: 8px;
    }

    .stat .value {
      font-size: 28px;
      font-weight: 700;
      color: var(--primary);
    }

    .filters {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 12px;
      margin-bottom: 14px;
    }

    select, input {
      width: 100%;
      border-radius: 12px;
      border: 1px solid #cfdceb;
      background: #fff;
      padding: 12px 14px;
      font-size: 14px;
      outline: none;
    }

    .summary-cards {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
      gap: 12px;
      margin-top: 12px;
    }

    .summary-card {
      border-radius: 18px;
      padding: 16px;
      color: #fff;
      box-shadow: var(--shadow);
      position: relative;
      overflow: hidden;
    }

    .summary-card h3 {
      margin: 0 0 10px;
      font-size: 18px;
    }

    .summary-card .big {
      font-size: 36px;
      font-weight: 700;
      line-height: 1;
      margin-bottom: 10px;
    }

    .summary-card .tags {
      display: flex;
      flex-wrap: wrap;
      gap: 6px;
    }

    .tag {
      display: inline-flex;
      align-items: center;
      border-radius: 999px;
      padding: 6px 10px;
      background: rgba(255,255,255,.18);
      font-size: 12px;
      backdrop-filter: blur(4px);
    }

    .tipo-0 { background: linear-gradient(135deg, #dc2626, #ef4444); }
    .tipo-1 { background: linear-gradient(135deg, #f59e0b, #f97316); }
    .tipo-2 { background: linear-gradient(135deg, #8b5cf6, #6366f1); }
    .tipo-3 { background: linear-gradient(135deg, #0ea5e9, #2563eb); }
    .tipo-4 { background: linear-gradient(135deg, #16a34a, #22c55e); }

    .flow-lanes {
      display: grid;
      grid-template-columns: repeat(5, minmax(220px, 1fr));
      gap: 14px;
      overflow-x: auto;
      padding-bottom: 4px;
    }

    .lane {
      min-width: 220px;
      background: #f9fbfe;
      border: 1px solid #dde8f5;
      border-radius: 18px;
      padding: 12px;
    }

    .lane h3 {
      margin: 0 0 12px;
      font-size: 15px;
      color: var(--primary);
      text-align: center;
      padding-bottom: 10px;
      border-bottom: 1px dashed #cfe0f2;
    }

    .step {
      border-radius: 16px;
      padding: 12px;
      margin-bottom: 12px;
      border: 1px solid #dbe7f3;
      background: white;
      cursor: pointer;
      transition: transform .15s ease, box-shadow .15s ease;
    }

    .step:hover {
      transform: translateY(-2px);
      box-shadow: 0 10px 18px rgba(15, 76, 129, .10);
    }

    .step.active {
      outline: 3px solid rgba(59,130,246,.22);
    }

    .step-title {
      font-weight: 700;
      margin-bottom: 6px;
      font-size: 14px;
    }

    .step-detail {
      font-size: 13px;
      color: #4b5563;
      line-height: 1.45;
    }

    .step-type {
      display: inline-block;
      margin-top: 8px;
      font-size: 11px;
      font-weight: 700;
      letter-spacing: .04em;
      border-radius: 999px;
      padding: 5px 8px;
      color: #fff;
    }

    .start, .end, .success { background: rgba(22,163,74,.12); }
    .start .step-type, .end .step-type, .success .step-type { background: var(--success); }
    .decision { background: rgba(245,158,11,.10); }
    .decision .step-type { background: var(--warning); }
    .alert { background: rgba(220,38,38,.08); }
    .alert .step-type { background: var(--danger); }
    .warning { background: rgba(139,92,246,.10); }
    .warning .step-type { background: #7c3aed; }
    .process .step-type { background: var(--secondary); }

    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 14px;
    }

    th, td {
      text-align: left;
      padding: 12px 10px;
      border-bottom: 1px solid #e8eef5;
      vertical-align: top;
    }

    th {
      color: var(--primary);
      background: #f7fbff;
      position: sticky;
      top: 0;
      z-index: 1;
    }

    .table-wrap {
      max-height: 540px;
      overflow: auto;
      border: 1px solid #e6edf5;
      border-radius: 16px;
    }

    .pill {
      display: inline-flex;
      align-items: center;
      padding: 5px 10px;
      border-radius: 999px;
      font-size: 12px;
      font-weight: 700;
    }

    .pill.p1 { background: rgba(220,38,38,.12); color: #b91c1c; }
    .pill.p3 { background: rgba(245,158,11,.14); color: #b45309; }
    .pill.p10 { background: rgba(59,130,246,.12); color: #1d4ed8; }
    .pill.p30 { background: rgba(22,163,74,.12); color: #15803d; }
    .pill.pd { background: rgba(107,114,128,.15); color: #374151; }

    .legend {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      margin-top: 12px;
    }

    .legend-item {
      display: flex;
      align-items: center;
      gap: 8px;
      font-size: 13px;
      color: #4b5563;
    }

    .dot {
      width: 14px;
      height: 14px;
      border-radius: 50%;
    }

    .muted { color: var(--muted); }
    .empty {
      padding: 24px;
      text-align: center;
      color: var(--muted);
    }

    .footer-note {
      font-size: 12px;
      color: var(--muted);
      margin-top: 12px;
    }

    @media (max-width: 900px) {
      .hero h1 { font-size: 26px; }
      .container { width: min(100% - 18px, 1400px); }
    }
  </style>
</head>
<body>
  <div class="container">
    <section class="hero">
      <h1>Fluxo de Sinistro Interativo</h1>
      <p>
        Painel interativo para visualização do fluxo macro de sinistro, tipos de ocorrência,
        responsáveis e prazos documentais. Use os filtros para consultar os documentos por
        tipo de sinistro, responsável e prazo.
      </p>
    </section>

    <section class="panel">
      <h2>Indicadores gerais</h2>
      <div class="stats">
        <div class="stat">
          <div class="label">Total de registros documentais</div>
          <div class="value" id="total-registros"></div>
        </div>
        <div class="stat">
          <div class="label">Tipos de sinistro</div>
          <div class="value" id="total-tipos"></div>
        </div>
        <div class="stat">
          <div class="label">Responsáveis</div>
          <div class="value" id="total-responsaveis"></div>
        </div>
        <div class="stat">
          <div class="label">Prazos distintos</div>
          <div class="value" id="total-prazos"></div>
        </div>
      </div>
    </section>

    <section class="panel">
      <h2>Resumo por tipo de sinistro</h2>
      <div class="summary-cards" id="summary-cards"></div>
    </section>

    <section class="panel">
      <h2>Fluxo macro do processo</h2>
      <div class="flow-lanes" id="flow-lanes"></div>
      <div class="legend">
        <div class="legend-item"><span class="dot" style="background:#16a34a"></span> Início / fim / encerrado</div>
        <div class="legend-item"><span class="dot" style="background:#3b82f6"></span> Processo</div>
        <div class="legend-item"><span class="dot" style="background:#f59e0b"></span> Decisão</div>
        <div class="legend-item"><span class="dot" style="background:#dc2626"></span> Etapa crítica</div>
        <div class="legend-item"><span class="dot" style="background:#7c3aed"></span> Tratativa especial</div>
      </div>
    </section>

    <section class="panel">
      <h2>Matriz documental e prazos</h2>
      <div class="filters">
        <select id="filtro-tipo">
          <option value="">Todos os tipos de sinistro</option>
        </select>
        <select id="filtro-responsavel">
          <option value="">Todos os responsáveis</option>
        </select>
        <select id="filtro-prazo">
          <option value="">Todos os prazos</option>
        </select>
        <input id="filtro-texto" type="text" placeholder="Pesquisar documento..." />
      </div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Tipo de Sinistro</th>
              <th>Documento</th>
              <th>Responsável</th>
              <th>Prazo</th>
            </tr>
          </thead>
          <tbody id="tabela-body"></tbody>
        </table>
      </div>
      <div class="footer-note" id="contador-registros"></div>
    </section>
  </div>

  <script>
    const FLOW_STEPS = __FLOW_STEPS__;
    const DOCUMENT_ROWS = __DOCUMENT_ROWS__;
    const SUMMARY = __SUMMARY__;
    const STATS = __STATS__;

    const lanesOrder = ["LOJAS / CDS", "TRANSPORTADORA", "TRANSPORTE", "REGULADORA", "SEGURADORA"];

    function escapeHtml(text) {
      return String(text)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/\"/g, "&quot;")
        .replace(/'/g, "&#039;");
    }

    function classPrazo(prazo) {
      const p = String(prazo).toUpperCase();
      if (p.includes("DIÁRIO") || p.includes("DIARIO")) return "pd";
      if (p.includes("30")) return "p30";
      if (p.includes("10")) return "p10";
      if (p.includes("3")) return "p3";
      if (p.includes("1")) return "p1";
      return "pd";
    }

    function renderStats() {
      document.getElementById("total-registros").textContent = STATS.total_registros;
      document.getElementById("total-tipos").textContent = STATS.tipos.length;
      document.getElementById("total-responsaveis").textContent = STATS.responsaveis.length;
      document.getElementById("total-prazos").textContent = STATS.prazos.length;
    }

    function renderSummary() {
      const container = document.getElementById("summary-cards");
      const tipos = Object.keys(SUMMARY);
      container.innerHTML = tipos.map((tipo, i) => {
        const item = SUMMARY[tipo];
        return `
          <div class="summary-card tipo-${i % 5}">
            <h3>${escapeHtml(tipo)}</h3>
            <div class="big">${item.total_documentos}</div>
            <div class="muted" style="color:rgba(255,255,255,.92); margin-bottom:10px;">documentos / exigências</div>
            <div class="tags">
              <span class="tag">1 dia: ${item.criticos_1_dia}</span>
              <span class="tag">3 dias: ${item.criticos_3_dias}</span>
              <span class="tag">10+ dias: ${item.longos_10_30}</span>
              <span class="tag">Responsáveis: ${item.responsaveis.length}</span>
            </div>
          </div>
        `;
      }).join('');
    }

    function renderFlow() {
      const container = document.getElementById("flow-lanes");
      const grouped = {};
      lanesOrder.forEach(l => grouped[l] = []);
      FLOW_STEPS.forEach(step => {
        if (!grouped[step.lane]) grouped[step.lane] = [];
        grouped[step.lane].push(step);
      });

      container.innerHTML = lanesOrder.map(lane => `
        <div class="lane">
          <h3>${escapeHtml(lane)}</h3>
          ${grouped[lane].map(step => `
            <div class="step ${step.type}" data-step-id="${escapeHtml(step.id)}">
              <div class="step-title">${escapeHtml(step.title)}</div>
              <div class="step-detail">${escapeHtml(step.detail)}</div>
              <span class="step-type">${escapeHtml(step.type.toUpperCase())}</span>
            </div>
          `).join('')}
        </div>
      `).join('');

      document.querySelectorAll('.step').forEach(el => {
        el.addEventListener('click', () => {
          document.querySelectorAll('.step').forEach(s => s.classList.remove('active'));
          el.classList.add('active');
        });
      });
    }

    function popularFiltros() {
      const tipo = document.getElementById('filtro-tipo');
      const responsavel = document.getElementById('filtro-responsavel');
      const prazo = document.getElementById('filtro-prazo');

      STATS.tipos.forEach(v => tipo.insertAdjacentHTML('beforeend', `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`));
      STATS.responsaveis.forEach(v => responsavel.insertAdjacentHTML('beforeend', `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`));
      STATS.prazos.forEach(v => prazo.insertAdjacentHTML('beforeend', `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`));
    }

    function filtrarDocumentos() {
      const tipo = document.getElementById('filtro-tipo').value.trim().toUpperCase();
      const responsavel = document.getElementById('filtro-responsavel').value.trim().toUpperCase();
      const prazo = document.getElementById('filtro-prazo').value.trim().toUpperCase();
      const texto = document.getElementById('filtro-texto').value.trim().toUpperCase();

      const filtrados = DOCUMENT_ROWS.filter(row => {
        const okTipo = !tipo || row.tipo.toUpperCase() === tipo;
        const okResp = !responsavel || row.responsavel.toUpperCase() === responsavel;
        const okPrazo = !prazo || row.prazo.toUpperCase() === prazo;
        const okTexto = !texto || [row.tipo, row.documento, row.responsavel, row.prazo]
          .join(' ')
          .toUpperCase()
          .includes(texto);
        return okTipo && okResp && okPrazo && okTexto;
      });

      renderTabela(filtrados);
    }

    function renderTabela(rows) {
      const tbody = document.getElementById('tabela-body');
      const contador = document.getElementById('contador-registros');

      if (!rows.length) {
        tbody.innerHTML = `<tr><td colspan="4" class="empty">Nenhum registro encontrado para os filtros aplicados.</td></tr>`;
        contador.textContent = '0 registro(s) encontrado(s).';
        return;
      }

      tbody.innerHTML = rows.map(row => `
        <tr>
          <td>${escapeHtml(row.tipo)}</td>
          <td>${escapeHtml(row.documento)}</td>
          <td>${escapeHtml(row.responsavel)}</td>
          <td><span class="pill ${classPrazo(row.prazo)}">${escapeHtml(row.prazo)}</span></td>
        </tr>
      `).join('');

      contador.textContent = `${rows.length} registro(s) encontrado(s).`;
    }

    function init() {
      renderStats();
      renderSummary();
      renderFlow();
      popularFiltros();
      renderTabela(DOCUMENT_ROWS);

      ['filtro-tipo', 'filtro-responsavel', 'filtro-prazo'].forEach(id => {
        document.getElementById(id).addEventListener('change', filtrarDocumentos);
      });
      document.getElementById('filtro-texto').addEventListener('input', filtrarDocumentos);
    }

    init();
  </script>
</body>
</html>
'''


def build_html(output_file="fluxo_sinistro_interativo.html"):
    rows = [
        {
            "tipo": tipo,
            "documento": documento,
            "responsavel": responsavel,
            "prazo": prazo,
            "prazo_ordem": prazo_ordem(prazo),
        }
        for tipo, documento, responsavel, prazo in DOCUMENT_ROWS
    ]

    html = HTML_TEMPLATE
    html = html.replace("__FLOW_STEPS__", json.dumps(FLOW_STEPS, ensure_ascii=False, indent=2))
    html = html.replace("__DOCUMENT_ROWS__", json.dumps(rows, ensure_ascii=False, indent=2))
    html = html.replace("__SUMMARY__", json.dumps(SUMMARY, ensure_ascii=False, indent=2))
    html = html.replace("__STATS__", json.dumps(STATS, ensure_ascii=False, indent=2))

    output_path = Path(output_file)
    output_path.write_text(html, encoding="utf-8")
    return output_path.resolve()


if __name__ == "__main__":
    arquivo = build_html()
    print(f"HTML gerado com sucesso em: {arquivo}")
