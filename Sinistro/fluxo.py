# Script de Fluxo de Sinistro - Carrefour / Ezze Seguros
# Baseado nos documentos: "Documentos.pdf", "Procedimento de Abertura" e "Prazos de Analise"

class FluxoSinistro:
    def __init__(self):
        # Definição dos prazos gerais de liquidação e análise [22, 23]
        self.prazos_liquidacao = {
            "Fast Track (até R$ 100k)": {
                "Analise_Docs_Relatorio": "05 dias após último documento",
                "Pagamento": "10 dias após recibo de quitação"
            },
            "Medium/Large Loss (> R$ 100k)": {
                "Analise_e_Cobertura": "30 dias (Regra SUSEP) - média de 15 dias",
                "Pagamento": "10 dias após recibo assinado"
            }
        }

        # Mapeamento de documentos por área e prazo [1-19]
        self.documentacao_base = {
            "TRANSPORTE": [
                {"doc": "MDF-e", "prazo": "1 dia"},
                {"doc": "CT-e", "prazo": "1 dia"},
                {"doc": "Romaneio", "prazo": "1 dia"},
                {"doc": "Boletim de Viagem", "prazo": "1 dia"},
                {"doc": "PPT - Análise do Sinistro", "prazo": "10 dias"},
                {"doc": "Atualização Google Sheet / KPIs", "prazo": "Diário"}
            ],
            "BRK (Gerenciadora)": [
                {"doc": "Autorização de Embarque / Monitoramento", "prazo": "3 dias"},
                {"doc": "Histórico de Posições e Alertas", "prazo": "3 dias"},
                {"doc": "Relatório de Comandos e Mensagens", "prazo": "3 dias"},
                {"doc": "Liberação de Motorista/Veículo", "prazo": "3 dias"}
            ],
            "TRANSPORTADORA": [
                {"doc": "CNH / CRLV / ANTT", "prazo": "3 dias"},
                {"doc": "Ficha Cadastral", "prazo": "3 dias"},
                {"doc": "Discos de Tacógrafo", "prazo": "3 dias"}
            ],
            "LOJA / CD": [
                {"doc": "Lista detalhada do prejuízo", "prazo": "3 dias"},
                {"doc": "Laudo de recebimento da Qualidade (RT)", "prazo": "3 dias"},
                {"doc": "Evidências de Descarte (Laudo e Fotos)", "prazo": "3 dias"}
            ]
        }

    def obter_procedimento_abertura(self):
        """Retorna o passo a passo para abertura do sinistro [20]"""
        return [
            "1. Coletar dados: Placa, Tipo de Sinistro e Contato do Responsável.",
            "2. Ligar para Reguladora Global (0800-772-1233).",
            "3. Obter o número do sinistro.",
            "4. Informar previsão do técnico para acompanhamento da descarga (obrigatório para Temperatura/Acidente).",
            "5. Conferência de 100% da carga pela loja com presença do perito."
        ]

    def gerar_fluxo_por_tipo(self, tipo_sinistro):
        """Gera o detalhamento documental por tipo de sinistro [1-19, 24, 25]"""
        print(f"\n--- FLUXO DE SINISTRO: {tipo_sinistro.upper()} ---")
        
        # Procedimento inicial
        print("\nPROCEDIMENTO DE ABERTURA:")
        for passo in self.obter_procedimento_abertura():
            print(passo)

        # Documentos Específicos
        print("\nDOCUMENTAÇÃO E ÁREAS RESPONSÁVEIS:")
        for area, docs in self.documentacao_base.items():
            print(f"[{area}]")
            for item in docs:
                print(f" - {item['doc']} (Prazo: {item['prazo']})")
        
        # Adicionais específicos por tipo
        if "TEMPERATURA" in tipo_sinistro.upper():
            print(" - [TRANSPORTE] Gráfico de análise de temperatura (4 dias) [1]")
            print(" - [BRK] Histórico de Temperatura (3 dias) [2]")
        elif "ROUBO" in tipo_sinistro.upper() or "FURTO" in tipo_sinistro.upper():
            print(" - [BRK] Histórico de Posições da Isca (3 dias) [6]")
            print(" - [FINANCEIRO] NFs (10 dias) [5]")
            print(" - [AUTORIDADE] Boletim de Ocorrência [25]")

        print("\nPRAZOS DE LIQUIDAÇÃO ACORDADOS:")
        for categoria, prazos in self.prazos_liquidacao.items():
            print(f"{categoria}: Pagamento em {prazos['Pagamento']}")

# Execução do Especialista
fluxo = FluxoSinistro()
fluxo.gerar_fluxo_por_tipo("Variação de Temperatura")
fluxo.gerar_fluxo_por_tipo("Roubo / Furto")