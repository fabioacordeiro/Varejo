import plotly.graph_objects as go

def gerar_fluxo_plotly():
    # Definição dos nós do fluxo (Áreas e Etapas)
    # Baseado em: Procedimento de Abertura [5] e Prazos [6]
    label = [
        "1. Ocorrência (Sinistro)",                  # 0
        "2. Abertura (Torre de Controle)",           # 1 - Tel: 0800-772-1233 [5]
        "Área: TRANSPORTE (Prazo: 1-10 dias)",       # 2 - Docs: MDF-e, CT-e, PPT [1]
        "Área: BRK / Gerenciadora (Prazo: 3 dias)",  # 3 - Docs: Monitoramento, Isca [2]
        "Área: TRANSPORTADORA (Prazo: 3 dias)",      # 4 - Docs: CNH, CRLV, Tacógrafo [3]
        "Área: LOJA / CD (Prazo: 3 dias)",           # 5 - Docs: Laudo RT, Fotos Descarte [4]
        "3. Análise / Regulação",                    # 6 - 05 dias (Fast) / 30 dias (Large) [6, 7]
        "4. Pagamento (Financeiro)"                  # 7 - 10 dias após recibo [6]
    ]

    # Definição das conexões (De onde -> Para onde)
    source = [8-13]
    target = [8-14]
    # Grossura das linhas (simbólico)
    value =  [8, 11]

    # Cores personalizadas
    node_colors = ["orange", "yellow", "blue", "green", "grey", "red", "lightgrey", "darkgreen"]

    # Criação do Diagrama de Sankey
    fig = go.Figure(data=[go.Sankey(
        node = dict(
          pad = 15,
          thickness = 20,
          line = dict(color = "black", width = 0.5),
          label = label,
          color = node_colors
        ),
        link = dict(
          source = source,
          target = target,
          value = value,
          hovertemplate = 'Fluxo de Processo Documental<extra></extra>'
        ))])

    fig.update_layout(
        title_text="Fluxo de Sinistro Carrefour / Ezze Seguros<br><sub>Baseado em Procedimentos de Abertura e Tabelas de Documentos</sub>",
        font_size=12
    )

    # Salva como HTML e abre no navegador
    fig.write_html("fluxo_sinistro_plotly.html")
    fig.show()

if __name__ == "__main__":
    gerar_flow_info = """
    DETALHAMENTO DO FLUXO (Citações das Fontes):
    - Abertura: Realizada pela Torre de Controle via 0800-772-1233 [5].
    - Transporte: MDF-e e CT-e (1 dia), PPT de análise (10 dias) [1].
    - BRK: Histórico de posições e alertas (3 dias) [2].
    - Loja/CD: Laudo RT e fotos de descarte (3 dias) [4].
    - Liquidação Fast Track: Análise em 5 dias e Pagamento em 10 dias [6].
    - Liquidação Large Loss: Análise em até 30 dias (SUSEP) [7].
    """
    print(gerar_flow_info)
    gerar_fluxo_plotly()