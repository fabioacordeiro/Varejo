import pdfplumber
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
import re

def extract_temperature_data(pdf_path):
    """
    Extrai dados de temperatura do PDF Thermo King
    """
    data = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            # Tenta extrair tabelas primeiro
            tables = page.extract_tables()
            
            if tables:
                for table in tables:
                    # Procura pelo cabeçalho para identificar colunas
                    header_row = None
                    date_col = None
                    time_col = None
                    discharge_col = None
                    
                    for row_idx, row in enumerate(table):
                        if row and any('Date' in str(cell) or 'Data' in str(cell) for cell in row if cell):
                            header_row = row_idx
                            # Identifica índices das colunas
                            for col_idx, cell in enumerate(row):
                                if cell:
                                    cell_lower = str(cell).lower()
                                    if 'date' in cell_lower or 'data' in cell_lower:
                                        date_col = col_idx
                                    elif 'time' in cell_lower or 'hora' in cell_lower:
                                        time_col = col_idx
                                    elif 'discharge' in cell_lower and 'display' in cell_lower:
                                        discharge_col = col_idx
                            break
                    
                    # Se encontrou cabeçalho, processa as linhas de dados
                    if header_row is not None and discharge_col is not None:
                        for row in table[header_row + 1:]:
                            if row and len(row) > max(date_col or 0, time_col or 0, discharge_col):
                                try:
                                    # Extrai data
                                    date_str = str(row[date_col]).strip() if date_col is not None else None
                                    # Extrai hora
                                    time_str = str(row[time_col]).strip() if time_col is not None else None
                                    # Extrai temperatura
                                    temp_str = str(row[discharge_col]).strip() if discharge_col is not None else None
                                    
                                    if date_str and time_str and temp_str:
                                        # Remove caracteres não numéricos da temperatura (exceto - para negativos)
                                        temp_clean = re.sub(r'[^\d-]', '', temp_str)
                                        if temp_clean:
                                            temp = int(temp_clean)
                                            
                                            # Combina data e hora
                                            datetime_str = f"{date_str} {time_str}"
                                            dt = datetime.strptime(datetime_str, '%d/%m/%Y %H:%M')
                                            
                                            data.append({
                                                'Data/Hora': dt,
                                                'Temperatura (°C)': temp
                                            })
                                except (ValueError, IndexError) as e:
                                    continue
            
            # Método alternativo: busca por texto bruto
            text = page.extract_text()
            if text and not data:  # Só usa se não encontrou dados nas tabelas
                lines = text.split('\n')
                for line in lines:
                    # Procura padrão: DD/MM/YYYY HH:MM ... número
                    match = re.search(r'(\d{2}/\d{2}/\d{4})\s+(\d{2}:\d{2}).*?Discharge Display.*?(-?\d+)', line)
                    if match:
                        date_str = match.group(1)
                        time_str = match.group(2)
                        temp = int(match.group(3))
                        
                        try:
                            datetime_str = f"{date_str} {time_str}"
                            dt = datetime.strptime(datetime_str, '%d/%m/%Y %H:%M')
                            data.append({
                                'Data/Hora': dt,
                                'Temperatura (°C)': temp
                            })
                        except ValueError:
                            continue
    
    return pd.DataFrame(data)

def create_excel_and_chart(df, excel_path='temperatura_thermoking.xlsx'):
    """
    Cria planilha Excel e gráfico com as temperaturas
    """
    if df.empty:
        print("❌ Nenhum dado foi extraído do PDF.")
        return
    
    # Remove duplicatas e ordena
    df = df.drop_duplicates().sort_values('Data/Hora').reset_index(drop=True)
    
    # Salva no Excel
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Temperaturas')
        
        # Ajusta largura das colunas
        worksheet = writer.sheets['Temperaturas']
        worksheet.column_dimensions['A'].width = 20
        worksheet.column_dimensions['B'].width = 18
    
    print(f"✅ Planilha criada: {excel_path}")
    print(f"   Total de registros: {len(df)}")
    
    # Cria o gráfico
    fig, ax = plt.subplots(figsize=(18, 9))
    
    # Plota a linha de temperatura
    ax.plot(df['Data/Hora'], df['Temperatura (°C)'], 
            color='#1f77b4', linewidth=2.5, label='Temperatura', 
            marker='o', markersize=4, markerfacecolor='white', 
            markeredgewidth=1.5, zorder=5)
    
    # Destaca faixa CONGELADO (-15 a -12°C)
    ax.axhspan(-15, -12, alpha=0.25, color='#4A90E2', 
               label='Faixa Congelado (-15 a -12°C)', zorder=1)
    
    # Destaca faixa RESFRIADO (0 a 5°C)
    ax.axhspan(0, 5, alpha=0.25, color='#50C878', 
               label='Faixa Resfriado (0 a 5°C)', zorder=1)
    
    # Períodos específicos
    periodo1_inicio = datetime.strptime('14/11/2025 09:00', '%d/%m/%Y %H:%M')
    periodo1_fim = datetime.strptime('14/11/2025 09:44', '%d/%m/%Y %H:%M')
    periodo2_inicio = datetime.strptime('14/11/2025 17:49', '%d/%m/%Y %H:%M')
    periodo2_fim = datetime.strptime('14/11/2025 18:20', '%d/%m/%Y %H:%M')
    
    # Destaca períodos com faixas verticais
    ax.axvspan(periodo1_inicio, periodo1_fim, alpha=0.2, color='#FF6B6B', 
               label='Período 1 (09:00-09:44)', zorder=2)
    ax.axvspan(periodo2_inicio, periodo2_fim, alpha=0.2, color='#FFA500', 
               label='Período 2 (17:49-18:20)', zorder=2)
    
    # Adiciona linhas de referência
    ax.axhline(y=-12, color='blue', linestyle='--', linewidth=1, alpha=0.5)
    ax.axhline(y=-15, color='blue', linestyle='--', linewidth=1, alpha=0.5)
    ax.axhline(y=0, color='green', linestyle='--', linewidth=1, alpha=0.5)
    ax.axhline(y=5, color='green', linestyle='--', linewidth=1, alpha=0.5)
    
    # Configurações do gráfico
    ax.set_xlabel('Data/Hora', fontsize=13, fontweight='bold')
    ax.set_ylabel('Temperatura (°C)', fontsize=13, fontweight='bold')
    ax.set_title('Monitoramento de Temperatura - Thermo King\nCampo: Discharge Display', 
                 fontsize=15, fontweight='bold', pad=20)
    
    ax.grid(True, alpha=0.3, linestyle='--', linewidth=0.7, which='both')
    ax.legend(loc='best', fontsize=10, framealpha=0.95, shadow=True)
    
    # Formata eixo X
    import matplotlib.dates as mdates
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%d/%m %H:%M'))
    plt.xticks(rotation=45, ha='right')
    
    # Ajusta limites do eixo Y para melhor visualização
    y_min = df['Temperatura (°C)'].min() - 2
    y_max = df['Temperatura (°C)'].max() + 2
    ax.set_ylim(y_min, y_max)
    
    plt.tight_layout()
    
    # Salva o gráfico
    chart_path = 'grafico_temperatura_thermoking.png'
    plt.savefig(chart_path, dpi=300, bbox_inches='tight', facecolor='white')
    print(f"✅ Gráfico salvo: {chart_path}")
    
    plt.show()
    
    # Estatísticas detalhadas
    print("\n" + "="*70)
    print("📊 ESTATÍSTICAS GERAIS")
    print("="*70)
    print(f"Período analisado: {df['Data/Hora'].min().strftime('%d/%m/%Y %H:%M')} até {df['Data/Hora'].max().strftime('%d/%m/%Y %H:%M')}")
    print(f"Temperatura mínima: {df['Temperatura (°C)'].min()}°C")
    print(f"Temperatura máxima: {df['Temperatura (°C)'].max()}°C")
    print(f"Temperatura média: {df['Temperatura (°C)'].mean():.2f}°C")
    print(f"Desvio padrão: {df['Temperatura (°C)'].std():.2f}°C")
    
    # Análise por faixa
    congelado = df[(df['Temperatura (°C)'] >= -15) & (df['Temperatura (°C)'] <= -12)]
    resfriado = df[(df['Temperatura (°C)'] >= 0) & (df['Temperatura (°C)'] <= 5)]
    fora_faixa = df[~df.index.isin(congelado.index) & ~df.index.isin(resfriado.index)]
    
    print(f"\n🧊 CONGELADO (-15 a -12°C): {len(congelado)} registros ({len(congelado)/len(df)*100:.1f}%)")
    if len(congelado) > 0:
        print(f"   Média: {congelado['Temperatura (°C)'].mean():.2f}°C")
    
    print(f"\n❄️  RESFRIADO (0 a 5°C): {len(resfriado)} registros ({len(resfriado)/len(df)*100:.1f}%)")
    if len(resfriado) > 0:
        print(f"   Média: {resfriado['Temperatura (°C)'].mean():.2f}°C")
    
    print(f"\n⚠️  FORA DAS FAIXAS: {len(fora_faixa)} registros ({len(fora_faixa)/len(df)*100:.1f}%)")
    
    # Análise dos períodos destacados
    periodo1 = df[(df['Data/Hora'] >= periodo1_inicio) & (df['Data/Hora'] <= periodo1_fim)]
    periodo2 = df[(df['Data/Hora'] >= periodo2_inicio) & (df['Data/Hora'] <= periodo2_fim)]
    
    print("\n" + "="*70)
    print("🔴 ANÁLISE DOS PERÍODOS DESTACADOS")
    print("="*70)
    
    print(f"\nPeríodo 1 (14/11/2025 09:00-09:44):")
    if len(periodo1) > 0:
        print(f"  Registros: {len(periodo1)}")
        print(f"  Temp. média: {periodo1['Temperatura (°C)'].mean():.2f}°C")
        print(f"  Temp. min/max: {periodo1['Temperatura (°C)'].min()}°C / {periodo1['Temperatura (°C)'].max()}°C")
    else:
        print("  ⚠️ Nenhum registro encontrado neste período")
    
    print(f"\nPeríodo 2 (14/11/2025 17:49-18:20):")
    if len(periodo2) > 0:
        print(f"  Registros: {len(periodo2)}")
        print(f"  Temp. média: {periodo2['Temperatura (°C)'].mean():.2f}°C")
        print(f"  Temp. min/max: {periodo2['Temperatura (°C)'].min()}°C / {periodo2['Temperatura (°C)'].max()}°C")
    else:
        print("  ⚠️ Nenhum registro encontrado neste período")
    
    print("\n" + "="*70)

# Execução principal
if __name__ == "__main__":
    print("\n" + "="*70)
    print("🌡️  EXTRAÇÃO DE DADOS DE TEMPERATURA - THERMO KING")
    print("="*70 + "\n")
    
    # Caminho do PDF
    pdf_path = r"C:\Fabio\Desenvolvimento\Varejo\Temperatura\Temp_eq_Thermoking.pdf"
    
    print(f"📄 Processando: {pdf_path}\n")
    
    # Extrai dados
    df = extract_temperature_data(pdf_path)
    
    if not df.empty:
        print(f"✅ {len(df)} registros extraídos com sucesso!\n")
        
        # Cria Excel e gráfico
        create_excel_and_chart(df)
    else:
        print("❌ Nenhum dado foi encontrado no PDF.")
        print("\n💡 Sugestões:")
        print("   1. Verifique se o campo 'Discharge Display' existe no PDF")
        print("   2. Execute o script de diagnóstico novamente")
        print("   3. Compartilhe uma amostra do texto extraído para ajuste")