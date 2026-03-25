import pdfplumber
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
import re

def extract_temperature_data(pdf_path):
    """
    Extrai dados de temperatura do PDF Thermo King
    Campo: Discharge Display (coluna 9)
    """
    all_data = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            tables = page.extract_tables()
            
            for table in tables:
                if not table or len(table) < 2:
                    continue
                
                # Procura pelo cabeçalho com "Discharge Display"
                header_row_idx = None
                for idx, row in enumerate(table[:5]):  # Verifica primeiras 5 linhas
                    if row and any(cell and 'Discharge' in str(cell) and 'Display' in str(cell) for cell in row):
                        header_row_idx = idx
                        break
                
                if header_row_idx is None:
                    continue
                
                # Identifica a coluna "Discharge Display"
                header = table[header_row_idx]
                discharge_display_col = None
                hora_col = None
                
                for col_idx, cell in enumerate(header):
                    if cell:
                        cell_str = str(cell).replace('\n', ' ')
                        if 'Discharge' in cell_str and 'Display' in cell_str:
                            discharge_display_col = col_idx
                        elif 'Hora' in cell_str or 'Time' in cell_str:
                            hora_col = col_idx
                
                if discharge_display_col is None or hora_col is None:
                    continue
                
                # Extrai os dados
                for row in table[header_row_idx + 1:]:
                    if not row or len(row) <= max(discharge_display_col, hora_col):
                        continue
                    
                    try:
                        # Extrai data/hora
                        datetime_str = str(row[hora_col]).strip()
                        
                        # Extrai temperatura (Discharge Display)
                        temp_str = str(row[discharge_display_col]).strip()
                        
                        # Valida e converte
                        if datetime_str and temp_str and datetime_str != 'None':
                            # Remove caracteres não numéricos da temperatura (mantém - para negativos)
                            temp_clean = re.sub(r'[^\d.-]', '', temp_str)
                            
                            if temp_clean and temp_clean not in ['', '-', '.']:
                                temp = float(temp_clean)
                                
                                # Parse da data/hora
                                dt = datetime.strptime(datetime_str, '%m/%d/%Y %H:%M')
                                
                                all_data.append({
                                    'Data/Hora': dt,
                                    'Temperatura (°C)': temp
                                })
                    except (ValueError, IndexError) as e:
                        continue
            
            print(f"✓ Página {page_num}/{len(pdf.pages)} processada - {len(all_data)} registros até agora")
    
    return pd.DataFrame(all_data)

def create_excel_and_chart(df, excel_path='temperatura_thermoking.xlsx'):
    """
    Cria planilha Excel e gráfico com as temperaturas
    """
    if df.empty:
        print("❌ Nenhum dado foi extraído do PDF.")
        return
    
    # Remove duplicatas e ordena
    df = df.drop_duplicates().sort_values('Data/Hora').reset_index(drop=True)
    
    # Separa Data e Hora em colunas diferentes
    df_export = df.copy()
    df_export['Data'] = df_export['Data/Hora'].dt.strftime('%d/%m/%Y')
    df_export['Hora'] = df_export['Data/Hora'].dt.strftime('%H:%M')
    df_export = df_export[['Data', 'Hora', 'Temperatura (°C)']]
    
    # Salva no Excel
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Temperaturas')
        
        # Ajusta largura das colunas
        worksheet = writer.sheets['Temperaturas']
        worksheet.column_dimensions['A'].width = 15
        worksheet.column_dimensions['B'].width = 10
        worksheet.column_dimensions['C'].width = 20
    
    print(f"\n✅ Planilha criada: {excel_path}")
    print(f"   Total de registros: {len(df)}")
    
    # Cria o gráfico
    fig, ax = plt.subplots(figsize=(20, 10))
    
    # Plota a linha de temperatura
    ax.plot(df['Data/Hora'], df['Temperatura (°C)'], 
            color='#1f77b4', linewidth=2, label='Temperatura (Discharge Display)', 
            marker='o', markersize=3, markerfacecolor='white', 
            markeredgewidth=1, zorder=5)
    
    # Destaca faixa CONGELADO (-15 a -12°C)
    ax.axhspan(-15, -12, alpha=0.25, color='#4A90E2', 
               label='Faixa Congelado (-15 a -12°C)', zorder=1)
    
    # Destaca faixa RESFRIADO (0 a 5°C)
    ax.axhspan(0, 5, alpha=0.25, color='#50C878', 
               label='Faixa Resfriado (0 a 5°C)', zorder=1)
    
    # Períodos específicos
    try:
        periodo1_inicio = datetime.strptime('14/11/2025 09:00', '%d/%m/%Y %H:%M')
        periodo1_fim = datetime.strptime('14/11/2025 09:44', '%d/%m/%Y %H:%M')
        periodo2_inicio = datetime.strptime('14/11/2025 17:49', '%d/%m/%Y %H:%M')
        periodo2_fim = datetime.strptime('14/11/2025 18:20', '%d/%m/%Y %H:%M')
        
        # Destaca períodos com faixas verticais
        ax.axvspan(periodo1_inicio, periodo1_fim, alpha=0.2, color='#FF6B6B', 
                   label='Período 1 (14/11 09:00-09:44)', zorder=2)
        ax.axvspan(periodo2_inicio, periodo2_fim, alpha=0.2, color='#FFA500', 
                   label='Período 2 (14/11 17:49-18:20)', zorder=2)
    except:
        print("⚠️ Períodos destacados fora do intervalo de dados")
    
    # Adiciona linhas de referência
    ax.axhline(y=-12, color='blue', linestyle='--', linewidth=1, alpha=0.5)
    ax.axhline(y=-15, color='blue', linestyle='--', linewidth=1, alpha=0.5)
    ax.axhline(y=0, color='green', linestyle='--', linewidth=1, alpha=0.5)
    ax.axhline(y=5, color='green', linestyle='--', linewidth=1, alpha=0.5)
    
    # Configurações do gráfico
    ax.set_xlabel('Data/Hora', fontsize=14, fontweight='bold')
    ax.set_ylabel('Temperatura (°C)', fontsize=14, fontweight='bold')
    ax.set_title('Monitoramento de Temperatura - Thermo King\nCampo: Discharge Display', 
                 fontsize=16, fontweight='bold', pad=20)
    
    ax.grid(True, alpha=0.3, linestyle='--', linewidth=0.7, which='both')
    ax.legend(loc='best', fontsize=11, framealpha=0.95, shadow=True)
    
    # Formata eixo X
    import matplotlib.dates as mdates
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%d/%m\n%H:%M'))
    plt.xticks(rotation=0, ha='center')
    
    # Ajusta limites do eixo Y
    y_min = df['Temperatura (°C)'].min() - 3
    y_max = df['Temperatura (°C)'].max() + 3
    ax.set_ylim(y_min, y_max)
    
    plt.tight_layout()
    
    # Salva o gráfico
    chart_path = 'grafico_temperatura_thermoking.png'
    plt.savefig(chart_path, dpi=300, bbox_inches='tight', facecolor='white')
    print(f"✅ Gráfico salvo: {chart_path}")
    
    plt.show()
    
    # Estatísticas detalhadas
    print("\n" + "="*80)
    print("📊 ESTATÍSTICAS GERAIS")
    print("="*80)
    print(f"Período: {df['Data/Hora'].min().strftime('%d/%m/%Y %H:%M')} até {df['Data/Hora'].max().strftime('%d/%m/%Y %H:%M')}")
    print(f"Temperatura mínima: {df['Temperatura (°C)'].min():.1f}°C")
    print(f"Temperatura máxima: {df['Temperatura (°C)'].max():.1f}°C")
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
    if len(fora_faixa) > 0:
        print(f"   Média: {fora_faixa['Temperatura (°C)'].mean():.2f}°C")
    
    # Análise dos períodos destacados
    try:
        periodo1_inicio = datetime.strptime('14/11/2025 09:00', '%d/%m/%Y %H:%M')
        periodo1_fim = datetime.strptime('14/11/2025 09:44', '%d/%m/%Y %H:%M')
        periodo2_inicio = datetime.strptime('14/11/2025 17:49', '%d/%m/%Y %H:%M')
        periodo2_fim = datetime.strptime('14/11/2025 18:20', '%d/%m/%Y %H:%M')
        
        periodo1 = df[(df['Data/Hora'] >= periodo1_inicio) & (df['Data/Hora'] <= periodo1_fim)]
        periodo2 = df[(df['Data/Hora'] >= periodo2_inicio) & (df['Data/Hora'] <= periodo2_fim)]
        
        print("\n" + "="*80)
        print("🔴 ANÁLISE DOS PERÍODOS DESTACADOS")
        print("="*80)
        
        print(f"\nPeríodo 1 (14/11/2025 09:00-09:44):")
        if len(periodo1) > 0:
            print(f"  Registros: {len(periodo1)}")
            print(f"  Temp. média: {periodo1['Temperatura (°C)'].mean():.2f}°C")
            print(f"  Temp. min/max: {periodo1['Temperatura (°C)'].min():.1f}°C / {periodo1['Temperatura (°C)'].max():.1f}°C")
        else:
            print("  ⚠️ Nenhum registro encontrado neste período")
        
        print(f"\nPeríodo 2 (14/11/2025 17:49-18:20):")
        if len(periodo2) > 0:
            print(f"  Registros: {len(periodo2)}")
            print(f"  Temp. média: {periodo2['Temperatura (°C)'].mean():.2f}°C")
            print(f"  Temp. min/max: {periodo2['Temperatura (°C)'].min():.1f}°C / {periodo2['Temperatura (°C)'].max():.1f}°C")
        else:
            print("  ⚠️ Nenhum registro encontrado neste período")
    except:
        pass
    
    print("\n" + "="*80)

# Execução principal
if __name__ == "__main__":
    print("\n" + "="*80)
    print("🌡️  EXTRAÇÃO DE DADOS DE TEMPERATURA - THERMO KING")
    print("="*80 + "\n")
    
    # Caminho do PDF
    pdf_path = r"C:\Fabio\Desenvolvimento\Varejo\Temperatura\Temp_eq_Thermoking.pdf"
    
    print(f"📄 Processando: {pdf_path}\n")
    
    # Extrai dados
    df = extract_temperature_data(pdf_path)
    
    if not df.empty:
        print(f"\n✅ Extração concluída! Total: {len(df)} registros\n")
        
        # Cria Excel e gráfico
        create_excel_and_chart(df)
    else:
        print("❌ Nenhum dado foi encontrado no PDF.")