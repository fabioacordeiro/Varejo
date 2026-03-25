import pdfplumber
import pandas as pd
import matplotlib.pyplot as plt
import re
from datetime import datetime

# --- Configuration ---
PDF_FILE = "C:\Fabio\Desenvolvimento\Varejo\Temperatura\Temp_omini.pdf"
EXCEL_OUTPUT = "C:\Fabio\Desenvolvimento\Varejo\Temperatura\dados_temperatura_omini.xlsx"
GRAPH_OUTPUT = "C:\Fabio\Desenvolvimento\Varejo\Temperatura\grafico_temperatura_omini.png"

# Temperature ranges to highlight
HIGHLIGHT_TEMP_RED = (-15, -12)
HIGHLIGHT_TEMP_ORANGE = (0, 5)

# Time periods to highlight
HIGHLIGHT_PERIOD_GREEN_START = datetime(2025, 11, 4, 2, 0)
HIGHLIGHT_PERIOD_GREEN_END = datetime(2025, 11, 4, 3, 36)
HIGHLIGHT_PERIOD_PURPLE_START = datetime(2025, 11, 4, 12, 0)
HIGHLIGHT_PERIOD_PURPLE_END = datetime(2025, 11, 4, 12, 30)

# --- 1. Extract Data from PDF ---
print(f"Lendo PDF: {PDF_FILE}...")
all_rows = []

with pdfplumber.open(PDF_FILE) as pdf:
    for page_num, page in enumerate(pdf.pages, 1):
        tables = page.extract_tables()
        
        for table in tables:
            if not table or len(table) < 2:
                continue
            
            # Skip header row (first row)
            for row in table[1:]:
                if len(row) >= 22:
                    all_rows.append(row)
        
        if page_num % 20 == 0:
            print(f"  Processadas {page_num} páginas...")

print(f"Total de linhas extraídas: {len(all_rows)}")

# --- 2. Process Data ---
processed_data = []

for row_idx, row in enumerate(all_rows):
    try:
        # Extract date/time (column 2 - index 2)
        date_str = row[2].replace('\n', '').strip()
        
        # Parse datetime
        dt_object = pd.to_datetime(date_str, format='%d/%m/%Y %H:%M:%S', errors='coerce')
        
        if pd.isna(dt_object):
            continue
        
        # Extract temperatures from columns 6-10 (Temperatura 1-5)
        temp_columns = [6, 7, 8, 9, 10]
        
        for idx, col_idx in enumerate(temp_columns, 1):
            temp_str = row[col_idx].replace('\n', '').strip()
            
            # Skip empty values (-)
            if temp_str == '-' or not temp_str:
                continue
            
            # Convert comma to dot and parse
            try:
                temp_value = float(temp_str.replace(',', '.'))
            except ValueError:
                continue
            
            # Extract other data
            placa = row[0].strip()
            velocidade_str = row[15].replace('\n', '').strip()
            velocidade = None
            
            if 'km/h' in velocidade_str:
                try:
                    velocidade = float(velocidade_str.replace('km/h', '').replace(',', '.').strip())
                except:
                    pass
            
            estado = row[19].replace('\n', ' ').strip()
            localizacao = row[12].replace('\n', ' ').strip()[:100]
            
            processed_data.append({
                'Data/Hora': dt_object,
                'Temperatura (°C)': temp_value,
                'Sensor': f'Sensor {idx}',
                'Placa': placa,
                'Localização': localizacao,
                'Velocidade (km/h)': velocidade,
                'Estado': estado
            })
    
    except Exception as e:
        continue

# Create DataFrame
df_final = pd.DataFrame(processed_data)

if df_final.empty:
    print("\n❌ Nenhum dado válido extraído.")
    exit()

df_final = df_final.sort_values(by='Data/Hora').reset_index(drop=True)

print(f"\n✅ Registros processados: {len(df_final)}")
print(f"📊 Temperatura mínima: {df_final['Temperatura (°C)'].min():.1f}°C")
print(f"📊 Temperatura máxima: {df_final['Temperatura (°C)'].max():.1f}°C")
print(f"📅 Período: {df_final['Data/Hora'].min()} até {df_final['Data/Hora'].max()}")

# Count points in temperature ranges
red_points = df_final[
    (df_final['Temperatura (°C)'] >= HIGHLIGHT_TEMP_RED[0]) &
    (df_final['Temperatura (°C)'] <= HIGHLIGHT_TEMP_RED[1])
]
orange_points = df_final[
    (df_final['Temperatura (°C)'] >= HIGHLIGHT_TEMP_ORANGE[0]) &
    (df_final['Temperatura (°C)'] <= HIGHLIGHT_TEMP_ORANGE[1])
]

print(f"🔴 Pontos entre {HIGHLIGHT_TEMP_RED[0]}°C e {HIGHLIGHT_TEMP_RED[1]}°C: {len(red_points)}")
print(f"🟠 Pontos entre {HIGHLIGHT_TEMP_ORANGE[0]}°C e {HIGHLIGHT_TEMP_ORANGE[1]}°C: {len(orange_points)}")

# --- 3. Save to Excel ---
try:
    df_final.to_excel(EXCEL_OUTPUT, index=False)
    print(f"\n💾 Excel salvo: {EXCEL_OUTPUT}")
except Exception as e:
    print(f"❌ Erro ao salvar Excel: {e}")

# --- 4. Generate Plot ---
print("\n📈 Gerando gráfico...")

fig, ax = plt.subplots(figsize=(18, 9))

# Main temperature line
ax.plot(df_final['Data/Hora'], df_final['Temperatura (°C)'], 
        label='Temperatura', color='blue', linewidth=1.5, alpha=0.6)

# Add horizontal bands for temperature ranges
# Red band: -15°C to -12°C
ax.axhspan(HIGHLIGHT_TEMP_RED[0], HIGHLIGHT_TEMP_RED[1], 
           color='red', alpha=0.15, zorder=1,
           label=f'🔴 Faixa Crítica: {HIGHLIGHT_TEMP_RED[0]}°C a {HIGHLIGHT_TEMP_RED[1]}°C')

# Orange band: 0°C to 5°C
ax.axhspan(HIGHLIGHT_TEMP_ORANGE[0], HIGHLIGHT_TEMP_ORANGE[1], 
           color='orange', alpha=0.15, zorder=1,
           label=f'🟠 Faixa Alerta: {HIGHLIGHT_TEMP_ORANGE[0]}°C a {HIGHLIGHT_TEMP_ORANGE[1]}°C')

# Highlight points in temperature ranges with larger markers
if not red_points.empty:
    ax.scatter(red_points['Data/Hora'], red_points['Temperatura (°C)'], 
               color='red', s=100, zorder=5, marker='o', edgecolors='darkred', linewidths=2,
               label=f'Pontos Críticos ({len(red_points)} registros)')

if not orange_points.empty:
    ax.scatter(orange_points['Data/Hora'], orange_points['Temperatura (°C)'], 
               color='orange', s=100, zorder=5, marker='o', edgecolors='darkorange', linewidths=2,
               label=f'Pontos Alerta ({len(orange_points)} registros)')

# Highlight time periods with vertical spans
ax.axvspan(HIGHLIGHT_PERIOD_GREEN_START, HIGHLIGHT_PERIOD_GREEN_END, 
           color='green', alpha=0.25, zorder=0,
           label=f'🟢 Período 1: {HIGHLIGHT_PERIOD_GREEN_START.strftime("%d/%m %H:%M")}-{HIGHLIGHT_PERIOD_GREEN_END.strftime("%H:%M")}')

ax.axvspan(HIGHLIGHT_PERIOD_PURPLE_START, HIGHLIGHT_PERIOD_PURPLE_END, 
           color='purple', alpha=0.25, zorder=0,
           label=f'🟣 Período 2: {HIGHLIGHT_PERIOD_PURPLE_START.strftime("%d/%m %H:%M")}-{HIGHLIGHT_PERIOD_PURPLE_END.strftime("%H:%M")}')

# Chart configuration
ax.set_title('Monitoramento de Temperatura - Placa FWG9696\nDestaques: Faixas Críticas e Períodos Específicos', 
             fontsize=15, fontweight='bold', pad=20)
ax.set_xlabel('Data/Hora', fontsize=13, fontweight='bold')
ax.set_ylabel('Temperatura (°C)', fontsize=13, fontweight='bold')
ax.grid(True, linestyle='--', alpha=0.4, linewidth=0.8)
ax.legend(loc='best', fontsize=10, framealpha=0.9)
plt.xticks(rotation=45, ha='right')
plt.tight_layout()

# Save plot
try:
    plt.savefig(GRAPH_OUTPUT, dpi=200, bbox_inches='tight')
    print(f"📊 Gráfico salvo: {GRAPH_OUTPUT}")
except Exception as e:
    print(f"❌ Erro ao salvar gráfico: {e}")

plt.show()

print("\n✅ Processamento concluído!")
print(f"\n📋 Resumo:")
print(f"   - Total de registros: {len(df_final)}")
print(f"   - Registros na faixa crítica (-15°C a -12°C): {len(red_points)}")
print(f"   - Registros na faixa de alerta (0°C a 5°C): {len(orange_points)}")