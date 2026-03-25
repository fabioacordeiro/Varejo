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
                if len(row) >= 22:  # Ensure row has all columns
                    all_rows.append(row)
        
        if page_num % 20 == 0:
            print(f"  Processadas {page_num} páginas...")

print(f"Total de linhas extraídas: {len(all_rows)}")

# --- DEBUG: Show first 3 rows ---
print("\n=== DEBUG: Primeiras 3 linhas ===")
for i, row in enumerate(all_rows[:3]):
    print(f"\nLinha {i+1}:")
    print(f"  [0] Placa: '{row[0]}'")
    print(f"  [2] Data: '{row[2]}'")
    print(f"  [6] Temp1: '{row[6]}'")
    print(f"  [7] Temp2: '{row[7]}'")
    print(f"  [8] Temp3: '{row[8]}'")
    print(f"  [9] Temp4: '{row[9]}'")
    print(f"  [10] Temp5: '{row[10]}'")
    print(f"  [15] Velocidade: '{row[15]}'")

# --- 2. Process Data ---
processed_data = []
errors = {'date': 0, 'temp': 0, 'other': 0}

for row_idx, row in enumerate(all_rows):
    try:
        # Extract date/time (column 2 - index 2)
        date_str = row[2].replace('\n', ' ').strip()
        
        # Parse datetime
        dt_object = pd.to_datetime(date_str, format='%d/%m/%Y %H:%M:%S', errors='coerce')
        
        if pd.isna(dt_object):
            errors['date'] += 1
            if row_idx < 5:  # Debug first 5 failures
                print(f"\n⚠️ Falha ao parsear data na linha {row_idx}: '{date_str}'")
            continue
        
        # Extract temperatures from columns 6-10 (Temperatura 1-5)
        temp_columns = [6, 7, 8, 9, 10]
        found_temp = False
        
        for idx, col_idx in enumerate(temp_columns, 1):
            temp_str = row[col_idx].replace('\n', '').strip()
            
            # Skip empty values (-)
            if temp_str == '-' or not temp_str:
                continue
            
            # Convert comma to dot and parse
            try:
                temp_value = float(temp_str.replace(',', '.'))
                found_temp = True
            except ValueError:
                errors['temp'] += 1
                if row_idx < 5:
                    print(f"⚠️ Falha ao parsear temperatura na linha {row_idx}, coluna {col_idx}: '{temp_str}'")
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
            localizacao = row[12].replace('\n', ' ').strip()[:100]  # Limit location length
            
            processed_data.append({
                'Data/Hora': dt_object,
                'Temperatura (°C)': temp_value,
                'Sensor': f'Sensor {idx}',
                'Placa': placa,
                'Localização': localizacao,
                'Velocidade (km/h)': velocidade,
                'Estado': estado
            })
        
        if not found_temp:
            errors['temp'] += 1
    
    except Exception as e:
        errors['other'] += 1
        if row_idx < 5:
            print(f"⚠️ Erro na linha {row_idx}: {e}")
        continue

print(f"\n=== ERROS DE PROCESSAMENTO ===")
print(f"Erros de data: {errors['date']}")
print(f"Erros de temperatura: {errors['temp']}")
print(f"Outros erros: {errors['other']}")

# Create DataFrame
df_final = pd.DataFrame(processed_data)

if df_final.empty:
    print("\n❌ Nenhum dado válido extraído após processamento.")
    print("Verifique os erros acima para identificar o problema.")
    exit()

df_final = df_final.sort_values(by='Data/Hora').reset_index(drop=True)

print(f"\n✅ Registros processados: {len(df_final)}")
print(f"📊 Temperatura mínima: {df_final['Temperatura (°C)'].min():.1f}°C")
print(f"📊 Temperatura máxima: {df_final['Temperatura (°C)'].max():.1f}°C")
print(f"📅 Período: {df_final['Data/Hora'].min()} até {df_final['Data/Hora'].max()}")

# --- 3. Save to Excel ---
try:
    df_final.to_excel(EXCEL_OUTPUT, index=False)
    print(f"\n💾 Excel salvo: {EXCEL_OUTPUT}")
except Exception as e:
    print(f"❌ Erro ao salvar Excel: {e}")

# --- 4. Generate Plot ---
print("\n📈 Gerando gráfico...")

fig, ax = plt.subplots(figsize=(16, 8))

# Main temperature line
ax.plot(df_final['Data/Hora'], df_final['Temperatura (°C)'], 
        label='Temperatura', color='blue', linewidth=1, alpha=0.7)

# Highlight temperature ranges
red_points = df_final[
    (df_final['Temperatura (°C)'] >= HIGHLIGHT_TEMP_RED[0]) &
    (df_final['Temperatura (°C)'] <= HIGHLIGHT_TEMP_RED[1])
]
if not red_points.empty:
    ax.scatter(red_points['Data/Hora'], red_points['Temperatura (°C)'], 
               color='red', s=50, zorder=5, 
               label=f'🔴 {HIGHLIGHT_TEMP_RED[0]}°C a {HIGHLIGHT_TEMP_RED[1]}°C ({len(red_points)} pts)')

orange_points = df_final[
    (df_final['Temperatura (°C)'] >= HIGHLIGHT_TEMP_ORANGE[0]) &
    (df_final['Temperatura (°C)'] <= HIGHLIGHT_TEMP_ORANGE[1])
]
if not orange_points.empty:
    ax.scatter(orange_points['Data/Hora'], orange_points['Temperatura (°C)'], 
               color='orange', s=50, zorder=5,
               label=f'🟠 {HIGHLIGHT_TEMP_ORANGE[0]}°C a {HIGHLIGHT_TEMP_ORANGE[1]}°C ({len(orange_points)} pts)')

# Highlight time periods
ax.axvspan(HIGHLIGHT_PERIOD_GREEN_START, HIGHLIGHT_PERIOD_GREEN_END, 
           color='green', alpha=0.2, 
           label=f'🟢 {HIGHLIGHT_PERIOD_GREEN_START.strftime("%H:%M")}-{HIGHLIGHT_PERIOD_GREEN_END.strftime("%H:%M")}')

ax.axvspan(HIGHLIGHT_PERIOD_PURPLE_START, HIGHLIGHT_PERIOD_PURPLE_END, 
           color='purple', alpha=0.2,
           label=f'🟣 {HIGHLIGHT_PERIOD_PURPLE_START.strftime("%H:%M")}-{HIGHLIGHT_PERIOD_PURPLE_END.strftime("%H:%M")}')

# Chart configuration
ax.set_title('Variação de Temperatura - Placa FWG9696', fontsize=14, fontweight='bold')
ax.set_xlabel('Data/Hora', fontsize=12)
ax.set_ylabel('Temperatura (°C)', fontsize=12)
ax.grid(True, linestyle='--', alpha=0.5)
ax.legend(loc='best', fontsize=10)
plt.xticks(rotation=45, ha='right')
plt.tight_layout()

# Save plot
try:
    plt.savefig(GRAPH_OUTPUT, dpi=150)
    print(f"📊 Gráfico salvo: {GRAPH_OUTPUT}")
except Exception as e:
    print(f"❌ Erro ao salvar gráfico: {e}")

plt.show()

print("\n✅ Processamento concluído!")