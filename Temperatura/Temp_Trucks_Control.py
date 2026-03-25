# -*- coding: utf-8 -*-
import re
import math
import datetime as dt
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.dates import DateFormatter
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from pathlib import Path

# =========================
# CONFIGURAÇÕES
# =========================
PDF_PATH = "Temp_TRUCKS_CONTROL.pdf"
OUT_XLSX = "Temperatura_Trucks_Control.xlsx"
OUT_PNG  = "Grafico_TemperaturaTrucks_Control.png"

# Mapeamento: O PDF usa "Sensor 1", "Sensor 2", etc.
# Vamos mapear Sensor 1 -> T1, Sensor 2 -> T2, Sensor 3 -> T3
APURADA_PRIORIDADE = ["T1", "T2", "T3", "TS"]

# =========================
# FUNÇÕES AUXILIARES
# =========================
def _parse_float_cell(txt: str):
    if not txt: return None
    s = str(txt).strip().replace('"', '').replace(',', '.')
    try:
        return float(s)
    except:
        return None

def escolher_temperatura_apurada(row):
    for col in APURADA_PRIORIDADE:
        v = row.get(col)
        if v is not None and not (isinstance(v, float) and math.isnan(v)):
            return v
    return None

# =========================
# EXTRAÇÃO DO PDF
# =========================
def extrair_do_pdf(pdf_path: str) -> pd.DataFrame:
    registros = []
    # Regex para data (DD/MM/AAAA) e hora (HH:MM:SS)
    RE_DATE = re.compile(r"(\d{2}/\d{2}/\d{4})")
    RE_TIME = re.compile(r"(\d{2}:\d{2}:\d{2})")

    import pdfplumber
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            lines = text.splitlines()
            
            for i, line in enumerate(lines):
                # Limpa a linha de lixo comum no PDF
                clean_line = line.replace('"', '').replace(',', ' ').strip()
                
                date_match = RE_DATE.search(clean_line)
                time_match = RE_TIME.search(clean_line)
                
                if date_match and time_match:
                    dh_str = f"{date_match.group(1)} {time_match.group(1)}"
                    dh = dt.datetime.strptime(dh_str, "%d/%m/%Y %H:%M:%S")
                    
                    # Heurística para este PDF: A temperatura (Sensor 1) 
                    # geralmente é o último ou penúltimo valor numérico da linha
                    parts = clean_line.split()
                    nums = []
                    for p in parts:
                        val = _parse_float_cell(p)
                        if val is not None and -50 < val < 60: # Filtro de temperatura plausível
                            nums.append(val)
                    
                    if nums:
                        # No relatório Brasil Risk, a temperatura vem após Latitude/Longitude
                        # Pegamos o último valor como Sensor 1 (T1)
                        registros.append({
                            "DataHora": dh,
                            "T1": nums[-1], 
                            "T2": None, "T3": None, "TS": None
                        })

    df = pd.DataFrame(registros).sort_values("DataHora").reset_index(drop=True)
    df["Temperatura_Apurada"] = df.apply(escolher_temperatura_apurada, axis=1)
    return df

# =========================
# GERAÇÃO DE SAÍDAS (Mantido do seu original)
# =========================
def criar_grafico_png(df, out_png):
    if df.empty: return
    fig, ax = plt.subplots(figsize=(12, 6))
    ax.plot(df["DataHora"], df["Temperatura_Apurada"], label="Temp. Sensor 1", color='blue')
    ax.axhspan(-40, -12, color='red', alpha=0.1, label="Faixa Temp área Congelada")
    ax.axhspan(0, 5, color='green', alpha=0.1, label="Faixa Temp área Refrigerada")
    ax.set_title("Análise de Variação de Temperatura - Trucks Control")
    ax.xaxis.set_major_formatter(DateFormatter("%d/%m %H:%M"))
    plt.xticks(rotation=45)
    plt.legend()
    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    fig.savefig(out_png, dpi=150)
    plt.close(fig)

# Execução
if __name__ == "__main__":
    df_dados = extrair_do_pdf(PDF_PATH)
    if not df_dados.empty:
        criar_grafico_png(df_dados, OUT_PNG)
        df_dados.to_excel(OUT_XLSX, index=False)
        print(f"Sucesso! {len(df_dados)} registros processados.")
    else:
        print("Nenhum dado extraído.")