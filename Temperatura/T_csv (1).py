# -*- coding: utf-8 -*-
import argparse
from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
import sys

# ====== AJUSTE AQUI OS SEUS PADRÕES ======
DEFAULT_CSV   = r"C:\\Fabio\\Desenvolvimento\\Varejo\\Temperatura\\PDQ9B50 - Temperatura.csv"
DEFAULT_EXCEL = r"C:\\Fabio\\Desenvolvimento\\Varejo\\Temperatura\\Temperatura_saida_csv.xlsx"
DEFAULT_IMG   = r"C:\\Fabio\\Desenvolvimento\\Varejo\\Temperatura\\Grafico_Temperatura_csv.png"
DEFAULT_TEMP_COL   = "T1"              # coluna de temperatura
DEFAULT_DATE_COL   = "Data Posição"    # coluna de data/tempo (eixo X)

# Faixas de temperatura para destacar
LOW_A, LOW_B = -18.0, -12.0
HI_A,  HI_B  = 0.0,   5.0

# Períodos a destacar no gráfico (interpretei 1554 como 15:54)
# Formato de strings: "DD/MM/AAAA HH:MM"
HIGHLIGHT_PERIODS = [
    ("04/11/2025 02:00", "04/11/2025 03:36"),
    ("04/11/2025 12:00", "04/11/2025 12:30"),
]
# =========================================

ENCODINGS = ["utf-8", "latin-1", "cp1252"]
PRIORITY_TEMP_COLS = ["T1","T2","T3","TD1","TD2","TD3","TS","US"]

def read_csv_smart(csv_path: Path) -> pd.DataFrame:
    last_err = None
    for enc in ENCODINGS:
        try:
            # sep=None com engine="python" faz autodetecção do separador
            return pd.read_csv(csv_path, sep=None, engine="python", encoding=enc)
        except Exception as e:
            last_err = e
    raise RuntimeError(f"Falha ao ler CSV: {last_err}")

def pick_temperature_column(df: pd.DataFrame, user_choice=None):
    if user_choice and user_choice in df.columns:
        return user_choice
    for c in PRIORITY_TEMP_COLS:
        if c in df.columns:
            return c
    lowered = {c: str(c).lower() for c in df.columns}
    keys = ["temp","temperatura","temperature","°c","celsius"]
    cands = [c for c, low in lowered.items() if any(k in low for k in keys)]
    if cands:
        return cands[0]
    # fallback: coluna mais numérica
    numeric = []
    for c in df.columns:
        s = pd.to_numeric(df[c].astype(str).str.replace(",",".",regex=False), errors="coerce")
        if s.notna().sum() >= max(3, int(0.2*len(s))):
            numeric.append((c, s.notna().sum()))
    if numeric:
        numeric.sort(key=lambda t:t[1], reverse=True)
        return numeric[0][0]
    raise RuntimeError("Não foi possível identificar a coluna de temperatura.")

def to_float(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series.astype(str).str.replace(",",".",regex=False), errors="coerce")

def pick_x_axis(df: pd.DataFrame, user_date_col=None):
    # 1) coluna informada
    if user_date_col and user_date_col in df.columns:
        dt = pd.to_datetime(df[user_date_col], errors="coerce", dayfirst=True)
        if dt.notna().any():
            return dt
    # 2) Colunas comuns
    for c in ["Data Posição","Data","Date","Timestamp","Datetime","Data Chegada","Hora","Time"]:
        if c in df.columns:
            dt = pd.to_datetime(df[c], errors="coerce", dayfirst=True)
            if dt.notna().sum() >= max(3, int(0.2*len(dt))):
                return dt
    # 3) primeira coluna parseável
    for c in df.columns:
        dt = pd.to_datetime(df[c], errors="coerce", dayfirst=True)
        if dt.notna().sum() >= max(3, int(0.2*len(dt))):
            return dt
    # 4) índice
    return pd.RangeIndex(start=1, stop=len(df)+1)

def save_excel(df: pd.DataFrame, path: Path):
    ext = path.suffix.lower()
    if ext == ".xlsx":
        with pd.ExcelWriter(path, engine="openpyxl") as wr:
            df.to_excel(wr, index=False, sheet_name="Temperatura")
    else:
        # se extensão não for .xlsx, salva como .xlsx
        with pd.ExcelWriter(path.with_suffix(".xlsx"), engine="openpyxl") as wr:
            df.to_excel(wr, index=False, sheet_name="Temperatura")

def highlight_band_y(ax, y0, y1, alpha=0.15):
    ax.axhspan(y0, y1, alpha=alpha)

def highlight_periods_x(ax, x_series, periods, alpha=0.15):
    """
    periods: lista de tuplas (start_str, end_str) no formato 'DD/MM/AAAA HH:MM'
    x_series: série usada no eixo X (idealmente datetime)
    """
    # Se o eixo X não for datetime, não há como posicionar as janelas por data/hora
    # Nesses casos, a função apenas avisa no console e não desenha as faixas.
    if not (hasattr(x_series, "dt") or str(x_series.dtype).startswith("datetime")):
        print("⚠️ Eixo X não é datetime — não é possível sombrear períodos por data/hora.")
        return

    for start_str, end_str in periods:
        start = pd.to_datetime(start_str, dayfirst=True, errors="coerce")
        end   = pd.to_datetime(end_str,   dayfirst=True, errors="coerce")
        if pd.isna(start) or pd.isna(end):
            print(f"⚠️ Período inválido: '{start_str}' → '{end_str}' (ignorado)")
            continue
        if end < start:
            start, end = end, start
        ax.axvspan(start, end, alpha=0.18)

def plot_temperature_image(x, y, img_path: Path, temp_col: str):
    plt.figure(figsize=(12,5))
    ax = plt.gca()

    # Faixas de temperatura (Y)
    highlight_band_y(ax, LOW_A, LOW_B, alpha=0.15)
    highlight_band_y(ax, HI_A,  HI_B,  alpha=0.15)

    # Linha laranja = temperatura
    ax.plot(x, y, linewidth=1.8, color="orange", label=f"Temperatura apurada ({temp_col})")

    # Faixas de períodos (X) — só funciona se x for datetime
    highlight_periods_x(ax, x, HIGHLIGHT_PERIODS, alpha=0.18)

    ax.set_title("Temperatura — Faixas: [-18, -12] e [0, 5] °C; Linha laranja = apurada")
    ax.set_xlabel("Leitura")
    ax.set_ylabel("°C")
    ax.grid(True)
    ax.legend()
    plt.tight_layout()
    plt.savefig(img_path, dpi=150)
    plt.close()

def main():
    ap = argparse.ArgumentParser(description="CSV → Excel + Gráfico com faixas e períodos destacados")
    ap.add_argument("--csv", default=DEFAULT_CSV)
    ap.add_argument("--excel", default=DEFAULT_EXCEL)
    ap.add_argument("--img", default=DEFAULT_IMG)
    ap.add_argument("--temp-col", default=DEFAULT_TEMP_COL)
    ap.add_argument("--date-col", default=DEFAULT_DATE_COL)
    args = ap.parse_args()

    csv_path = Path(args.csv)
    if not csv_path.exists():
        print(f"Arquivo CSV não encontrado: {csv_path}")
        sys.exit(1)

    # 1) Ler CSV
    df = read_csv_smart(csv_path)

    # 2) Coluna temperatura
    temp_col = pick_temperature_column(df, args.temp_col)

    # 3) Série de temperatura (float)
    y = to_float(df[temp_col])

    # 4) Eixo X (idealmente datetime)
    x = pick_x_axis(df, args.date_col)

    # 5) Excel (mantém dados + coluna "Temperatura_Apurada")
    out_df = df.copy()
    out_df["Temperatura_Apurada"] = y
    excel_path = Path(args.excel); excel_path.parent.mkdir(parents=True, exist_ok=True)
    save_excel(out_df, excel_path)

    # 6) Gráfico com destaques
    img_path = Path(args.img); img_path.parent.mkdir(parents=True, exist_ok=True)
    plot_temperature_image(x, y, img_path, temp_col)

    print("--------------------------------------------------")
    print(f"✅ Excel salvo em:  {excel_path}")
    print(f"✅ Imagem salva em: {img_path}")
    print(f"ℹ️ Coluna usada:    {temp_col}")
    # Se x não for datetime, avisar que períodos não foram sombreados
    if not (hasattr(x, "dt") or str(x.dtype).startswith("datetime")):
        print("⚠️ Observação: Eixo X não é datetime; os períodos (20/10/2025 ...) não puderam ser sombreados.")
    else:
        print("🟦 Períodos destacados no eixo X: ", HIGHLIGHT_PERIODS)
    print("--------------------------------------------------")

if __name__ == "__main__":
    main()