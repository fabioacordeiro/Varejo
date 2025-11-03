# -*- coding: utf-8 -*-
import argparse
from pathlib import Path
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import sys

# ====== AJUSTE AQUI OS SEUS PADRÕES ======
DEFAULT_CSV   = r"C:\\Fabio\\Desenvolvimento\\Varejo\\Temperatura\\PDQ9B50 - Temperatura.csv"
DEFAULT_EXCEL = r"C:\\Fabio\\Desenvolvimento\\Varejo\\Temperatura\\Temperatura_saida_csv.xlsx"
DEFAULT_IMG   = r"C:\\Fabio\\Desenvolvimento\\Varejo\\Temperatura\\Grafico_Temperatura_csv.png"
DEFAULT_TEMP_COL = "T1"              # mude se quiser usar T2/T3/TD1...
DEFAULT_DATE_COL = "Data Posição"    # ou deixe None
# =========================================

ENCODINGS = ["utf-8", "latin-1", "cp1252"]
PRIORITY_TEMP_COLS = ["T1","T2","T3","TD1","TD2","TD3","TS","US"]

def read_csv_smart(csv_path: Path) -> pd.DataFrame:
    last_err = None
    for enc in ENCODINGS:
        try:
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
    if cands: return cands[0]
    numeric = []
    for c in df.columns:
        s = pd.to_numeric(df[c].astype(str).str.replace(",",".",regex=False), errors="coerce")
        if s.notna().sum() >= max(3, int(0.2*len(s))):
            numeric.append((c, s.notna().sum()))
    if numeric:
        numeric.sort(key=lambda t:t[1], reverse=True)
        return numeric[0][0]
    raise RuntimeError("Não foi possível identificar a coluna de temperatura.")

def normalize_float(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s.astype(str).str.replace(",",".",regex=False), errors="coerce")

def pick_x_axis(df: pd.DataFrame, user_date_col=None):
    if user_date_col and user_date_col in df.columns:
        dt = pd.to_datetime(df[user_date_col], errors="coerce", dayfirst=True)
        if dt.notna().any(): return dt
    for c in ["Data Posição","Data","Date","Timestamp","Datetime","Data Chegada","Hora","Time"]:
        if c in df.columns:
            dt = pd.to_datetime(df[c], errors="coerce", dayfirst=True)
            if dt.notna().sum() >= max(3, int(0.2*len(dt))):
                return dt
    for c in df.columns:
        dt = pd.to_datetime(df[c], errors="coerce", dayfirst=True)
        if dt.notna().sum() >= max(3, int(0.2*len(dt))):
            return dt
    return pd.RangeIndex(start=1, stop=len(df)+1)

def save_excel(df: pd.DataFrame, path: Path):
    ext = path.suffix.lower()
    if ext == ".xls":
        with pd.ExcelWriter(path, engine="xlwt") as wr:
            df.to_excel(wr, index=False, sheet_name="Temperatura")
    else:
        with pd.ExcelWriter(path if ext==".xlsx" else path.with_suffix(".xlsx"), engine="openpyxl") as wr:
            df.to_excel(wr, index=False, sheet_name="Temperatura")

def plot_temperature_image(x, y, img_path: Path,
                           low_a=-18, low_b=-12, hi_a=0, hi_b=5):
    import matplotlib.pyplot as plt
    plt.figure(figsize=(12,5))
    ax = plt.gca()
    ax.axhspan(low_a, low_b, alpha=0.15)
    ax.axhspan(hi_a,  hi_b,  alpha=0.15)
    ax.plot(x, y, linewidth=1.8, color="orange", label="Temperatura apurada")
    ax.set_title("Temperatura — Faixas: [-18, -12] e [0, 5] °C; Linha laranja = apurada")
    ax.set_xlabel("Leitura"); ax.set_ylabel("°C"); ax.grid(True); ax.legend()
    plt.tight_layout(); plt.savefig(img_path, dpi=150); plt.close()

def main():
    ap = argparse.ArgumentParser(description="CSV → Excel + Gráfico")
    ap.add_argument("--csv", default=DEFAULT_CSV)
    ap.add_argument("--excel", default=DEFAULT_EXCEL)
    ap.add_argument("--img", default=DEFAULT_IMG)
    ap.add_argument("--temp-col", default=DEFAULT_TEMP_COL)
    ap.add_argument("--date-col", default=DEFAULT_DATE_COL)
    ap.add_argument("--low-a", type=float, default=-18.0)
    ap.add_argument("--low-b", type=float, default=-12.0)
    ap.add_argument("--hi-a", type=float, default=0.0)
    ap.add_argument("--hi-b", type=float, default=5.0)
    args = ap.parse_args()

    df = read_csv_smart(Path(args.csv))
    temp_col = pick_temperature_column(df, args.temp_col)
    y = normalize_float(df[temp_col])
    x = pick_x_axis(df, args.date_col)

    out_df = df.copy()
    out_df["Temperatura_Apurada"] = y

    excel_path = Path(args.excel); excel_path.parent.mkdir(parents=True, exist_ok=True)
    save_excel(out_df, excel_path)

    img_path = Path(args.img); img_path.parent.mkdir(parents=True, exist_ok=True)
    plot_temperature_image(x, y, img_path, args.low_a, args.low_b, args.hi_a, args.hi_b)

    print("✅ Pronto!")
    print(f"Excel: {excel_path}")
    print(f"Imagem: {img_path}")
    print(f"Coluna usada: {temp_col}")

if __name__ == "__main__":
    main()