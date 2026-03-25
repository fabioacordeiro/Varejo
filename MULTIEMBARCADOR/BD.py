# -*- coding: utf-8 -*-
"""
Compila todos os CSVs de uma raiz em um único banco SQLite (tabela 'registros'),
filtra pelos CNPJs informados e exporta um Excel único usando o engine 'openpyxl'.

Requisitos:
    pip install pandas openpyxl

Opcional (se quiser o banco SQLite):
    (já incluso no Python padrão)
"""
import re
import sqlite3
from pathlib import Path

import pandas as pd

# ==============================
# CONFIGURAÇÕES
# ==============================
# Ajuste a raiz para varrer as pastas e subpastas com CSVs
ROOT_DIR = Path(r"C:\\Fabio\Desenvolvimento\\Varejo\\MULTIEMBARCADOR")  # <-- ajuste
# Saídas
OUTPUT_DB = Path(r"C:\\Fabio\Desenvolvimento\\Varejo\\MULTIEMBARCADOR\dados_compilados.db")  # <-- ajuste
OUTPUT_XLSX = Path(r"C:\\Fabio\Desenvolvimento\\Varejo\\MULTIEMBARCADOR\\resultado_filtrado.xlsx")  # <-- ajuste

# CNPJs alvo (normalizados: somente dígitos)
ALVOS = {"63073266002047", "63073266002470"}

# ==============================
# FUNÇÕES
# ==============================

def garantir_engine_openpyxl():
    """Garante que openpyxl esteja disponível para o pandas ExcelWriter."""
    try:
        import openpyxl  # noqa: F401
    except ImportError as exc:
        raise RuntimeError(
            "O pacote 'openpyxl' não está instalado. "
            "Instale com: pip install openpyxl"
        ) from exc


def normaliza_cnpj(valor: str) -> str | None:
    """Mantém somente dígitos; retorna 14 dígitos ou None."""
    if pd.isna(valor):
        return None
    s = re.sub(r"\D", "", str(valor))
    return s if len(s) == 14 else None


def detectar_colunas_cnpj(df: pd.DataFrame) -> list[str]:
    """Detecta colunas de CNPJ por nome e, se necessário, por padrão dos dados."""
    candidatos = []
    for col in df.columns:
        col_norm = re.sub(r"[^a-z0-9]", "", str(col).lower())
        if "cnpj" in col_norm:
            candidatos.append(col)

    # fallback por padrão dos dados (amostra)
    if not candidatos and not df.empty:
        for col in df.columns:
            amostra = df[col].astype(str).head(50).tolist()
            hits = sum(
                1
                for x in amostra
                if re.search(r"\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}", x)
                or re.fullmatch(r"\d{14}", re.sub(r"\D", "", x))
            )
            if hits >= max(3, int(len(amostra) * 0.2)):
                candidatos.append(col)

    return candidatos


def tenta_ler_csv(caminho: Path) -> pd.DataFrame | None:
    """
    Tenta ler um CSV com combinações de separador/encoding.
    Lê tudo como texto (dtype=str) para evitar DtypeWarning/zeros à esquerda.
    """
    tentativas = [
        {"sep": ";", "encoding": "utf-8", "dtype": str, "low_memory": False},
        {"sep": ",", "encoding": "utf-8", "dtype": str, "low_memory": False},
        {"sep": ";", "encoding": "latin1", "dtype": str, "low_memory": False},
        {"sep": ",", "encoding": "latin1", "dtype": str, "low_memory": False},
    ]
    for opt in tentativas:
        try:
            return pd.read_csv(caminho, **opt)
        except Exception:
            continue
    return None


def compilar_csvs(raiz: Path) -> pd.DataFrame:
    """Varre a raiz por *.csv, empilha e adiciona metadados de origem."""
    dfs = []
    arquivos = list(raiz.rglob("*.csv"))
    print(f"[INFO] CSVs encontrados: {len(arquivos)}")
    for arq in arquivos:
        df = tenta_ler_csv(arq)
        if df is None or df.empty:
            print(f"[WARN] Ignorando (falha/vasio): {arq}")
            continue
        df.columns = [str(c).strip() for c in df.columns]
        df["__arquivo_origem"] = str(arq.name)
        df["__pasta_origem"] = str(arq.parent)
        dfs.append(df)

    if not dfs:
        print("[INFO] Nenhum CSV válido lido.")
        return pd.DataFrame()

    combined = pd.concat(dfs, ignore_index=True, sort=False)
    print(f"[INFO] Linhas compiladas: {len(combined)} | Colunas: {len(combined.columns)}")
    return combined


def salvar_sqlite(df: pd.DataFrame, caminho_db: Path):
    """Salva DataFrame no SQLite (tabela 'registros') e cria índice simples."""
    if df.empty:
        print("[INFO] DataFrame vazio; não salvando no SQLite.")
        return
    caminho_db.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(caminho_db) as conn:
        df.to_sql("registros", conn, if_exists="replace", index=False)
        try:
            conn.execute(
                "CREATE INDEX IF NOT EXISTS idx_registros_arquivo ON registros(__arquivo_origem)"
            )
        except Exception:
            pass
        conn.commit()
    print(f"[OK] Banco salvo em: {caminho_db.resolve()}")


def filtrar_por_cnpj(df: pd.DataFrame, alvos: set[str]) -> pd.DataFrame:
    """Cria CNPJ_NORMALIZADO e retorna filtrado pelos CNPJs alvo."""
    if df.empty:
        return df.copy()

    cols_cnpj = detectar_colunas_cnpj(df)
    print(f"[INFO] Colunas de CNPJ detectadas: {cols_cnpj if cols_cnpj else 'nenhuma'}")

    def extrai_cnpj_linha(row):
        for c in cols_cnpj:
            v = normaliza_cnpj(row.get(c))
            if v:
                return v
        return None

    df = df.copy()
    df["CNPJ_NORMALIZADO"] = df.apply(extrai_cnpj_linha, axis=1) if cols_cnpj else None
    filtrado = df[df["CNPJ_NORMALIZADO"].isin(alvos)] if "CNPJ_NORMALIZADO" in df.columns else df.iloc[0:0]
    print(f"[INFO] Linhas filtradas pelos CNPJs-alvo: {len(filtrado)}")
    return filtrado


def exportar_excel_openpyxl(df_filtrado: pd.DataFrame, df_full: pd.DataFrame, saida_xlsx: Path):
    """
    Exporta para Excel usando engine 'openpyxl'.
    Por requisito do pedido: gera somente uma planilha com os dados filtrados.
    Caso filtrado esteja vazio, exporta o combinado completo para facilitar diagnóstico.
    """
    garantir_engine_openpyxl()
    saida_xlsx.parent.mkdir(parents=True, exist_ok=True)

    df_export = df_filtrado if not df_filtrado.empty else df_full
    nome_aba = "Dados"

    with pd.ExcelWriter(saida_xlsx, engine="openpyxl") as writer:
        df_export.to_excel(writer, sheet_name=nome_aba, index=False)

    print(f"[OK] Excel gerado em: {saida_xlsx.resolve()} (aba: {nome_aba})")


# ==============================
# MAIN
# ==============================

def main():
    # 1) Compilar CSVs
    combined = compilar_csvs(ROOT_DIR)

    # 2) Salvar no SQLite
    salvar_sqlite(combined, OUTPUT_DB)

    # 3) Filtrar pelos CNPJs alvo
    filtrado = filtrar_por_cnpj(combined, ALVOS)

    # 4) Exportar Excel usando openpyxl
    exportar_excel_openpyxl(filtrado, combined, OUTPUT_XLSX)

    # Resumo
    print("======== RESUMO ========")
    print(f"Linhas compiladas: {len(combined)}")
    print(f"Linhas filtradas : {len(filtrado)}")
    print(f"SQLite: {OUTPUT_DB}")
    print(f"Excel : {OUTPUT_XLSX}")

if __name__ == "__main__":
    main()