# -*- coding: utf-8 -*-
"""
Exporta tabelas/consultas do Access para um único Excel (.xlsx), 1 aba por objeto.
- Rota A (COM/Access) com pywin32: máxima fidelidade (TransferSpreadsheet).
- Rota B (ODBC) com pyodbc+pandas: robusto e sem o Access aberto.

USO (exemplos):
1) Exportar TABELAS específicas via COM (recomendado):
   python Exportar.py --accdb "C:\Fabio\CARREFOUR\Cadastro de Fornecedores Oficial.accdb" --out "C:\Fabio\Desenvolvimento\Varejo\Access\export.xlsx" --use-com --tables "CADASTRO FORNECEDORES, Tabela1"

2) Exportar TODAS as tabelas via COM:
   python access_to_excel.py --accdb "C:\dados\base.accdb" --out "C:\dados\export.xlsx" --use-com --all-tables

3) Exportar objetos (tabelas/consultas) via ODBC:
   python access_to_excel.py --accdb "C:\dados\base.accdb" --out "C:\dados\export.xlsx" --tables "Tabela1" --queries "MinhaConsulta"

4) Exportar uma CONSULTA via ODBC forçando SQL:
   python access_to_excel.py --accdb "C:\dados\base.accdb" --out "C:\dados\export.xlsx" --sql "SELECT * FROM [MinhaConsulta]"
"""

import argparse
import sys
from pathlib import Path

# Dependências opcionais (checadas em runtime)
try:
    import win32com.client as win32
    from win32com.client import constants as ac
except Exception:
    win32 = None
    ac = None

import pandas as pd

# ---------- Utilidades gerais ----------
def sanitize_sheet_name(name: str) -> str:
    # Limita 31 chars e remove inválidos para sheet do Excel
    invalid = '[]:*?/\\'
    s = ''.join(ch for ch in name if ch not in invalid)
    return s[:31] if s else 'Sheet'

def ensure_parent(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)

# ---------- ROTA A: COM/Access (pywin32) ----------
def export_via_com(accdb_path: Path, out_xlsx: Path, table_names=None, query_names=None,
                   all_tables=False, all_queries=False, overwrite=True):
    if win32 is None or ac is None:
        raise RuntimeError("pywin32 não está disponível. Instale com: pip install pywin32")

    ensure_parent(out_xlsx)

    # Abrir Access
    app = win32.Dispatch("Access.Application")
    app.Visible = False
    app.OpenCurrentDatabase(str(accdb_path))

    try:
        # Enumerar se necessário
        db = app.CurrentDb()
        tables = table_names[:] if table_names else []
        queries = query_names[:] if query_names else []

        if all_tables:
            for tdef in db.TableDefs:
                name = tdef.Name
                # Ignorar tabelas de sistema
                if name.startswith("MSys"):
                    continue
                tables.append(name)

        if all_queries:
            for qdef in db.QueryDefs:
                name = qdef.Name
                # Ignorar queries de sistema
                if name.startswith("~") or name.startswith("MSys"):
                    continue
                queries.append(name)

        # Remover duplicatas preservando ordem
        def uniq(seq):
            seen = set()
            out = []
            for x in seq:
                if x not in seen:
                    seen.add(x)
                    out.append(x)
            return out

        tables = uniq(tables)
        queries = uniq(queries)

        # Se arquivo existe e overwrite:
        if out_xlsx.exists() and overwrite:
            out_xlsx.unlink()

        # Exportar: cada objeto para uma aba.
        # Dica: TransferSpreadsheet com Range="NomeAba!"
        # Excel12Xml => .xlsx
        SpreadsheetType = ac.acSpreadsheetTypeExcel12Xml  # 10

        # Primeiro crie o arquivo exportando a primeira aba, depois vá acrescentando
        first = True

        # Função helper: exportar um objeto (tabela/consulta) para uma aba específica
        def export_object(obj_name: str):
            sheet = sanitize_sheet_name(obj_name)
            # Se não é a primeira exportação e o arquivo já existe, o Access "acrescenta" dados
            # na aba indicada pelo parâmetro Range. Para garantir nova aba com cabeçalhos,
            # usamos Range=f"{sheet}!"
            app.DoCmd.TransferSpreadsheet(
                ac.TransferType.acExport,  # exportar
                SpreadsheetType,
                obj_name,                  # tabela/consulta
                str(out_xlsx),             # destino
                True,                      # HasFieldNames
                f"{sheet}!"                # Range (nome da aba)
            )

        # Exportar tabelas
        for name in tables:
            export_object(name)
            first = False

        # Exportar consultas
        for name in queries:
            export_object(name)
            first = False

        if not tables and not queries:
            raise RuntimeError("Nada para exportar. Informe --tables/--queries ou use --all-tables/--all-queries.")

    finally:
        try:
            app.CloseCurrentDatabase()
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass

# ---------- ROTA B: ODBC (pyodbc + pandas) ----------
def connect_access_odbc(accdb_path: Path):
    import pyodbc
    conn_str = (
        r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
        rf"Dbq={accdb_path};"
        r"ExtendedAnsiSQL=1;"
    )
    return pyodbc.connect(conn_str, autocommit=True)

def export_via_odbc(accdb_path: Path, out_xlsx: Path, table_names=None, query_names=None,
                    sql_list=None, overwrite=True):
    ensure_parent(out_xlsx)
    mode = "w"
    if out_xlsx.exists() and overwrite:
        out_xlsx.unlink()

    # Abrir conexão
    conn = connect_access_odbc(accdb_path)

    try:
        with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:

            # Exportar tabelas
            for tname in (table_names or []):
                sheet = sanitize_sheet_name(tname)
                df = pd.read_sql_query(f"SELECT * FROM [{tname}]", conn)
                df.to_excel(writer, index=False, sheet_name=sheet)

            # Exportar consultas (tratadas como SELECT * FROM [consulta])
            for qname in (query_names or []):
                sheet = sanitize_sheet_name(qname)
                df = pd.read_sql_query(f"SELECT * FROM [{qname}]", conn)
                df.to_excel(writer, index=False, sheet_name=sheet)

            # Exportar SQLs arbitrárias (cada uma vira uma aba)
            for i, sql in enumerate(sql_list or [], start=1):
                sheet = sanitize_sheet_name(f"SQL_{i}")
                df = pd.read_sql_query(sql, conn)
                df.to_excel(writer, index=False, sheet_name=sheet)

    finally:
        conn.close()

# ---------- CLI ----------
def parse_args():
    p = argparse.ArgumentParser(description="Exportar Access → Excel (.xlsx) sem perda de dados.")
    p.add_argument("--accdb", required=True, help="Caminho do banco Access (.accdb ou .mdb)")
    p.add_argument("--out", required=True, help="Caminho do Excel de saída (.xlsx)")
    p.add_argument("--tables", default="", help="Lista de tabelas separadas por vírgula")
    p.add_argument("--queries", default="", help="Lista de consultas separadas por vírgula")
    p.add_argument("--sql", action="append", help="SQL(s) arbitrária(s) para exportar (pode repetir a flag)")
    p.add_argument("--all-tables", action="store_true", help="Exportar TODAS as tabelas")
    p.add_argument("--all-queries", action="store_true", help="Exportar TODAS as consultas")
    p.add_argument("--use-com", action="store_true", help="Usar COM/Access (pywin32). Se ausente, usa ODBC.")
    p.add_argument("--no-overwrite", action="store_true", help="Não sobrescrever o arquivo de saída se já existir")
    return p.parse_args()

def main():
    args = parse_args()
    accdb_path = Path(args.accdb)
    out_xlsx = Path(args.out)
    if not accdb_path.exists():
        print(f"❌ Arquivo Access não encontrado: {accdb_path}")
        sys.exit(1)

    tables = [s.strip() for s in args.tables.split(",") if s.strip()]
    queries = [s.strip() for s in args.queries.split(",") if s.strip()]
    overwrite = not args.no_overwrite

    try:
        if args.use_com:
            export_via_com(
                accdb_path, out_xlsx,
                table_names=tables,
                query_names=queries,
                all_tables=args.all_tables,
                all_queries=args.all_queries,
                overwrite=overwrite
            )
        else:
            export_via_odbc(
                accdb_path, out_xlsx,
                table_names=tables,
                query_names=queries,
                sql_list=args.sql or [],
                overwrite=overwrite
            )
        print("✅ Exportação concluída.")
        print(f"➡️  Excel gerado em: {out_xlsx}")
    except Exception as e:
        print("❌ Erro na exportação:")
        print(e)
        sys.exit(2)

if __name__ == "__main__":
    main()