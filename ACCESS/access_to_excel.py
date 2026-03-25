# -*- coding: utf-8 -*-
# como usar
# utilize o prompt abaixo: Copie a linha inteira meno "#"
# python access_to_excel.py --accdb "C:\\Fabio\\CARREFOUR\\Cadastro de Fornecedores Oficial.accdb" --out "C:\\Fabio\\Desenvolvimento\\Varejo\\Access\\Plan_transportadoras.xlsx"

"""
Exporta tabelas/consultas do Access para um único Excel (.xlsx), 1 aba por objeto.
- Rota A (COM/Access) com pywin32: máxima fidelidade (TransferSpreadsheet).
- Rota B (ODBC) com pyodbc+pandas: robusto e sem o Access aberto.

USO (exemplos):
1) Exportar TABELAS específicas via COM (recomendado):
   python access_to_excel.py --accdb "C:\\Fabio\\CARREFOUR\\Cadastro de Fornecedores Oficial.accdb" --out "C:\\Fabio\\Desenvolvimento\\Varejo\\Access\\export.xlsx" --use-com --tables "CADASTRO FORNECEDORES, Tabela1"
   python access_to_excel.py --accdb "C:\\Fabio\\CARREFOUR\\Cadastro de Fornecedores Oficial.accdb" --out "C:\\Fabio\\Desenvolvimento\\Varejo\\Access\\Plan_transportadoras.xlsx"
2) Exportar TODAS as tabelas via COM:
   python access_to_excel.py --accdb "C:\dados\base.accdb" --out "C:\dados\export.xlsx" --use-com --all-tables

3) Exportar objetos (tabelas/consultas) via ODBC:
   python access_to_excel.py --accdb "C:\dados\base.accdb" --out "C:\dados\export.xlsx" --tables "Tabela1" --queries "MinhaConsulta"

4) Exportar uma CONSULTA via ODBC forçando SQL:
   python access_to_excel.py --accdb "C:\dados\base.accdb" --out "C:\dados\export.xlsx" --sql "SELECT * FROM [MinhaConsulta]"
"""

# -*- coding: utf-8 -*-
# -*- coding: utf-8 -*-
# -*- coding: utf-8 -*-
import argparse
import sys
from pathlib import Path
import win32com.client as win32

EXCEL_TYPE_XLSX = 10  # Excel12Xml (.xlsx)
AC_EXPORT = 1         # acExport

def sanitize_sheet_name_base(name: str) -> str:
    """
    Padrão B: troca espaços por '_', remove '$' e caracteres inválidos, limita a 31.
    """
    name = str(name).replace("$", "").replace(" ", "_")
    invalid = '[]:*?/\\'
    name = ''.join(ch for ch in name if ch not in invalid).strip()
    return name[:31] if name else "Sheet"

def uniquify_sheet_name(base: str, used: set) -> str:
    base = sanitize_sheet_name_base(base)
    if base not in used:
        used.add(base)
        return base
    i = 2
    while True:
        suffix = f"_{i}"
        cand = (base[:31 - len(suffix)] + suffix) if len(base) + len(suffix) > 31 else base + suffix
        if cand not in used:
            used.add(cand)
            return cand
        i += 1

def export_all_tables_and_queries(accdb_path: Path, out_xlsx: Path, overwrite: bool = True):
    # Preparar destino
    if overwrite and out_xlsx.exists():
        out_xlsx.unlink()
    out_xlsx.parent.mkdir(parents=True, exist_ok=True)

    # Inicia Access (sem mexer em .Visible para evitar erro DAO360)
    app = win32.Dispatch("Access.Application")
    app.OpenCurrentDatabase(str(accdb_path))

    try:
        db = app.CurrentDb()

        # Tabelas (ignora MSys*)
        tables = [tdef.Name for tdef in db.TableDefs if not str(tdef.Name).startswith("MSys")]
        # Consultas (ignora internas)
        queries = [qdef.Name for qdef in db.QueryDefs
                   if not str(qdef.Name).startswith("~") and not str(qdef.Name).startswith("MSys")]

        if not tables and not queries:
            raise RuntimeError("Nenhuma tabela/consulta encontrada no banco.")

        used_sheets = set()

        def export_object(obj_name: str, prefix: str = ""):
            base = f"{prefix}{obj_name}" if prefix else obj_name
            sheet = uniquify_sheet_name(base, used_sheets)

            # 1ª tentativa: Range com "Aba!"
            try:
                app.DoCmd.TransferSpreadsheet(
                    TransferType=AC_EXPORT,
                    SpreadsheetType=EXCEL_TYPE_XLSX,
                    TableName=obj_name,
                    FileName=str(out_xlsx),
                    HasFieldNames=True,
                    Range=f"{sheet}!"
                )
                return
            except Exception:
                # 2ª: Range sem "!"
                try:
                    app.DoCmd.TransferSpreadsheet(
                        TransferType=AC_EXPORT,
                        SpreadsheetType=EXCEL_TYPE_XLSX,
                        TableName=obj_name,
                        FileName=str(out_xlsx),
                        HasFieldNames=True,
                        Range=sheet
                    )
                    return
                except Exception:
                    # 3ª: nome ainda mais curto e “limpo”
                    mini = sanitize_sheet_name_base(sheet.replace("_", ""))
                    mini = (mini[:28] + "_x") if len(mini) > 28 else (mini if mini else "Sh")
                    app.DoCmd.TransferSpreadsheet(
                        TransferType=AC_EXPORT,
                        SpreadsheetType=EXCEL_TYPE_XLSX,
                        TableName=obj_name,
                        FileName=str(out_xlsx),
                        HasFieldNames=True,
                        Range=f"{mini}!"
                    )

        # Exporta tabelas (sem prefixo)
        for t in tables:
            export_object(t, prefix="")

        # Exporta consultas (prefixo para diferenciar)
        for q in queries:
            export_object(q, prefix="Q_")

    finally:
        try:
            app.CloseCurrentDatabase()
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass

def main():
    parser = argparse.ArgumentParser(
        description="Exporta TODAS as TABELAS e CONSULTAS do Access para um único .xlsx (1 aba por objeto)."
    )
    parser.add_argument("--accdb", required=True, help="Caminho do arquivo .accdb/.mdb")
    parser.add_argument("--out",   required=True, help="Caminho do .xlsx de saída")
    args = parser.parse_args()

    accdb = Path(args.accdb)
    out_xlsx = Path(args.out)

    if not accdb.exists():
        print(f"❌ Arquivo Access não encontrado: {accdb}")
        sys.exit(1)

    try:
        export_all_tables_and_queries(accdb, out_xlsx, overwrite=True)
        print("✅ Exportação concluída com sucesso.")
        print(f"➡️  Excel gerado em: {out_xlsx}")
    except Exception as e:
        print("❌ Erro na exportação:")
        print(e)
        sys.exit(2)

if __name__ == "__main__":
    main()