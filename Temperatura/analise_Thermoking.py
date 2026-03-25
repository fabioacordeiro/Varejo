import pdfplumber

PDF_FILE = "C:\\Fabio\\Desenvolvimento\\Varejo\\Temperatura\\Temp_eq_Thermoking.pdf"

print("=== DIAGNÓSTICO DO PDF ===\n")

with pdfplumber.open(PDF_FILE) as pdf:
    print(f"Total de páginas: {len(pdf.pages)}\n")
    
    for page_num, page in enumerate(pdf.pages, 1):
        print(f"\n--- PÁGINA {page_num} ---")
        
        # Extrai texto bruto
        text = page.extract_text()
        print("\n[TEXTO BRUTO - Primeiras 500 caracteres]:")
        print(text[:500] if text else "Nenhum texto extraído")
        
        # Extrai tabelas
        tables = page.extract_tables()
        print(f"\n[TABELAS ENCONTRADAS]: {len(tables)}")
        
        for i, table in enumerate(tables, 1):
            print(f"\n  Tabela {i}:")
            print(f"  - Linhas: {len(table)}")
            print(f"  - Colunas: {len(table[0]) if table else 0}")
            
            # Mostra as primeiras 5 linhas
            print("\n  Primeiras 5 linhas:")
            for row_idx, row in enumerate(table[:5]):
                print(f"    Linha {row_idx}: {row}")
        
        if page_num == 1:  # Só primeira página para não poluir
            break

print("\n=== FIM DO DIAGNÓSTICO ===")