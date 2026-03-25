import pdfplumber
import re

PDF_FILE = r"C:\Fabio\Desenvolvimento\Varejo\Temperatura\Temp_eq_Thermoking.pdf"

print("="*80)
print("🔍 DIAGNÓSTICO DETALHADO DO PDF")
print("="*80)

with pdfplumber.open(PDF_FILE) as pdf:
    print(f"\n📄 Total de páginas: {len(pdf.pages)}\n")
    
    # Analisa apenas a primeira página em detalhes
    page = pdf.pages[0]
    
    print("\n" + "="*80)
    print("📋 MÉTODO 1: EXTRAÇÃO DE TABELAS")
    print("="*80)
    
    tables = page.extract_tables()
    print(f"\nTotal de tabelas encontradas: {len(tables)}")
    
    for table_idx, table in enumerate(tables, 1):
        print(f"\n--- TABELA {table_idx} ---")
        print(f"Dimensões: {len(table)} linhas x {len(table[0]) if table else 0} colunas")
        
        # Mostra as primeiras 10 linhas
        print("\nPrimeiras 10 linhas:")
        for row_idx, row in enumerate(table[:10]):
            print(f"  Linha {row_idx}: {row}")
        
        # Procura por "Discharge Display" no cabeçalho
        print("\n🔎 Procurando 'Discharge Display' no cabeçalho:")
        for row_idx, row in enumerate(table[:5]):
            for col_idx, cell in enumerate(row):
                if cell and 'discharge' in str(cell).lower():
                    print(f"  ✅ Encontrado na Linha {row_idx}, Coluna {col_idx}: '{cell}'")
    
    print("\n" + "="*80)
    print("📋 MÉTODO 2: EXTRAÇÃO DE TEXTO BRUTO")
    print("="*80)
    
    text = page.extract_text()
    
    if text:
        lines = text.split('\n')
        print(f"\nTotal de linhas de texto: {len(lines)}")
        
        print("\n🔎 Primeiras 30 linhas do texto:")
        for i, line in enumerate(lines[:30], 1):
            print(f"  {i:2d}: {line}")
        
        print("\n🔎 Procurando linhas com 'Discharge Display':")
        discharge_lines = [line for line in lines if 'discharge' in line.lower() and 'display' in line.lower()]
        
        if discharge_lines:
            print(f"\n✅ Encontradas {len(discharge_lines)} linhas com 'Discharge Display':")
            for i, line in enumerate(discharge_lines[:10], 1):
                print(f"  {i}: {line}")
        else:
            print("  ❌ Nenhuma linha encontrada com 'Discharge Display'")
        
        print("\n🔎 Procurando linhas com padrão de data (DD/MM/YYYY):")
        date_lines = [line for line in lines if re.search(r'\d{2}/\d{2}/\d{4}', line)]
        
        if date_lines:
            print(f"\n✅ Encontradas {len(date_lines)} linhas com datas:")
            for i, line in enumerate(date_lines[:10], 1):
                print(f"  {i}: {line}")
        else:
            print("  ❌ Nenhuma linha encontrada com padrão de data")
    
    print("\n" + "="*80)
    print("📋 MÉTODO 3: EXTRAÇÃO DE TEXTO COM LAYOUT")
    print("="*80)
    
    text_layout = page.extract_text(layout=True)
    
    if text_layout:
        lines_layout = text_layout.split('\n')
        print(f"\nTotal de linhas (com layout): {len(lines_layout)}")
        
        print("\n🔎 Primeiras 30 linhas (preservando espaçamento):")
        for i, line in enumerate(lines_layout[:30], 1):
            print(f"  {i:2d}: |{line}|")
    
    print("\n" + "="*80)
    print("📋 MÉTODO 4: ANÁLISE DE PALAVRAS-CHAVE")
    print("="*80)
    
    keywords = ['discharge', 'display', 'temperature', 'temp', 'date', 'time', 'hora', 'data']
    
    print("\n🔎 Buscando palavras-chave no texto completo:")
    for keyword in keywords:
        count = text.lower().count(keyword) if text else 0
        status = "✅" if count > 0 else "❌"
        print(f"  {status} '{keyword}': {count} ocorrências")
    
    print("\n" + "="*80)
    print("📋 MÉTODO 5: EXTRAÇÃO DE PALAVRAS (WORDS)")
    print("="*80)
    
    words = page.extract_words()
    print(f"\nTotal de palavras extraídas: {len(words)}")
    
    # Procura por "Discharge" e "Display"
    discharge_words = [w for w in words if 'discharge' in w['text'].lower()]
    display_words = [w for w in words if 'display' in w['text'].lower()]
    
    print(f"\n🔎 Palavras com 'Discharge': {len(discharge_words)}")
    for w in discharge_words[:5]:
        print(f"  Texto: '{w['text']}' | Posição: x={w['x0']:.1f}, y={w['top']:.1f}")
    
    print(f"\n🔎 Palavras com 'Display': {len(display_words)}")
    for w in display_words[:5]:
        print(f"  Texto: '{w['text']}' | Posição: x={w['x0']:.1f}, y={w['top']:.1f}")

print("\n" + "="*80)
print("✅ DIAGNÓSTICO CONCLUÍDO")
print("="*80)
print("\n💡 Próximos passos:")
print("   1. Analise a saída acima")
print("   2. Identifique onde está 'Discharge Display'")
print("   3. Compartilhe as linhas relevantes para ajuste do script")