import pdfplumber

pdf_path = r"C:\\Fabio\\Desenvolvimento\\Varejo\\Temperatura\\Temp_sem_cabecalho.pdf"

with pdfplumber.open(pdf_path) as pdf:
    for i, page in enumerate(pdf.pages[:3], start=1):
        text = page.extract_text()
        print(f"Página {i}:")
        print(repr(text[:500] if text else text))
        print("-" * 80)