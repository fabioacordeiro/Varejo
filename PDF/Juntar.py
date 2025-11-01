# pip install PyPDF2
import time
from pathlib import Path
from PyPDF2 import PdfMerger

def main():
    inicio = time.time()
    nome_programa = "Juntar PDF"
    print("_"*50)
    print(f"{nome_programa:^50}")
    print("_"*50)

    # Pastas (ajuste se necessário)
    PASTA_ORIGINAIS = Path(r"C:\Fabio\Desenvolvimento\Varejo\PDF\Juntar_pdf\PDFS_ORIGINAIS")
    PASTA_ALTERADOS = Path(r"C:\Fabio\Desenvolvimento\Varejo\PDF\Juntar_pdf\PDFS_ALTERADOS")
    PASTA_ALTERADOS.mkdir(parents=True, exist_ok=True)

    print(f"Caminho PDFS_ORIGINAIS: {PASTA_ORIGINAIS}")
    print(f"Caminho PDFS_ALTERADOS: {PASTA_ALTERADOS}")

    # Coletar PDFs (ordenados por nome) e limitar a 4
    pdfs = sorted(PASTA_ORIGINAIS.glob("*.pdf"), key=lambda p: p.name.casefold())
    if not pdfs:
        print("Nenhum PDF encontrado em PDFS_ORIGINAIS.")
        return

    pdfs = pdfs[:80]  # pegar no máximo 4 arquivos
    print("Arquivos selecionados para juntar (na ordem):")
    for p in pdfs:
        print(" -", p.name)

    saida = PASTA_ALTERADOS / "Juntar.pdf"
    with PdfMerger() as merger:
        for p in pdfs:
            merger.append(str(p))  # aceita caminho de arquivo
        merger.write(str(saida))

    print(f"OK! Arquivo gerado: {saida}")
    dur = time.time() - inicio
    h, rem = divmod(int(dur), 3600)
    m, s = divmod(rem, 60)
    print(f"Tempo total: {h:02d}:{m:02d}:{s:02d}")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("ERRO:", e.__class__.__name__, "-", e)