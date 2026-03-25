import os
import re
import fitz
import cv2
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from PIL import Image
from rapidocr_onnxruntime import RapidOCR

PDF_PATH = r"C:\\Fabio\\Desenvolvimento\\Varejo\\Temperatura\\Temp_sem_cabecalho.pdf"
XLSX_SAIDA = r"C:\\Fabio\\Desenvolvimento\\Varejo\\Temperatura\\temperaturas_extraidas.xlsx"
PNG_GRAFICO = r"C:\\Fabio\\Desenvolvimento\\Varejo\\Temperatura\\grafico_temperatura_novo.png"

ocr_engine = RapidOCR()


def normalizar_texto(txt):
    if txt is None:
        return ""
    txt = str(txt).replace("\n", " ").replace("\r", " ")
    txt = re.sub(r"\s+", " ", txt)
    return txt.strip()


def pdf_para_imagens(pdf_path, dpi=220):
    doc = fitz.open(pdf_path)
    imagens = []
    zoom = dpi / 72
    mat = fitz.Matrix(zoom, zoom)

    for i in range(len(doc)):
        page = doc.load_page(i)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img_cv = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
        imagens.append((i + 1, img_cv))

    doc.close()
    return imagens


def preprocessar_pagina(img_bgr):
    gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
    gray = cv2.GaussianBlur(gray, (3, 3), 0)
    gray = cv2.resize(gray, None, fx=1.5, fy=1.5, interpolation=cv2.INTER_CUBIC)
    _, th = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return cv2.cvtColor(th, cv2.COLOR_GRAY2BGR)


def ocr_com_boxes(img_bgr):
    result, _ = ocr_engine(img_bgr)
    itens = []

    if not result:
        return itens

    for item in result:
        try:
            box = item[0]
            text = normalizar_texto(item[1])

            if not text:
                continue

            xs = [p[0] for p in box]
            ys = [p[1] for p in box]

            x1, x2 = min(xs), max(xs)
            y1, y2 = min(ys), max(ys)

            itens.append({
                "text": text,
                "x1": x1,
                "x2": x2,
                "y1": y1,
                "y2": y2,
                "xc": (x1 + x2) / 2,
                "yc": (y1 + y2) / 2,
                "h": y2 - y1
            })
        except Exception:
            pass

    return itens


def localizar_header(itens):
    """
    Procura os rótulos das colunas 'Data' e 'Temp. 1'.
    """
    x_data = None
    x_temp1 = None
    y_header = None

    for item in itens:
        txt = item["text"].lower().replace(" ", "")
        if txt == "data":
            x_data = item["xc"]
            y_header = item["yc"]
        elif "temp.1" in txt or "temp1" in txt or "templ" in txt:
            x_temp1 = item["xc"]
            if y_header is None:
                y_header = item["yc"]

    if x_data is None or x_temp1 is None:
        raise Exception("Não foi possível localizar as colunas 'Data' e 'Temp. 1' no cabeçalho.")

    return x_data, x_temp1, y_header


def agrupar_blocos_em_registros(itens, y_header):
    """
    Agrupa os blocos abaixo do cabeçalho em registros.
    Como a data está em duas linhas, agrupamos por faixas verticais maiores.
    """
    dados = [i for i in itens if i["yc"] > y_header + 10]
    if not dados:
        return []

    dados = sorted(dados, key=lambda z: z["yc"])

    alturas = [i["h"] for i in dados if i["h"] > 0]
    altura_media = np.median(alturas) if alturas else 20
    tolerancia = max(18, int(altura_media * 1.8))

    grupos = []
    atual = [dados[0]]

    for item in dados[1:]:
        media_y = sum(x["yc"] for x in atual) / len(atual)
        if abs(item["yc"] - media_y) <= tolerancia:
            atual.append(item)
        else:
            grupos.append(atual)
            atual = [item]

    if atual:
        grupos.append(atual)

    return grupos


def extrair_data_do_grupo(grupo, x_data):
    """
    Procura textos próximos da coluna Data e monta:
    dd-mm-yyyy hh:mm:ss
    """
    faixa = [g for g in grupo if abs(g["xc"] - x_data) < 130]

    if not faixa:
        return ""

    faixa = sorted(faixa, key=lambda z: z["yc"])
    texto_total = " ".join(g["text"] for g in faixa)
    texto_total = normalizar_texto(texto_total)

    data_match = re.search(r"\d{2}-\d{2}-\d{4}", texto_total)
    hora_match = re.search(r"\d{2}:\d{2}:\d{2}", texto_total)

    if data_match and hora_match:
        return f"{data_match.group(0)} {hora_match.group(0)}"

    # fallback: procurar separadamente
    data_txt = ""
    hora_txt = ""

    for item in faixa:
        txt = item["text"]
        if not data_txt:
            m = re.search(r"\d{2}-\d{2}-\d{4}", txt)
            if m:
                data_txt = m.group(0)
        if not hora_txt:
            m = re.search(r"\d{2}:\d{2}:\d{2}", txt)
            if m:
                hora_txt = m.group(0)

    if data_txt and hora_txt:
        return f"{data_txt} {hora_txt}"

    return ""


def extrair_temp1_do_grupo(grupo, x_temp1):
    """
    Procura o valor da temperatura na faixa da coluna Temp. 1.
    Espera valores como -12, -13, -14 ou -12.5
    """
    faixa = [g for g in grupo if abs(g["xc"] - x_temp1) < 120]

    if not faixa:
        return None

    faixa = sorted(faixa, key=lambda z: (abs(z["xc"] - x_temp1), z["yc"]))

    for item in faixa:
        txt = item["text"].replace(",", ".")
        m = re.search(r"-?\d+(?:\.\d+)?", txt)
        if m:
            try:
                return float(m.group(0))
            except Exception:
                pass

    # fallback: pega último número da faixa
    texto_total = " ".join(g["text"] for g in faixa).replace(",", ".")
    nums = re.findall(r"-?\d+(?:\.\d+)?", texto_total)
    if nums:
        try:
            return float(nums[-1])
        except Exception:
            pass

    return None


def processar_pdf(pdf_path):
    paginas = pdf_para_imagens(pdf_path)
    registros = []

    for num_pagina, img in paginas:
        print(f"Processando página {num_pagina}...")

        img_proc = preprocessar_pagina(img)
        itens = ocr_com_boxes(img_proc)

        if not itens:
            print(f"Nenhum texto encontrado na página {num_pagina}.")
            continue

        x_data, x_temp1, y_header = localizar_header(itens)
        grupos = agrupar_blocos_em_registros(itens, y_header)

        for grupo in grupos:
            data = extrair_data_do_grupo(grupo, x_data)
            temp1 = extrair_temp1_do_grupo(grupo, x_temp1)

            # só guarda se tiver ao menos data ou temperatura
            if data or temp1 is not None:
                registros.append({
                    "Data": data,
                    "Temp. 1": temp1,
                    "Pagina": num_pagina
                })

    if not registros:
        raise Exception("Nenhum dado foi extraído do PDF.")

    df = pd.DataFrame(registros)

    # remove linhas ruins
    df["Data"] = df["Data"].astype(str).str.strip()
    df = df[(df["Data"] != "") | (df["Temp. 1"].notna())].copy()

    # converte Data
    df["Data"] = pd.to_datetime(df["Data"], format="%d-%m-%Y %H:%M:%S", errors="coerce")

    # remove linhas sem data válida
    df = df[df["Data"].notna()].copy()

    # ordena
    df = df.sort_values("Data").reset_index(drop=True)

    return df


def salvar_excel(df, caminho_saida):
    with pd.ExcelWriter(caminho_saida, engine="openpyxl", datetime_format="dd/mm/yyyy hh:mm:ss") as writer:
        df.to_excel(writer, sheet_name="Dados", index=False)
        ws = writer.sheets["Dados"]

        larguras = {
            "Data": 22,
            "Temp. 1": 12,
            "Pagina": 10
        }

        for idx, col in enumerate(df.columns, start=1):
            letra = ws.cell(row=1, column=idx).column_letter
            ws.column_dimensions[letra].width = larguras.get(col, 15)


def gerar_grafico(df, caminho_png):
    plt.figure(figsize=(14, 6))
    plt.plot(df["Data"], df["Temp. 1"], marker="o")
    plt.xlabel("Data e Hora")
    plt.ylabel("Temperatura")
    plt.title("Gráfico de Temperatura - Temp. 1")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig(caminho_png, dpi=150)
    plt.close()


if __name__ == "__main__":
    if not os.path.exists(PDF_PATH):
        raise FileNotFoundError(f"PDF não encontrado: {PDF_PATH}")

    df = processar_pdf(PDF_PATH)

    if df.empty:
        raise Exception("A extração terminou, mas não houve registros válidos.")

    salvar_excel(df, XLSX_SAIDA)
    gerar_grafico(df, PNG_GRAFICO)

    print(f"Excel gerado com sucesso: {XLSX_SAIDA}")
    print(f"Gráfico gerado com sucesso: {PNG_GRAFICO}")