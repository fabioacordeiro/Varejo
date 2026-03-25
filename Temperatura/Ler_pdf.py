import os
import re
import fitz
import cv2
import numpy as np
import pandas as pd
from PIL import Image
from rapidocr_onnxruntime import RapidOCR

PDF_PATH = r"C:\\Fabio\\Desenvolvimento\\Varejo\\Temperatura\\Temp_sem_cabecalho.pdf"
XLSX_SAIDA = r"C:\\Fabio\\Desenvolvimento\\Varejo\\Temperatura\\dados_extraidos.xlsx"

COLUNAS = [
    "Data",
    "Velocidade",
    "Ignição",
    "Bloqueio",
    "Bateria",
    "latitude",
    "longitude",
    "Estado",
    "Municipio",
    "Rua",
    "Bairro",
    "Numero",
    "Sateilite",
    "Memoria",
    "GPS",
    "Ponto Ref.",
    "Temp. 1",
    "Temp. 2",
    "Temp. 3",
]

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
    return cv2.cvtColor(gray, cv2.COLOR_GRAY2BGR)


def ocr_com_boxes(img_bgr):
    result, _ = ocr_engine(img_bgr)
    itens = []

    if not result:
        return itens

    for item in result:
        try:
            box = item[0]
            text = normalizar_texto(item[1])
            score = item[2] if len(item) > 2 else 0

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
                "score": score
            })
        except Exception:
            pass

    return itens


def agrupar_em_linhas(itens, tolerancia=18):
    if not itens:
        return []

    itens = sorted(itens, key=lambda z: z["yc"])
    linhas = []

    for item in itens:
        colocado = False
        for linha in linhas:
            media_y = sum(x["yc"] for x in linha) / len(linha)
            if abs(item["yc"] - media_y) <= tolerancia:
                linha.append(item)
                colocado = True
                break

        if not colocado:
            linhas.append([item])

    for linha in linhas:
        linha.sort(key=lambda z: z["x1"])

    linhas.sort(key=lambda grupo: sum(x["yc"] for x in grupo) / len(grupo))
    return linhas


def extrair_ancoras_header(linhas):
    """
    Localiza a linha do cabeçalho e guarda a posição X aproximada de cada coluna.
    """
    linha_header = None

    for linha in linhas:
        texto = " ".join(x["text"].lower() for x in linha)
        if "velocidade" in texto and ("ignição" in texto or "ignicao" in texto):
            linha_header = linha
            break

    if not linha_header:
        raise Exception("Cabeçalho da tabela não encontrado na página.")

    header_text = " ".join(x["text"] for x in linha_header)

    # Posições fixas estimadas pelo layout do PDF do exemplo
    # Caso mude um pouco, ainda tende a funcionar bem.
    ancoras = {
        "Data": 90,
        "Velocidade": 210,
        "Ignição": 265,
        "Bloqueio": 320,
        "Bateria": 380,
        "latitude": 450,
        "longitude": 560,
        "Estado": 650,
        "Municipio": 760,
        "Rua": 930,
        "Bairro": 1080,
        "Numero": 1170,
        "Sateilite": 1230,
        "Memoria": 1285,
        "GPS": 1340,
        "Ponto Ref.": 1405,
        "Temp. 1": 1485,
        "Temp. 2": 1540,
        "Temp. 3": 1595,
    }

    return linha_header, ancoras, header_text


def linha_eh_dado(linha):
    texto = " ".join(x["text"] for x in linha)
    return bool(re.search(r"\d{2}-\d{2}-\d{4}", texto))


def alocar_textos_nas_colunas(linha, ancoras):
    registro = {c: "" for c in COLUNAS}

    centros = [(col, x) for col, x in ancoras.items()]

    acumulado = {c: [] for c in COLUNAS}

    for item in linha:
        txt = item["text"]
        xc = item["xc"]

        melhor_col = min(centros, key=lambda p: abs(xc - p[1]))[0]
        acumulado[melhor_col].append(txt)

    for col in COLUNAS:
        registro[col] = normalizar_texto(" ".join(acumulado[col]))

    return registro


def ajustar_campo_data(registro):
    """
    Junta a data/hora e remove fragmentações do OCR.
    """
    texto = registro["Data"]
    texto = texto.replace("UTC—03", "UTC-03").replace("UTC -03", "UTC-03")
    texto = normalizar_texto(texto)

    m = re.search(r"(\d{2}-\d{2}-\d{4})\s+(\d{2}:\d{2}:\d{2})", texto)
    if m:
        registro["Data"] = f"{m.group(1)} {m.group(2)}"
    return registro


def normalizar_registro(reg):
    reg = {k: normalizar_texto(v) for k, v in reg.items()}

    reg = ajustar_campo_data(reg)

    mapa_sim_nao = {
        "Sim": "Sim",
        "SIm": "Sim",
        "SIM": "Sim",
        "Nao": "Não",
        "NAO": "Não",
        "Não": "Não",
    }

    for col in ["Ignição", "Bloqueio", "Sateilite", "Memoria", "GPS"]:
        if reg[col] in mapa_sim_nao:
            reg[col] = mapa_sim_nao[reg[col]]

    if reg["Estado"]:
        reg["Estado"] = reg["Estado"].replace("5P", "SP").replace("Sp", "SP").upper()

    for col in ["latitude", "longitude", "Temp. 1", "Temp. 2", "Temp. 3"]:
        reg[col] = reg[col].replace(",", ".").replace(" ", "")

    reg["Numero"] = re.sub(r"[^\d]", "", reg["Numero"])

    return reg


def converter_tipos(df):
    if "Data" in df.columns:
        df["Data"] = pd.to_datetime(df["Data"], format="%d-%m-%Y %H:%M:%S", errors="coerce")

    for col in ["Velocidade", "Bateria", "latitude", "longitude", "Numero", "Temp. 1", "Temp. 2", "Temp. 3"]:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.replace(",", ".", regex=False)
                .replace("", pd.NA)
            )
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


def ajustar_largura_colunas(ws, df):
    larguras = {
        "Data": 22,
        "Velocidade": 12,
        "Ignição": 12,
        "Bloqueio": 12,
        "Bateria": 10,
        "latitude": 14,
        "longitude": 14,
        "Estado": 10,
        "Municipio": 22,
        "Rua": 28,
        "Bairro": 20,
        "Numero": 10,
        "Sateilite": 12,
        "Memoria": 12,
        "GPS": 10,
        "Ponto Ref.": 20,
        "Temp. 1": 10,
        "Temp. 2": 10,
        "Temp. 3": 10,
        "Pagina": 10,
    }

    for idx, col in enumerate(df.columns, start=1):
        letra = ws.cell(row=1, column=idx).column_letter
        ws.column_dimensions[letra].width = larguras.get(col, 15)


def processar_pdf(pdf_path, xlsx_saida):
    paginas = pdf_para_imagens(pdf_path)
    registros = []

    for num_pagina, img in paginas:
        print(f"Processando página {num_pagina}...")

        img_proc = preprocessar_pagina(img)
        itens = ocr_com_boxes(img_proc)

        if not itens:
            print(f"Nenhum texto encontrado na página {num_pagina}.")
            continue

        linhas = agrupar_em_linhas(itens, tolerancia=18)
        _, ancoras, _ = extrair_ancoras_header(linhas)

        for linha in linhas:
            if not linha_eh_dado(linha):
                continue

            reg = alocar_textos_nas_colunas(linha, ancoras)
            reg = normalizar_registro(reg)

            if reg["Data"] or (reg["latitude"] and reg["longitude"]):
                reg["Pagina"] = num_pagina
                registros.append(reg)

    if not registros:
        raise Exception("Nenhum registro foi extraído do PDF.")

    df = pd.DataFrame(registros)

    for col in COLUNAS:
        if col not in df.columns:
            df[col] = ""

    df = df[COLUNAS + ["Pagina"]]
    df = converter_tipos(df)

    with pd.ExcelWriter(xlsx_saida, engine="openpyxl", datetime_format="dd/mm/yyyy hh:mm:ss") as writer:
        df.to_excel(writer, sheet_name="Dados", index=False)
        ws = writer.sheets["Dados"]
        ajustar_largura_colunas(ws, df)

    print(f"Planilha gerada com sucesso: {xlsx_saida}")


if __name__ == "__main__":
    if not os.path.exists(PDF_PATH):
        raise FileNotFoundError(f"PDF não encontrado: {PDF_PATH}")

    processar_pdf(PDF_PATH, XLSX_SAIDA)