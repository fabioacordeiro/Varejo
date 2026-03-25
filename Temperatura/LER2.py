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


def preprocessar(img_bgr, ampliar=2):
    gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
    gray = cv2.GaussianBlur(gray, (3, 3), 0)
    gray = cv2.resize(gray, None, fx=ampliar, fy=ampliar, interpolation=cv2.INTER_CUBIC)
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
                "yc": (y1 + y2) / 2
            })
        except Exception:
            pass

    return itens


def recortar_colunas_interesse(img_bgr):
    """
    Recorta apenas:
    - coluna Data
    - coluna Temp. 1

    Coordenadas em proporção da página.
    Ajuste fino se necessário.
    """
    h, w = img_bgr.shape[:2]

    # área útil da tabela
    y1 = int(h * 0.20)
    y2 = int(h * 0.93)

    # coluna Data
    data_x1 = int(w * 0.06)
    data_x2 = int(w * 0.16)

    # coluna Temp. 1
    temp_x1 = int(w * 0.80)
    temp_x2 = int(w * 0.86)

    crop_data = img_bgr[y1:y2, data_x1:data_x2].copy()
    crop_temp = img_bgr[y1:y2, temp_x1:temp_x2].copy()

    return {
        "crop_data": crop_data,
        "crop_temp": crop_temp,
        "offset_y": y1
    }


def extrair_registros_data(itens_data):
    """
    Lê a coluna Data.
    Cada registro possui:
    - data
    - hora
    - UTC

    Mas só usamos data + hora.
    """
    if not itens_data:
        return []

    itens_data = sorted(itens_data, key=lambda z: z["yc"])

    datas = []
    horas = []

    for item in itens_data:
        txt = item["text"]

        m_data = re.search(r"\d{2}-\d{2}-\d{4}", txt)
        if m_data:
            datas.append({
                "valor": m_data.group(0),
                "y": item["yc"]
            })

        m_hora = re.search(r"\d{2}:\d{2}:\d{2}", txt)
        if m_hora:
            horas.append({
                "valor": m_hora.group(0),
                "y": item["yc"]
            })

    registros = []

    for d in datas:
        hora_encontrada = ""
        menor_dist = 999999

        for h in horas:
            dist = h["y"] - d["y"]
            if 0 < dist < 80 and dist < menor_dist:
                menor_dist = dist
                hora_encontrada = h["valor"]

        if hora_encontrada:
            registros.append({
                "Data": f"{d['valor']} {hora_encontrada}",
                "y_ref": d["y"]
            })

    return registros


def extrair_registros_temp(itens_temp):
    """
    Lê a coluna Temp. 1.
    Busca números negativos como -12, -13, -14 ou decimais.
    """
    if not itens_temp:
        return []

    itens_temp = sorted(itens_temp, key=lambda z: z["yc"])
    registros = []

    for item in itens_temp:
        txt = item["text"].replace(",", ".")
        m = re.search(r"-?\d+(?:\.\d+)?", txt)
        if m:
            try:
                valor = float(m.group(0))
                registros.append({
                    "Temp. 1": valor,
                    "y_ref": item["yc"]
                })
            except Exception:
                pass

    return registros


def casar_data_com_temp(reg_datas, reg_temps):
    """
    Casa cada data com a temperatura mais próxima verticalmente.
    """
    resultado = []

    for d in reg_datas:
        melhor_temp = None
        menor_dist = 999999

        for t in reg_temps:
            dist = abs(t["y_ref"] - d["y_ref"])
            if dist < 35 and dist < menor_dist:
                menor_dist = dist
                melhor_temp = t

        resultado.append({
            "Data": d["Data"],
            "Temp. 1": melhor_temp["Temp. 1"] if melhor_temp else None
        })

    return resultado


def processar_pdf(pdf_path):
    paginas = pdf_para_imagens(pdf_path)
    todos = []

    for num_pagina, img in paginas:
        print(f"Processando página {num_pagina}...")

        recortes = recortar_colunas_interesse(img)

        crop_data = preprocessar(recortes["crop_data"], ampliar=2)
        crop_temp = preprocessar(recortes["crop_temp"], ampliar=2)

        itens_data = ocr_com_boxes(crop_data)
        itens_temp = ocr_com_boxes(crop_temp)

        reg_datas = extrair_registros_data(itens_data)
        reg_temps = extrair_registros_temp(itens_temp)

        registros_pagina = casar_data_com_temp(reg_datas, reg_temps)

        for r in registros_pagina:
            r["Pagina"] = num_pagina
            todos.append(r)

    if not todos:
        raise Exception("Nenhum dado foi extraído do PDF.")

    df = pd.DataFrame(todos)

    df["Data"] = pd.to_datetime(df["Data"], format="%d-%m-%Y %H:%M:%S", errors="coerce")
    df = df[df["Data"].notna()].copy()

    df["Temp. 1"] = pd.to_numeric(df["Temp. 1"], errors="coerce")

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
    df_plot = df.dropna(subset=["Temp. 1"]).copy()

    if df_plot.empty:
        print("Aviso: não há valores de temperatura válidos para gerar o gráfico.")
        return

    plt.figure(figsize=(14, 6))
    plt.plot(df_plot["Data"], df_plot["Temp. 1"], marker="o")
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