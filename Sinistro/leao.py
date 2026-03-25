# -*- coding: utf-8 -*-
from PIL import Image, ImageDraw, ImageFont
import os

img_path = r"c:\\Fabio\\Desenvolvimento\\Varejo\\Sinistro\\leao.png"
out_path = r"c:\\Fabio\\Desenvolvimento\\Varejo\\Sinistro\\leao4.png"

img = Image.open(img_path).convert("RGBA")
w, h = img.size
draw = ImageDraw.Draw(img)

# Função para localizar uma fonte com suporte a acentuação no Windows
def carregar_fonte(tamanho=24, negrito=False):
    fontes_teste = []

    if negrito:
        fontes_teste = [
            r"C:\Windows\Fonts\arialbd.ttf",
            r"C:\Windows\Fonts\segoeuib.ttf",
            r"C:\Windows\Fonts\calibrib.ttf",
            "DejaVuSans-Bold.ttf",
        ]
    else:
        fontes_teste = [
            r"C:\Windows\Fonts\arial.ttf",
            r"C:\Windows\Fonts\segoeui.ttf",
            r"C:\Windows\Fonts\calibri.ttf",
            "DejaVuSans.ttf",
        ]

    for caminho in fontes_teste:
        try:
            return ImageFont.truetype(caminho, tamanho)
        except:
            continue

    return ImageFont.load_default()

font_title = carregar_fonte(18, negrito=True)
font_text = carregar_fonte(14, negrito=False)

# Dados dos leões
lion = [
    ("Leão Asiático", "Peso: 160–190 kg\nCompr.: ~2,7 m\nOrigem: Índia (Gir)"),
    ("Leão Africano", "Peso: 190–250 kg\nCompr.: ~3,0 m\nOrigem: África Subsaariana"),
    ("Leão do Atlas", "Peso: até 270 kg\nCompr.: ~3,3 m\nOrigem: Norte da África"),
    ("Leão das Cavernas", "Peso: 300–350 kg\nCompr.: ~3,5 m\nOrigem: Europa/Ásia"),
    ("Leão Americano", "Peso: 350–420 kg\nCompr.: ~3,7 m\nOrigem: América do Norte"),
    ("Panthera fossilis", "Peso: 250–320 kg\nCompr.: ~3,3 m\nOrigem: Europa"),
]

# Mantém a ordem correta dos leões
lions = lion

# Posições aproximadas
positions = [int(w * p) for p in [0.09, 0.23, 0.37, 0.52, 0.67, 0.82]]
y = int(h * 0.05)

for (title, text), x in zip(lions, positions):
    box_w = 300
    box_h = 120
    x0 = x - box_w // 2
    y0 = y
    x1 = x + box_w // 2
    y1 = y + box_h

    # evita cortar caixa fora da imagem
    if x0 < 0:
        x1 += abs(x0)
        x0 = 0
    if x1 > w:
        excesso = x1 - w
        x0 -= excesso
        x1 = w

    draw.rectangle(
        [x0, y0, x1, y1],
        outline=(0, 0, 0, 255),
        width=2,
        fill=(255, 255, 255, 230)
    )

    draw.text((x0 + 10, y0 + 6), title, fill=(0, 0, 0, 255), font=font_title)
    draw.multiline_text(
        (x0 + 10, y0 + 42),
        text,
        fill=(0, 0, 0, 255),
        font=font_text,
        spacing=4
    )

img.save(out_path)
print(f"Imagem salva em: {out_path}")