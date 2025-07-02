# pip install pytesseract
# pip install tesseract
# Desenvolvido por Fábio A Cordeiro em 11/04/2025

import pytesseract
from PIL import Image



# Aponta para o caminho onde o Tesseract foi instalado
pytesseract.pytesseract.tesseract_cmd = r'C:\\Fabio\\Desenvolvimento\\Varejo\\Extrair_textotesseract.exe'

# Caminho da imagem após reset
image_path = "C:\\Fabio\\Desenvolvimento\\Varejo\\Extrair_texto\\Imagem.jpeg"


# Abrir imagem e aplicar OCR
image = Image.open(image_path)
texto_extraido = pytesseract.image_to_string(image, lang="Ingles")

# Mostrar o texto extraído
texto_extraido