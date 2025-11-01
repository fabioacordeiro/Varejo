#Todo dia acessando os mesmos endereços de browser
#pip install time

import webbrowser
import time


# Caminho do Google Chrome (ajuste conforme sua instalação)
chrome_path = "C:/Program Files/Google/Chrome/Application/chrome.exe %s"

# Lista de URLs
urls = [
    "https://mail.google.com/mail/u/0/#inbox",
    "https://drive.google.com/drive/my-drive",
    "https://www.docusign.com/pt-br/logout",
    "https://docs.google.com/spreadsheets/d/1A0ucvV0AB4-evtQI2atCe5oObwjsfE0Jwc_JJWiP0Xs/edit?gid=0#gid=0",
    "https://url.de.m.mimecastprotect.com/s/a0oRC79o2xumQV3XqC8fZioLZzZ?domain=1drv.ms",
    "https://carrefour.elaw.com.br/logout.elaw",
    "https://grupocarrefour.multiembarcador.com.br/#Home",
    "https://TB.multitms.com.br/#home",
    "https://br2.brasilrisk.com.br/Account/Login?ReturnUrl=%2fHome%2fListar",
    "https://account.docusign.com/oauth/auth/?login_hint=fabio_cordeiro%40carrefour.com"
]

# Definir Chrome como navegador
webbrowser.register('chrome', None, webbrowser.BackgroundBrowser("C:/Program Files/Google/Chrome/Application/chrome.exe"))

# Abrir cada URL em uma nova aba
for url in urls:
    webbrowser.get('chrome').open_new_tab(url)
    time.sleep(1)  # pequeno intervalo para não sobrecarregar



