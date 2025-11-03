import pyautogui
import time
import os

# Tempo de intervalo entre comandos
pyautogui.PAUSE = 1.2
time.sleep(3)
# Caminho do navegador Google Chrome no Windows
#chrome_path = r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"

# URL desejada
#url = "https://apps.docusign.com/send/home"

# 1) Abre o Chrome
#os.startfile(chrome_path)
#time.sleep(3)

# 2) Digita o endereço do DocuSign
#pyautogui.write(url)
#pyautogui.press("enter")
print("Clicar em Iniciar")
pyautogui.click(x=652, y=616)
time.sleep(1)
print("Clicar em Envelopes")
pyautogui.click(x=687, y=681)
time.sleep(1)
print("Clicar em Enviar Envelopes")
pyautogui.click(x=859, y=679)
time.sleep(2)
print("Clicar em UPLOAD")
pyautogui.click(x=859, y=679)
time.sleep(3)
print("Clicar em Procurar")
pyautogui.click(x=701, y=455)
print("informar o caminho na pasta")
time.sleep(3)
pyautogui.click(x=404, y=45)

# Caminho do navegador Google Chrome no Windows
path_01 = r"C:\Fabio\CARREFOUR\PGR\PGR_TB\PGR_TB"
pyautogui.write(path_01)
pyautogui.press("enter")
print("Selecionar o arquivo")
pyautogui.click(x=596, y=137)
time.sleep(0.870)
print("Clicar em OK")
pyautogui.click(x=1183, y=618)
pyautogui.scroll(5)
print("Verificar.............")
time.sleep(5)

pyautogui.click(x=360, y=469)
end_fabio = r"Fabio Alves Cordeiro"
pyautogui.write(end_fabio)
pyautogui.click(x=333, y=594)
e_fabio = r"Fabio_Cordeiro@carrefour.com"
pyautogui.write(e_fabio)
pyautogui.scroll(10)
time.sleep(1)