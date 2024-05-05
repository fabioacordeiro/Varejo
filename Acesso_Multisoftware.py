#pip install paperclip
#pip install pyautogui
#pip install webbrowser

import pyautogui
import webbrowser
import pyperclip
from time import sleep
from datetime import datetime, timedelta, date
import os # pastas e caminhos
import shutil #para copiar arquivos

data_machine = date.today()
deltad = timedelta(days=30)
data_anterior = data_machine-deltad
dataf = datetime.strftime(data_machine, '%d/%m/%Y')
datai =  datetime.strftime(data_anterior, '%d/%m/%Y')

print(60*'-')
print(f'Data Atual sistema:{data_machine}')
print(f'Data anterior sistema:{data_anterior}')
print(f'Data Inicial formatada:{datai}')
print(f'Data Final formatada:{dataf}')
print('Automação de Processo - Torre de Controle')
print('Analista: Fábio Alves Cordeiro')
print('Acesso ao Multiembarcador')
print(60*'-')
sleep(1)
# Entrar no site Multiembarcador
webbrowser.open("https://grupocarrefour.multiembarcador.com.br/#Home")
# colar o endereço do gmail e dar um ENTER
sleep(28)

# Digitando user e senha
pyautogui.click(x=827, y=675)
pyautogui.write("Fabio_cordeiro")
pyautogui.click(x=837, y=683)
pyautogui.write("Fac@2023")
pyautogui.click(x=1053, y=534)
pyautogui.press("enter")

sleep(9)
#Acessando o menu do relatório
pyautogui.click(x=563, y=53)
pyautogui.write("https://grupocarrefour.multiembarcador.com.br/#Relatorios/Cargas/Carga")
pyautogui.press("enter")
#sleep(6)
#pyautogui.click(x=1276, y=380)
sleep(6)

print('Processo finalizado com sucesso!')
