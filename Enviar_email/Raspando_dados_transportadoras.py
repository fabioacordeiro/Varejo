#pip install webdriver-manager
#pip install ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import pandas as pd
# Configuração do WebDriver
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)

# Função para coletar dados de uma única página
def coletar_dados_da_pagina():
    transportadoras = []
    # Localize os elementos que contêm os dados das transportadoras
    cards = driver.find_elements(By.CLASS_NAME, "box in partialResults track by $index")  # Substitua "element-class" pela classe correta
    for card in cards:
        nome = card.find_element(By.CLASS_NAME, "m-boxCompany__A__info__name").text  # Substitua "nome-class" pela classe correta
        print(type(nome))
        endereco = card.find_element(By.CLASS_NAME, "Whatsapp").text  # Substitua "endereco-class" pela classe correta
        print(type(endereco))
        telefone = card.find_element(By.CLASS_NAME, "Whatsapp").text  # Substitua "telefone-class" pela classe correta
        print(type(telefone))
        celular = card.find_element(By.CLASS_NAME, "Whatsapp").text  # Substitua "telefone-class" pela classe correta
        print(type(celular))
        email = card.find_element(By.CLASS_NAME, "E-mail").text  # Substitua "telefone-class" pela classe correta
        print(type(email))
        transportadoras.append({
            'Nome': nome,
            'Endereço': endereco,
            'Telefone': telefone,
            'Celular':celular,
            'Email':email
        })
    return transportadoras

# Acessar o site
url = "https://www.transvias.com.br/transportadoras/estados/sao-paulo"
driver.get(url)
time.sleep(3)  # Esperar carregar a página

todas_transportadoras = []

# Coletar dados da primeira página
todas_transportadoras.extend(coletar_dados_da_pagina())

# Navegar pelas páginas e coletar dados
while True:
    try:
        next_button = driver.find_element(By.LINK_TEXT, "box in partialResults track by $index")  # Substitua pelo texto do botão de próxima página
        next_button.click()
        time.sleep(3)  # Esperar carregar a próxima página
        todas_transportadoras.extend(coletar_dados_da_pagina())
        
        
        
        
    except:
        break  # Sai do loop quando não encontrar mais o botão de próxima página

# Fechar o WebDriver
driver.quit()

# Converter os dados em um DataFrame do Pandas
df = pd.DataFrame(todas_transportadoras)

# Gravar os dados em uma planilha Excel
df.to_excel("transportadoras.xlsx", index=False)

print("Dados gravados com sucesso em 'transportadoras.xlsx'")