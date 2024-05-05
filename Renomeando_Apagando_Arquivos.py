# os + shutil - Copiando arquivos com Python
# Vamos copiar arquivos de uma pasta para outra.
# Copiar -> shutil.copy
import os
import shutil
import time

HOME = os.path.expanduser('~')

# DESKTOP = os.path.join(HOME, 'Desktop')
# PASTA_ORIGINAL = os.path.join(DESKTOP, 'EXEMPLO')
# NOVA_PASTA = os.path.join(DESKTOP, 'NOVA_PASTA')
PASTA_ORIGINAL = ('C:\\Users\\User\\Downloads')
NOVA_PASTA = ('C:\\Fabio\\CARREFOUR\\Automacao1')
print(PASTA_ORIGINAL)
print(NOVA_PASTA)
time.sleep(0.5)
os.makedirs(NOVA_PASTA, exist_ok=True)

for root, dirs, files in os.walk(PASTA_ORIGINAL):
    for dir_ in dirs:
        for file in files:
                #Relatório_de_Cargas_10-10-2023_11-59.csv
                name = file
                if 'Relatório_de_Cargas_' in name:
                    os.rename
                    print(file)



'''
for root, dirs, files in os.walk(PASTA_ORIGINAL):
    for dir_ in dirs:
        caminho_novo_diretorio = os.path.join(
            root.replace(PASTA_ORIGINAL, NOVA_PASTA), dir_
        )
        os.makedirs(caminho_novo_diretorio, exist_ok=True)

    for file in files:
        caminho_arquivo = os.path.join(root, file)
        caminho_novo_arquivo = os.path.join(
            root.replace(PASTA_ORIGINAL, NOVA_PASTA), file
        )
        shutil.copy(caminho_arquivo, caminho_novo_arquivo)
'''