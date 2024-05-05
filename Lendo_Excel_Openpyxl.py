# openpyxl para arquivos Excel xlsx, xlsm, xltx e xltm (instalação)
# Com essa biblioteca será possível ler e escrever dados em células
# específicas, formatar células, inserir gráficos,
# criar fórmulas, adicionar imagens e outros elementos gráficos às suas
# planilhas. Ela é útil para automatizar tarefas envolvendo planilhas do
# Excel, como a criação de relatórios e análise de dados e/ou facilitando a
# manipulação de grandes quantidades de informações.
# Instalação necessária: pip install openpyxl
# Documentação: https://openpyxl.readthedocs.io/en/stable/

# pip install openpyxl

from pathlib import Path

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from openpyxl import Workbook, load_workbook

#Abrindo a planilha DDR.xlsx.
planilha = load_workbook("DDR.xlsx")

aba_ativa = planilha.active

ver = 'Normal'
#A programação abaixo o python vai ler cada célula da coluna"i"
#e verificar se é nula, se for vai inserir na planilha na célula a 
# palavra Normal e imprimir na tela a palavra Normal se não for vai
# ignorar a célula.
for celula in aba_ativa["I"]:
    if  (celula.value) is None:
        celula.value = ver
        print(ver)
	
#Salvando em uma nova planilha.
planilha.save("DDR1.xlsx")

print('Arquivo criado e Salvo')
