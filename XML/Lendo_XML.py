import xml.etree.ElementTree as ET
tree = ET.parse('Exemplo.xml')
root = tree.getroot()
tree.write("Exemplo1.xml")
#print(ET.tostring(root))
#imprimir todas os campos encontrados como Chave
print('='*80)
tree = ET.parse('Exemplo1.xml')
root = tree.getroot()
for desc in tree.findall(".//descricao"):
    print(desc.text)

for id, desc in enumerate(tree.findall(".//Id")):
    desc.set('id', str(id))

tree.write('Exemplo_2.xml')
descricao = tree.find('descricao[@id="0"]').text
descricao
print('Fim')


