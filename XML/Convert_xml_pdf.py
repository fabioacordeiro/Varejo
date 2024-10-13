# pip install fpdf
import xml.etree.ElementTree as ET
from fpdf import FPDF

# Função para ler o arquivo XML e extrair informações
def ler_xml_nfe('29240897422620014887553010001217531250113585-procNFe.xml'):
    tree = ET.parse('C:\\Fabio\\Desenvolvimento\\Varejo\\XML')
    root = tree.getroot()

    # Namespace do XML da NF-e
    ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

    # Extraindo informações básicas
    ide = root.find('.//nfe:ide', ns)
    emit = root.find('.//nfe:emit', ns)
    dest = root.find('.//nfe:dest', ns)
    total = root.find('.//nfe:total', ns)
    
    data = {
        'numero_nfe': ide.find('nfe:nNF', ns).text,
        'data_emissao': ide.find('nfe:dhEmi', ns).text,
        'emitente': emit.find('nfe:xNome', ns).text,
        'cnpj_emitente': emit.find('nfe:CNPJ', ns).text,
        'destinatario': dest.find('nfe:xNome', ns).text,
        'cnpj_destinatario': dest.find('nfe:CNPJ', ns).text,
        'valor_total': total.find('.//nfe:vNF', ns).text,
    }

    return data

# Função para criar o PDF
def criar_pdf_nfe(data, output_path):
    pdf = FPDF()
    pdf.add_page()

    # Título
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(200, 10, "Nota Fiscal Eletrônica (NF-e)", ln=True, align='C')

    # Adicionando informações da NF-e
    pdf.set_font("Arial", size=12)
    pdf.ln(10)
    pdf.cell(200, 10, f"Número da NF-e: {data['numero_nfe']}", ln=True)
    pdf.cell(200, 10, f"Data de Emissão: {data['data_emissao']}", ln=True)
    pdf.cell(200, 10, f"Emitente: {data['emitente']} (CNPJ: {data['cnpj_emitente']})", ln=True)
    pdf.cell(200, 10, f"Destinatário: {data['destinatario']} (CNPJ: {data['cnpj_destinatario']})", ln=True)
    pdf.cell(200, 10, f"Valor Total: R$ {data['valor_total']}", ln=True)

    # Salvar o PDF
    pdf.output(output_path)

# Caminho do arquivo XML
xml_path = 'C:\\Fabio\\Desenvolvimento\\Varejo\\XML\\29240897422620014887553010001217531250113585-procNFe.xml'
# Caminho do arquivo PDF de saída
pdf_path = 'nfe_output.pdf'

# Extrair dados do XML
dados_nfe = ler_xml_nfe(xml_path)
# Criar o PDF
criar_pdf_nfe(dados_nfe, pdf_path)

print(f"PDF gerado com sucesso em: {pdf_path}")