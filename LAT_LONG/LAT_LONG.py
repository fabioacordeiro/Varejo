#pip install pandas geopy openpyxl

import pandas as pd
from geopy.geocoders import Nominatim

# Configurar o nome do arquivo e a aba a ser lida
excel_file = 'BASE_LOJAS.xlsx'
sheet_name = 'BASE_LOJAS'

# Carregar a planilha
df = pd.read_excel(excel_file, sheet_name=sheet_name)

# Inicializar o geolocalizador
geolocator = Nominatim(user_agent="geoapi", timeout=15)

# Função para obter latitude e longitude
def get_lat_long(address):
    try:
        location = geolocator.geocode(address)
        if location:
            return location.latitude, location.longitude
        else:
            return None, None
    except Exception as e:
        print(f"Erro ao buscar o endereço {address}: {e}")
        return None, None

# Adicionar colunas de latitude e longitude no DataFrame
df['Latitude'], df['Longitude'] = zip(*df['ENDEREÇO'].apply(get_lat_long))

# Exibir os resultados
print(df)

# Salvar a planilha com as novas colunas
df.to_excel('enderecos_com_coordenadas.xlsx', index=False)
print("Planilha salva como 'enderecos_com_coordenadas.xlsx'")