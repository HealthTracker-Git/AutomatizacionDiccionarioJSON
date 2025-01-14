import pandas as pd

file_path = 'DICCIONARIO_SERIE_A_2009.xls'
df = pd.read_excel(file_path, sheet_name='AO2', engine='xlrd')  # Cambia 'AO2' por el nombre de tu hoja

# Imprimir los datos
print(df)
