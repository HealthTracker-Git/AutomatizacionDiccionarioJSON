#%%
import pandas as pd
import numpy as np
# Leer archivo .xlsx

#%% ver tabla excel A01
file_path = 'DICCIONARIO_SERIE_A_2009.xlsx'
sheet_name = 'A01'

# Usar openpyxl como motor para .xlsx
df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')

# Mostrar las primeras filas del DataFrame
print(df.head())

#%%


file_path = 'DICCIONARIO_SERIE_A_2009.xlsx'

# Listar los nombres de las hojas
with pd.ExcelFile(file_path, engine='openpyxl') as xls:
    print(xls.sheet_names)

# %%
# Cargar el archivo Excel
file_path = 'DICCIONARIO_SERIE_A_2009.xlsx'
df = pd.read_excel(file_path, sheet_name='A02', header=None)

# Buscar las filas que contienen la información clave
servicio_salud_row = df[df[1] == "SERVICIO DE SALUD"].index[0]
comuna_row = df[df[1] == "COMUNA:  - (  )"].index[0]
establecimiento_row = df[df[1] == "ESTABLECIMIENTO:  - (  )"].index[0]
mes_row = df[df[1] == "MES:  - (  )"].index[0]
ano_row = df[df[1] == "AÑO: 2009"].index[0]
seccion_title_row = df[df[1] == "CONTROLES DE SALUD"].index[0]
seccion_name_row = df[df[1].str.contains("SECCIÓN A: CONTROLES DE SALUD DE LA MUJER", na=False)].index[0]

# Extracción de la información de la tabla
table_start_row = seccion_name_row + 2  # La tabla empieza 2 filas después del nombre de la sección

# Extraer la tabla relevante
table = df.iloc[table_start_row:]

# Filtrar las columnas para obtener solo las de interés
columns_of_interest = ['TIPO DE CONTROL', 'PROFESIONAL', 'TOTAL', 'BENEFICIARIOS', 'SEXO']
table_filtered = table[[0, 1, 2, 4, 5]]  # Dependiendo de cómo estén organizadas las columnas en tu Excel

# Mostrar la tabla filtrada
print(table_filtered)
# %%
