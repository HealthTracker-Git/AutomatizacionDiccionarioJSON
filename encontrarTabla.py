#%%
import pandas as pd

def get_value_from_position(df, row_idx, col_idx):
    """
    Obtiene el valor de una celda específica en un DataFrame dado su índice de fila y columna.
    
    Parámetros:
    df (pd.DataFrame): El DataFrame del cual obtener el valor.
    row_idx (int): El índice de la fila.
    col_idx (int): El índice de la columna.
    
    Retorna:
    valor: El valor de la celda especificada.
    """
    try:
        value = df.iloc[row_idx, col_idx]
        return value
    except IndexError:
        return "Posición fuera de rango"
def extract_rectangle(df, start_row, start_col, end_row, end_col):
    """
    Extrae un rectángulo de un DataFrame dado las coordenadas de inicio y final.
    
    Parámetros:
    df (pd.DataFrame): El DataFrame original.
    start_row (int): Índice de la fila de inicio.
    start_col (int): Índice de la columna de inicio.
    end_row (int): Índice de la fila final.
    end_col (int): Índice de la columna final.
    
    Retorna:
    pd.DataFrame: Un nuevo DataFrame que contiene el rectángulo de celdas extraído.
    """
    # Asegurarnos de que las coordenadas de inicio y final estén dentro de los límites
    if start_row < 0 or start_col < 0 or end_row >= df.shape[0] or end_col >= df.shape[1]:
        return "Coordenadas fuera de rango"
    
    # Extraer el bloque rectangular
    rectangle_df = df.iloc[start_row:end_row+1, start_col:end_col+1]
    return rectangle_df


#%% ejemplo1
# Ejemplo de uso
file_path = 'DICCIONARIO_SERIE_A_2009.xlsx'
df = pd.read_excel(file_path, sheet_name='A02', header=None)

# Ingresar posición
row_idx = 5  # Fila
col_idx = 1  # Columna

# Obtener el valor de la celda
value = get_value_from_position(df, row_idx, col_idx)
print(f"El valor en la posición ({row_idx}, {col_idx}) es: {value}")

#%% ejemplo2 obtener la tabla

# Ejemplo de uso
file_path = 'DICCIONARIO_SERIE_A_2009.xlsx'
df = pd.read_excel(file_path, sheet_name='A02', header=None)

# Ingresar coordenadas de inicio y final
start_row = 7  # Fila de inicio
start_col = 0  # Columna de inicio
end_row = 15    # Fila final
end_col = 5    # Columna final

# Obtener el rectángulo
rectangle = extract_rectangle(df, start_row, start_col, end_row, end_col)

# Imprimir el nuevo DataFrame
print("Rectángulo extraído del DataFrame:")
print(rectangle)

# %% ejemplo3 detectar la tabla
import numpy as np

def expand_to_rectangle(df, start_row, start_col):
    """
    Expande un rectángulo desde una coordenada dada hacia arriba y hacia la derecha,
    deteniéndose cuando se encuentra un NaN en las columnas o un '01010403' en las filas.
    
    Parámetros:
    df (pd.DataFrame): El DataFrame original.
    start_row (int): Índice de la fila de inicio.
    start_col (int): Índice de la columna de inicio.
    
    Retorna:
    pd.DataFrame: El DataFrame resultante con el rectángulo extraído.
    """
    # Inicializar las coordenadas de expansión
    row = start_row
    col = start_col
    
    # Expandirse hacia la derecha hasta encontrar un NaN
    while col < df.shape[1] and pd.notna(df.iloc[row, col]):
        col += 1
    
    # Expandirse hacia abajo hasta encontrar el valor '01010403'
    while row < df.shape[0] and df.iloc[row, start_col] != '01010403':
        row += 1

    # Extraer el bloque rectangular desde el inicio hasta las posiciones encontradas
    rectangle_df = df.iloc[start_row:row, start_col:col]
    
    return rectangle_df

# Ejemplo de uso
file_path = 'DICCIONARIO_SERIE_A_2009.xlsx'
df = pd.read_excel(file_path, sheet_name='A02', header=None)

# Ingresar coordenadas de inicio
start_row = 7  # Fila de inicio
start_col = 0  # Columna de inicio

# Obtener el rectángulo
rectangle = expand_to_rectangle(df, start_row, start_col)

# Imprimir el nuevo DataFrame
print("Rectángulo expandido desde la coordenada inicial:")
print(rectangle)

# %% ejemplo4 alt
import pandas as pd

# Simulamos el DataFrame
file_path = 'DICCIONARIO_SERIE_A_2009.xlsx'
df = pd.read_excel(file_path, sheet_name='A02', header=None)

# Establecemos las filas que contienen el encabezado y los límites de la tabla
start_row = 17  # Fila donde comienza la tabla (fila de encabezados)
end_row = 23  # Fila donde terminan los datos (última fila con datos)

# Extraemos la información de las filas y columnas de la tabla
table_df = df.iloc[start_row:end_row+1, :]  # Asegúrate de incluir la última fila

# Imprimir el DataFrame resultante
print("Tabla extraída:")
print(table_df)

# %% EJEMPLO 4 EN ESETEROIDES
import pandas as pd
import numpy as np

def obtener_tablas(df, start_row):
    tablas = []
    current_table = []
    
    # Iterar desde la fila inicial hacia abajo
    row = start_row
    while row < len(df):
        cell_value = df.iloc[row, 0]  # Valor de la celda en la primera columna
        
        # Verificar si la celda contiene un "SECCIÓN" (con o sin tilde)
        if isinstance(cell_value, str) and cell_value.lower().startswith("sección"):
            if current_table:  # Si ya tenemos una tabla, la añadimos
                tablas.append(pd.DataFrame(current_table))
                current_table = []  # Resetear la tabla actual
            # Continuar con la siguiente fila, ignoramos "SECCIÓN"
            row += 1
            continue
        
        # Verificar si la celda es NaN, lo que indica el fin de la última tabla
        if pd.isna(cell_value):
            if current_table:  # Si ya hay una tabla, añadirla y terminar
                tablas.append(pd.DataFrame(current_table))
            break
        
        # Agregar fila a la tabla actual
        current_table.append(df.iloc[row, :])
        row += 1
    
    # Si quedó alguna tabla pendiente, agregarla
    if current_table:
        tablas.append(pd.DataFrame(current_table))
    
    return tablas

# Simulamos la lectura del archivo Excel
file_path = 'DICCIONARIO_SERIE_A_2009.xlsx'
df = pd.read_excel(file_path, sheet_name='A02', header=None)

# Coordenada de inicio
start_row = 17

# Obtener las tablas
tablas = obtener_tablas(df, start_row)

# Mostrar las tablas obtenidas
for idx, tabla in enumerate(tablas):
    print(f"Tabla {idx + 1}:")
    print(tabla)
    print("\n" + "="*40 + "\n")


# %% EJEMPLO 5 SECCION
#%% EJEMPLO 6 TAMAÑO DE SECCION CON NOMBRE (DE SIGUIENTE SECCION)
import re
import pandas as pd

def quitar_tildes(texto):
    """
    Elimina las tildes de las letras del texto.
    
    Args:
        texto (str): El texto del que se eliminarán las tildes.
    
    Returns:
        str: El texto sin tildes.
    """
    # Diccionario con las letras acentuadas y sus equivalentes sin acento
    acentos = {
        'á': 'a', 'é': 'e', 'í': 'i', 'ó': 'o', 'ú': 'u',
        'Á': 'A', 'É': 'E', 'Í': 'I', 'Ó': 'O', 'Ú': 'U'
    }
    
    # Usamos una expresión regular para reemplazar las letras con tildes por las equivalentes sin tildes
    return re.sub(r'[áéíóúÁÉÍÓÚ]', lambda x: acentos[x.group(0)], texto)

def obtener_texto_y_filas_hasta_seccion(df, col_idx, start_row=0):
    """
    Obtiene el texto que empieza por 'SECCIÓN' y el número de filas leídas
    hasta que se encuentra un texto que empieza por 'SECCIÓN' (con o sin tilde).
    
    Args:
        df (pd.DataFrame): El DataFrame de donde se extraerá la columna.
        col_idx (int): El índice de la columna que se quiere leer.
        start_row (int, opcional): La fila en la que empezar a leer (default es 0).
        
    Returns:
        list: Un arreglo con el texto que empieza con 'SECCIÓN' y el número de filas leídas hasta esa sección.
    """
    fila_contada = 0
    texto_seccion = None
    
    # Iterar por todas las filas de la columna desde la fila de inicio
    for row in df.iloc[start_row:, col_idx]:
        # Verificar si la fila es un texto que empieza con 'SECCIÓN'
        if isinstance(row, str) and quitar_tildes(row).lower().startswith('seccion'):
            texto_seccion = row  # Guardar el texto de la sección
            break
        
        fila_contada += 1
    
    return [texto_seccion, fila_contada]


# Ejemplo de uso
file_path = 'DICCIONARIO_SERIE_A_2009.xlsx'
df = pd.read_excel(file_path, sheet_name='A05', header=None)

# Coordenada de columna a extraer (ej. columna 0)
col_idx = 1

# Fila de inicio (por ejemplo, la fila 5)
start_row = 16

# Llamar a la función para obtener el texto y el número de filas hasta encontrar una sección
resultado = obtener_texto_y_filas_hasta_seccion(df, col_idx, start_row)

# Imprimir el resultado
print(f'Texto encontrado: {resultado[0]}')
print(resultado[1])

# %% NORMALIZAR UN TEXTO
import re

def normalizar_texto(texto):
    """
    Normaliza el texto convirtiéndolo a mayúsculas, cambiando los ':' por '-', 
    eliminando símbolos como ';' y reemplazando los espacios por '_', 
    pero si hay un '- ' (guion seguido de espacio) lo cambia por un '_', 
    y no agrega el guion bajo al final si el último carácter es un espacio.
    
    Args:
        texto (str): El texto a normalizar.
        
    Returns:
        str: El texto normalizado.
    """

    # Convertir a mayúsculas
    texto = texto.upper()
    
    # Eliminar saltos de línea
    texto = texto.replace("\n", "")

    # Reemplazar ':' por guion '-'
    texto = texto.replace(":", "-")
    
    # Eliminar símbolos ';' y otros caracteres no deseados
    texto = re.sub(r'[^\w\s:-]', '', texto)
    
    # Reemplazar espacios por guion bajo, pero si hay '- ' se cambia por '_'
    texto = texto.replace(" ", "_").replace("-_", "-")  # Primero, cambia espacios a _ y luego ajusta el caso específico
    
    # Eliminar el guion bajo al final si el último carácter es un espacio
    if texto.endswith("_"):
        texto = texto[:-1]
    
    return texto

# Ejemplo de uso
texto_original = "cool-text \n"
texto_normalizado = normalizar_texto(texto_original)
print(texto_normalizado)


# %% ELIMINAR FILAS Y COLUMNAS NULA
import pandas as pd
import numpy as np

def eliminar_nulas(df):
    """
    Elimina columnas y filas completamente nulas o con valores 0 de un DataFrame.

    Args:
        df (pd.DataFrame): El DataFrame del que se eliminarán columnas y filas nulas o con valores 0.

    Returns:
        pd.DataFrame: Un DataFrame limpio sin columnas ni filas completamente nulas o con valores 0.
    """
    # Reemplazar los valores 0 por NaN para que también sean considerados nulos
    #df = df.replace(0, np.nan)
    
    # Eliminar columnas que están completamente vacías o nulas
    df = df.dropna(axis=1, how='all')
    
    # Eliminar filas que están completamente vacías o nulas
    df = df.dropna(axis=0, how='all')
    
    return df


# Ejemplo de uso
data = {
    'A': [1, 2, None],
    'B': [None, None, None],
    'C': [4, 0, 6],
    'D': [None, None, None]
}
df = pd.DataFrame(data)

print("Antes de limpiar:")
print(df)

df_limpio = eliminar_nulas(df)

print("\nDespués de limpiar:")
print(df_limpio)


# %% MAIN intento 1
# Ejemplo de uso
# Simulamos el DataFrame
file_path = 'DICCIONARIO_SERIE_A_2009.xlsx'
df = pd.read_excel(file_path, sheet_name='A01', header=None, dtype=str)

start_row = 7
titulo = get_value_from_position(df, (start_row - 1), 1)
titulo_normalizado = normalizar_texto(titulo) #Normalizamos el titulo para guardar en .csv

resultado = obtener_texto_y_filas_hasta_seccion(df,1,start_row)
#print(resultado[1])



#print(titulo_normalizado)
tabla = extract_rectangle(df, start_row, 0,(start_row + resultado[1] - 1), 20)
tabla_limpia = eliminar_nulas(tabla)

# Guardar el DataFrame en un archivo .csv
tabla_limpia.to_csv(f'{titulo_normalizado}.csv', index=False, quoting=1)
#tabla_limpia.to_excel(f'{titulo_normalizado}.xlsx', index=False)  # quoting=1 => csv.QUOTE_ALL
#tabla_limpia.to_json(f'{titulo_normalizado}.json', orient='records', lines=True)
tabla_limpia



# %% CREAR CARPETA
import os

def crear_carpeta(ruta):
    """
    Crea una carpeta en la ruta especificada si no existe.

    Args:
        ruta (str): La ruta de la carpeta que se desea crear.
    """
    try:
        os.makedirs(ruta, exist_ok=True)  # Crea la carpeta y todas las subcarpetas necesarias
        print(f"Carpeta creada: {ruta}")
    except Exception as e:
        print(f"Error al crear la carpeta: {e}")

# Ejemplo de uso
crear_carpeta("TEST1")

# %% MAIN intento 2 (lee TODAS las tablas siempre que tengan el SECCION como separador)
import math
# Ejemplo de uso
# Simulamos el DataFrame
file_path = 'DICCIONARIO_SERIE_A_2009.xlsx'
df = pd.read_excel(file_path, sheet_name='A03', header=None, dtype=str)
df = eliminar_nulas(df)
titulo_carpeta = get_value_from_position(df, 5, 1)
titulo_carpeta_normalizado = normalizar_texto(titulo_carpeta)
crear_carpeta(f"archivos-normalizados/{titulo_carpeta_normalizado}/")
#print(titulo_carpeta_normalizado)
#Valor de inicio
start_row = 7
resultado = ["x", 1]
while resultado[1] != 0:
    #print("holamundo")
    resultado = obtener_texto_y_filas_hasta_seccion(df, 1, start_row)
    if resultado[1] != 0:
        titulo = get_value_from_position(df, (start_row - 1), 1)
        titulo_normalizado = normalizar_texto(titulo)
        #titulo_normalizado = normalizar_texto(titulo)
        #print(resultado)

        #print(resultado)
        tabla = extract_rectangle(df, start_row, 0, (start_row + resultado[1]- 1), 20)
        tabla_limpia = eliminar_nulas(tabla)
        tabla_limpia.to_excel(f"archivos-normalizados/{titulo_carpeta_normalizado}/{titulo_normalizado}.xlsx", index=False)
    
        start_row += resultado[1] + 1
    




# %% NOMBRE DE SHEETS (ARREGLO)
import pandas as pd

def filtrar_sheets_con_A(sheets):
    # Filtrar y devolver solo los nombres que comienzan con 'A'
    return [sheet for sheet in sheets if sheet.startswith('A')]

# Ejemplo de uso
#sheets = ['NOMBRE', 'A01', 'A02', 'A03', 'A04', 'A05', 'A06', 'A07', 'A08', 'A09', 'A10', 'A11', 'A19', 'A22', 'A23', 'A24', 'A25', 'A26', 'A27', 'control', 'MACROS']

# Llamar a la función
#sheets_con_A = filtrar_sheets_con_A(sheets)

# Imprimir el resultado
#print(sheets_con_A)


# Ruta del archivo Excel
file_path = 'DICCIONARIO_SERIE_A_2009.xlsx'

# Cargar el archivo Excel
xls = pd.ExcelFile(file_path)

# Obtener los nombres de todas las hojas (tablas)
nombres_hojas = xls.sheet_names
nombres_hojas_normalizados = filtrar_sheets_con_A(nombres_hojas)

# Imprimir los nombres de las hojas
print(nombres_hojas_normalizados)

# %% MAIN intento 3 (Ahora itera por todos los Sheets de un archivo .xlsx)
import math

file_path = 'DICCIONARIO_SERIE_A_2009.xlsx'
# Cargar el archivo Excel
xls = pd.ExcelFile(file_path)

# Obtener los nombres de todas las hojas (tablas)
nombres_hojas = xls.sheet_names
nombres_hojas_normalizados = filtrar_sheets_con_A(nombres_hojas)
for sheet in nombres_hojas_normalizados:

    df = pd.read_excel(file_path, sheet_name=sheet, header=None, dtype=str)
    df = eliminar_nulas(df)
    titulo_carpeta = get_value_from_position(df, 5, 1)
    titulo_carpeta_normalizado = normalizar_texto(titulo_carpeta)
    crear_carpeta(f"archivos-normalizados/{titulo_carpeta_normalizado}/")
    
    #Inicializamos los valores
    start_row = 7   #Siempre Empieza pro la linea 7
    resultado = ["x", 1]    #Se inicializan los valores para poder entrar en el while
    while resultado[1] != 0:
        
        resultado = obtener_texto_y_filas_hasta_seccion(df, 1, start_row)
        if resultado[1] != 0:
            titulo = get_value_from_position(df, (start_row - 1), 1)
            titulo_normalizado = normalizar_texto(titulo)
            #titulo_normalizado = normalizar_texto(titulo)
            #print(resultado)

            #print(resultado)
            tabla = extract_rectangle(df, start_row, 0, (start_row + resultado[1]- 1), 20)
            tabla_limpia = eliminar_nulas(tabla)
            tabla_limpia.to_excel(f"archivos-normalizados/{titulo_carpeta_normalizado}/{titulo_normalizado}.xlsx", index=False)
        
            start_row += resultado[1] + 1
            last_file = f"archivos-normalizados/{titulo_carpeta_normalizado}/{titulo_normalizado}.xlsx"
# xls2 = pd.ExcelFile(last_file)
# # Eliminar la última fila
# xls2 = xls2.drop(df.index[-1])



# %% Rellenar una columna

def rellenar_celdas(df):
    """
    Rellena las celdas vacías hacia abajo y hacia la derecha en un DataFrame.
    
    Args:
        df (pd.DataFrame): El DataFrame original con celdas vacías.
    
    Returns:
        pd.DataFrame: El DataFrame modificado con celdas rellenadas.
    """
    df = df.copy()  # Crear una copia para no modificar el original
    
    # Lógica para rellenar las celdas hacia abajo
    for col in df.columns:
        for row in range(len(df) - 1):  # Hasta la penúltima fila
            if pd.notna(df.at[row, col]) and pd.isna(df.at[row + 1, col]):
                df.at[row + 1, col] = df.at[row, col]
    
    # Lógica para rellenar hacia la derecha
    for row in range(len(df)):  # Iterar por cada fila
        for col in range(len(df.columns) - 1):  # Hasta la penúltima columna
            if pd.notna(df.iloc[row, col]) and pd.isna(df.iloc[row, col + 1]):
                df.iloc[row, col + 1] = df.iloc[row, col]
    
    return df




file_path = r"C:\Users\benja\OneDrive\Escritorio\Programas\PROGRAMACION\practica1\preProyecto2-REM\archivos-normalizados\REM-A01___CONTROLES_DE_SALUD\SECCIÓN_E-CONTROLES_DE_EMBARAZO_CON_PAREJA_FAMILIAR_U_OTRO.xlsx"
df = pd.read_excel(file_path)


print("DataFrame original:")
print(df)

# Lógica para rellenar las celdas hacia abajo
# for col in df.columns:
#     for row in range(len(df) - 1):  # Hasta la penúltima fila
#         # Si la celda actual tiene valor y la de abajo no, copia el valor
#         if pd.notna(df.at[row, col]) and pd.isna(df.at[row + 1, col]):
#             df.at[row + 1, col] = df.at[row, col]
#         # Si la celda actual tiene valor y la de abajo también, se detiene y pasa a la siguiente columna

# # Lógica para rellenar hacia la derecha
# for row in range(len(df)):  # Iterar por cada fila
#     for col in range(len(df.columns) - 1):  # Hasta la penúltima columna
#         # Si la celda actual tiene valor y la celda a la derecha está vacía, copia el valor
#         if pd.notna(df.iloc[row, col]) and pd.isna(df.iloc[row, col + 1]):
#             df.iloc[row, col + 1] = df.iloc[row, col]
df_modificado = rellenar_celdas(df)


print("\nDataFrame actualizado:")
print(df_modificado)


# %% largo de una tabla
import pandas as pd

def obtener_numero_columnas(df):
    """
    Devuelve el número entero de columnas en un DataFrame.

    Args:
        df (pd.DataFrame): El DataFrame del que se quieren contar las columnas.

    Returns:
        int: El número de columnas en el DataFrame.
    """
    return df.shape[1]

# Ejemplo de uso
data = {
    'A': [1, 2, 3],
    'B': [4, 5, 6],
    'C': [7, 8, 9]
}
df = pd.DataFrame(data)

numero_columnas = obtener_numero_columnas(df)
print(f"El DataFrame tiene {numero_columnas} columnas.")


# %%
file_path = 'DICCIONARIO_SERIE_A_2009.xlsx'
df = pd.read_excel(file_path, sheet_name='A03', header=None, dtype=str)
print(obtener_numero_columnas(df))
# %%
