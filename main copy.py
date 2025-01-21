#%%
import pandas as pd
import numpy as np
import re
import os
import math

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
    - df (pd.DataFrame): El DataFrame original.
    - start_row (int): Índice de la fila de inicio.
    - start_col (int): Índice de la columna de inicio.
    - end_row (int): Índice de la fila final.
    - end_col (int): Índice de la columna final.
    
    Retorna:
    - pd.DataFrame: Un nuevo DataFrame que contiene el rectángulo de celdas extraído.
    
    Excepciones:
    - ValueError: Si las coordenadas están fuera de rango o son inconsistentes.
    """
    # Validar límites
    if start_row < 0:
        raise ValueError(f"start_row ({start_row}) está fuera de los límites (debe ser >= 0).")
    if start_col < 0:
        raise ValueError(f"start_col ({start_col}) está fuera de los límites (debe ser >= 0).")
    if end_row >= df.shape[0]:
        raise ValueError(f"end_row ({end_row}) excede el límite de filas del DataFrame ({df.shape[0] - 1}).")
    if end_col >= df.shape[1]:
        raise ValueError(f"end_col ({end_col}) excede el límite de columnas del DataFrame ({df.shape[1] - 1}).")
    if end_row < start_row:
        raise ValueError(f"end_row ({end_row}) es menor que start_row ({start_row}), lo que es inconsistente.")
    if end_col < start_col:
        raise ValueError(f"end_col ({end_col}) es menor que start_col ({start_col}), lo que es inconsistente.")
    
    # Extraer el bloque rectangular
    rectangle_df = df.iloc[start_row:end_row + 1, start_col:end_col + 1]
    return rectangle_df

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

def quitar_tildes(texto):
    """
    Elimina las tildes de las letras del texto.
    
    Args:
        texto (str): El texto del que se eliminarán las tildes.
    
    Returns:
        str: El texto sin tildes.
    """
    # Diccionario con las letras acentuadas y sus equivalentes sin acento
    if not isinstance(texto, str):
        raise TypeError(f"Se esperaba una cadena, pero se recibió {type(texto).__name__}")
  
    
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
        list: Un arreglo con el texto que empieza con 'SECCIÓN', el número de filas leídas hasta esa sección, un booleano indicando si es el titulo de una subseccion o no.
    """
    fila_contada = 0
    texto_seccion = None
    isSubSeccion = False
    row = None  # Inicializar la variable para usarla después del bucle

    # Iterar por todas las filas de la columna desde la fila de inicio
    for row in df.iloc[start_row:, col_idx]:
        # Verificar si la fila es un texto que empieza con 'SECCIÓN'
        if isinstance(row, str) and quitar_tildes(row).lower().startswith('seccion'):
            texto_seccion = row  # Guardar el texto de la sección
            break
        
        fila_contada += 1

    # Verificar si es una subsección en base al valor de la última fila procesada
    if fila_contada == 0 and row is not None and isinstance(row, str) and quitar_tildes(row).lower().startswith('seccion'):
        isSubSeccion = True
        return [texto_seccion, fila_contada, isSubSeccion]
    elif row is not None:
        return [texto_seccion, fila_contada, isSubSeccion]
    else:
        return ["ESTO ES EL FIN", fila_contada, isSubSeccion]

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

def eliminar_nulas(df):
    """
    Elimina columnas y filas completamente nulas o con valores 0 de un DataFrame.

    Args:
        df (pd.DataFrame): El DataFrame del que se eliminarán columnas y filas nulas o con valores 0.

    Returns:
        pd.DataFrame: Un DataFrame limpio sin columnas ni filas completamente nulas o con valores 0.
    """
    # Reemplazar los valores 0 por NaN para que también sean considerados nulos
    df = df.replace(0, np.nan)
    
    # Eliminar columnas que están completamente vacías o nulas
    df = df.dropna(axis=1, how='all')
    
    # Eliminar filas que están completamente vacías o nulas
    df = df.dropna(axis=0, how='all')
    
    return df
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

def filtrar_sheets_con_A(sheets):
    # Filtrar y devolver solo los nombres que comienzan con 'A'
    return [sheet for sheet in sheets if sheet.startswith('A')]

def obtener_numero_columnas(df):
    """
    Devuelve el número entero de columnas en un DataFrame.

    Args:
        df (pd.DataFrame): El DataFrame del que se quieren contar las columnas.

    Returns:
        int: El número de columnas en el DataFrame.
    """
    return df.shape[1]

# %% Obtener todos los archivos Excel de una carpeta 
def obtener_nombres_archivos_excel(ruta_carpeta):
    """
    Obtiene los nombres de todos los archivos .xlsx y .xlsm en una carpeta.

    Args:
        ruta_carpeta (str): Ruta de la carpeta donde buscar los archivos.

    Returns:
        list: Una lista con los nombres de los archivos encontrados (incluye extensión).
    """
    # Verificar si la ruta es válida
    if not os.path.exists(ruta_carpeta):
        raise FileNotFoundError(f"La ruta proporcionada no existe: {ruta_carpeta}")

    # Filtrar los archivos con las extensiones deseadas
    archivos_excel = [
        archivo for archivo in os.listdir(ruta_carpeta)
        if archivo.endswith(('.xlsx', '.xlsm'))
    ]

    return archivos_excel

# Ejemplo de uso
ruta = "Files/"
archivos = obtener_nombres_archivos_excel(ruta)
print(archivos)

#%% Ultima fila de un texto
def find_last_occurrence(df, text):
    """
    Encuentra la coordenada de la última fila en el DataFrame donde se encuentra el texto dado (insensible a mayúsculas/minúsculas).

    Args:
        df (pd.DataFrame): DataFrame donde buscar.
        text (str): El texto que se desea buscar.

    Returns:
        tuple: Una tupla con (fila, columna) de la última ocurrencia del texto.
              Si no se encuentra, devuelve (None, None).
    """
    last_row, last_col = None, None
    text = text.lower() # Convertir el texto de búsqueda a minúsculas una sola vez

    for row_idx in range(len(df)):
        for col_idx in range(len(df.columns)):
            value = df.iat[row_idx, col_idx]
            if isinstance(value, str) and text in value.lower(): # Convertir el valor de la celda a minúsculas para la comparación
                last_row, last_col = row_idx, col_idx

    return (last_row, last_col)


# Ejemplo de uso
file_path = "Files/DICCIONARIO CODIGOS SA_23_V1.4.xlsm"

df = pd.read_excel(file_path, sheet_name='A04', header=None, dtype=str)

fila, columna = find_last_occurrence(df, 'COL01')
print(f"Última ocurrencia de 'manzana': Fila {fila}, Columna {columna}")




# %% Encontrar un valor x

def find_first_occurrence(df, target):
    """
    Busca la primera aparición de un string en un DataFrame y devuelve sus coordenadas.
    
    Parámetros:
    df (pd.DataFrame): El DataFrame donde buscar.
    target (str): El string que se desea buscar.
    
    Retorna:
    tuple: Una tupla con las coordenadas (fila, columna) de la primera aparición.
           Retorna None si el string no se encuentra e imprime un mensaje de error.
    """
    for row_idx, row in df.iterrows():
        for col_idx, value in row.items():
            if value == target:
                return (row_idx, col_idx)
    
    # Si no se encuentra el texto, imprimir un mensaje de error
    #print(f"Error: No se encontró el texto '{target}' en la tabla.")
    return None


# Ejemplo de uso
file_path = 'DICCIONARIO_SERIE_A_2009.xlsx'
df = pd.read_excel(file_path, sheet_name='A01', header=None)

# Buscar el string 'manzana'
resultado = find_first_occurrence(df, 'COL1')
row_indx,col_indx  = resultado 
print(f"Primera ocurrencia: {resultado}")  # Output: Primera ocurrencia: (0, 'A')
valor = get_value_from_position(df, row_indx, col_indx)
print(valor)

# %% Obtener la ultima columna de una tabla
def find_last_col_to_right(df, start_row, start_col):
    """
    Busca hacia la derecha desde una coordenada específica en un DataFrame y devuelve 
    el número de la columna del último valor que comienza con 'COL'.
    
    Parámetros:
    df (pd.DataFrame): El DataFrame donde buscar.
    start_row (int): Índice de la fila desde donde comenzar.
    start_col (int): Índice de la columna desde donde comenzar.
    
    Retorna:
    int: Número de la columna del último valor que empieza con 'COL'. 
         Si no se encuentra ningún valor válido, devuelve None.
    """
    #Verificar que la fila y columna inicial estén dentro del rango del DataFrame
    if start_row < 0 or start_row >= df.shape[0] or start_col < 0 or start_col >= df.shape[1]:
        raise ValueError("Coordenadas fuera de rango")
    
    # Inicializar el índice de la última columna válida
    last_valid_col = None
    
    # Recorrer las columnas hacia la derecha desde la columna inicial
    for col_idx in range(start_col, df.shape[1]):
        value = str(df.iloc[start_row, col_idx])  # Obtener el valor como string
        if value.startswith("COL"):
            last_valid_col = col_idx  # Actualizar la última columna válida
        else:
            break  # Detener el proceso si el valor no comienza con 'COL'
    
    return last_valid_col

# Ejemplo de uso
file_path = 'DICCIONARIO_SERIE_A_2009.xlsx'
df = pd.read_excel(file_path, sheet_name='A01', header=None)

# Probar la función
resultado = find_last_col_to_right(df, 9, 3)
print(f"Última columna con 'COL': {resultado}")  # Output: Última columna con 'COL': 3

#%% primera coincidencia
def buscar_primera_coincidencia(df, texto):
    """
    Busca la primera coincidencia de un texto en un DataFrame.

    Args:
        df (pd.DataFrame): El DataFrame donde buscar.
        texto (str): El texto a buscar.

    Returns:
        tuple or None: Una variable con la coordenada (fila) si se encuentra el texto,
                       o None si no se encuentra.
    """
    for fila_idx, fila in df.iterrows():
        for col_idx, valor in fila.items():
            if isinstance(valor, str) and texto.lower() in valor.lower():
                return fila_idx
    return None

#%% Encontrar el titulo
import re
import unicodedata

def remove_accents(input_str):
    """
    Elimina las tildes de un string dado.
    
    Parámetros:
    input_str (str): El string del cual se eliminarán las tildes.
    
    Retorna:
    str: El string sin tildes.
    """
    return ''.join(
        c for c in unicodedata.normalize('NFD', input_str)
        if unicodedata.category(c) != 'Mn'
    )

def titulo_finder(df, partial_title):
    """
    Busca la coordenada de la fila y columna donde se encuentra un título
    que comienza con una parte específica, ignorando tildes y mayúsculas.
    
    Parámetros:
    df (pd.DataFrame): El DataFrame donde buscar.
    partial_title (str): La primera parte del título a buscar.
    
    Retorna:
    tuple: Una tupla con las coordenadas (fila, columna) del título encontrado.
           Retorna None si no se encuentra.
    """
    # Eliminar espacio y tildes del partial_title
    #partial_title = partial_title.strip().lower()
    partial_title_normalized = remove_accents(partial_title).lower().strip()
    

    for row_idx, row in df.iterrows():
        for col_idx, value in row.items():
            if isinstance(value, str):
                # Eliminar tildes y convertir a minúsculas para comparar
                value_normalized = remove_accents(value).lower().strip()
                if value_normalized.startswith(partial_title_normalized):
                    return (row_idx, col_idx)
    
    # Si no se encuentra el título
    print(f"Error: No se encontró ningún título que comience con '{partial_title}' en la tabla.")
    return None


# Crear un ejemplo de DataFrame
data = {
    'Col1': ['RaM-A01: Detalle', 'Datos', 'REM-02: Informe'],
    'Col2': ['Resumen', 'ram: Prestaciones y otros', 'N/A']
}
df = pd.DataFrame(data)

# Buscar el título que empieza con 'REM'
result = titulo_finder(df, " rém")
print(result)  # Devuelve la coordenada de la primera coincidencia


# %% ENCONTRAR ULTIMO COL1 / COL01
import pandas as pd
import re

def encontrar_ultimo_col01(df):
    """
    Encuentra las coordenadas de la última ocurrencia de 'COL01' (y variantes como 'COL1', 'col01', etc.)
    en el DataFrame utilizando expresiones regulares.
    
    Parameters:
    df (pandas.DataFrame): DataFrame de entrada
    
    Returns:
    tuple: (fila, columna) de la última ocurrencia de COL01
           (-1, -1) si no se encuentra
    """
    ultima_fila = -1
    ultima_columna = -1
    # Expresión regular para encontrar cualquier variante de 'COL01'
    patron = re.compile(r'col0?1', re.IGNORECASE)
    
    # Iteramos por todas las celdas del DataFrame
    for fila in range(len(df)):
        for columna in range(len(df.columns)):
            # Convertimos el valor a string y comparamos con la expresión regular
            valor = str(df.iloc[fila, columna]).strip()
            if patron.match(valor):
                # Actualizamos las coordenadas si encontramos una ocurrencia
                # que está en una columna posterior o en la misma columna pero fila posterior
                if (columna > ultima_columna or 
                    (columna == ultima_columna and fila > ultima_fila)):
                    ultima_fila = fila
                    ultima_columna = columna
    
    return ultima_fila, ultima_columna

# Example DataFrame
path_file = "archivos-normalizados/REM-A30_AR-ATENCIÓN_Y_ORIENTACIÓN_DE_SALUD_A_DISTANCIA_HD/SECCION_H-ORIENTACIÓN_TELEFÓNICA_EN_SALUD.xlsx"
seccion_h = pd.read_excel(path_file)

# Para encontrar las coordenadas del último COL01
fila, columna = encontrar_ultimo_col01(seccion_h)
print(f"Último COL01 encontrado en: fila {fila}, columna {columna}")

# Find furthest occurrence
#row, col = find_furthest_col_coordinate(seccion_h, 'COL1')
#print(f"Furthest COL1 found at: row {row}, column {col}")

# Find all occurrences (for verification)
#all_coords = find_all_coordinates(df, 'COL1')
#print(f"All occurrences: {all_coords}")

# path_file ="archivos-normalizados/REM-A30_AR-ATENCIÓN_Y_ORIENTACIÓN_DE_SALUD_A_DISTANCIA_HD/SECCION_H-ORIENTACIÓN_TELEFÓNICA_EN_SALUD.xlsx"
# seccion_h = pd.read_excel(path_file)
# coordenadas = find_highest_col(seccion_h, "col01")
# print(coordenadas)
# print(seccion_h.columns.tolist())

#%% MAIN COMPLETO


Archivos_SA_diccionarios = obtener_nombres_archivos_excel("FILES/")
for DiccionarioAño in Archivos_SA_diccionarios:
    
    file_path = f"{DiccionarioAño}"
    #crear_carpeta(DiccionarioAño)

    #file_path = 'SA-2010_CODIGOS.xlsx'
    nombre_carpeta_principal = file_path.rsplit('.', 1)[0]
    direccion_principal_out = f"archivos-normalizados/{nombre_carpeta_principal}"
    crear_carpeta(direccion_principal_out)
    # Cargar el archivo Excel
    xls = pd.ExcelFile(f"FILES/{DiccionarioAño}")

    # Obtener los nombres de todas las hojas (tablas)
    nombres_hojas = xls.sheet_names
    nombres_hojas_normalizados = filtrar_sheets_con_A(nombres_hojas)
    print(nombres_hojas_normalizados)
    for sheet in nombres_hojas_normalizados:
        print(sheet)
        df = pd.read_excel(f"FILES/{DiccionarioAño}", sheet_name=sheet, header=None, dtype=str)
        table_widht = df.shape[1]

        
        titulo_row, titulo_col = titulo_finder(df, "REM")
        titulo_carpeta = get_value_from_position(df, titulo_row, titulo_col)
        titulo_carpeta_normalizado = normalizar_texto(titulo_carpeta)
        crear_carpeta(f"archivos-normalizados/{nombre_carpeta_principal}/{titulo_carpeta_normalizado}/")

        #Valor de inicio
        start_row = buscar_primera_coincidencia(df, "SECCIÓN") + 1 
        #print(start_row)
        resultado = ["x", 1, False]

        while resultado[1] != 0 or resultado[2] == True:    # El largo de una columna es diferente de 0 o es el titulo de una sub seccion
            resultado = obtener_texto_y_filas_hasta_seccion(df, 1, start_row)
            #print(resultado[2])
            if resultado[1] != 0 or resultado[2] == True:   # El largo de una columna es diferente de 0 o es el titulo de una sub seccion
                titulo = get_value_from_position(df, (start_row - 1), 1)
                titulo_normalizado = normalizar_texto(titulo)
                #print(start_row, (start_row + resultado[1] - 1))
                #print((table_widht-1))
                if resultado[2] == False:   #Es el titulo de una sub Seccion? (False)
                    #print(titulo_normalizado)
                    tabla = extract_rectangle(df, start_row, 0, (start_row + resultado[1] - 1), (table_widht-1)) #agarra el tamaño total de columnas y le resta 1
                    
                    coordenadas = find_first_occurrence(tabla, "COL1")
                    if coordenadas is None:
                        # Intentar con "COL01" si no se encontró "COL1"
                        coordenadas = find_first_occurrence(tabla, "COL01")

                    if coordenadas is not None:
                        row_COL, col_COL = coordenadas
                        #print(f"Coordenadas encontradas: fila {row_COL}, columna {col_COL}")
                    else:
                        print("No se encontró ninguna coincidencia para COL1 o COL01")

                    last_col = find_last_col_to_right(df, row_COL, col_COL ) #OCUPAR DF original para obtener las cordenadas absolutas
                    tabla = extract_rectangle(tabla, 0, 0, (resultado[1] - 1), last_col)
                    tabla_limpia2 = eliminar_nulas(tabla)
                    tabla_limpia2.to_excel(f"{direccion_principal_out}/{titulo_carpeta_normalizado}/{titulo_normalizado}.xlsx", index=False)
                else: # (True)
                    #crear_carpeta(f"{direccion_principal_out}/{titulo_carpeta_normalizado}/{titulo_normalizado}")
                    print(" ")
            

                start_row += resultado[1] + 1
                last_file = f"{direccion_principal_out}/{titulo_carpeta_normalizado}/{titulo_normalizado}.xlsx"
        last_file_df = pd.read_excel(last_file)
        output_file = last_file

        coordenadas = encontrar_ultimo_col01(last_file_df) # Busca coordenadas de "col1", "COL1", "Col1", etc.

        if coordenadas != (None, None):
            row_COL, col_COL = coordenadas
            last_file_df = last_file_df.iloc[:row_COL + 1]
            try:
                last_file_df.to_excel(output_file, index=False)
                #print(f"Archivo guardado en {output_file}")
            except Exception as e:
                print(f"Error al guardar el archivo: {e}")
        else:
            print("No se encontró ninguna coincidencia para 'COL1' o 'COL01'.")

#%% MAIN testeo individual
import math
# # Ejemplo de uso
file_path = "FILES/DICCIONARIO_SERIE_A_2009.xlsx"

df = pd.read_excel(file_path, sheet_name='A06', header=None, dtype=str)
table_widht = df.shape[1]


# titulo_carpeta = get_value_from_position(df, 5, 1)
#titulo_carpeta_normalizado = normalizar_texto(titulo_carpeta)
# valocococo = get_value_from_position(df, 5, 1)
# valocococo = valocococo.strip()
# print(valocococo)
titulo_row, titulo_col = titulo_finder(df, "REM")
titulo_carpeta = get_value_from_position(df, titulo_row, titulo_col)
titulo_carpeta_normalizado = normalizar_texto(titulo_carpeta)
crear_carpeta(f"archivos-normalizados/{titulo_carpeta_normalizado}/")

#Valor de inicio
start_row, start_col = titulo_finder(df, "SECCIÓN") 
start_row += 1 
#print(start_row)
resultado = ["x", 1, False]

#row, col = find_first_occurrence(df, "COL01")
#print(row, col)


while resultado[1] != 0 or resultado[2] == True:    # El largo de una columna es diferente de 0 o es el titulo de una sub seccion
    resultado = obtener_texto_y_filas_hasta_seccion(df, 1, start_row)
    #print(resultado[2])
    if resultado[1] != 0 or resultado[2] == True:   # El largo de una columna es diferente de 0 o es el titulo de una sub seccion
        titulo = get_value_from_position(df, (start_row - 1), 1)
        titulo_normalizado = normalizar_texto(titulo)
        #print(start_row, (start_row + resultado[1] - 1))
        #print((table_widht-1))
        if resultado[2] == False:   #Es el titulo de una sub Seccion? (False)
            print(titulo_normalizado)
            tabla = extract_rectangle(df, start_row, 0, (start_row + resultado[1] - 1), (table_widht-1)) #agarra el tamaño total de columnas y le resta 1
            
            coordenadas = find_first_occurrence(tabla, "COL1")
            if coordenadas is None:
                # Intentar con "COL01" si no se encontró "COL1"
                coordenadas = find_first_occurrence(tabla, "COL01")

            if coordenadas is not None:
                row_COL, col_COL = coordenadas
                #print(f"Coordenadas encontradas: fila {row_COL}, columna {col_COL}")
            else:
                print("No se encontró ninguna coincidencia para COL1 o COL01")

            last_col = find_last_col_to_right(df, row_COL, col_COL ) #OCUPAR DF original para obtener las cordenadas absolutas
            tabla = extract_rectangle(tabla, 0, 0, (resultado[1] - 1), last_col)
            tabla_limpia2 = eliminar_nulas(tabla)
            tabla_limpia2.to_excel(f"archivos-normalizados/{titulo_carpeta_normalizado}/{titulo_normalizado}.xlsx", index=False)
        else: # (True)
            crear_carpeta(f"archivos-normalizados/{titulo_carpeta_normalizado}/{titulo_normalizado}")
       

        start_row += resultado[1] + 1
#         last_file = f"archivos-normalizados/{titulo_carpeta_normalizado}/{titulo_normalizado}.xlsx"
# last_file_df = pd.read_excel(last_file)
# output_file = last_file

# coordenadas = encontrar_ultimo_col01(last_file_df) # Busca coordenadas de "col1", "COL1", "Col1", etc.

# if coordenadas != (None, None):
#     row_COL, col_COL = coordenadas
#     last_file_df = last_file_df.iloc[:row_COL + 1]
#     try:
#         last_file_df.to_excel(output_file, index=False)
#       #print(f"Archivo guardado en {output_file}")
#     except Exception as e:
#         print(f"Error al guardar el archivo: {e}")
# else:
#     print("No se encontró ninguna coincidencia para 'COL1' o 'COL01'.")

# %% MAIN SEMI-COmpleto (todas las hojas de un año)

# Archivos_SA_diccionarios = obtener_nombres_archivos_excel("FILES/")
# for DiccionarioAño in Archivos_SA_diccionarios:
    
#file_path = f"{DiccionarioAño}"
#crear_carpeta(DiccionarioAño)

file_path = 'SA-13_V1.3_CODIGOS.xlsx'
nombre_carpeta_principal = file_path.rsplit('.', 1)[0]
direccion_principal_out = f"archivos-normalizados/{nombre_carpeta_principal}"
crear_carpeta(direccion_principal_out)
# Cargar el archivo Excel
xls = pd.ExcelFile(f"FILES/{file_path}")

# Obtener los nombres de todas las hojas (tablas)
nombres_hojas = xls.sheet_names
nombres_hojas_normalizados = filtrar_sheets_con_A(nombres_hojas)
print(nombres_hojas_normalizados)
for sheet in nombres_hojas_normalizados:
    print(sheet)
    df = pd.read_excel(f"FILES/{file_path}", sheet_name=sheet, header=None, dtype=str)
    table_widht = df.shape[1]

    
    titulo_row, titulo_col = titulo_finder(df, "REM")
    titulo_carpeta = get_value_from_position(df, titulo_row, titulo_col)
    titulo_carpeta_normalizado = normalizar_texto(titulo_carpeta)
    crear_carpeta(f"archivos-normalizados/{nombre_carpeta_principal}/{titulo_carpeta_normalizado}/")

    #Valor de inicio
    start_row = buscar_primera_coincidencia(df, "SECCIÓN") + 1 
    #print(start_row)
    resultado = ["x", 1, False]

    while resultado[1] != 0 or resultado[2] == True:    # El largo de una columna es diferente de 0 o es el titulo de una sub seccion
        resultado = obtener_texto_y_filas_hasta_seccion(df, 1, start_row)
        #print(resultado[2])
        if resultado[1] != 0 or resultado[2] == True:   # El largo de una columna es diferente de 0 o es el titulo de una sub seccion
            titulo = get_value_from_position(df, (start_row - 1), 1)
            titulo_normalizado = normalizar_texto(titulo)
            #print(start_row, (start_row + resultado[1] - 1))
            #print((table_widht-1))
            if resultado[2] == False:   #Es el titulo de una sub Seccion? (False)
                #print(titulo_normalizado)
                tabla = extract_rectangle(df, start_row, 0, (start_row + resultado[1] - 1), (table_widht-1)) #agarra el tamaño total de columnas y le resta 1
                
                coordenadas = find_first_occurrence(tabla, "COL1")
                if coordenadas is None:
                    # Intentar con "COL01" si no se encontró "COL1"
                    coordenadas = find_first_occurrence(tabla, "COL01")

                if coordenadas is not None:
                    row_COL, col_COL = coordenadas
                    #print(f"Coordenadas encontradas: fila {row_COL}, columna {col_COL}")
                else:
                    print("No se encontró ninguna coincidencia para COL1 o COL01")

                last_col = find_last_col_to_right(df, row_COL, col_COL ) #OCUPAR DF original para obtener las cordenadas absolutas
                tabla = extract_rectangle(tabla, 0, 0, (resultado[1] - 1), last_col)
                tabla_limpia2 = eliminar_nulas(tabla)
                tabla_limpia2.to_excel(f"{direccion_principal_out}/{titulo_carpeta_normalizado}/{titulo_normalizado}.xlsx", index=False)
            else: # (True)
                #crear_carpeta(f"{direccion_principal_out}/{titulo_carpeta_normalizado}/{titulo_normalizado}")
                print(" ")
        

            start_row += resultado[1] + 1
# %%
