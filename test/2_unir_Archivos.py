import pandas as pd
import os

# Directorio donde están tus archivos de Excel
directorio = r"Z:\Autocad Files\PROYECTOS 2024\RESIDENCIA LOTE 10\Mueble_de_Cocina\Solidos\Materiales"

# Lista de columnas que quieres extraer
columnas_interes = ['QTY', 'DESCRIPTION', 'Largo', 'Ancho', 'Espesor']

# DataFrame vacío para almacenar todos los datos
df_consolidado = pd.DataFrame()

# Obtener la lista de archivos de Excel en el directorio
archivos_excel = [f for f in os.listdir(directorio) if f.endswith('.xlsx')]

# Iterar sobre cada archivo de Excel
for archivo in archivos_excel:
    ruta_archivo = os.path.join(directorio, archivo)
    
    try:
        # Leer el archivo Excel
        df = pd.read_excel(ruta_archivo, engine='openpyxl')
        
        # Verificar qué columnas están disponibles en el DataFrame
        columnas_disponibles = [col for col in columnas_interes if col in df.columns]
        
        # Seleccionar solo las columnas de interés (si existen en el archivo)
        df_seleccionado = df[columnas_disponibles].dropna(how='all')
        
        # Añadir una columna con el nombre del archivo (sin la extensión)
        df_seleccionado['Archivo'] = os.path.splitext(archivo)[0]
        
        # Añadir los datos al DataFrame consolidado
        df_consolidado = pd.concat([df_consolidado, df_seleccionado], ignore_index=True)
    
    except Exception as e:
        print(f"Error al procesar {archivo}: {e}")

# Reordenar las columnas para que 'Archivo' sea la primera, si está presente
columnas_finales = ['Archivo'] + columnas_interes
columnas_finales = [col for col in columnas_finales if col in df_consolidado.columns]
df_consolidado = df_consolidado[columnas_finales]

# Guardar el DataFrame consolidado en un nuevo archivo Excel
ruta_salida = os.path.join(directorio, 'Union_Archivos.xlsx')
try:
    df_consolidado.to_excel(ruta_salida, index=False, engine='openpyxl')
    print(f"Se ha creado el archivo consolidado: {ruta_salida}")
except Exception as e:
    print(f"Error al guardar el archivo consolidado: {e}")