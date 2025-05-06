import pandas as pd
import os

# Directorio donde están tus archivos de Excel
directorio = r'Z:\Autocad Files\PROYECTOS 2024\RESIDENCIA LOTE 10\Mueble_de_Cocina\Solidos\Materiales'

# Obtener la lista de archivos de Excel en el directorio
archivos_excel = [f for f in os.listdir(directorio) if f.endswith('.xlsx')]

# Iterar sobre cada archivo de Excel
for archivo in archivos_excel:
    ruta_archivo = os.path.join(directorio, archivo)
    
    try:
        # Leer el archivo Excel
        df = pd.read_excel(ruta_archivo, engine='openpyxl')
        
        # Reemplazar las letras "mm" en todo el DataFrame
        df = df.applymap(lambda x: x.replace('mm', '') if isinstance(x, str) else x)
        
        # Guardar el archivo Excel modificado
        df.to_excel(ruta_archivo, index=False, engine='openpyxl')
        print(f"Se han eliminado las letras 'mm' de {archivo}.")
    
    except Exception as e:
        print(f"Error al procesar {archivo}: {e}")

print("Proceso completado.")
