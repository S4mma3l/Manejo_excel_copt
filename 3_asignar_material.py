import pandas as pd
import os

# Directorio donde están tus archivos de Excel
directorio = r"Z:\Autocad Files\PROYECTOS 2024\RESIDENCIA LOTE 10\Mueble_de_Cocina\Solidos\Materiales"

# Obtener la lista de archivos de Excel en el directorio
archivos_excel = [f for f in os.listdir(directorio) if f.endswith('.xlsx')]

# Mostrar la lista de archivos y pedir al usuario que elija uno
print("Archivos disponibles:")
for i, archivo in enumerate(archivos_excel):
    print(f"{i + 1}. {archivo}")

# Pedir al usuario que elija un archivo
try:
    opcion = int(input("Elige el número del archivo que deseas procesar: "))

    # Validar la opción seleccionada
    if 1 <= opcion <= len(archivos_excel):
        archivo_seleccionado = archivos_excel[opcion - 1]
        ruta_archivo = os.path.join(directorio, archivo_seleccionado)
        
        # Mapeo de espesores a materiales
        mapa_espesores = {
            16: 'MELAMINA 16 MM',
            18: 'EUCALIPTO 18 MM',
            20: 'MADERA MELINA',
            17: 'MELAMINA 18 MM',
            22: 'MADERA MELINA'
        }
        
        # Leer el archivo Excel seleccionado
        df = pd.read_excel(ruta_archivo, engine='openpyxl')
        
        # Verificar que la columna 'Espesor' exista
        if 'Espesor' in df.columns:
            # Crear la nueva columna 'Material' basada en el espesor
            df['Material'] = df['Espesor'].map(mapa_espesores)
        else:
            print("La columna 'Espesor' no se encuentra en el archivo.")
            exit()
        
        # Guardar el archivo Excel modificado en el mismo directorio
        nombre_archivo_modificado = f'Asignacion_Materiales_{archivo_seleccionado}'
        ruta_salida = os.path.join(directorio, nombre_archivo_modificado)
        
        try:
            df.to_excel(ruta_salida, index=False, engine='openpyxl')
            print(f"Se ha añadido la columna 'Material' y se ha guardado el archivo modificado: {ruta_salida}")
        except Exception as e:
            print(f"Error al guardar el archivo modificado: {e}")
    
    else:
        print("Opción no válida. Por favor, selecciona un número de la lista.")
        
except ValueError:
    print("Entrada no válida. Por favor, ingresa un número entero.")

