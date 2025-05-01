# Gestor de Listas COPT

Una herramienta en Python para procesar archivos Excel, consolidar datos, asignar materiales y generar listas para COPT.

## Características

*   Elimina automáticamente "mm" de los archivos fuente.
*   Consolida datos de múltiples archivos Excel en `Fusion.xlsx`.
*   Asigna materiales basado en la columna 'Espesor'.
*   Inserta los datos procesados en una hoja 'LISTA' de un archivo destino.
*   Interfaz gráfica moderna creada con Tkinter y ttkbootstrap.

## Uso

1.  Ejecuta el programa (el `.exe` o `python tu_script.py`).
2.  **Paso 1:** Haz clic en "Procesar Carpeta y Asignar Materiales". Selecciona la carpeta con los archivos Excel fuente. El programa creará/actualizará `Fusion.xlsx` en esa misma carpeta.
3.  **Paso 2:** Haz clic en "Insertar Datos en Hoja Destino". Selecciona primero el archivo `Fusion.xlsx` generado y luego el archivo Excel destino que contiene la hoja `LISTA` (a partir de la fila 13).

## Dependencias (si se ejecuta desde .py)

*   pandas
*   openpyxl
*   ttkbootstrap

Puedes instalarlas con: `pip install pandas openpyxl ttkbootstrap`

