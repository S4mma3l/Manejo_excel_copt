import xlwings as xw

# Ruta del archivo .xls que deseas convertir
xls_file_path = r'Z:\Autocad Files\PROYECTOS 2024\RESIDENCIA LOTE 10\Mueble_de_Cocina\Solidos\Materiales\Lista_Mueble_.xls'

# Ruta donde guardarás el nuevo archivo .xlsx
xlsx_file_path = r'Z:\Autocad Files\PROYECTOS 2024\RESIDENCIA LOTE 10\Mueble_de_Cocina\Solidos\Materiales\Lista_Mueble_converted.xlsx'

# Abrir el archivo .xls con Excel
app = xw.App(visible=False)  # Mantener Excel en segundo plano
wb = app.books.open(xls_file_path)

# Guardar como .xlsx
wb.save(xlsx_file_path)

# Cerrar el archivo y la aplicación Excel
wb.close()
app.quit()

print(f"Archivo convertido exitosamente a {xlsx_file_path}")
