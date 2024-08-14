import os
import pandas as pd
import argparse

# Configuración de argparse
parser = argparse.ArgumentParser(description='Script para eliminar archivos según una lista en Excel.')
parser.add_argument('--base', type=str, default='', help='Ruta base para las rutas relativas.')
parser.add_argument('--borrar', type=str, default=r'C:\Users\juan.vermejo\Documents\CPNO\Duplicados.xlsx', 
                    help='Ruta del archivo de Excel que contiene la lista de archivos a borrar.')

# Obtener los argumentos
args = parser.parse_args()

# Ruta del archivo de Excel
excel_file = args.borrar
ruta_base = args.base

# Leer la hoja "Informes x borrar" del archivo de Excel
df = pd.read_excel(excel_file, sheet_name='Informes x borrar')

print(df)

# Iterar sobre cada fila en el DataFrame
for index, row in df.iterrows():
    ruta_relativa = row['Ruta relativa']
    archivo_a_borrar = row['Borrar']

    # Construir la ruta completa del archivo
    ruta_completa = os.path.join(ruta_base, ruta_relativa, archivo_a_borrar, " - Informes CPNO.xlsx")

    # Verificar si el archivo existe y luego eliminarlo
    if os.path.isfile(ruta_completa):
        os.remove(ruta_completa)
        print(f"Archivo {ruta_completa} eliminado.")
    else:
        print(f"Archivo {ruta_completa} no encontrado.")
