import xlwings as xw
import os
import argparse

# Configuración de argumentos de línea de comandos
parser = argparse.ArgumentParser(description='Procesar archivos de Excel en una carpeta.')
parser.add_argument('--path', type=str, default=r'C:\Users\juan.vermejo\Documents\CPNO\Pruebas\Masivo',
                    help='Ruta de la carpeta que contiene los archivos de Excel. Por defecto, se utiliza la ruta predeterminada.')
args = parser.parse_args()

# Ruta de la carpeta que contiene los archivos de Excel
carpeta = args.path

# Itera sobre cada archivo en la carpeta
for archivo in os.listdir(carpeta):
    if archivo.endswith('.xlsx'):
        ruta_archivo = os.path.join(carpeta, archivo)
        
        with xw.App(visible=False) as app:
            wb = xw.Book(ruta_archivo)

            # Verifica si la hoja "HojaLegal" existe
            if 'HojaLegal' in [hoja.name for hoja in wb.sheets]:
                hoja = wb.sheets['HojaLegal']

                
                hoja.api.Unprotect('123456')

                # Inicializa variables para almacenar las filas que contienen "Detalle" y "Total"
                fila_detalle = None
                fila_total = None

                # Obtiene todos los valores de la columna C
                columna_c = hoja.range('C1:C' + str(hoja.cells.last_cell.row)).value

                # Itera sobre los valores para encontrar las filas correspondientes
                for fila, valor in enumerate(columna_c, start=1):
                    if valor == "Detalle":
                        fila_detalle = fila
                    elif valor == "Total":
                        fila_total = fila
                        break  # Si encontramos "Total", podemos detener la búsqueda

                # Verifica que se encontraron las filas necesarias
                if fila_detalle is not None and fila_total is not None:
                    # Define la fórmula que suma el rango en la columna G
                    rango_suma = f'G{fila_detalle + 1}:G{fila_total - 1}'
                    celda_formula = hoja.range(f'G{fila_total}')
                    formula = f'=SUM({rango_suma})'
                    
                    # Inserta la fórmula en la celda correspondiente
                    celda_formula.value = formula
                    print(f'Fórmula insertada en G{fila_total} en el archivo {archivo}: {formula}')
                else:
                    print(f"No se encontraron todas las filas necesarias para insertar la fórmula en el archivo {archivo}.")

                # Guarda los cambios en el archivo
                wb.save()
            else:
                print(f"La hoja 'HojaLegal' no existe en el archivo {archivo}. No se realizaron cambios.")
