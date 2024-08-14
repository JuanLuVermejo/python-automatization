import openpyxl
import xlwings as xw
from datetime import datetime
import os
import math

def leer_celda(ruta_archivo, nombre_hoja, celda):
    try:
        libro = openpyxl.load_workbook(ruta_archivo)
        hoja = libro[nombre_hoja]
        valor_celda = hoja[celda].value
        return valor_celda
    except FileNotFoundError:
        return "El archivo no fue encontrado."
    except KeyError:
        return "La hoja especificada no existe."
    except Exception as e:
        return f"Ocurrió un error: {e}"

def transformar_fecha(fecha_str):
    try:
        mes, año = fecha_str.split('.')
        fecha_transformada = f"{año}-{mes.zfill(2)}-01 00:00:00"
        return fecha_transformada
    except ValueError:
        return "Formato de fecha inválido."

def convertir_fecha_excel(fecha_str):
    try:
        fecha = datetime.strptime(fecha_str, '%Y-%m-%d %H:%M:%S')
        fecha_transformada = fecha.strftime('%Y-%m-%d %H:%M:%S')
        return fecha_transformada
    except ValueError:
        return "Formato de fecha inválido."

def eliminar_filas(ruta_archivo, valor_buscado, nombre_hoja, contraseña):
    try:
        app = xw.App(visible=False)
        libro = xw.Book(ruta_archivo)
        hoja = libro.sheets[nombre_hoja]

        if hoja.api.ProtectContents:
            hoja.api.Unprotect(contraseña)

        celdas = hoja.range('A1:A' + str(hoja.cells.last_cell.row))

        consecutivos_vacios = 0  # Contador para celdas consecutivas vacías

        for fila, celda in enumerate(celdas, start=1):
            if celda.value is None or str(celda.value).strip() == "":
                consecutivos_vacios += 1
            else:
                consecutivos_vacios = 0  # Resetear el contador si se encuentra una celda no vacía

            if consecutivos_vacios >= 2:  # Dos celdas consecutivas vacías
                print(f"Dos celdas consecutivas vacías encontradas en el archivo '{ruta_archivo}'. Saltando a siguiente informe.")
                libro.close()
                app.quit()
                return "Saltado debido a celdas consecutivas vacías."

            if str(celda.value) == str(valor_buscado):
                fila_fin = min(fila + 10, hoja.cells.last_cell.row)
                print(f"Eliminando filas de la {fila + 1} a {fila_fin - 1} en la hoja '{nombre_hoja}'.")
                hoja.api.Rows(f"{fila + 1}:{fila_fin}").Delete()

                hoja.api.Protect(contraseña)
                libro.save()
                libro.close()
                app.quit()
                return f"Se eliminaron filas debajo de la fila {fila} en la hoja '{nombre_hoja}'."

        libro.close()
        app.quit()
        return "El valor no fue encontrado en la columna A."
    except FileNotFoundError:
        return "El archivo no fue encontrado."
    except KeyError:
        return "La hoja especificada no existe."
    except Exception as e:
        return f"Ocurrió un error: {e}"

def procesar_archivos_en_carpeta(carpeta, nombre_hoja_calculos, celda, nombre_hoja_bd, contraseña):
    for archivo in os.listdir(carpeta):
        if archivo.endswith('.xlsx') or archivo.endswith('.xlsm'):
            ruta_archivo = os.path.join(carpeta, archivo)
            
            valor = leer_celda(ruta_archivo, nombre_hoja_calculos, celda)
            
            if not valor or str(valor).lower() == 'nan' or (isinstance(valor, float) and math.isnan(valor)):
                print(f"Archivo: {archivo} - Valor en la celda {celda} es inválido o 'nan'. Saltando este archivo.")
                continue

            fecha_transformada = transformar_fecha(valor)
            fecha_convertida = convertir_fecha_excel(fecha_transformada)
            
            print(f"El valor de la celda {celda} en la hoja '{nombre_hoja_calculos}' es: {fecha_convertida}")
            resultado = eliminar_filas(ruta_archivo, fecha_convertida, nombre_hoja_bd, contraseña)
            print(f"Archivo: {archivo} - {resultado}")

# Parámetros
carpeta = r'C:\Users\juan.vermejo\OneDrive - ACI PROYECTOS S.A.S. SUCURSAL DEL PERU\Documentos\CPNO - ACI\Informes Locales\Juanlu'
nombre_hoja_calculos = 'Hoja de Calculos'
celda = 'W2'
nombre_hoja_bd = 'BDDetalle'
contraseña = '123456'

procesar_archivos_en_carpeta(carpeta, nombre_hoja_calculos, celda, nombre_hoja_bd, contraseña)
