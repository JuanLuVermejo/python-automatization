import os
import xlwings as xw
import argparse

# Configuración de argumentos de línea de comandos
parser = argparse.ArgumentParser(description="Exportar hojas de Excel como PDF")
parser.add_argument("--path", type=str, help="Ruta de la carpeta con archivos Excel", default=r"C:\Users\juan.vermejo\Documents\CPNO\Pruebas\Masivo\Actualizados")
args = parser.parse_args()

# Ruta de la carpeta, ya sea proporcionada o la ruta por defecto
carpeta_excel = args.path

# Itera sobre todos los archivos en la carpeta
for archivo in os.listdir(carpeta_excel):
    if archivo.endswith(".xlsx"):  # Filtra solo los archivos Excel
        ruta_excel = os.path.join(carpeta_excel, archivo)
        ruta_pdf = ruta_excel.replace(".xlsx", ".pdf")  # Genera el nombre del PDF en la misma carpeta

        try:
            # Abre el archivo de Excel
            wb = xw.Book(ruta_excel)

            # Selecciona la hoja "HojaLegal"
            hoja = wb.sheets['HojaLegal']

            # Exporta la hoja como PDF
            hoja.api.ExportAsFixedFormat(0, ruta_pdf)

            print(f"Exportación a PDF realizada con éxito: {ruta_pdf}")

        except Exception as e:
            print(f"Ocurrió un error con el archivo {archivo}: {e}")

        finally:
            # Cierra el libro de Excel sin guardar cambios
            if 'wb' in locals():
                wb.close()
