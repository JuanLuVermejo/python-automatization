import os
import argparse
import win32com.client as win32

# Función para bloquear hojas
def lock_sheets(file_path, sheet_names, password):
    try:
        # Iniciar una instancia de Excel en segundo plano
        excel = win32.DispatchEx('Excel.Application')
        excel.Visible = False  # Asegurar que Excel no sea visible
        excel.DisplayAlerts = False  # Deshabilitar las alertas de Excel

        # Abrir el archivo
        workbook = excel.Workbooks.Open(file_path)
        
        # Bloquear las hojas
        for sheet_name in sheet_names:
            try:
                sheet = workbook.Sheets(sheet_name)
                if not sheet.ProtectContents:
                    sheet.Protect(Password=password)
            except Exception as e:
                print(f"Error al proteger la hoja {sheet_name} en el archivo {file_path}: {e}")
        
        # Guardar el archivo con las modificaciones
        workbook.Save()
        workbook.Close(SaveChanges=True)
        excel.Quit()
        
        print(f"Archivo modificado: {file_path}")
        
    except Exception as e:
        print(f"Error al modificar {file_path}: {e}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Bloquear hojas en archivos de Excel.')
    parser.add_argument('--path', type=str, default=r'C:\Users\juan.vermejo\Documents\CPNO\00. Informes por Analizar', help='Ruta de la carpeta que contiene los archivos de Excel.')
    parser.add_argument('--password', type=str, default='123456', help='Contraseña para proteger las hojas.')
    args = parser.parse_args()

    folder_path = args.path
    password = args.password
    sheets_to_lock = ["BD", "BDDetalle", "Hoja de Calculos", "BDPercy"]

    # Obtener la lista de archivos a modificar
    files_to_modify = [os.path.join(folder_path, filename) for filename in os.listdir(folder_path) if filename.endswith('.xlsx')]

    # Procesar cada archivo secuencialmente
    for file_path in files_to_modify:
        lock_sheets(file_path, sheets_to_lock, password)

    print("Bloqueo completado.")
