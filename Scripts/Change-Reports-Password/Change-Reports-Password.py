import xlwings as xw
import os
import argparse

# Argument parser para obtener la nueva contraseña desde la línea de comandos
def get_arguments():
    parser = argparse.ArgumentParser(description="Desbloquear y bloquear hojas en archivos Excel.")
    parser.add_argument('--newPass', type=str, default="123456", help="Nueva contraseña para bloquear las hojas (por defecto es '123456').")
    return parser.parse_args()

# Función que desbloquea y vuelve a bloquear las hojas
def process_excel(file_path, new_password, app):
    try:
        # Abrir el archivo de Excel en segundo plano
        wb = app.books.open(file_path)
        # Imprimir el nombre y la ruta del archivo que está siendo procesado
        print(f"Procesando archivo: '{file_path}'")
        
        # Lista de hojas objetivo
        sheet_names = ["BD", "BDDetalle", "Hoja de Calculos", "BDPercy", "BDDimensiones"]
        
        # Bandera para saber si se realizó alguna modificación
        modified = False
        
        for sheet_name in sheet_names:
            try:
                # Desbloquear la hoja si está protegida
                sheet = wb.sheets[sheet_name]
                if sheet.api.ProtectContents:
                    sheet.api.Unprotect(Password="123456")
                    print(f"Hoja '{sheet_name}' desbloqueada.")
                    modified = True
                
                # Volver a proteger la hoja con la nueva contraseña
                sheet.api.Protect(Password=new_password)
                print(f"Hoja '{sheet_name}' bloqueada con nueva contraseña.")
            
            except Exception as e:
                print(f"No se pudo procesar la hoja '{sheet_name}' en el archivo '{file_path}': {e}")
        
        # Guardar los cambios solo si se realizó alguna modificación
        if modified:
            wb.save()
            print(f"Modificaciones guardadas en el archivo '{file_path}'.\n")
        else:
            print(f"No se realizaron modificaciones en el archivo '{file_path}'.\n")
        
        # Cerrar el libro
        wb.close()
    
    except Exception as e:
        print(f"Error al procesar el archivo '{file_path}': {e}")

# Función para procesar todos los archivos en las carpetas
def process_folder(folder_path, new_password):
    # Crear una instancia de Excel en segundo plano
    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False
    
    # Recorrer todas las carpetas y archivos dentro de la carpeta principal
    try:
        for subdir, _, files in os.walk(folder_path):
            for file in files:
                if file.endswith(".xlsx"):
                    file_path = os.path.join(subdir, file)
                    process_excel(file_path, new_password, app)
    finally:
        # Cerrar la aplicación de Excel
        app.quit()

if __name__ == "__main__":
    args = get_arguments()
    new_password = args.newPass
    folder_path = r"C:\Users\juan.vermejo\Documents\CPNO\Pruebas\Masivo"
    process_folder(folder_path, new_password)
