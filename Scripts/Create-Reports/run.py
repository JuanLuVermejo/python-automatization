import os
import win32com.client as win32
import pandas as pd
import time

# Ruta y configuración inicial
id_pc = r'C:\Users\juan.vermejo'
base_ruta = r'Documents\CPNO'
nube = r'C:\Users\juan.vermejo\OneDrive - ACI PROYECTOS S.A.S. SUCURSAL DEL PERU\Documentos\CPNO - ACI'
base_ruta_pc = os.path.join(id_pc, base_ruta)
base_percy_path = os.path.join(nube, 'Mst - Consolidado de Informes.xlsx')
print(base_percy_path)
macro_file_path = os.path.join(base_ruta_pc, 'Macro - Informes CPNO (NO BORRAR).xlsm')
base_path = os.path.join(base_ruta_pc, '00. Informes por Analizar')
template_path = os.path.join(base_ruta_pc, 'Plantilla - Informes CPNO.xlsx')


# Pausa para asegurar que todos los archivos externos están listos antes de proceder
time.sleep(5)

# Leer el archivo Excel y extraer datos necesarios
try:
    df = pd.read_excel(base_percy_path, sheet_name='A Trabajar')

    # Verificar que las columnas necesarias están en el DataFrame
    required_columns = ['Cuenta Contrato', 'FECHA CORTE']
    if not all(col in df.columns for col in required_columns):
        raise ValueError("Las columnas requeridas no están disponibles en el DataFrame.")

    # Convertir 'FECHA CORTE' al formato MM-YYYY, asignar None si falla
    df['FECHA CORTE'] = pd.to_datetime(df['FECHA CORTE'], errors='coerce').dt.strftime('%m.%Y').replace('NaT', None)

    # Convertir 'Cuenta Contrato' a int, asignar None si falla
    df['Cuenta Contrato'] = pd.to_numeric(df['Cuenta Contrato'], errors='coerce').astype('Int64')

    # Eliminar filas donde 'Cuenta Contrato' es None para asegurar la integridad de los datos
    df.dropna(subset=['Cuenta Contrato'], inplace=True)

    # Convertir los datos limpios a lista para procesamiento posterior
    cuenta_contrato = df[['Cuenta Contrato', 'FECHA CORTE']].astype(str).values.tolist()

except Exception as e:
    print(f"Error al leer o procesar el archivo Excel: {e}")
    raise

# Imprimir las primeras 10 filas para verificación
print(cuenta_contrato[:10])

# Verificar si el archivo de macros existe
if not os.path.exists(macro_file_path):
    raise FileNotFoundError(f"El archivo de macros no se encuentra en la ruta especificada: {macro_file_path}")

# Configurar la aplicación Excel, abrirla en modo invisible y desactivar alertas
try:
    excel = win32.DispatchEx('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False

    # Abrir el libro que contiene las macros
    try:
        macro_workbook = excel.Workbooks.Open(macro_file_path)
    except Exception as e:
        raise Exception(f"No se pudo abrir el archivo de macros: {e}")

    # Iterar sobre cada cuenta y fecha
    for cc in cuenta_contrato:
        cc_str = f"{cc[0]} - {cc[1]}"  # Concatenar cuenta y fecha para evitar conflictos en los nombres de archivos

        # Crear ruta completa para el nuevo archivo
        new_file_path = os.path.join(base_path, f'{cc_str} - Informes CPNO.xlsx')

        # Abrir la plantilla y asignar valores
        try:
            data_workbook = excel.Workbooks.Open(template_path)
        except Exception as e:
            print(f"No se pudo abrir la plantilla: {e}")
            continue  # Saltar a la siguiente iteración si no se puede abrir la plantilla

        data_sheet = data_workbook.Sheets('Hoja de Calculos')
        data_sheet.Activate()
        data_sheet.Range('C2').Value = cc[0]  # Asignar cuenta contrato a la celda C2
        data_sheet.Range('W2').Value = cc[1]  # Asignar fecha corte a la celda W2

        # Ejecutar macro que procesa los informes
        try:
            excel.Run("'Macro - Informes CPNO (NO BORRAR).xlsm'!IteradorCPNO")
        except Exception as e:
            print(f"No se pudo ejecutar la macro: {e}")
            data_workbook.Close(False)  # Cerrar sin guardar si ocurre un error
            continue  # Saltar a la siguiente iteración si la macro falla
        
        # Intentar proteger las hojas especificadas
        try:
            sheet_BD = data_workbook.Sheets("BD")
            sheet_BDDimensiones = data_workbook.Sheets("BDDetalle")
            sheet_BD.Protect('123456')
            sheet_BDDimensiones.Protect('123456')
        except Exception as e:
            print(f"Ocurrió un error al proteger las hojas: {e}")
        
        # Guardar y cerrar el archivo con la información actualizada
        try:
            data_workbook.SaveAs(new_file_path, FileFormat=51)  # 51 corresponde al formato .xlsx
            data_workbook.Close()
        except Exception as e:
            print(f"Error al guardar el archivo {new_file_path}: {e}")
            data_workbook.Close(False)  # Cerrar sin guardar si ocurre un error

    # Cerrar Excel y liberar recursos
    macro_workbook.Close()
    excel.Quit()
    del excel

except Exception as e:
    print(f"Error al manejar archivos de Excel: {e}")
    if 'excel' in locals():
        excel.Quit()
    raise