import os
import pandas as pd

# Define la ruta de la carpeta raíz que contiene las subcarpetas
carpeta_raiz = r'C:\Users\juan.vermejo\OneDrive - Gas Natural de Lima y Callao S.A. (GNLC)\Reporteria\97. Informes CPNO'
ruta_salida = r'C:\Users\juan.vermejo\Documents\CPNO\Consolidado_BDPercy.xlsx'

# Inicializamos una lista para almacenar los datos de cada archivo
datos_consolidados = []

# Recorremos todas las subcarpetas y archivos en la carpeta raíz
for ruta_directorio, subdirs, archivos in os.walk(carpeta_raiz):
    for archivo in archivos:
        # Filtra solo archivos de Excel
        if archivo.endswith('.xlsx'):
            ruta_archivo = os.path.join(ruta_directorio, archivo)
            
            try:
                # Lee el archivo Excel y obtiene la hoja "BDPercy", comenzando desde la fila 3
                df = pd.read_excel(ruta_archivo, sheet_name='BDPercy', header=2)
                
                # Filtramos solo la primera fila de datos de la hoja
                if not df.empty:
                    # Creamos un diccionario con los datos y añadimos el nombre de la carpeta y del archivo
                    datos_fila = df.iloc[0].to_dict()
                    datos_fila['Carpeta'] = os.path.basename(ruta_directorio)
                    datos_fila['Archivo'] = archivo
                    datos_consolidados.append(datos_fila)
            
            except Exception as e:
                print(f"Error en el archivo {archivo} en la ruta {ruta_directorio}: {e}")

# Crear un DataFrame con los datos consolidados
df_consolidado = pd.DataFrame(datos_consolidados)

# Guardar el DataFrame consolidado en un archivo Excel
df_consolidado.to_excel(ruta_salida, index=False)

print(f"Consolidación completada. Archivo guardado en {ruta_salida}")
