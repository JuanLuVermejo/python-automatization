import os
import pandas as pd

# Ruta de la carpeta con los archivos Excel
carpeta_excel = r'C:\Users\juan.vermejo\Documents\CPNO\Pruebas\Masivo\Comedores Populares'

# Diccionario de mapeo de encabezados equivalentes
mapeo_columnas = {
    "GIRO": "Clase de PS",
    "TARIFA SAP": "Tarifa (AGOSTO)",
    "Fecha Alta": "Fecha Habilitacion",
    "CLIENTE": "Nombres y Apellidos"
}

# Inicializa un DataFrame vacío para almacenar toda la información
df_final = pd.DataFrame()

# Itera sobre los archivos de la carpeta
for archivo in os.listdir(carpeta_excel):
    if archivo.endswith('.xlsx'):  # Filtra solo archivos .xlsx
        ruta_archivo = os.path.join(carpeta_excel, archivo)
        
        # Extraer la "Cuenta Contrato" del nombre del archivo
        cuenta_contrato = archivo.split(' - ')[0]
        
        # Lee el archivo Excel en la hoja 'BD' con encabezados en la fila 2 (índice 1 en pandas)
        df = pd.read_excel(ruta_archivo, sheet_name='BD', header=1)
        
        # Renombrar las columnas según el mapeo definido
        df = df.rename(columns=mapeo_columnas)
        
        # Filtrar la fila donde la columna "CC" coincida con la "Cuenta Contrato"
        fila_filtrada = df[df['CC'] == int(cuenta_contrato)]
        
        # Verificar si se encontró la fila
        if not fila_filtrada.empty:
            # Concatenar la fila filtrada con el DataFrame final
            df_final = pd.concat([df_final, fila_filtrada], ignore_index=True)
        else:
            print(f"No se encontró la Cuenta Contrato {cuenta_contrato} en el archivo {archivo}")

# Ruta donde se guardará el DataFrame final
ruta_salida = r'C:\Users\juan.vermejo\Documents\CPNO\Pruebas\Masivo\Resultados\Mst - lecturas cluster comedores 2 (Inf Originales).xlsx'

# Guardar el DataFrame final en un archivo Excel
df_final.to_excel(ruta_salida, index=False)

# Confirmación final
print(f"Archivo guardado exitosamente en: {ruta_salida}")
