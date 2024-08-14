import xlwings as xw
import pandas as pd
import os

# Rutas de archivos
clientes_file = r'C:\Users\juan.vermejo\Documents\CPNO\Mst - Consolidado de Informes.xlsx'
calculadora_file = r'C:\Users\juan.vermejo\Documents\CPNO\Pruebas\Formato de calculo CPNO_vv1.7_CC (Actualizado Julio) TMP.xlsx'
archivo_progreso = r'C:\Users\juan.vermejo\Documents\CPNO\Pruebas\clientes_actualizados.xlsx'
indice_progreso_file = r'C:\Users\juan.vermejo\Documents\CPNO\Pruebas\indice_progreso.txt'

# Cargar el archivo de clientes usando pandas
clientes_df = pd.read_excel(clientes_file, sheet_name='Informes Realizados', header=2)

# Filtrar los registros donde 'Categoria Tarifaria', 'Mes Inicio', y 'Date created' no estén vacíos
clientes_df = clientes_df.dropna(subset=['Categoria Tarifaria', 'Mes Inicio', 'Date created'])

# Convertir fechas a formato dd/mm/yyyy como texto
clientes_df['Mes Inicio'] = pd.to_datetime(clientes_df['Mes Inicio'], errors='coerce').dt.strftime('%d/%m/%Y')
clientes_df['Date created'] = pd.to_datetime(clientes_df['Date created'], errors='coerce') + pd.offsets.MonthEnd(0)

# Convertir 'Date created' a texto en formato dd/mm/yyyy
clientes_df['Date created'] = clientes_df['Date created'].dt.strftime('%d/%m/%Y')

# Convertir 'Mes Inicio' a cadena y filtrar registros para excluir aquellos con año 1900
clientes_df['Mes Inicio'] = clientes_df['Mes Inicio'].astype(str)
clientes_df = clientes_df[~clientes_df['Mes Inicio'].str.contains('1900', na=False)]

# Añadir nuevas columnas para los resultados si no existen
if 'intComp' not in clientes_df.columns:
    clientes_df['intComp'] = None
if 'intMor' not in clientes_df.columns:
    clientes_df['intMor'] = None

def guardar_progreso(indice):
    """Guarda el índice de la última fila procesada en un archivo."""
    with open(indice_progreso_file, 'w') as file:
        file.write(str(indice))

def leer_indice_progreso():
    """Lee el índice de la última fila procesada desde el archivo, si existe."""
    if os.path.exists(indice_progreso_file):
        with open(indice_progreso_file, 'r') as file:
            contenido = file.read().strip()
            if contenido.isdigit():  # Verifica si el contenido es un número
                return int(contenido)
    return -1  # Devuelve -1 si el archivo no existe o el contenido no es un número válido

# Leer el índice de la última fila procesada
indice_progreso = leer_indice_progreso()

# Función para procesar cada cliente
def procesar_cliente(index, fila_cliente):
    try:
        # Obtener los datos del cliente desde el DataFrame
        dato1 = fila_cliente['Categoria Tarifaria']
        dato2 = fila_cliente['Mes Inicio']
        dato3 = fila_cliente['Date created']
        dato4 = fila_cliente['Volumen a Recuperar']

        # Usar xlwings para abrir el archivo de la calculadora y copiar los datos
        app = xw.App(visible=False)
        wb = app.books.open(calculadora_file)
        sheet = wb.sheets['calculadora']

        # Pegar los datos en las celdas correspondientes de la hoja calculadora
        sheet.range('C3').value = dato1
        sheet.range('D3').value = "'" + dato2  # Agregar comilla simple para asegurarse de que se pegue como texto
        sheet.range('E3').value = "'" + dato3  # Agregar comilla simple para asegurarse de que se pegue como texto
        sheet.range('F3').value = dato4

        # Forzar el cálculo de la hoja
        sheet.api.Calculate()

        # Obtener los resultados de las fórmulas
        intComp = sheet.range('N3').value
        intMor = sheet.range('O3').value

        # Guardar los resultados en el DataFrame de clientes
        clientes_df.at[index, 'intComp'] = intComp
        clientes_df.at[index, 'intMor'] = intMor

        # Imprimir los resultados en la terminal
        print(f"Cliente {index + 1}: Categoria Tarifaria = {dato1}, Mes Inicio = {dato2}, Date created (fin de mes) = {dato3}, Volumen a Recuperar = {dato4}, intComp = {intComp}, intMor = {intMor}")

        # Cerrar el libro sin guardar cambios
        wb.close()
        app.quit()
    except Exception as e:
        print(f"Error procesando el cliente en el índice {index}: {e}")

# Recorrer cada fila del DataFrame de clientes y procesar los datos
for index, fila in clientes_df.iterrows():
    if index > indice_progreso:  # Empezar desde el índice guardado
        procesar_cliente(index, fila)
        if (index + 1) % 10 == 0:  # Guardar progreso cada 10 clientes procesados
            guardar_progreso(index)

# Guardar el DataFrame de clientes con los resultados actualizados en un nuevo archivo Excel al finalizar
guardar_progreso(len(clientes_df) - 1)
clientes_df.to_excel(archivo_progreso, index=False)
