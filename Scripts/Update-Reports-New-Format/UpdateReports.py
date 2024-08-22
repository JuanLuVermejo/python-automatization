import xlwings as xw

def obtener_valores_y_rango(archivo_excel):
    try:
        # Intentar abrir el archivo Excel
        wb = xw.Book(archivo_excel)
        
        # Seleccionar la hoja "Hoja de Calculos"
        hoja = wb.sheets['Hoja de Calculos']

        # Inicializar variables para almacenar los resultados
        valor_total_consumo = None
        fila_item = None
        fila_total_m3h = None
        muestra_tipica_valores = None
        tipo_de_calculo = None
        cuenta_contrato = None
        rango_celdas = None
        ultimo_valor = None
        tiempo_trabajo = None
        dias_trabajo_por_mes = None

        # Recopilar el valor de la celda C2
        cuenta_contrato = hoja.range('C2').value
        print(f"Valor de 'N° Cuenta Contrato' (C2): {cuenta_contrato}")

        # Buscar el valor en la columna B
        for cell in hoja.range('B1:B1000'):  # Asumiendo que no hay más de 1000 filas
            if cell.value in ["TOTAL CONSUMO RECUPERADO", "TOTAL CONSUMO RELIQUIDACIÓN"]:
                valor_total_consumo = hoja.range(f'D{cell.row}').value
                print(f"Valor de 'TOTAL CONSUMO': {valor_total_consumo}")
            elif cell.value == "Item":
                fila_item = cell.row
            elif cell.value == "Total m3/h":
                fila_total_m3h = cell.row
            elif cell.value == "Muestra Tipica":
                muestra_tipica_valores = (hoja.range(f'C{cell.row}').value, hoja.range(f'D{cell.row}').value)
                print(f"Valores de 'Muestra Tipica': C{cell.row}={muestra_tipica_valores[0]}, D{cell.row}={muestra_tipica_valores[1]}")
            elif cell.value == "Tiempo de trabajo (h)":
                tiempo_trabajo = hoja.range(f'F{cell.row}').value
                print(f"Valor de 'Tiempo de trabajo (h)': {tiempo_trabajo}")
            elif cell.value == "Días de Trabajo por Mes":
                dias_trabajo_por_mes = hoja.range(f'F{cell.row}').value
                print(f"Valor de 'Días de Trabajo por Mes': {dias_trabajo_por_mes}")

        # Validar que se hayan encontrado las filas necesarias
        if fila_item is None:
            raise ValueError("No se encontró 'Item' en la columna B.")
        if fila_total_m3h is None:
            raise ValueError("No se encontró 'Total m3/h' en la columna B.")

        # Definir el rango basado en las filas encontradas
        rango_celdas = hoja.range(f'C{fila_item + 1}:D{fila_total_m3h - 1}').value
        print(f"Rango de celdas desde 'Item' hasta 'Total m3/h': {rango_celdas}")

        # Validar si existe la hoja "BDDetalle"
        sheet_names = [sheet.name.strip() for sheet in wb.sheets]
        
        if 'BDDetalle' in sheet_names:
            hoja_detalle = wb.sheets['BDDetalle']
            
            # Obtener el último valor no vacío en la columna A
            valores_columna_a = hoja_detalle.range('A1:A1000').value  # Asumiendo que no hay más de 1000 filas
            
            for valor in reversed(valores_columna_a):
                if valor not in (None, '', 'Null'):
                    ultimo_valor = valor
                    break

            print(f"Último valor en la columna A de 'BDDetalle': {ultimo_valor}")
        else:
            print(f"La hoja 'BDDetalle' no existe en este archivo. Hojas disponibles: {sheet_names}")

    except Exception as e:
        print(f"Error: {e}")
    finally:
        # Asegurarse de cerrar el archivo Excel si fue abierto correctamente
        try:
            wb.close()
        except UnboundLocalError:
            pass

    # Retornar los valores recopilados
    return (valor_total_consumo, rango_celdas, muestra_tipica_valores, 
            tipo_de_calculo, cuenta_contrato, tiempo_trabajo, dias_trabajo_por_mes, ultimo_valor)

# Definir la ruta al archivo de Excel
archivo_excel = r'C:\Users\juan.vermejo\Documents\CPNO\Pruebas\Masivo\3035430 - 05.2024 - Informes CPNO.xlsx'

# Ejecutar la función y obtener los valores
(val_total_consumo, rango_celdas, muestra_tipica, tipo_calculo, cuenta_contrato, 
 tiempo_trabajo, dias_trabajo, ultimo_valor) = obtener_valores_y_rango(archivo_excel)
