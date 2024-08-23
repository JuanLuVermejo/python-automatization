import os
import xlwings as xw
import datetime
from pathlib import Path

def obtener_valores_y_rango(archivo_excel):
    try:
        app = xw.App(visible=False)  # Ejecutar Excel en segundo plano
        wb = app.books.open(archivo_excel)
        hoja = wb.sheets['Hoja de Calculos']

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
        medidor_referencia_sap = None

        cuenta_contrato = hoja.range('C2').value
        print(f"Valor de 'N° Cuenta Contrato' (C2): {cuenta_contrato}")

        for cell in hoja.range('B1:B1000'):
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

        if fila_item is None:
            raise ValueError("No se encontró 'Item' en la columna B.")
        if fila_total_m3h is None:
            raise ValueError("No se encontró 'Total m3/h' en la columna B.")

        rango_celdas = hoja.range(f'C{fila_item + 2}:D{fila_total_m3h - 1}').value
        print(f"Rango de celdas desde 'Item' hasta 'Total m3/h': {rango_celdas}")

        for cell in hoja.range('B2:Z2'):
            if cell.value == "Tipo de Calculo":
                tipo_de_calculo = hoja.range(f'{cell.offset(0, 1).address}').value
                print(f"Valor de 'Tipo de Calculo': {tipo_de_calculo}")
            elif cell.value == "Medidor Referencia SAP":
                medidor_referencia_sap = hoja.range(f'{cell.offset(0, 1).address}').value
                print(f"Valor de 'Medidor Referencia SAP': {medidor_referencia_sap}")

        sheet_names = [sheet.name.strip() for sheet in wb.sheets]
        if 'BDDetalle' in sheet_names:
            hoja_detalle = wb.sheets['BDDetalle']
            valores_columna_a = hoja_detalle.range('A1:A1000').value
            
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
        try:
            wb.close()
            app.quit()
        except UnboundLocalError:
            pass

    return (valor_total_consumo, rango_celdas, muestra_tipica_valores, tipo_de_calculo, cuenta_contrato, tiempo_trabajo, dias_trabajo_por_mes, ultimo_valor, medidor_referencia_sap)

def insertar_datos_y_guardar(archivo_plantilla, cuenta_contrato, valor_total_consumo, rango_celdas, muestra_tipica_valores, tipo_de_calculo, ultimo_valor, tiempo_trabajo, dias_trabajo_por_mes, medidor_referencia_sap, archivo_original, macros):
    try:
        wb_plantilla = macros.app.books.open(archivo_plantilla)
        hoja_plantilla = wb_plantilla.sheets['Hoja de Calculos']

        # Convertir último valor a formato MM.YYYY
        if isinstance(ultimo_valor, (datetime.datetime, datetime.date)):
            ultimo_valor = ultimo_valor.strftime('%m.%Y')

        hoja_plantilla.range('C2').value = int(cuenta_contrato)
        hoja_plantilla.range('A1').value = valor_total_consumo
        hoja_plantilla.range('C10').value = rango_celdas
        hoja_plantilla.range('C28').value = muestra_tipica_valores[0]
        hoja_plantilla.range('D28').value = muestra_tipica_valores[1]
        hoja_plantilla.range('T2').value = tipo_de_calculo
        hoja_plantilla.range('W2').value = ultimo_valor
        hoja_plantilla.range('F10').value = tiempo_trabajo
        hoja_plantilla.range('F11').value = dias_trabajo_por_mes
        hoja_plantilla.range('P2').value = medidor_referencia_sap

        carpeta_original = os.path.dirname(archivo_original)
        carpeta_actualizados = os.path.join(carpeta_original, 'Actualizados')
        if not os.path.exists(carpeta_actualizados):
            os.makedirs(carpeta_actualizados)

        nombre_archivo_3 = f"{int(cuenta_contrato)} - {ultimo_valor} - Informes CPNO.xlsx"
        ruta_archivo_3 = os.path.join(carpeta_actualizados, nombre_archivo_3)


        # Ejecutar macros en el archivo de macros
        macros.macro('BuscarObjetivo')()
        macros.macro('IteradorCPNO')()
        print("Macros ejecutadas correctamente.")

        hoja_plantilla.range('A1').value = None

        wb_plantilla.save(ruta_archivo_3)
        print(f"Archivo guardado como: {ruta_archivo_3}")

    except Exception as e:
        print(f"Error: {e}")
    finally:
        try:
            wb_plantilla.close()
        except UnboundLocalError:
            pass

def procesar_archivo(archivo_excel, archivo_plantilla, macros):
    (val_total_consumo, rango_celdas, muestra_tipica, tipo_calculo, cuenta_contrato, 
     tiempo_trabajo, dias_trabajo, ultimo_valor, medidor_referencia_sap) = obtener_valores_y_rango(archivo_excel)

    insertar_datos_y_guardar(
        archivo_plantilla, cuenta_contrato, val_total_consumo, rango_celdas, muestra_tipica, 
        tipo_calculo, ultimo_valor, tiempo_trabajo, dias_trabajo, medidor_referencia_sap, archivo_excel, macros
    )

def procesar_todos_los_archivos(carpeta_entrada, archivo_plantilla):
    archivos_excel = [f for f in Path(carpeta_entrada).glob('*.xlsx') if 'Plantilla' not in f.name]
    
    # Abrir el archivo de macros una vez
    app_macros = xw.App(visible=False)  # Ejecutar Excel en segundo plano
    wb_macros = app_macros.books.open(r"C:\Users\juan.vermejo\Documents\CPNO\Pruebas\Macro - Buscar Objetivo Ajuste CPNO.xlsm")
    
    try:
        for archivo in archivos_excel:
            procesar_archivo(archivo, archivo_plantilla, wb_macros)

    finally:
        wb_macros.close()
        app_macros.quit()
        print("Archivo de macros cerrado.")

# Ruta de la carpeta de entrada y de la plantilla
carpeta_entrada = r'C:\Users\juan.vermejo\Documents\CPNO\Pruebas\Masivo'
archivo_plantilla = r"C:\Users\juan.vermejo\Documents\CPNO\Plantilla - Informes CPNO.xlsx"

# Procesar todos los archivos en la carpeta
procesar_todos_los_archivos(carpeta_entrada, archivo_plantilla)