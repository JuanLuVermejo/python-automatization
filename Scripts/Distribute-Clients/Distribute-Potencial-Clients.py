## BOLSA GERALDINE

import pandas as pd

# Ruta del archivo Excel original y la nueva ruta donde se guardará el archivo
archivo_excel = r"C:\Users\juan.vermejo\Documents\CPNO\BolsaLibre.xlsx"
nuevo_archivo = r"C:\Users\juan.vermejo\Documents\CPNO\BolsaLibre_Asignado.xlsx"

# Leer el archivo Excel usando pandas
df = pd.read_excel(archivo_excel, sheet_name="Cruce")

# Filtrar los registros que tienen FALSO en la columna "Cerrado con Deposito?"
df_filtrado = df[df['Cerrado con Deposito?'] == False]
df_filtrado = df[df['Esta en agenda IBR?'] == False]
df_filtrado = df[df['Estatus Gestión Comercial'] != "Se firmó acuerdo"]


# Aplicar más filtros:
# 1. Eliminar registros con valores vacíos en "InformesElaborados.Ingreso CPNO esperado"
# 2. Excluir registros donde "InformesElaborados.Ingreso CPNO esperado" sea igual a 590 o 1180
df_filtrado = df_filtrado[df_filtrado['InformesElaborados.Ingreso CPNO esperado'].notna()]  # Filtrar valores no vacíos
df_filtrado = df_filtrado[~df_filtrado['InformesElaborados.Ingreso CPNO esperado'].isin([590, 1180])]  # Excluir 590 y 1180

# Verificar si el DataFrame no está vacío después del filtrado
if df_filtrado.empty:
    print("No se encontraron registros que cumplan con los criterios de filtrado.")
else:
    # Nombres de las personas para asignar clientes
    personas = ['Jhonny Llontop', 'Manuel Huambachano', 'Fraciny Minaya', 'Jackeline Ferro']

    # Ordenar los clientes por el aporte dinerario "InformesElaborados.Ingreso CPNO esperado" de mayor a menor
    df_filtrado = df_filtrado.sort_values(by='InformesElaborados.Ingreso CPNO esperado', ascending=False).reset_index(drop=True)

    # Número de personas y número total de clientes
    num_personas = len(personas)
    num_clientes = len(df_filtrado)

    # Inicializar listas para almacenar el total de aportes de cada persona
    aportes_personas = {persona: 0 for persona in personas}
    asignaciones = []  # Lista para guardar las asignaciones de cada cliente

    # Distribuir los clientes de manera equitativa pero balanceando el dinero
    for i in range(num_clientes):
        # Calcular qué persona tiene actualmente el menor total de aportes
        persona_asignada = min(aportes_personas, key=aportes_personas.get)

        # Asignar el cliente a esa persona
        asignaciones.append(persona_asignada)

        # Sumar el aporte del cliente al total acumulado de la persona asignada
        aportes_personas[persona_asignada] += df_filtrado.loc[i, 'InformesElaborados.Ingreso CPNO esperado']

    # Agregar la columna 'Persona' con las asignaciones al DataFrame
    df_filtrado['Persona'] = asignaciones

    # Guardar el DataFrame con la nueva columna en un archivo Excel nuevo
    df_filtrado.to_excel(nuevo_archivo, index=False)

    print(f"Asignación completada y archivo guardado en: {nuevo_archivo}")
