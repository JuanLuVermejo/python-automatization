# Procesamiento de Clientes con Pandas y xlwings

Este script en Python está diseñado para procesar y actualizar datos de clientes almacenados en un archivo de Excel, realizando cálculos con una plantilla de Excel externa y guardando los resultados en un nuevo archivo Excel.

## Requisitos

- Python 3.x
- Librerías:
  - `pandas`
  - `xlwings`

### Instalación de Librerías

Para instalar las librerías necesarias, puedes utilizar `pip`. Abre una terminal y ejecuta el siguiente comando:

```bash
pip install pandas xlwings

```

## Descripción General

El script sigue los siguientes pasos principales:

1. **Carga de Datos:**
   - Se carga un archivo de Excel con información de clientes.
   - Se filtran las filas que contienen datos incompletos en las columnas `Categoria Tarifaria`, `Mes Inicio`, y `Date created`.
   - Las fechas se formatean al formato `dd/mm/yyyy`.

2. **Preparación del DataFrame:**
   - Se añaden columnas para almacenar los resultados (`intComp`, `intMor`) si no existen.

3. **Procesamiento de Clientes:**
   - Para cada cliente, se abren los datos en una plantilla de Excel utilizando `xlwings`.
   - Se insertan los datos en la hoja de cálculo y se calculan automáticamente los valores necesarios.
   - Los resultados se almacenan en el DataFrame original.

4. **Guardado del Progreso:**
   - Se guarda periódicamente el índice del último cliente procesado, lo que permite reanudar el procesamiento en caso de interrupciones.
   - Al finalizar, se guarda el DataFrame actualizado en un nuevo archivo Excel.

## Funcionalidades Clave

- **Gestión de Progreso:**
  - El script permite guardar y leer el progreso del procesamiento de clientes, lo que facilita la reanudación en caso de interrupciones.

- **Automatización de Excel:**
  - Utiliza `xlwings` para abrir, modificar y calcular automáticamente los resultados en una plantilla de Excel sin intervención manual.

- **Manejo de Errores:**
  - Se implementa control de errores para asegurar que el procesamiento continúe aun cuando se encuentra un error con un cliente en particular.

## Resultados

Al finalizar, el script genera un nuevo archivo Excel que contiene los datos originales de los clientes junto con los resultados calculados (`intComp` e `intMor`), listos para ser utilizados en análisis posteriores o reportes.

Este script es ideal para el procesamiento masivo y automatizado de datos de clientes, especialmente en entornos donde la precisión y la eficiencia son críticas.
