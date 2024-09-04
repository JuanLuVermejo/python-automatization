# Procesamiento de Clientes con Pandas y xlwings

Este script en Python está diseñado para procesar datos de clientes almacenados en un archivo de Excel, realizar cálculos utilizando una plantilla de Excel externa, y guardar los resultados en un nuevo archivo Excel. El código hace uso de las bibliotecas `pandas` para la manipulación de datos y `xlwings` para la automatización de Excel.

## Requisitos

- Python 3.x
- Bibliotecas: `pandas`, `xlwings`
- Archivos de Excel en las rutas especificadas

## Estructura del Código

### 1. Importación de Librerías

```python
import xlwings as xw
import pandas as pd
import os