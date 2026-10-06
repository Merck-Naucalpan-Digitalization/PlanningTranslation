

from calendar import monthrange
from datetime import datetime, timedelta
from pathlib import Path
from tokenize import group

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
productos=pd.read_excel("Fresh export 6.xlsx", sheet_name="Forecast balanceado")
catalogo_dc= pd.read_excel("PRODUCTOS DC.xlsx", sheet_name="Catalogo")
PROD_TIME = {"DC1": 27, "DC2": 24}   # horas por lote
CLEAN_TIME = 8                       # horas de limpieza por corrida
MAX_CAMPAIGN_lotes= 3                # lotes maximos por campana (mismo granel)
MAX_CAMPAIGN_JER=3*(115200)

HORA_INICIO_DEFECTO = 7               # hora del primer arranque del mes
ESQUEMAS = ("24/7", "24/5")
REDONDEOS = ("ceil", "floor", "round")

# ---------------------------------------------------------------------------
# Parametros de negocio (la app puede pasar otros)
# ---------------------------------------------------------------------------
PROD_TIME = {"DC1": 27, "DC2": 24}   # horas por lote
CLEAN_TIME = 8                        # horas de limpieza por corrida
MAX_CAMPAIGN_lotes= 3                      # lotes maximos por campana (mismo granel)
MAX_CAMPAIGN_JER=3*(115200)

# Si True, se limpia tambien despues de la ultima corrida del mes.
CLEANING_COUNTS_LAST_RUN = False

HORA_INICIO_DEFECTO = 7               # hora del primer arranque del mes
ESQUEMAS = ("24/7", "24/5")
REDONDEOS = ("ceil", "floor", "round")

# Nombres que comparte con planning_balanceo.py
COL_GRANEL = "Nombre granel"
COL_DC = "DC"
HOJA_LOTES = "Lotes por producto"
FORMATO_MES = "%m-%Y"

# Formato del Excel de salida (mismo estilo que el libro de balanceo)
FUENTE = "Arial"
VERDE_MERCK = "0F5A3C"
FORMATO_LOTES = "#,##0"
FORMATO_HORAS = "#,##0.0"
FORMATO_FECHA = "dd/mm/yyyy hh:mm"
FORMATO_MES = "%m-%Y"
def columnas_de_mes(columnas):
    """Devuelve los encabezados de mes que son demanda futura.
    Se recorren los encabezados de derecha a izquierda y se guardan los que
    pandas pueda leer como fecha, hasta topar con el primero que ya no es
    futuro. Como el forecast trae dos bloques de meses (cajas y piezas), este
    recorrido inverso cae en el segundo bloque, que es el de piezas.
    """
    hoy = pd.Timestamp.today()
    seleccion = []
    for columna in reversed(list(columnas)):
        fecha = pd.to_datetime(columna, errors="coerce")
        if pd.isna(fecha):
            continue
        seleccion.append(columna)
        if (fecha.year, fecha.month) <= (hoy.year, hoy.month):
            break
    seleccion.reverse()

    fechas = pd.to_datetime(pd.Series(seleccion), errors="coerce")
    if not seleccion or fechas.dt.year.min() < 2000:
        raise ValueError("No se detectaron encabezados de mes validos en el "
                         "forecast. Revise que los encabezados de mes sean "
                         "fechas de Excel.")
    meses = [f.strftime(FORMATO_MES) for f in fechas]
    return seleccion, meses

def mes_a_anio_mes(mes):
    """Convierte la etiqueta de mes del balanceo ('08-2026') en (2026, 8)."""
    fecha = pd.to_datetime(mes, format=FORMATO_MES, errors="coerce")
    if pd.isna(fecha):
        fecha = pd.to_datetime(mes, errors="coerce")
    if pd.isna(fecha):
        raise ValueError(f"No se pudo interpretar el mes '{mes}'. "
                         f"Se espera el formato MM-AAAA.")
    return int(fecha.year), int(fecha.month)
def inicio_por_defecto(mes):
    """Primer arranque del mes, a las 07:00 del dia 1."""
    anio, numero = mes_a_anio_mes(mes)
    return datetime(anio, numero, 1, HORA_INICIO_DEFECTO, 0)
def union_concentracion(productos,catalogo_dc):
    productos = productos.merge(
    catalogo_dc[["SKUMERCK", "camp"]].drop_duplicates(subset=["SKUMERCK"]),
    on="SKUMERCK", how="left")
    productos.insert(5, "camp", productos.pop("camp"))
    return productos

def campaign_manage(productos,meses):
    #dc1 campaign
    productos= productos.sort_values(["Nombre granel", "DC",],
                              ascending=[True, True])
    count = 0
    for mes in meses:
        print(mes)
        for dc in productos.groupby("DC"):
            print(dc)
            for fam in productos.groupby("Nombre granel"):
                print(fam)
                for prod in productos.groupby("SKUMERCK"):
                    print(prod)


union_concentracion(productos,catalogo_dc)
seleccion, meses = columnas_de_mes(productos.columns)
prod=campaign_manage(productos,meses)
