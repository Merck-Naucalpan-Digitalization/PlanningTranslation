from pathlib import Path
ruta_fcst = Path("C:\\Users\\x286384\\MerckGroup\\Digitalización - Documents\\DigitalProjectsLibrary\\Planning\\FCST Merck Septiembre 26.xlsx")
ruta_catalogo = Path("C:\\Users\\x286384\\MerckGroup\\Digitalización - Documents\\DigitalProjectsLibrary\\Planning\\PRODUCTOS DC.xlsx")
hoja_fcst="September 2026"
hoja_catalogo="Catalogo"
import planning_balanceo as pb
import planning_produccion as pp
import planning_acondi as pa
###planning balanceo
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
TAMANO_LOTE = 115_200      # jeringas por lote
FORMATO_MES = "%m-%Y"
agg_dict = {}
tol_jer=100

FUENTE = "Arial"
VERDE_MERCK = "0F5A3C"
FORMATO_LOTES = "#,##0.00"
FORMATO_PIEZAS = "#,##0"

###planning Produccion
from calendar import month, monthrange
from datetime import datetime, timedelta                 # jeringas por lote
PROD_TIME = {"DC1": 27, "DC2": 24}    # horas por lote
CLEAN_TIME = 8                        # horas de limpieza al cerrar una campana
# Si True, se limpia tambien despues de la ultima campana del mes.
LIMPIEZA_AL_CIERRE = False
HORA_INICIO_DEFECTO = 7               # hora del primer arranque del mes
ESQUEMA_BASE = "24/5"                 # el que siempre se planea
ESQUEMA_OPCIONAL = "24/7"             # solo si las horas alcanzan
ESQUEMAS = (ESQUEMA_BASE, ESQUEMA_OPCIONAL)
LINEAS = ("DC1", "DC2")
# Nombres que comparte con planning_balanceo.py
#esto se tiene que cambiar a id cols a medida de lo posible. 
id_cols = ['SKUMERCK','Description','DC','Nombre granel',"Linea"]
COL_GRANEL = "Nombre granel"
COL_DC = "DC"
COL_SKU = "SKUMERCK"
COL_DESC = "Description"
HOJA_FORECAST = "Forecast balanceado"
FORMATO_MES = "%m-%Y"
# Formato del Excel de salida (mismo estilo que el libro de balanceo)

FORMATO_PIEZAS = "#,##0"
FORMATO_HORAS = "#,##0.0"
FORMATO_FECHA = "dd/mm/yyyy hh:mm"

#planning Acondi
from itertools import combinations

# ---------------------------------------------------------------------------
# Parametros de negocio (la app puede pasar otros)
# ---------------------------------------------------------------------------
TAMANO_LOTE = 115_200            # jeringas por lote
HORAS_POR_LOTE = 6.5             # horas de acondicionamiento por lote
HORAS_CAMBIO_FORMATO = 9         # horas que cuesta un cambio de formato
UMBRAL_CAMBIO = 10               # relation > UMBRAL_CAMBIO pide cambio
RELATION_DESCONOCIDA = 40


LINEAS_ACONDI = ("L2", "L4")
N_EXACTO = 10                    # hasta este tamano se busca el orden optimo

# Nombres de columna que comparte con los otros modulos
COL_LINEA = "Linea"
HOJA_MD = "MD"
HOJA_MD_ENCABEZADO = 2           # la hoja MD trae dos renglones de titulo
FORMATO_MES = pp.FORMATO_MES

# Atributos de formato que definen la relation entre dos SKU
COLS_FORMATO = ["Cantidad de Blisters", "Medida de instructivo",
                "Medida de blister", "Medida de estuche"]
COLS_MD = [COL_SKU] + COLS_FORMATO + ["familia", COL_LINEA]

# Como viene marcada cada linea en la hoja MD
MARCAS_LINEA = {"L2": ("L2", "L2/L4"), "L4": ("L4", "L2/L4")}
FAMILIA_DC = "DC"

# Formato del Excel de salida (mismo estilo que los otros libros)

# Colores del cronograma
COLOR_BLOQUE_A = "DDEBF7"        # azul muy claro
COLOR_BLOQUE_B = "FCE4D6"
COLOR_LIMPIEZA = "D9D9D9"        # naranja muy claro
COLOR_CAMBIO = "D9D9D9"          # gris: cambio de formato
COLOR_FIN_SEMANA = "BFBFBF"      # gris oscuro: la linea esta detenida

ETIQUETA_CAMBIO = "Cambio de formato"
ETIQUETA_FIN_SEMANA = "Fin de semana"
COLS_CRONOGRAMA = ["Nombre", COL_SKU, "Secuencia", "Jeringas", "Lotes",
                   "Horas", "Relation previa",
                   "Fecha y hora de inicio", "Fecha y hora de finalizacion"]

fcst, catalogo = pb.leer_entradas(ruta_fcst, hoja_fcst, ruta_catalogo, hoja_catalogo)
demanda, meses=pb.preparar_demanda(fcst)
detalle, agg_dict=pb.detalle_por_sku(catalogo, demanda, meses)
def balanceo_1(detalle, meses, lote=TAMANO_LOTE):
    log_balanceo = pd.DataFrame(columns=['SKUMERCK', 'Description', 'Nombre granel','DC', 'fecha', 'log', 'Units'] + meses)
    expandido = pb.duplicar_por_dc(detalle, meses)
    fam_dc2 = expandido.loc[expandido["DC"] == "DC2", id_cols[3]].unique()
    
    for mes in meses:
        for fam in fam_dc2:
            filas = expandido.index[(expandido["DC"] == "DC2") & (expandido["Nombre granel"] == fam)]
            total = expandido.loc[filas,mes].sum()
            residuo_fam = total % lote
            #if (total%lote)<5 or (total%lote)==0 or (total%lote)>=(lote-5):
            #    continue
            
            residuo_valido = not (residuo_fam < 5 or residuo_fam == 0 or residuo_fam >= (lote - 5))

            mask_dc2 = (expandido[id_cols[3]] == fam) & (expandido["DC"] == "DC2")
            #if residuo_fam != 0:
            if residuo_valido:
                residuo = residuo_fam
                candidates = expandido[
                ((expandido[mes] >= residuo_fam-5)) & 
                (expandido[id_cols[3]] == fam) &
                (expandido[id_cols[2]] == 'DC2')][mes]
                if not candidates.empty:
                    min_idx = candidates.idxmin()
                    sku_seleccionado = expandido.loc[min_idx, id_cols[0]]
                    if expandido.loc[min_idx, mes] < residuo:
                        residuo_rnd = residuo - expandido.loc[min_idx, mes]
                        residuo=expandido.loc[min_idx,mes]

                    # Se resta el residuo de la fila para balancear DC2
                    before_dc2 = expandido.loc[[min_idx]] 
                    expandido.loc[min_idx, mes] = expandido.loc[min_idx, mes] - residuo    
                    after_dc2 = expandido.loc[[min_idx]].copy()           
                    # Find the same SKU but with DC1
                    dc1_mask = ((expandido[id_cols[0]] == sku_seleccionado) & (expandido[id_cols[2]] == 'DC1'))
                    before_dc1 = expandido.loc[dc1_mask].copy()
                    expandido.loc[dc1_mask, mes] += residuo
                    after_dc1 = expandido.loc[dc1_mask].copy()

                    columnas_a_cero = [m for m in meses if m != mes]
                    for df, label in zip([before_dc2, before_dc1, after_dc2, after_dc1],
                        ['1. before_dc2', '1. before_dc1', '1. after_dc2', '1. after_dc1']):
                        df['log'] = label
                        df['fecha'] = mes
                        df[columnas_a_cero] = 0
                    # Append all to log
                    log_balanceo = pd.concat([log_balanceo, before_dc2, before_dc1, after_dc2, after_dc1], ignore_index=True)
            total_dc2 = expandido.loc[filas, mes].sum()
            if total_dc2 != 0 and abs(total_dc2 / lote) < ((lote - 10) / lote):
                skus_fam = expandido.loc[filas, "SKUMERCK"].unique()
                for sku in skus_fam:
                    idx_dc2 = expandido.index[
                        (expandido["DC"] == "DC2") &
                        (expandido["SKUMERCK"] == sku) &
                        (expandido["Nombre granel"] == fam)]
                    idx_dc1 = expandido.index[
                        (expandido["DC"] == "DC1") &
                        (expandido["SKUMERCK"] == sku) &
                        (expandido["Nombre granel"] == fam)]
                    piezas_dc2 = expandido.at[idx_dc2[0], mes]
                    expandido.at[idx_dc1[0], mes] += piezas_dc2
                    expandido.at[idx_dc2[0], mes] = 0.0
    
    for mes in meses:
        if log_balanceo[mes].sum() == 0:
            log_balanceo = log_balanceo.drop(columns=[mes])
    residuos = expandido.copy()
    #correciones de unidades
    
    for mes in meses:
        for  sku in expandido["SKUMERCK"].unique():
            idx_dc2 = expandido.index[
                    (expandido["SKUMERCK"] == sku) &
                    (expandido[mes] > 0)]
            for idx in idx_dc2:
                unidad=expandido.at[idx, mes] % 10
                if unidad == 0:
                    continue
                print(f"SKU: {sku}, Mes: {mes},DC: {expandido.at[idx, 'DC']}, Unidad a corregir: {unidad}")
                if unidad > 0 and unidad < 6:
                    expandido.at[idx, mes] = expandido.at[idx, mes] - unidad
                if unidad > 5:
                    expandido.at[idx, mes] = expandido.at[idx, mes] + (10 - unidad)
    excedente = expandido[meses] % lote
    residuos[meses] = excedente
    
    
    
    expandido.to_csv("balanceado.csv", index=False)
    log_balanceo.to_csv("log_balanceo.csv", index=False)
    residuos.to_csv("residuos.csv", index=False)

    return expandido, log_balanceo,residuos                
#swap implementation for balancing DC2 to DC1 and converting to lots               
detalle_balanceado, log_balanceo, residuos = balanceo_1(detalle, meses,lote=TAMANO_LOTE)
lotes_producto=lotes_producto = pb.consolidar_por_producto(detalle_balanceado, meses, lote=TAMANO_LOTE)
lotes_producto.to_csv("lotes_producto.csv", index=False)
print(lotes_producto)