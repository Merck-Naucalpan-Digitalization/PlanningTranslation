#testing gorunds
import os
import queue
import subprocess
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path
import pandas as pd

# Importacion de funciones y variables de la logica.
import planning_balanceo as pb
import planning_produccion as pp
ruta_fcst = Path("C:\\Users\\x286384\\MerckGroup\\Digitalización - Documents\\DigitalProjectsLibrary\\Planning\\FCST Merck Septiembre 26.xlsx")
ruta_catalogo = Path("C:\\Users\\x286384\\MerckGroup\\Digitalización - Documents\\DigitalProjectsLibrary\\Planning\\PRODUCTOS DC.xlsx")
hoja_fcst = "September 2026"
hoja_catalogo = "Catalogo"
hoja_paros='Paros Programados'
ruta_salida = Path("C:\\Users\\x286384\\MerckGroup\\Digitalización - Documents\\DigitalProjectsLibrary\\Planning\\Joe\\integrated multi\\Export Graveyard\\Forecast_test.xlsx")
ruta_salida_2 = Path("C:\\Users\\x286384\\MerckGroup\\Digitalización - Documents\\DigitalProjectsLibrary\\Planning\\Joe\\integrated multi\\Export Graveyard\\Prod_plan_test.xlsx")
TAMANO_LOTE = 115_200                 # jeringas por lote
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
COL_GRANEL = "Nombre granel"
COL_DC = "DC"
COL_SKU = "SKUMERCK"
COL_DESC = "Description"
HOJA_FORECAST = "Forecast balanceado"
FORMATO_MES = "%m-%Y"
# Formato del Excel de salida (mismo estilo que el libro de balanceo)
FUENTE = "Arial"
VERDE_MERCK = "0F5A3C"
FORMATO_PIEZAS = "#,##0"
FORMATO_HORAS = "#,##0.0"
FORMATO_FECHA = "dd/mm/yyyy hh:mm"
horas_paro=0
#pb.ejecutar(ruta_fcst, hoja_fcst, ruta_catalogo, hoja_catalogo, ruta_salida,lote=TAMANO_LOTE, mes=None)
#pp.ejecutar(ruta_salida, ruta_salida_2, mes=None, horas_paro=None,fecha_inicio_dc1=None, fecha_inicio_dc2=None, lote=TAMANO_LOTE,hoja_forecast=pp.HOJA_FORECAST)

def all_plans():
    lote=TAMANO_LOTE
    fore    = pp.obtener_forecast(ruta_salida,"Forecast balanceado")
    meses_2 = pp.meses_disponibles(fore)
    cronogramas_records={}
    planes= {}
    for mes in meses_2:
        base, meses = pp.preparar_mes(fore, mes, lote)
        if base.empty:
            raise ValueError(f"El mes '{mes}' no tiene demanda por planear.")
        anio, numero = pp.mes_a_anio_mes(mes)
        inicio_mes = datetime(anio, numero, 1, HORA_INICIO_DEFECTO, 0)
        inicios = {"DC1": inicio_mes,"DC2": inicio_mes}    
        for esquema in ESQUEMAS:
            df, movimientos, estado = pp.trasladar_para_ajustar(base, mes, esquema, horas_paro=0)
            df = pp.armar_campanas(df)
            inicios_esquema = {dc: pp._avanzar_tiempo(inicios[dc], 0, esquema) for dc in LINEAS}
            cronogramas = {dc: pp.generar_cronograma(df, dc, inicios_esquema[dc], esquema)
                            for dc in LINEAS if (df[COL_DC] == dc).any()}
            cronogramas_records[mes + esquema]=cronogramas
            planes[esquema] = {"detalle": df, "movimientos": movimientos,
                            "estado": estado, "cronogramas": cronogramas,
                            "inicios": inicios_esquema}
    return cronogramas_records   # <-- se guarda para el resumen

#TABLA RESUMEN 24/5 Y 24/7
fore = pp.obtener_forecast(ruta_salida,"Forecast balanceado")
meses=pp.meses_disponibles(fore)
cronogramas_records=all_plans()
filas_resumen = []
horas_paro=pd.read_excel(ruta_catalogo,sheet_name=hoja_paros)
horas_paro["Mes"] = horas_paro["Mes"].dt.strftime(FORMATO_MES)

def resumen_esquemas(meses, LINEAS,ESQUEMAS,cronograma_records,):
    for mes in meses:
        for dc in LINEAS:
            for esquema in ESQUEMAS:
                cro = cronogramas_records[mes + esquema][dc]
                if cro is None or cro.empty:
                    continue
                inicios = cro["Fecha y hora de inicio"].dropna()
                fines = cro["Fecha y hora de finalizacion"].dropna()
                filas_resumen.append({
                    "Mes": mes,
                    "DC": dc,
                    "Esquema": esquema,
                    "Fecha inicio": inicios.iloc[0] if len(inicios) else None,
                    "Fecha fin": fines.iloc[-1] if len(fines) else None,
                    "Horas Produccion": cro["Horas"].sum(),
                    "Horas limpieza": cro.loc[cro["Nombre"] == "Limpieza", "Horas"].sum(),
                    "horas paro": horas_paro.loc[(horas_paro["Mes"] == mes) & (horas_paro["DC"] == dc),"Hrs"].sum(),})

    planes_resumen = pd.DataFrame(filas_resumen, columns=[
        "Mes", "DC", "Esquema", "Fecha inicio", "Fecha fin",
        "Horas Produccion", "Horas limpieza", "horas paro"])
    return(planes_resumen)


def make_or_break(dir):
    ''' crea las carpetas nuevas o checa si ya existen donde se guardaron'''
    try:
        os.mkdir(dir)  # Single folder
        print(f"Folder '{dir}' created.")
    except FileExistsError:
        pass

def gen_all_plans_xlsx(tipo,gen_fut,meses,ruta_salida,cronogramas_records,LINEAS=LINEAS,ESQUEMAS=ESQUEMAS):
    '''manda todos los futuros planes a su propio excel'''    
    gen_fut = True
    if gen_fut == True:
        main_folder = Path(ruta_salida)
        make_or_break(main_folder)
        for mes in meses:
            mes_folder = main_folder / mes
            make_or_break(mes_folder)
            for dc in LINEAS:
                dc_folder = mes_folder / dc
                make_or_break(dc_folder)
                for esquema in ESQUEMAS:
                    etiqueta = esquema.replace("/", "")      # 24/5 -> 245
                    file_name = dc_folder / f"{mes}_{etiqueta}_{dc}.xlsx"
                    with pd.ExcelWriter(file_name, engine="openpyxl") as writer:
                        df = cronogramas_records[mes + esquema][dc]
                        etiqueta = esquema.replace("/", "")      # 24/5 -> 245
                        nombre_hoja = f"{tipo}_{mes}_{dc}_{etiqueta}"
                        df.to_excel(writer, sheet_name=nombre_hoja, index=False)
                        hoja = writer.book[nombre_hoja]
                        pp.dar_formato(hoja, df, pp.FORMATO_PIEZAS)
                        pp.pintar_cronograma(hoja, df)


##Acondi_prototype
#mas o menos esto seria la preparación antes de iniciar los ciclos por mes para orden y demas etc.
MD=pd.read_excel(ruta_catalogo,sheet_name='MD',header=2)
MD=MD[['SKU','Cantidad de Blisters','Medida de instructivo','Medida de blister','Medida de estuche','familia','Linea']]
MD = MD.rename(columns={"SKU": "SKUMERCK"})
MD['SKUMERCK'] = MD['SKUMERCK'].astype(str)
MD['SKUMERCK'].astype(str).str.strip()
MD=MD[MD['familia']=='DC'].sort_values('SKUMERCK',)
MD_L4=MD[MD['Linea'].isin(['L4','L2/L4'])]
MD_L4=MD_L4.dropna(subset=['SKUMERCK'])
MD_L2=MD[MD['Linea'].isin(['L2','L2/L4'])]
MD_L2=MD_L2.dropna(subset=['SKUMERCK'])
from itertools import product
cols = ["Cantidad de Blisters", "Medida de instructivo", "Medida de blister", "Medida de estuche"]
sku_list = MD_L2['SKUMERCK'].tolist()


Matrizstorage=[]
for df in [MD_L4, MD_L2]:
    relation = pd.DataFrame(columns=['sku_1', 'sku_2', 'relation'])
    if df is MD_L4:
        name="L4"
    else:
        name="L2"
    for sku_1, sku_2 in product(df.index, repeat=2):
        if sku_2 < sku_1:
            continue
        diffs = (df.loc[sku_1, cols] != df.loc[sku_2, cols]).sum()
        t_relation = diffs * 10
        relation = pd.concat([relation, pd.DataFrame({'sku_1': [df.loc[sku_1, 'SKUMERCK']], 'sku_2': [df.loc[sku_2, 'SKUMERCK']], 'relation': [t_relation]})], ignore_index=True)
    if df is MD_L2:
        L2_weight= relation
    else:
        L4_weight= relation
acondi_fore=fore.copy()
acondi_fore.drop(acondi_fore['DC'], inplace=True)
acondi_fore.groupby('SKUMERCK').sum()
acondi_fore= acondi_fore.sort_values('SKUMERCK')


#MD.to_csv("MD_DC.csv", index=False)



fore=fore.drop(['Nombre granel','Description','Units'],axis=1)
fore['SKUMERCK'] = fore['SKUMERCK'].astype(str)

fore['SKUMERCK'].astype(str).str.strip()
print(MD)
#relationship matrix