"""
planning_balanceo.py
Balanceo de demanda de jeringas entre las lineas DC1 y DC2 (sitio Naucalpan).

Como funciona, en cuatro pasos:
    1. Del forecast se toman las columnas de meses futuros (el bloque de piezas).
    2. Cada SKU se cruza con el catalogo para saber su familia de granel y si
       corre en DC2 (ambas banderas de linea encendidas) o solo en DC1.
    3. La demanda se suma por familia y linea. DC2 se deja en lotes completos y
       el remanente en piezas se pasa al bloque DC1 de la misma familia.
    4. Todo se divide entre el tamano de lote y se exporta a un Excel.

Requisitos: pip install pandas openpyxl
"""

from pathlib import Path
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# --- Parametros de negocio (la app puede pasar otros) ---
TAMANO_LOTE = 115_200      # jeringas por lote
FORMATO_MES = "%m-%Y"
agg_dict = {}
# --- Nombres de columna ---
id_cols = ['SKUMERCK','Description']

# --- Formato del Excel de salida ---
FUENTE = "Arial"
VERDE_MERCK = "0F5A3C"
FORMATO_LOTES = "#,##0.00"
FORMATO_PIEZAS = "#,##0"

#Joe:la funcion log solo sirve para notificar al usuario, no es necesaria
def log(mensaje):
    """Unico punto de salida de mensajes. La app de escritorio reemplaza esta
    funcion por la suya para mostrarlos en la bitacora de la ventana."""
    print(mensaje)
# ==============================================================================
# 1. Lectura
# ==============================================================================
def leer_entradas(ruta_fcst, hoja_fcst, ruta_catalogo, hoja_catalogo):
    """Lee los dos Excel y revisa que traigan lo minimo necesario."""
    fcst = pd.read_excel(ruta_fcst, sheet_name=hoja_fcst)
    fcst[id_cols[0]]=fcst[id_cols[0]].astype(str).str.strip()
    catalogo = pd.read_excel(ruta_catalogo, sheet_name=hoja_catalogo)
    catalogo[id_cols[0]]=catalogo[id_cols[0]].astype(str).str.strip()
    log(f"Forecast: {len(fcst)} filas | Catalogo: {len(catalogo)} filas")
    return fcst, catalogo

# ==============================================================================
# 2. Ventana de meses
# ==============================================================================
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


def listar_meses(ruta_fcst, hoja_fcst):
    """Etiquetas de mes que ofrece la hoja, leyendo solo los encabezados.
    Sirve para llenar el combo de la app sin cargar todo el libro. Si la hoja
    no trae fechas validas regresa una lista vacia.
    """
    encabezados = pd.read_excel(ruta_fcst, sheet_name=hoja_fcst, nrows=0).columns
    try:
        _, meses = columnas_de_mes(encabezados)
    except ValueError:
        return []
    return meses

def preparar_demanda(fcst):
    """Recorta el forecast a SKU + meses futuros y renombra los encabezados.
    Regresa (demanda, meses), donde `meses` son etiquetas tipo '08-2026'.
    """

    columnas, meses = columnas_de_mes(fcst.columns)
    demanda = fcst[[id_cols[0]] + columnas].dropna(subset=[id_cols[0]]).copy()
    demanda.columns = [id_cols[0]] + meses
    demanda[meses] = demanda[meses].astype(float)
    log(f"Ventana de demanda: {len(meses)} meses ({meses[0]} a {meses[-1]})")
    return demanda, meses
# ==============================================================================
# 3. Cruce con el catalogo y consolidacion
# ==============================================================================
def detalle_por_sku(catalogo, demanda, meses, fcst=None,agg_dict=None):
    """Una fila por SKU con su familia, su linea y su demanda en piezas.
    Esta tabla es la base de todo: el consolidado sale de agruparla, asi que
    las dos cuadran por construccion.
    """
    
    tabla = catalogo[[id_cols[0], "Nombre granel", "Linea Granel 1", "Linea Granel 2"]].merge(
        demanda, on=id_cols[0], how="left")
    tabla[meses] = tabla[meses].astype(float).fillna(0.0)
    if agg_dict is None:
        agg_dict = {col: 'sum' for col in catalogo.columns if col not in ['Nombre granel', 'Linea Granel 1', 'Linea Granel 2']}

    # Las dos banderas del catalogo se colapsan en una sola etiqueta de linea.
    #JOE: esta parte me encanto, asi se queda. 
    corre_en_dc2 = (tabla["Linea Granel 1"] == 1) & (tabla["Linea Granel 2"] == 1)
    tabla["DC"] = corre_en_dc2.map({True: "DC2", False: "DC1"})
    tabla = tabla.drop(columns=["Linea Granel 1", "Linea Granel 2"])

    # Datos descriptivos del forecast, solo para que el reporte se lea mejor.
    descriptivas = []
    if fcst is not None:
        descriptivas = [c for c in (id_cols[1], "Units") if c in fcst.columns]
    if descriptivas:
        catalogo_desc = fcst[[id_cols[0]] + descriptivas].drop_duplicates(subset=[id_cols[0]])
        tabla = tabla.merge(catalogo_desc, on=id_cols[0], how="left")
    # la tabla total horizonte practicamente seria el balanceo por producto ya 
    # echa, solo necesita implementar ese codigo de balanceo para ser veridica. 
    tabla = tabla[["Nombre granel", "DC", id_cols[0]] + descriptivas + meses]
    tabla = tabla.sort_values(["Nombre granel", "DC"],
                              ascending=[True, True])
    log(f"Detalle por SKU: {len(tabla)} renglones")
    return tabla.reset_index(drop=True),agg_dict


# ==============================================================================
# 4. Balanceo DC2 -> DC1 y conversion a lotes
# ==============================================================================
def balancear(detalle, meses, lote=TAMANO_LOTE):
    """Cierra los lotes de DC2 pasando su remanente a DC1 y convierte a lotes.
    Por cada familia y cada mes:
        residuo = demanda_DC2 % lote
        DC2 se queda con lotes enteros y DC1 absorbe el residuo.
    Regresa (lotes, residuos). `residuos` esta en piezas y dice cuanto falta
    para cerrar el ultimo lote de cada fila.
    Se conserva como la version anterior del balanceo (por familia completa,
    no por producto). Sirve como hoja de referencia/comparacion en el Excel.
    Suma el detalle por familia de granel y linea. solia ser consolidado pero 
    era una funcion inecesaria por ahora, si acaso podria ser util cuando se quiera 
    automatizar al agrupamiento para mas familias y productos."""

    consolidado = detalle.groupby(["Nombre granel", "DC"], as_index=False)[meses].sum()
    log(f"Consolidado: {consolidado["Nombre granel"].nunique()} familias, "
        f"{(consolidado["DC"] == 'DC2').sum()} bloques DC2")
    
    tabla = consolidado.copy()
    #JOE: en ves de utilizar mask usa index y garantiza que no se repita nada, sirve igual
    for familia, grupo in tabla.groupby("Nombre granel"):
        filas_dc2 = grupo.index[grupo["DC"] == "DC2"]
        filas_dc1 = grupo.index[grupo["DC"] == "DC1"]
        if len(filas_dc2) > 1 or len(filas_dc1) > 1:
            raise ValueError(f"La familia '{familia}' tiene mas de una fila por "
                         "linea en el consolidado; revise el catalogo.")
        if len(filas_dc2) == 0:
            continue
        if len(filas_dc1) == 0:
            # Sin bloque DC1 no hay donde depositar el remanente: se deja
            # intacto para no perder demanda.
            log(f"Aviso: la familia '{familia}' tiene DC2 pero no DC1, "
                "no se transfiere su remanente.")
            continue
        origen, destino = filas_dc2[0], filas_dc1[0]
        for mes in meses:
            residuo = tabla.at[origen, mes] % lote
            tabla.at[origen, mes] -= residuo
            tabla.at[destino, mes] += residuo
    # Faltante en piezas para cerrar el ultimo lote, antes de dividir.
    '''
    residuos = tabla.copy()
    excedente = tabla[meses] % lote
    residuos[meses] = excedente
    '''
    tabla[meses] = (tabla[meses] / lote)
    
    return tabla#, residuos
# ==============================================================================
# 4. Balanceo por producto (SKU): DC2 -> DC1 y conversion a lotes
# ==============================================================================
def duplicar_por_dc(detalle, meses):
    """Expande el detalle (una fila por SKU) a una fila por SKU y por DC.
    Los SKU que corren en DC2 aportan una segunda fila: la fila original
    conserva su demanda en DC2, y se agrega una fila hermana en DC1 con
    demanda en cero, lista para recibir el remanente del balanceo. Los SKU
    que solo corren en DC1 no se duplican.
    """
    es_dc2 = detalle["DC"] == "DC2"
    fila_dc1_nueva = detalle.loc[es_dc2].copy()
    fila_dc1_nueva["DC"] = "DC1"
    fila_dc1_nueva[meses] = 0.0

    expandido = pd.concat([detalle, fila_dc1_nueva], ignore_index=True)
    expandido = expandido.sort_values(
        ["Nombre granel", id_cols[0], "DC"], ascending=[True, True, False]
    ).reset_index(drop=True)
    expandido.to_csv("expandido.csv", index=False)
    return expandido


def balancear_por_producto(detalle, meses, lote=TAMANO_LOTE):
    """Balanceo por producto: el remanente de cada SKU en DC2 se pasa al
    mismo SKU en DC1.

    Por cada SKU con capacidad en DC2 y cada mes:
        residuo = demanda_DC2 % lote
        DC2 se queda con lotes enteros y DC1 (mismo SKU) absorbe el residuo.
    Si el SKU no tiene fila en DC1 (no se puede fabricar ahi), el residuo se
    queda en DC2 sin transferir: no todo lo que sobra en DC2 se puede mover.

    Regresa (detalle_balanceado, log_balanceo):
        - detalle_balanceado: el detalle por SKU y DC ya con el remanente
          movido, en piezas.
        - log_balanceo: una fila por movimiento (SKU, mes) con el estado
          antes y despues en DC1 y DC2, para trazabilidad.
    """
    log_balanceo = pd.DataFrame(columns=['SKUMERCK', 'Description', 'Nombre granel', 'DC', 'fecha', 'log', 'Units'] + meses)
    expandido = duplicar_por_dc(detalle, meses)
    prod_con_dc2 = expandido.loc[expandido["DC"] == "DC2", id_cols[0]].unique()
    for prod in prod_con_dc2:
        for mes in meses:
            mask_dc2 = (expandido[id_cols[0]] == prod) & (expandido["DC"] == "DC2")
            sku_dc2 = expandido.loc[mask_dc2, mes].values[0]
            if sku_dc2 % lote != 0:
                residuo = sku_dc2 % lote
                candidates = expandido[
                (expandido[mes] > residuo) & 
                (expandido['SKUMERCK'] == prod) &
                (expandido['DC'] == 'DC2')][mes]
                if candidates.empty:
                    continue
                min_idx = candidates.idxmin()
                # Se resta el residuo de la fila para balancear DC2
                before_dc2 = expandido.loc[[min_idx]] 
                expandido.loc[min_idx, mes] = expandido.loc[min_idx, mes] - residuo    
                after_dc2 = expandido.loc[[min_idx]].copy()           
                # Find the same SKU but with DC1
                dc1_mask = (expandido['SKUMERCK'] == prod) & (expandido['DC'] == 'DC1')
                before_dc1 = expandido.loc[dc1_mask].copy()
                expandido.loc[dc1_mask, mes] += residuo
                after_dc1 = expandido.loc[dc1_mask].copy()
                columnas_a_cero = [m for m in meses if m != mes]
                for df, label in zip([before_dc2, before_dc1, after_dc2, after_dc1],
                    ['before_dc2', 'before_dc1', 'after_dc2', 'after_dc1']):
                    df['log'] = label
                    df['fecha'] = mes
                    df[columnas_a_cero] = 0
                # Append all to log
                log_balanceo = pd.concat([log_balanceo, before_dc2, before_dc1, after_dc2, after_dc1], ignore_index=True)
    for mes in meses:
        if log_balanceo[mes].sum() == 0:
            log_balanceo = log_balanceo.drop(columns=[mes])
    residuos = expandido.copy()
    excedente = expandido[meses] % lote
    residuos[meses] = excedente
    
    
    expandido.to_csv("balanceado.csv", index=False)
    log_balanceo.to_csv("log_balanceo.csv", index=False)
    residuos.to_csv("residuos.csv", index=False)

    return expandido, log_balanceo,residuos

def consolidar_por_producto(detalle_balanceado, meses, lote=TAMANO_LOTE):
    """Agrupa el detalle ya balanceado por producto en Nombre granel + DC.
    Reemplaza la vieja tabla "Lotes por familia": ahora el agrupamiento
    parte de un balanceo hecho por SKU, no por linea completa, asi que puede
    seguir mostrando un residuo en DC2 cuando ese residuo pertenece a un
    producto que no se puede fabricar en DC1.
    """
    agrupado = detalle_balanceado.drop(columns=[id_cols[0]] + [
        c for c in (id_cols[1], "Units") if c in detalle_balanceado.columns
    ]).groupby(["Nombre granel", "DC"], as_index=False)[meses].sum()
    agrupado[meses] = agrupado[meses] / lote
    return agrupado
# ==============================================================================
# 5. Exportacion a Excel
# ==============================================================================
def filtrar_mes(tabla, mes, meses, quitar=()):
    """Deja solo las columnas descriptivas y la del mes indicado."""
    descriptivas = [c for c in tabla.columns if c not in meses and c not in quitar]
    return tabla[descriptivas + [mes]].copy()
def armar_resumen(meses, mes, detalle, lotes, lote, origen):
    """Tabla de trazabilidad que encabeza el libro."""
    filas = [
        ("Fecha de ejecucion", pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("Jeringas por lote", f"{lote:,}"),
        ("Meses procesados", f"{len(meses)} ({meses[0]} a {meses[-1]})"),
        ("Familias de granel", str(lotes["Nombre granel"].nunique())),
        ("SKU considerados", str(len(detalle))),
        ("Archivo de forecast", origen["fcst"]),
        ("Archivo de catalogo", origen["catalogo"])]
    return pd.DataFrame(filas, columns=["Concepto", "Valor"])
def dar_formato(hoja, tabla, formato_numeros):
    """Encabezado verde, anchos de columna, filtro y formato numerico."""
    for celda in hoja[1]:
        celda.font = Font(name=FUENTE, size=10, bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor=VERDE_MERCK)
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    hoja.row_dimensions[1].height = 28
    for fila in hoja.iter_rows(min_row=2):
        for celda in fila:
            celda.font = Font(name=FUENTE, size=10)
            if isinstance(celda.value, (int, float)):
                celda.number_format = formato_numeros
    for i, columna in enumerate(tabla.columns, start=1):
        ancho_datos = int(tabla[columna].astype(str).str.len().max()) if len(tabla) else 0
        ancho = min(max(len(str(columna)) + 4, ancho_datos + 2, 10), 42)
        hoja.column_dimensions[get_column_letter(i)].width = ancho
    hoja.freeze_panes = "C2"
    hoja.auto_filter.ref = hoja.dimensions
def exportar_excel(ruta_salida, resumen, lotes, detalle, residuos,lotes_producto=None, log_balanceo=None):
    """Escribe el libro con las cuatro hojas del proceso."""
    ruta_salida = Path(ruta_salida).with_suffix(".xlsx")
    ruta_salida.parent.mkdir(parents=True, exist_ok=True)
    hojas = {
        "Resumen": (resumen, "@"),
        "Lotes por producto": (lotes_producto, FORMATO_LOTES),
        "Lotes por familia": (lotes, FORMATO_LOTES),
        "Forecast balanceado": (detalle, FORMATO_PIEZAS),
        "Residuos": (residuos, FORMATO_PIEZAS),
        "Log balanceo": (log_balanceo, FORMATO_PIEZAS),
    }

    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        for nombre, (tabla, formato) in hojas.items():
            tabla.to_excel(writer, sheet_name=nombre, index=False)
            dar_formato(writer.book[nombre], tabla, formato)
        writer.book["Resumen"].column_dimensions["B"].width = 60
    log(f"Excel generado: {ruta_salida}")
    return ruta_salida

# ==============================================================================
# 6. Proceso completo
# ==============================================================================
def ejecutar(ruta_fcst, hoja_fcst, ruta_catalogo, hoja_catalogo, ruta_salida,
             lote=TAMANO_LOTE, mes=None):
    """Corre todo el proceso y escribe el Excel.
    `mes` limita las hojas 'Detalle SKU' y 'Residuos' a ese mes; los lotes
    siempre se calculan sobre todo el horizonte.
    Regresa un diccionario con las tablas, los meses y los totales por linea.
    """
    log(f"Inicio | lote={lote:,} jeringas")
    fcst, catalogo = leer_entradas(ruta_fcst, hoja_fcst, ruta_catalogo, hoja_catalogo)
    demanda, meses = preparar_demanda(fcst)
    if mes and mes not in meses:
        raise ValueError(f"El mes '{mes}' no esta en el horizonte "
                         f"({meses[0]} a {meses[-1]}).")
    detalle, agg_dict = detalle_por_sku(catalogo, demanda, meses, fcst,agg_dict=None)
    # Balanceo por producto (nuevo): esta es la version que manda en el
    # "Detalle SKU" y en "Lotes por producto".
    
    detalle_balanceado, log_balanceo, residuos = balancear_por_producto(detalle, meses, lote)
    lotes_producto = consolidar_por_producto(detalle_balanceado, meses, lote)
    detalle.to_csv("detalle_salida.csv", index=False)
    lotes = balancear(detalle, meses, lote)
    detalle_salida = detalle_balanceado
    residuos_salida = residuos
    log_balanceo_salida = log_balanceo
    
    if mes:
        detalle_salida = filtrar_mes(detalle_balanceado, mes, meses)   # <-- antes: detalle
        detalle_salida = detalle_salida.sort_values(
            ["Nombre granel", "DC", mes], ascending=[True, True, False])
        residuos_salida = filtrar_mes(residuos, mes, meses)
        if len(log_balanceo):                                          # <-- nuevo
            log_balanceo_salida = log_balanceo[log_balanceo["fecha"] == mes].copy()
        log(f"Detalle y residuos recortados al mes {mes}")
    origen = {
        "fcst": f"{Path(ruta_fcst).name} (hoja '{hoja_fcst}')",
        "catalogo": f"{Path(ruta_catalogo).name} (hoja '{hoja_catalogo}')",}
    resumen = armar_resumen(meses, mes, detalle, lotes, lote, origen)
    exportar_excel(ruta_salida, resumen, lotes, detalle_salida, residuos_salida,lotes_producto=lotes_producto
               , log_balanceo=log_balanceo_salida)

    total_dc1 = float(lotes.loc[lotes["DC"] == "DC1", meses].to_numpy().sum())
    total_dc2 = float(lotes.loc[lotes["DC"] == "DC2", meses].to_numpy().sum())
    log(f"Fin | {total_dc1:,.2f} lotes en DC1 | {total_dc2:,.2f} lotes en DC2")

    return {
        "lotes": lotes,
        "lotes_producto": lotes_producto,      # <-- nuevo
        "detalle": detalle_balanceado,          # <-- antes era: detalle
        "residuos": residuos,
        "log_balanceo": log_balanceo,           # <-- nuevo
        "meses": meses,}