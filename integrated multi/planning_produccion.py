"""
planning_produccion.py
Paso 2 del proceso: plan de produccion de las lineas DC1 / DC2.
Parte de la hoja "Forecast balanceado" que exporta planning_balanceo.py y
responde tres preguntas para UN solo mes:
    1. Cuantas horas pide cada linea con la demanda de ese mes?
    2. Caben esas horas en el calendario, con esquema 24/7 y con 24/5?
    3. A que hora empieza y termina cada producto y cada limpieza?
Reglas de negocio modeladas:
- Una corrida por PRODUCTO, en el orden en que viene la hoja. No se agrupan
  lotes de tres en tres: cada renglon del forecast balanceado es su propia
  corrida, y solo los renglones consecutivos del mismo granel comparten
  campana (no se limpia entre ellos).
- Las cantidades del forecast balanceado NO se modifican. Lo unico que puede
  cambiar es en que linea se fabrica un producto, y solo cuando ese SKU
  aparece en las dos lineas (tiene equivalente DC1/DC2).
- Sin redondeos: las horas salen de las piezas tal como estan
  (piezas / tamano de lote * horas por lote).
- Tiempo de fabricacion por lote: DC1 = 27 h, DC2 = 24 h.
- Limpieza entre campanas: 8 h.
- Esquemas 24/7 y 24/5. El paro programado se descuenta del total en 24/7; en
  24/5 se intenta acomodar en fin de semana.
- Si en 24/5 no alcanzan las horas, ese plan no se genera.
Igual que en planning_balanceo.py, todos los mensajes salen por `log`, asi que
la app de escritorio puede reemplazar esa funcion por la suya.
Requisitos: pip install pandas openpyxl
"""

from calendar import month, monthrange
from datetime import datetime, timedelta
from pathlib import Path
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Parametros de negocio (la app puede pasar otros)
# ---------------------------------------------------------------------------
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

def log(mensaje):
    """Unico punto de salida de mensajes. La app de escritorio la reemplaza."""
    print(mensaje)
# ==============================================================================
# 1. Entrada: el forecast ya balanceado
# ==============================================================================
# de lo que estoy viendo de este bloque es solo el manejo de dato sy lectura, 
# si quisieramos integrar toda la logica a un mismo codigo mucho de esto podria eliminarse.}
# se me hace mejor idea dejar todo como esta por ahora o minimo en dos archivos diferentes para los pasos. 

#revisar si esta funcion se repite entre planning balanceo y este documento 
def mes_a_anio_mes(mes):
    """Convierte la etiqueta de mes del balanceo ('08-2026') en (2026, 8).
    revisar si e snecesario o existe mejor forma de realizar este codigo."""
    fecha = pd.to_datetime(mes, format=FORMATO_MES, errors="coerce")
    if pd.isna(fecha):
        fecha = pd.to_datetime(mes, errors="coerce")
    if pd.isna(fecha):
        raise ValueError(f"No se pudo interpretar el mes '{mes}'. "
                         f"Se espera el formato MM-AAAA.")
    return int(fecha.year), int(fecha.month)


def meses_disponibles(tabla):
    """Columnas de la tabla que son etiquetas de mes."""
    fijas = {COL_GRANEL, COL_DC, COL_SKU, COL_DESC, "Units", "camp"}
    meses = []
    for columna in tabla.columns:
        if columna in fijas:
            continue
        if pd.notna(pd.to_datetime(columna, format=FORMATO_MES, errors="coerce")):
            meses.append(columna)
    return meses

def obtener_forecast(origen, hoja=HOJA_FORECAST):
    """Acepta las tres formas en que puede llegar el resultado del balanceo.
        - dict devuelto por planning_balanceo.ejecutar() -> usa 'detalle'.
        - DataFrame ya armado.
        - ruta al .xlsx exportado -> lee la hoja 'Forecast balanceado'.
    El orden de los renglones se respeta tal como viene: es el que manda en
    el cronograma.
    """
    if isinstance(origen, dict):
        tabla = origen.get("detalle")
        if tabla is None:
            raise ValueError("El resultado del balanceo no trae 'detalle'.")
        return tabla.copy()
    if isinstance(origen, pd.DataFrame):
        return origen.copy()
    # Si no es dict ni DataFrame, es la ruta del libro que genero el balanceo.
    # Se convierte a Path antes de usar .name: cuando la ruta llega como
    # cadena, que es lo normal viniendo de la app, el atributo no existe.
    ruta = Path(origen)
    tabla = pd.read_excel(ruta, sheet_name=hoja)
    log(f"Forecast balanceado leido de {ruta.name} (hoja '{hoja}'): "
        f"{len(tabla)} renglones")
    return tabla

def preparar_mes(tabla, mes, lote=TAMANO_LOTE, process_type="sing"):
    """Recorta el forecast al mes pedido y calcula lotes y horas por producto.
    No se toca ninguna cantidad: las piezas pasan tal cual y los lotes salen
    de dividir entre el tamano de lote, sin redondear. Se conserva el orden
    original de la hoja porque de ahi sale la secuencia de fabricacion.
    """
    #REVISAR: sospecho que hay mucho bloat aqui. 

    meses = meses_disponibles(tabla)
    if process_type=="mult":
        columnas = [c for c in (COL_GRANEL, COL_DC, COL_SKU, COL_DESC, "camp") if c in tabla.columns]
        df = tabla[columnas + meses].copy()
        df[COL_DC] = df[COL_DC].astype(str).str.upper().str.strip()
        for mes in meses:
            df[mes] = pd.to_numeric(df[mes], errors="coerce").fillna(0.0)
        desconocidas = sorted(set(df[COL_DC]) - set(LINEAS))
        if desconocidas:
            raise ValueError(f"Lineas no reconocidas en el forecast: {desconocidas}. "f"Solo se planean {list(LINEAS)}.")
        return df, meses

    if process_type=="sing":
        if mes not in meses:
            raise ValueError(f"El mes '{mes}' no esta en el forecast balanceado "
                            f"({', '.join(map(str, meses)) or 'sin meses'}).")
        
        columnas = [c for c in (COL_GRANEL, COL_DC, COL_SKU, COL_DESC, "camp")
                    if c in tabla.columns]
        df = tabla[columnas + [mes]].copy()
        df = df.rename(columns={mes: "jeringas"})
        df["jeringas"] = pd.to_numeric(df["jeringas"], errors="coerce").fillna(0.0)
        df[COL_DC] = df[COL_DC].astype(str).str.upper().str.strip()    
        # Los renglones en cero no ocupan linea, pero no se borra nada del origen.
        df = df[df["jeringas"] > 0].copy()
        df["lotes"] = df["jeringas"] / lote
        df["horas_produccion"] = df.apply(
            lambda f: f["lotes"] * PROD_TIME[f[COL_DC]], axis=1)
        log(f"Mes {mes}: {len(df)} productos con demanda | "
            f"{df['jeringas'].sum():,.0f} jeringas | {df['lotes'].sum():,.2f} lotes")
        desconocidas = sorted(set(df[COL_DC]) - set(LINEAS))
        if desconocidas:
            raise ValueError(f"Lineas no reconocidas en el forecast: {desconocidas}. "
                            f"Solo se planean {list(LINEAS)}.")
    return df.reset_index(drop=True), meses
# ==============================================================================
# 2. Horas disponibles segun esquema y paros programados
# ==============================================================================
def horas_disponibles(year, month, esquema, horas_paro=0,fecha_inicio=None):
    """Horas de produccion disponibles en el mes para una linea."""
    # necesitamos verificar si esta linea se adapata en caso de qu eel tiempo disponilble se 
    # necesite calcular a partir de la fecha en la que se pidio. 
    
    #ESQUEMA_BASE = "24/5"
    #ESQUEMA_OPCIONAL = "24/7" 
    
    if fecha_inicio is not None:
        day = fecha_inicio.day
    else:
        day = 1

    primer_dia = datetime(year, month, day)
    dias_disp = monthrange(year, month)[1]-(day-1)
    dias_habiles = sum(1 for d in range(dias_disp)
                       if (primer_dia + timedelta(days=d)).weekday() < 5)
                       
    dias_finde = dias_disp - dias_habiles
    #asumiendo que no importe el dia que se empiece produccion, es a las 7 am de ese dia 
    # por esto se le resta 7 horas a las horas de producción.     
    horas_habiles = (dias_habiles * 24)-7
    horas_finde = dias_finde * 24    

    if esquema == ESQUEMA_BASE:
        #este es el de 24/5
        if horas_paro <= horas_finde:
            # El paro cabe completo en el fin de semana: no toca tiempo habil.
            disp = horas_habiles
        else:
            disp = horas_habiles - (horas_paro - horas_finde)
    elif esquema == ESQUEMA_OPCIONAL:
        #ya correccion a 24/7
        horas_habiles+=horas_finde
        disp = horas_habiles - horas_paro
        horas_finde=0
    else:
        raise ValueError(f"esquema debe ser uno de {ESQUEMAS}")
    detalle = {
            "esquema": esquema,
            "horas_habiles": horas_habiles,
            "horas_fin_de_semana_disponibles": horas_finde,
            "horas_paro_programado": horas_paro,
            "horas_disponibles_produccion": disp,}
    return disp, detalle

def tabla_horas_disponibles(mes, horas_paro):
    """Una fila por linea y esquema con el desglose del calendario."""
    anio, numero = mes_a_anio_mes(mes)
    filas = []
    for esquema in ESQUEMAS:
        for dc in LINEAS:
            _, detalle = horas_disponibles(anio, numero, esquema,
                                           horas_paro_de_linea(horas_paro, dc))
            filas.append({COL_DC: dc, **detalle})
    columnas = [COL_DC, "esquema", "horas_habiles",
                "horas_fin_de_semana_disponibles", "horas_paro_programado",
                "horas_disponibles_produccion"]
    tabla = pd.DataFrame(filas)
    return tabla[[c for c in columnas if c in tabla.columns]]
# Nota: esta tabla ya no va al libro. La hoja del Excel ahora es "Planes
# resumen" (ver armar_planes_resumen), que reporta lo que de verdad salio del
# cronograma en vez de lo que el calendario ofrecia. Se deja porque sirve para
# revisar a mano el desglose de un mes.
# ==============================================================================
# 3. Campanas: una corrida por producto
# ==============================================================================
#generar campañas lo esta haciendo mal, necesita ser por cada 3 lotes y si se pasa o ya no es la familia
#ya entra en la campaña que sigue 
def armar_campanas(df, process_type="sing", meses=None):
    """Numera campañas y divide productos que superan el límite de campaña.

    Una campaña puede contener como máximo 345,600 jeringas. Los renglones
    consecutivos del mismo granel y de la misma línea comparten campaña hasta
    alcanzar ese límite; cuando se alcanza, se abre otra campaña aunque el
    siguiente tramo siga siendo el mismo SKU o granel. Si una sola fila supera
    el límite, se divide internamente en varios renglones, conservando la suma
    original de jeringas, lotes y horas de producción.

    La descripción, el SKU y las demás columnas informativas se copian a cada
    tramo. Solo `jeringas`, `lotes`, `horas_produccion` y `campana` cambian para
    representar correctamente la división; todos los cálculos siguen siendo
    por SKU/registro y nunca por descripción.
    """
    if process_type == "mult":
        # Este modo no se utiliza en la ejecución mensual actual. Se conserva
        # la interfaz para no romper llamadas existentes.
        return df.copy()

    if process_type != "sing":
        raise ValueError("process_type debe ser 'sing' o 'mult'.")

    if df.empty:
        resultado = df.copy()
        resultado["campana"] = pd.Series(dtype="int64")
        return resultado.reset_index(drop=True)

    limite_campana = 345_600.0
    ordenado = df.sort_values(
        by=[COL_GRANEL, COL_DC, "jeringas"],
        ascending=[True, True, False],
        kind="stable",
    ).copy()

    filas = []
    campana = 0
    ultima_clave = None
    jeringas_en_campana = 0.0
    tolerancia = 1e-9

    for _, fila in ordenado.iterrows():
        clave = (fila[COL_DC], fila[COL_GRANEL])
        restantes = float(fila["jeringas"])

        # El cambio de línea o de granel siempre cierra la campaña anterior.
        if clave != ultima_clave:
            if ultima_clave is not None:
                campana += 1
            ultima_clave = clave
            jeringas_en_campana = 0.0

        while restantes > tolerancia:
            espacio = limite_campana - jeringas_en_campana
            if espacio <= tolerancia:
                campana += 1
                jeringas_en_campana = 0.0
                espacio = limite_campana

            tramo = min(restantes, espacio)
            nueva_fila = fila.copy()
            nueva_fila["jeringas"] = tramo
            nueva_fila["lotes"] = tramo / TAMANO_LOTE
            nueva_fila["horas_produccion"] = (
                nueva_fila["lotes"] * PROD_TIME[nueva_fila[COL_DC]]
            )
            nueva_fila["campana"] = campana
            filas.append(nueva_fila)

            restantes -= tramo
            jeringas_en_campana += tramo

            # Si el tramo llenó exactamente la campaña, la siguiente vuelta
            # abrirá una nueva campaña. No se incrementa aquí para evitar crear
            # una campaña vacía al cambiar de granel o de línea.

    resultado = pd.DataFrame(filas, columns=list(ordenado.columns) + ["campana"])
    
    return resultado.reset_index(drop=True)


def horas_por_linea(df):
    """Horas de produccion y de limpieza que pide cada linea."""
    filas = []
    for dc in LINEAS:
        bloque = df[df[COL_DC] == dc]
        if bloque.empty:
            continue
        campanas = bloque["campana"].nunique()
        limpiezas = campanas if LIMPIEZA_AL_CIERRE else max(campanas - 1, 0)
        horas_prod = float(bloque["horas_produccion"].sum())
        horas_limp = limpiezas * CLEAN_TIME
        filas.append({
            COL_DC: dc,
            "productos": len(bloque),
            "jeringas": float(bloque["jeringas"].sum()),
            "lotes": float(bloque["lotes"].sum()),
            "campanas": campanas,
            "horas_produccion": horas_prod,
            "limpiezas": limpiezas,
            "horas_limpieza": horas_limp,
            "horas_requeridas": horas_prod + horas_limp,
        })
    return pd.DataFrame(filas)


def horas_paro_de_linea(horas_paro, dc):
    """Las horas de paro de UNA linea, venga un dict o un solo numero.

    `horas_paro` llega como {"DC1": h, "DC2": h} desde la hoja de paros, pero
    tambien se acepta un escalar para aplicar lo mismo a las dos lineas, y
    None para no descontar nada.
    """
    if horas_paro is None:
        return 0.0
    if isinstance(horas_paro, dict):
        return float(horas_paro.get(dc, 0) or 0)
    return float(horas_paro)


def evaluar_esquema(df, mes, esquema, horas_paro):
    """Horas requeridas contra disponibles, por linea, para un esquema."""
    anio, numero = mes_a_anio_mes(mes)
    resumen = horas_por_linea(df)
    filas = []
    for _, fila in resumen.iterrows():
        dc = fila[COL_DC]
        # Cada linea descuenta SU paro, no la suma de las dos.
        disp, _ = horas_disponibles(anio, numero, esquema,
                                    horas_paro_de_linea(horas_paro, dc))
        registro = fila.to_dict()
        registro["esquema"] = esquema
        registro["horas_disponibles"] = disp
        registro["holgura_horas"] = disp - fila["horas_requeridas"]
        registro["factible"] = registro["holgura_horas"] >= 0
        filas.append(registro)
    return pd.DataFrame(filas)


# ==============================================================================
# 4. Traslados entre lineas (solo productos con equivalente DC1/DC2)
# ==============================================================================
def trasladar_para_ajustar(df, mes, esquema, horas_paro=None):
    """Mueve productos completos de la linea saturada a la otra.
    Solo son candidatos los SKU que aparecen en las dos lineas: son los que
    tienen equivalente y por lo tanto se pueden fabricar en cualquiera. Las
    cantidades no se tocan, unicamente cambia la etiqueta de linea.
    Se mueve el producto mas chico que alcance a resolver el sobrecupo, para
    alterar el plan lo menos posible.
    """
    df = df.copy()
    movimientos = []
    skus_en_ambas = {sku for sku, grupo in df.groupby(COL_SKU)
                     if grupo[COL_DC].nunique() > 1}
    if not skus_en_ambas:
        return df, pd.DataFrame(), evaluar_esquema(armar_campanas(df), mes, esquema, horas_paro)

    for _ in range(len(df)):                      # tope duro: nunca cicla
        estado = evaluar_esquema(armar_campanas(df), mes, esquema, horas_paro)
        saturadas = estado[~estado["factible"]]
        if saturadas.empty:
            break

        movido = False
        for _, fila in saturadas.sort_values("holgura_horas").iterrows():
            origen = fila[COL_DC]
            destino = "DC2" if origen == "DC1" else "DC1"
            libre = estado.loc[estado[COL_DC] == destino, "holgura_horas"]
            holgura_destino = float(libre.iloc[0]) if len(libre) else None
            if holgura_destino is None:
                # La otra linea no tiene carga: su holgura es todo el mes.
                anio, numero = mes_a_anio_mes(mes)
                holgura_destino, _ = horas_disponibles(
                    anio, numero, esquema,
                    horas_paro_de_linea(horas_paro, destino))
            if holgura_destino <= 0:
                continue
            candidatos = df[(df[COL_DC] == origen)
                            & (df[COL_SKU].isin(skus_en_ambas))].copy()
            if candidatos.empty:
                continue
            # Las horas cambian al cruzar de linea: 27 h en DC1, 24 h en DC2.
            candidatos["horas_en_destino"] = candidatos["lotes"] * PROD_TIME[destino]
            candidatos = candidatos[candidatos["horas_en_destino"] <= holgura_destino]
            if candidatos.empty:
                continue

            faltan = -fila["holgura_horas"]
            suficientes = candidatos[candidatos["horas_produccion"] >= faltan]
            elegido = (suficientes.nsmallest(1, "horas_produccion") if not suficientes.empty
                       else candidatos.nlargest(1, "horas_produccion"))
            idx = elegido.index[0]

            df.at[idx, COL_DC] = destino
            df.at[idx, "horas_produccion"] = float(elegido["horas_en_destino"].iloc[0])
            movimientos.append({
                COL_SKU: df.at[idx, COL_SKU],
                COL_GRANEL: df.at[idx, COL_GRANEL],
                "de": origen,
                "a": destino,
                "jeringas": df.at[idx, "jeringas"],
                "lotes": df.at[idx, "lotes"],
                "horas_en_destino": df.at[idx, "horas_produccion"],
                "motivo": f"{origen} excedia por {faltan:,.1f} h en {esquema}",
            })
            log(f"Traslado {esquema}: {df.at[idx, COL_SKU]} de {origen} a {destino} "
                f"({df.at[idx, 'jeringas']:,.0f} jeringas)")
            movido = True
            break

        if not movido:
            break

    estado = evaluar_esquema(armar_campanas(df), mes, esquema, horas_paro)
    return df, pd.DataFrame(movimientos), estado


# ==============================================================================
# 5. Cronograma
# ==============================================================================
def _avanzar_tiempo(inicio, horas, esquema):
    """Suma horas al reloj. En 24/7 el reloj nunca se detiene. En 24/5 la
    ventana productiva es lunes 07:00 -> sabado 00:00 (113 h).
    Un arranque en fin de semana se corre al lunes a HORA_INICIO_DEFECTO,
    pero un bloque que termina exacto en el corte se reporta en el corte:
    no se empuja al lunes si ya no quedan horas por consumir.
    """
    if esquema == ESQUEMA_OPCIONAL:          # 24/7: el reloj no se detiene
        return inicio + timedelta(hours=horas)

    actual, restante = inicio, horas

    # Normalizacion del arranque: si el reloj ya viene en fin de semana,
    # la produccion abre el lunes a las 07:00.
    if actual.weekday() >= 5:
        actual = (actual.replace(hour=HORA_INICIO_DEFECTO, minute=0,
                                 second=0, microsecond=0)
                  + timedelta(days=7 - actual.weekday()))

    while restante > 1e-9:
        limite = (actual.replace(hour=0, minute=0, second=0, microsecond=0)
                  + timedelta(days=5 - actual.weekday()))
        bloque = (limite - actual).total_seconds() / 3600
        if restante < bloque:
            actual += timedelta(hours=restante)
            restante = 0
        else:
            actual = limite
            restante -= bloque
            if restante > 1e-9:              # falta trabajo -> lunes 07:00
                actual += timedelta(days=2, hours=HORA_INICIO_DEFECTO)

    return actual
def _cabe_sin_pausa(tiempo, horas, esquema):
    """Horas que quedan hasta el sabado 00:00 y si el bloque cabe completo.
    En 24/7 siempre cabe. Devuelve (cabe, horas_hasta_el_corte, corte).
    """
    if esquema == ESQUEMA_OPCIONAL:
        return True, None, None
    corte = (tiempo.replace(hour=0, minute=0, second=0, microsecond=0)
             + timedelta(days=5 - tiempo.weekday()))
    disp = (corte - tiempo).total_seconds() / 3600
    return horas <= disp + 1e-9, disp, corte
def generar_cronograma(df, dc, fecha_inicio, esquema, lote=TAMANO_LOTE):
    """Una fila por LOTE y una fila 'Limpieza' al cerrar cada campana.
    En 24/5 un lote no se puede pausar: si no cabe completo antes del sabado
    a las 00:00, se inserta un renglon 'Fin de semana' y el lote abre el
    lunes a las 07:00.
    """
    df = df.copy()
    df["orden"] = range(len(df))
    bloque = df[df[COL_DC] == dc].sort_values("orden")
    filas = []
    tiempo = _avanzar_tiempo(fecha_inicio, 0, esquema)

    # --- Slot de la primera semana: la campana mas grande que alcance a
    # --- terminar antes del fin de semana pasa al lugar 1.
    if esquema == ESQUEMA_BASE and tiempo.weekday() <= 4:
        # El viernes cuenta completo: de lunes 07:00 a sabado 00:00 son 113 h.
        w1_hrs_disp = ((4 - tiempo.weekday() + 1) * 24) - 7

        camp_hours = bloque.groupby("campana")["horas_produccion"].sum()
        candidatas = camp_hours[camp_hours <= w1_hrs_disp]

        if not candidatas.empty:
            mejor_campana = candidatas.idxmax()
            primera = bloque["campana"].min()
            if mejor_campana != primera:
                bloque.loc[bloque["campana"] == primera, "campana"] = -1
                bloque.loc[bloque["campana"] == mejor_campana, "campana"] = primera
                bloque.loc[bloque["campana"] == -1, "campana"] = mejor_campana
                '''
                log(f"Reorden 24/5 en {dc}: campana {mejor_campana} "
                    f"({candidatas[mejor_campana]:,.1f} h) pasa al lugar "
                    f"{primera} para llenar las {w1_hrs_disp:,.0f} h "
                    "disponibles antes del fin de semana.")
                '''
                bloque = bloque.sort_values(["campana", "orden"])

    ultima_campana = bloque["campana"].max()

    for campana, grupo in bloque.groupby("campana", sort=True):
        for _, fila in grupo.iterrows():
            jeringas_totales = fila["jeringas"]
            horas_por_lote = PROD_TIME[dc]
            lotes_completos = int(jeringas_totales // lote)
            residuo = jeringas_totales - (lotes_completos * lote)
            tramos = [(lote, horas_por_lote)] * lotes_completos
            if residuo > 1e-9:
                tramos.append((residuo, (residuo / lote) * horas_por_lote))

            for jeringas_tramo, horas_tramo in tramos:
                # Lote que se abre es lote que se termina: si no cabe antes
                # del sabado, se para la linea y el lote abre el lunes.
                if esquema == ESQUEMA_BASE:
                    corte = (tiempo.replace(hour=0, minute=0, second=0,
                                            microsecond=0)
                             + timedelta(days=5 - tiempo.weekday()))
                    hrs_hasta_corte = (corte - tiempo).total_seconds() / 3600
                    if horas_tramo > hrs_hasta_corte + 1e-9:
                        reinicio = _avanzar_tiempo(corte, 0, esquema)
                        filas.append({
                            "Nombre": "Fin de semana",
                            COL_SKU: "",
                            "Campana": None,
                            "Jeringas": None,
                            "Lotes": None,
                            "Horas": None,
                            "Fecha y hora de inicio": None,
                            "Fecha y hora de finalizacion": None,
                        })
                        '''
                        log(f"Fin de semana en {dc}: quedaban "
                            f"{hrs_hasta_corte:,.1f} h y el lote pide "
                            f"{horas_tramo:,.1f} h. Abre el "
                            f"{reinicio:%d/%m %H:%M}.")'''
                        tiempo = reinicio

                fin = _avanzar_tiempo(tiempo, horas_tramo, esquema)
                filas.append({
                    "Nombre": fila[COL_GRANEL],
                    COL_SKU: fila[COL_SKU],
                    "Campana": campana,
                    "Jeringas": jeringas_tramo,
                    "Lotes": jeringas_tramo / lote,
                    "Horas": horas_tramo,
                    "Fecha y hora de inicio": tiempo,
                    "Fecha y hora de finalizacion": fin,
                })
                tiempo = fin

        if campana == ultima_campana and not LIMPIEZA_AL_CIERRE:
            continue

        # Normaliza el reloj antes de medir: si venia parado en fin de semana,
        # la limpieza arranca el lunes a las 07:00.
        tiempo_limpieza_inicio = _avanzar_tiempo(tiempo, 0, esquema)
        tiempo = tiempo_limpieza_inicio

        # La limpieza se absorbe en el fin de semana: si no cierra antes del
        # corte, se reporta terminando el viernes a medianoche y la campana
        # siguiente abre el lunes a las 07:00. Nunca cierra en lunes.
        cabe, disp, corte = _cabe_sin_pausa(tiempo, CLEAN_TIME, esquema)
        if cabe:
            fin_limpieza = _avanzar_tiempo(tiempo, CLEAN_TIME, esquema)
            tiempo = fin_limpieza
        else:
            tiempo_limpieza_inicio = _avanzar_tiempo(corte, 0, esquema)  # lunes 07:00
            tiempo = tiempo_limpieza_inicio
            fin_limpieza = _avanzar_tiempo(tiempo, CLEAN_TIME, esquema)
            tiempo = fin_limpieza

        filas.append({
            "Nombre": "Limpieza",
            COL_SKU: "",
            "Campana": None,
            "Jeringas": None,
            "Lotes": None,
            "Horas": CLEAN_TIME,
            "Fecha y hora de inicio": tiempo_limpieza_inicio,
            "Fecha y hora de finalizacion": fin_limpieza,
        })

        # Si la limpieza toco el corte, la linea queda parada hasta el lunes.
        if esquema == ESQUEMA_BASE and tiempo > fin_limpieza:
            filas.append({
                "Nombre": "Fin de semana",
                COL_SKU: "",
                "Campana": None,
                "Jeringas": None,
                "Lotes": None,
                "Horas": None,
                "Fecha y hora de inicio": None,
                "Fecha y hora de finalizacion": None,
            })

    return pd.DataFrame(filas, columns=["Nombre", COL_SKU, "Campana", "Jeringas", "Lotes", "Horas","Fecha y hora de inicio", "Fecha y hora de finalizacion"])

# --- Colores para el cronograma: alternan por bloque de campana ---
COLOR_BLOQUE_A = "E2EFDA"   # verde muy claro
COLOR_BLOQUE_B = "FFF2CC"   # amarillo muy claro
COLOR_LIMPIEZA = "D9D9D9"   # gris, para que la limpieza resalte como separador
COLOR_FIN_SEMANA = "BFBFBF" # gris mas oscuro: la linea esta detenida


def pintar_cronograma(hoja, tabla):
    """Colorea cada renglon del cronograma: alterna de color cada vez que
    cierra una campana (fila 'Limpieza'), pinta la limpieza en gris claro y
    el fin de semana en gris oscuro, para que se vea donde corta un bloque y
    donde la linea esta parada.
    """
    if "Nombre" not in tabla.columns or tabla.empty:
        return

    colores = (COLOR_BLOQUE_A, COLOR_BLOQUE_B)
    turno = 0
    for i, nombre in enumerate(tabla["Nombre"], start=2):  # fila 1 es encabezado
        if nombre == "Limpieza":
            color = COLOR_LIMPIEZA
        elif nombre == "Fin de semana":
            color = COLOR_FIN_SEMANA
        else:
            color = colores[turno % 2]
        relleno = PatternFill("solid", fgColor=color)
        for celda in hoja[i]:
            celda.fill = relleno
        if nombre == "Limpieza":
            turno += 1   # el siguiente bloque de produccion cambia de color
# ==============================================================================
# 5b. Horas de paro programado, resumen de esquemas y planes a futuro
#
# Lo que antes vivia en el banco de pruebas (testing_grounds) se mudo aqui.
# Son tres cosas distintas que se usaban juntas:
#   - leer_horas_paro: la hoja de paros del Excel de catalogo.
#   - armar_planes_resumen: una fila por mes/linea/esquema, sustituye a la
#     hoja de "Horas disponibles".
#   - planes_del_horizonte + gen_all_plans_xlsx: el plan de cada mes a futuro,
#     cada cronograma en su propio archivo.
# ==============================================================================
def leer_horas_paro(ruta_paros=None, hoja_paros=None, mes=None, tabla=None):
    """Horas de paro programado por linea, leidas de la hoja de paros.

    La hoja trae una fila por evento con Mes, DC y Hrs. Se suman las horas de
    las filas que coinciden con el mes pedido. Un mes sin filas no es un error:
    significa que no hay paro programado, o sea cero horas. Por eso el regreso
    siempre trae las dos lineas, aunque la hoja no las mencione.

    Se puede mandar `tabla` ya leida para no volver a abrir el Excel cuando se
    recorre todo el horizonte.
    """
    vacio = {dc: 0.0 for dc in LINEAS}
    if tabla is None:
        if not ruta_paros or not hoja_paros:
            return vacio
        tabla = pd.read_excel(ruta_paros, sheet_name=hoja_paros)
    if tabla is None or tabla.empty:
        return vacio

    faltan = [c for c in ("Mes", COL_DC, "Hrs") if c not in tabla.columns]
    if faltan:
        raise ValueError(f"La hoja de paros no tiene las columnas {faltan}. "
                         f"Se esperan 'Mes', '{COL_DC}' y 'Hrs'.")

    etiquetas = pd.to_datetime(tabla["Mes"], errors="coerce").dt.strftime(FORMATO_MES)
    horas = pd.to_numeric(tabla["Hrs"], errors="coerce").fillna(0.0)
    if mes is not None:
        filtro = etiquetas == mes
    else:
        filtro = pd.Series(True, index=tabla.index)

    resultado = dict(vacio)
    for dc in LINEAS:
        resultado[dc] = float(horas[filtro & (tabla[COL_DC] == dc)].sum())
    return resultado


def normalizar_tabla_paros(ruta_paros=None, hoja_paros=None):
    """Lee la hoja de paros una sola vez. Devuelve None si no se eligio."""
    if not ruta_paros or not hoja_paros:
        return None
    tabla = pd.read_excel(ruta_paros, sheet_name=hoja_paros)
    log(f"Hoja de paros '{hoja_paros}' leida: {len(tabla)} registros")
    return tabla


def armar_planes_resumen(planes_por_mes, tabla_paros=None):
    """Una fila por mes, linea y esquema con lo que de verdad salio del plan.

    Sustituye a la hoja de "Horas disponibles": en vez de lo que el calendario
    ofrecia, aqui se lee del cronograma ya construido cuando arranca, cuando
    termina, cuantas horas son de produccion y cuantas de limpieza. Las horas
    de paro se traen de la hoja de paros para poder comparar.

    `planes_por_mes` es {mes: {esquema: plan}}, con el plan tal como lo arma
    `ejecutar`. Los esquemas que no se generaron se saltan sin ruido.
    """
    filas = []
    for mes, planes in planes_por_mes.items():
        paros = leer_horas_paro(mes=mes, tabla=tabla_paros)
        for esquema in ESQUEMAS:
            plan = planes.get(esquema)
            if plan is None:
                continue
            for dc in LINEAS:
                cro = plan["cronogramas"].get(dc)
                if cro is None or cro.empty:
                    continue
                inicios = cro["Fecha y hora de inicio"].dropna()
                fines = cro["Fecha y hora de finalizacion"].dropna()
                limpieza = cro.loc[cro["Nombre"] == "Limpieza", "Horas"].sum()
                filas.append({
                    "Mes": mes,
                    COL_DC: dc,
                    "Esquema": esquema,
                    "Fecha inicio": inicios.iloc[0] if len(inicios) else None,
                    "Fecha fin": fines.iloc[-1] if len(fines) else None,
                    "Horas produccion": cro["Horas"].sum() - limpieza,
                    "Horas limpieza": limpieza,
                    "Horas paro": paros.get(dc, 0.0),
                })
    return pd.DataFrame(filas, columns=[
        "Mes", COL_DC, "Esquema", "Fecha inicio", "Fecha fin",
        "Horas produccion", "Horas limpieza", "Horas paro"])


def make_or_break(carpeta):
    """Crea la carpeta si no existe. Si ya estaba, no pasa nada."""
    carpeta = Path(carpeta)
    carpeta.mkdir(parents=True, exist_ok=True)
    return carpeta


def plan_de_un_mes(forecast, mes, horas_paro, lote=TAMANO_LOTE, inicios=None):
    """Arma los planes de los dos esquemas para un mes, sin exportar nada.

    Es el mismo cuerpo que usa `ejecutar`, aislado para poder repetirlo en
    todos los meses del horizonte sin duplicar reglas.
    """
    base, _ = preparar_mes(forecast, mes, lote)
    if base.empty:
        return None, base

    anio, numero = mes_a_anio_mes(mes)
    inicio_mes = datetime(anio, numero, 1, HORA_INICIO_DEFECTO, 0)
    inicios = inicios or {dc: inicio_mes for dc in LINEAS}

    planes = {}
    for esquema in ESQUEMAS:
        df, movimientos, estado = trasladar_para_ajustar(base, mes, esquema,
                                                         horas_paro)
        df = armar_campanas(df)
        if not bool(estado["factible"].all()):
            faltan = estado.loc[~estado["factible"]]
            detalle = " | ".join(f"{f[COL_DC]} excede por "
                                 f"{-f['holgura_horas']:,.1f} h"
                                 for _, f in faltan.iterrows())
            if esquema == ESQUEMA_OPCIONAL:
                log(f"{mes}: plan {esquema} no generado ({detalle})")
                planes[esquema] = None
                continue
            log(f"Aviso: {mes} plan {esquema} no cabe en el mes ({detalle}). "
                "Se entrega el cronograma para que se vea el desborde.")

        inicios_esquema = {dc: _avanzar_tiempo(inicios[dc], 0, esquema)
                           for dc in LINEAS}
        cronogramas = {dc: generar_cronograma(df, dc, inicios_esquema[dc],
                                              esquema)
                       for dc in LINEAS if (df[COL_DC] == dc).any()}
        planes[esquema] = {"detalle": df, "movimientos": movimientos,
                           "estado": estado, "cronogramas": cronogramas,
                           "inicios": inicios_esquema}
    return planes, base


def planes_del_horizonte(forecast, lote=TAMANO_LOTE, tabla_paros=None):
    """Recorre todos los meses del forecast balanceado y arma su plan.

    Un mes sin demanda se salta con un aviso: no es motivo para tumbar la
    corrida completa.
    """
    resultado = {}
    for mes in meses_disponibles(forecast):
        horas_paro = leer_horas_paro(mes=mes, tabla=tabla_paros)
        planes, _ = plan_de_un_mes(forecast, mes, horas_paro, lote)
        if planes is None:
            log(f"{mes}: sin demanda por planear, se omite")
            continue
        resultado[mes] = planes
    return resultado


def gen_all_plans_xlsx(carpeta_raiz, planes_por_mes, tipo="Prod",
                       mes_omitido=None):
    """Guarda SOLO el cronograma de cada plan a futuro, un Excel por archivo.

    Estos archivos no son el libro completo: traen el plan pelon, sin resumen
    ni traslados. El libro con todas las hojas es el principal, el del mes que
    se eligio en la app.

    El arbol lo comparte con el acondicionamiento, para que los dos planes del
    mismo mes y esquema queden en la misma carpeta:
        <raiz>/<mes>/<esquema>/<mes>_<esquema>_<linea>_<tipo>.xlsx
    Asi en '<raiz>/01-2027/245' conviven DC1, DC2, L2 y L4.

    `mes_omitido` salta el mes que ya se entrego completo en el libro
    principal, para no tener dos versiones del mismo plan.
    """
    raiz = make_or_break(carpeta_raiz)
    generados = []
    for mes, planes in planes_por_mes.items():
        if mes_omitido is not None and mes == mes_omitido:
            continue
        for esquema in ESQUEMAS:
            plan = planes.get(esquema)
            if plan is None:
                continue
            etiqueta = esquema.replace("/", "")          # 24/5 -> 245
            carpeta = make_or_break(raiz / mes / etiqueta)
            for dc in LINEAS:
                cro = plan["cronogramas"].get(dc)
                if cro is None or cro.empty:
                    continue
                archivo = carpeta / f"{mes}_{etiqueta}_{dc}_{tipo}.xlsx"
                nombre_hoja = f"{tipo}_{dc}_{etiqueta}"[:31]
                with pd.ExcelWriter(archivo, engine="openpyxl") as writer:
                    cro.to_excel(writer, sheet_name=nombre_hoja, index=False)
                    hoja = writer.book[nombre_hoja]
                    dar_formato(hoja, cro, FORMATO_PIEZAS)
                    pintar_cronograma(hoja, cro)
                generados.append(archivo)
    log(f"Cronogramas de produccion a futuro: {len(generados)} archivos en "
        f"{raiz}")
    return generados


    

# ==============================================================================
# 6. Exportacion a Excel
# ==============================================================================
def armar_resumen(mes, horas_paro, planes, inicios, lote, origen):
    """Tabla de trazabilidad que encabeza el libro."""
    filas = [
        ("Fecha de ejecucion", pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("Mes planeado", str(mes)),
        ("Jeringas por lote", f"{lote:,}"),
        ("Horas por lote", f"DC1 = {PROD_TIME['DC1']} h | DC2 = {PROD_TIME['DC2']} h"),
        ("Limpieza por campana", f"{CLEAN_TIME} h"),
        ("Horas de paro DC1", f"{horas_paro.get('DC1', 0):,.1f}"),
        ("Horas de paro DC2", f"{horas_paro.get('DC2', 0):,.1f}"),
        ("Inicio DC1", inicios["DC1"].strftime("%d/%m/%Y %H:%M")),
        ("Inicio DC2", inicios["DC2"].strftime("%d/%m/%Y %H:%M")),
        ("Origen del forecast", origen),
    ]
    for esquema in ESQUEMAS:
        plan = planes.get(esquema)
        if plan is None:
            filas.append((f"Plan {esquema}", "No generado: no alcanzan las horas"))
            continue
        estado = plan["estado"]
        for _, f in estado.iterrows():
            filas.append((
                f"Plan {esquema} - {f[COL_DC]}",
                f"{f['horas_requeridas']:,.1f} h de {f['horas_disponibles']:,.1f} h "
                f"| {f['campanas']} campanas | holgura {f['holgura_horas']:,.1f} h"))
        filas.append((f"Traslados {esquema}", str(len(plan["movimientos"]))))
    return pd.DataFrame(filas, columns=["Concepto", "Valor"])


def dar_formato(hoja, tabla, formato_numeros):
    """Encabezado verde, anchos, filtro y formato numerico/de fecha."""
    for celda in hoja[1]:
        celda.font = Font(name=FUENTE, size=10, bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor=VERDE_MERCK)
        celda.alignment = Alignment(horizontal="center", vertical="center",
                                    wrap_text=True)
    hoja.row_dimensions[1].height = 28
    for fila in hoja.iter_rows(min_row=2):
        for celda in fila:
            celda.font = Font(name=FUENTE, size=10)
            if isinstance(celda.value, datetime):
                celda.number_format = FORMATO_FECHA
            elif isinstance(celda.value, (int, float)):
                celda.number_format = formato_numeros
    for i, columna in enumerate(tabla.columns, start=1):
        ancho_datos = int(tabla[columna].astype(str).str.len().max()) if len(tabla) else 0
        ancho = min(max(len(str(columna)) + 4, ancho_datos + 2, 10), 42)
        hoja.column_dimensions[get_column_letter(i)].width = ancho
    hoja.freeze_panes = "A2"
    hoja.auto_filter.ref = hoja.dimensions


def exportar_excel(ruta_salida, resumen, planes_resumen, planes):
    """Escribe el libro: Resumen, Planes resumen, Traslados y cronogramas.

    La hoja "Horas disponibles" se cambio por "Planes resumen": en vez de lo
    que el calendario ofrecia en teoria, ahora se reporta lo que realmente
    salio de cada cronograma, con las horas de paro al lado para comparar.
    """
    ruta_salida = Path(ruta_salida).with_suffix(".xlsx")
    ruta_salida.parent.mkdir(parents=True, exist_ok=True)

    traslados = []
    for esquema, plan in planes.items():
        movimientos = plan["movimientos"]
        if len(movimientos):
            copia = movimientos.copy()
            copia.insert(0, "esquema", esquema)
            traslados.append(copia)
    tabla_traslados = (pd.concat(traslados, ignore_index=True) if traslados
                       else pd.DataFrame(columns=["esquema", COL_SKU, COL_GRANEL,
                                                  "de", "a", "jeringas", "lotes",
                                                  "horas_en_destino", "motivo"]))

    hojas = {
        "Resumen": (resumen, "@"),
        "Planes resumen": (planes_resumen, FORMATO_HORAS),
    }
    # Un cronograma por esquema y linea. Los esquemas que no caben no llegan aqui.
    for esquema, plan in planes.items():
        etiqueta = esquema.replace("/", "")
        for dc, cro in plan["cronogramas"].items():
            hojas[f"Cronograma {dc} {etiqueta}"] = (cro, FORMATO_PIEZAS)

    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        for nombre, (tabla, formato) in hojas.items():
            tabla = pd.DataFrame() if tabla is None else tabla
            if not len(tabla.columns):
                tabla = pd.DataFrame({"Sin registros": []})
            tabla.to_excel(writer, sheet_name=nombre, index=False)
            dar_formato(writer.book[nombre], tabla, formato)
            if nombre.startswith("Cronograma"):
                pintar_cronograma(writer.book[nombre], tabla)
        writer.book["Resumen"].column_dimensions["B"].width = 70

    log(f"Excel generado: {ruta_salida}")
    return ruta_salida


# ==============================================================================
# 7. Proceso completo
# ==============================================================================
def ejecutar(origen, ruta_salida, mes=None, horas_paro=None,
             fecha_inicio_dc1=None, fecha_inicio_dc2=None, lote=TAMANO_LOTE,
             hoja_forecast=HOJA_FORECAST, ruta_paros=None, hoja_paros=None,
             generar_futuros=False, carpeta_futuros=None):
    """Corre el plan de produccion de UN mes y escribe el Excel.

    `origen` es el resultado del balanceo: el dict que devuelve
    planning_balanceo.ejecutar(), un DataFrame del forecast balanceado, o la
    ruta del .xlsx que exporto el balanceo.
    `mes` es la etiqueta MM-AAAA; si no se manda, se usa el mes en curso.
    Siempre se planea 24/5. El plan 24/7 solo se genera si las horas alcanzan.

    Las horas de paro salen de la hoja de paros (`ruta_paros` + `hoja_paros`).
    Un mes sin filas en esa hoja vale cero horas de paro. `horas_paro` sigue
    aceptandose para forzar valores a mano; si viene, manda sobre la hoja.

    Con `generar_futuros` se exportan ademas los cronogramas de TODOS los meses
    del horizonte, uno por archivo, bajo `carpeta_futuros`. El libro principal
    del mes elegido se escribe siempre, marquese o no esa opcion.
    """
    mes = mes or (pd.Timestamp.today().strftime(FORMATO_MES))
    tabla_paros = normalizar_tabla_paros(ruta_paros, hoja_paros)
    if horas_paro is None:
        horas_paro = leer_horas_paro(mes=mes, tabla=tabla_paros)
        log(f"Horas de paro de la hoja | " + " | ".join(
            f"{dc}: {horas_paro.get(dc, 0):,.1f} h" for dc in LINEAS))
    else:
        horas_paro = dict(horas_paro)

    etiqueta_origen = ("resultado en memoria del balanceo"
                       if isinstance(origen, (dict, pd.DataFrame))
                       else f"{Path(origen).name} (hoja '{hoja_forecast}')")
    log(f"Inicio del plan | mes={mes} | origen: {etiqueta_origen}")

    forecast = obtener_forecast(origen, hoja_forecast)
    base, meses = preparar_mes(forecast, mes, lote)
    if base.empty:
        raise ValueError(f"El mes '{mes}' no tiene demanda por planear.")
    
    anio, numero = mes_a_anio_mes(mes)
    inicio_mes = datetime(anio, numero, 1, HORA_INICIO_DEFECTO, 0)
    
    inicios = {"DC1": fecha_inicio_dc1 or inicio_mes,
               "DC2": fecha_inicio_dc2 or inicio_mes}

    planes = {}
    for esquema in ESQUEMAS:
        df, movimientos, estado = trasladar_para_ajustar(base, mes, esquema,
                                                        horas_paro)
        df = armar_campanas(df)


        if not bool(estado["factible"].all()):
            faltan = estado.loc[~estado["factible"]]
            detalle = " | ".join(f"{f[COL_DC]} excede por "
                                f"{-f['holgura_horas']:,.1f} h"
                                for _, f in faltan.iterrows())
            if esquema == ESQUEMA_OPCIONAL:
                log(f"Plan {esquema} no generado: {detalle}")
                planes[esquema] = None
                continue
            log(f"Aviso: el plan {esquema} no cabe en el mes ({detalle}). "
                "Se entrega el cronograma para que se vea el desborde.")

        inicios_esquema = {dc: _avanzar_tiempo(inicios[dc], 0, esquema) for dc in LINEAS}   # <-- linea nueva
        
        cronogramas = {dc: generar_cronograma(df, dc, inicios_esquema[dc], esquema)
                    for dc in LINEAS if (df[COL_DC] == dc).any()}
        planes[esquema] = {"detalle": df, "movimientos": movimientos,
                        "estado": estado, "cronogramas": cronogramas,
                        "inicios": inicios_esquema}   # <-- se guarda para el resumen
        log(f"Plan {esquema} listo | " + " | ".join(
            f"{f[COL_DC]}: {f['horas_requeridas']:,.1f}/{f['horas_disponibles']:,.1f} h"
            for _, f in estado.iterrows()))

    generados = {e: p for e, p in planes.items() if p is not None}
    resumen = armar_resumen(mes, horas_paro, planes, inicios, lote,
                            etiqueta_origen)

    # Los planes a futuro se arman antes de exportar para que la hoja de
    # resumen del libro principal pueda abarcar todo el horizonte.
    planes_por_mes = {mes: planes}
    archivos_futuros = []
    if generar_futuros:
        log("Armando los planes del resto del horizonte...")
        planes_por_mes = planes_del_horizonte(forecast, lote, tabla_paros)
        planes_por_mes[mes] = planes          # el mes elegido manda
        raiz = carpeta_futuros or Path(ruta_salida).parent
        archivos_futuros = gen_all_plans_xlsx(raiz, planes_por_mes,
                                              mes_omitido=mes)

    planes_resumen = armar_planes_resumen(planes_por_mes, tabla_paros)
    ruta_excel = exportar_excel(ruta_salida, resumen, planes_resumen, generados)

    log(f"Fin del plan | esquemas generados: "
        f"{', '.join(generados) if generados else 'ninguno'}")

    return {
        "mes": mes,
        "base": base,
        "planes": planes,
        "planes_resumen": planes_resumen,
        "resumen": resumen,
        "ruta_excel": ruta_excel,
        "archivos_futuros": archivos_futuros,}