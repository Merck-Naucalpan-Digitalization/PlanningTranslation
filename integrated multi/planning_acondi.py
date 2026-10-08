"""
planning_acondi.py
Paso 3 del proceso: plan de ACONDICIONAMIENTO (empaque) en las lineas L2 / L4.

Parte del mismo "Forecast balanceado" que exporta planning_balanceo.py, pero
aqui ya no importa si la jeringa se fabrico en DC1 o DC2: el acondicionamiento
ve un solo numero por SKU y por mes, y lo reparte entre sus dos lineas de
empaque.

Como funciona, en cinco pasos:
    1. De la hoja "MD" del catalogo se sacan las relaciones entre SKU. Dos SKU
       se comparan en cuatro atributos de formato y su `relation` es
       (atributos distintos) * 10, asi que solo puede valer 0, 10, 20, 30 o 40.
    2. Del forecast balanceado se quita la columna DC y se suma la demanda por
       SKU, de modo que quede un solo renglon por SKU y mes.
    3. Por cada mes la demanda se reparte entre L2 y L4: primero los SKU
       exclusivos de cada linea y luego los compartidos, en bloques de un lote
       completo, hacia la linea que lleve menos carga.
    4. Cada linea se secuencia para que el plan pague el menor numero de
       cambios de formato posible (ver `secuenciar`).
    5. Se arma el cronograma hora por hora, el resumen por mes/linea/esquema y
       se exporta todo a Excel.

Reglas de negocio modeladas:
- Un lote de 115,200 jeringas se acondiciona en 6.5 h.
- Entre dos SKU consecutivos solo se paga cambio de formato si su `relation`
  es MAYOR a 10, y entonces cuesta 9 h fijas (da igual si es 20, 30 o 40).
- Por eso la relation NO es el costo en horas: lo que encarece el plan es
  cuantas transiciones cruzan el umbral, no cuanto suman. El secuenciador
  minimiza primero los cambios de formato y solo desempata por relation.
- Esquemas 24/5 y 24/7, con los mismos cortes de fin de semana que
  planning_produccion.py (se reutiliza su reloj).
- Las cantidades del forecast no se modifican: lo unico que se decide es en
  que linea y en que orden se acondiciona cada SKU.

Igual que los otros modulos, todos los mensajes salen por `log`, asi que la
app de escritorio puede reemplazar esa funcion por la suya.

Requisitos: pip install pandas openpyxl
"""

from datetime import datetime
from itertools import combinations
from pathlib import Path

import pandas as pd

import planning_produccion as pp

# ---------------------------------------------------------------------------
# Parametros de negocio (la app puede pasar otros)
# ---------------------------------------------------------------------------
TAMANO_LOTE = 115_200            # jeringas por lote
HORAS_POR_LOTE = 6.5             # horas de acondicionamiento por lote
HORAS_CAMBIO_FORMATO = 9         # horas que cuesta un cambio de formato
UMBRAL_CAMBIO = 10               # relation > UMBRAL_CAMBIO pide cambio
RELATION_DESCONOCIDA = 40

# ------------------------------------------------------------------------------
# NOTA DE RENDIMIENTO (pendiente, no tocar todavia)
#
# La hoja MD trae del orden de 3,000 relaciones. Hoy se cargan completas a un
# dict {(sku_a, sku_b): relation} en `mapa_relaciones`, y se consulta con
# `relacion()`. En numeros: 3,000 llaves de tupla son unos pocos cientos de KB,
# asi que la memoria no es el problema. Lo que puede doler es el secuenciador,
# que consulta el mapa O(n^2) veces por linea y por mes, multiplicado por los
# arranques del greedy y las pasadas de 2-opt.
#
# Antes de optimizar hay que medir. El orden sugerido cuando se llegue a eso:
#   1. Perfilar con cProfile una corrida de horizonte completo y ver si el
#      tiempo esta en la lectura del Excel, en el armado del mapa o en el
#      secuenciador. La sospecha es el tercero.
#   2. Si es el secuenciador: pasar de dict a una matriz de numpy indexada por
#      posicion, con los SKU traducidos a enteros una sola vez. Las consultas
#      dejan de hashear tuplas y 2-opt puede evaluar el delta vectorizado.
#   3. Si es la lectura: cachear el mapa por (ruta, hoja, fecha de modificacion)
#      para no releer el mismo Excel en cada mes del horizonte.
#   4. Si es el armado: construirlo con `to_dict` sobre el DataFrame ya filtrado
#      en vez de iterar renglon por renglon.
# Lo que NO conviene es adelantarse: hoy la version con dict es la legible, y
# sin medicion no se sabe si el cambio gana algo.
# ------------------------------------------------------------------------------
        # peor caso cuando un par no esta en MD
LINEAS_ACONDI = ("L2", "L4")
N_EXACTO = 10                    # hasta este tamano se busca el orden optimo

# Nombres de columna que comparte con los otros modulos
COL_SKU = pp.COL_SKU             # "SKUMERCK"
COL_DESC = pp.COL_DESC           # "Description"
COL_LINEA = "Linea"
COL_GRANEL = pp.COL_GRANEL       # "Nombre granel"
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
FORMATO_PIEZAS = pp.FORMATO_PIEZAS
FORMATO_HORAS = pp.FORMATO_HORAS

# Colores del cronograma
COLOR_BLOQUE_A = "DDEBF7"        # azul muy claro
COLOR_BLOQUE_B = "FCE4D6"        # naranja muy claro
COLOR_CAMBIO = "D9D9D9"          # gris: cambio de formato
COLOR_FIN_SEMANA = "BFBFBF"      # gris oscuro: la linea esta detenida

ETIQUETA_CAMBIO = "Cambio de formato"
ETIQUETA_FIN_SEMANA = "Fin de semana"

COLS_CRONOGRAMA = ["Nombre", COL_SKU, "Secuencia", "Jeringas", "Lotes",
                   "Horas", "Relation previa",
                   "Fecha y hora de inicio", "Fecha y hora de finalizacion"]


def log(mensaje):
    """Unico punto de salida de mensajes. La app de escritorio la reemplaza."""
    print(mensaje)


# ==============================================================================
# 1. Lectura de MD y relaciones entre SKU
# ==============================================================================
def leer_md(ruta_catalogo, hoja=HOJA_MD, encabezado=HOJA_MD_ENCABEZADO):
    """Lee la hoja MD del catalogo y la parte en las dos tablas de linea.

    Solo se queda con la familia DC. Un SKU marcado 'L2/L4' aparece en las dos
    tablas: es un SKU compartido y se puede acondicionar en cualquiera.
    Regresa (md, md_por_linea) donde md_por_linea es {'L2': df, 'L4': df}.
    """
    md = pd.read_excel(Path(ruta_catalogo), sheet_name=hoja, header=encabezado)
    md = md.rename(columns={"SKU": COL_SKU})

    faltantes = [c for c in COLS_MD if c not in md.columns]
    if faltantes:
        raise ValueError(
            f"La hoja '{hoja}' no trae las columnas {faltantes}. "
            f"Revise que el encabezado este en el renglon {encabezado + 1}.")

    md = md[COLS_MD].copy()
    md[COL_SKU] = md[COL_SKU].astype(str).str.strip()
    md = md.dropna(subset=[COL_SKU])
    md = md[~md[COL_SKU].isin(("", "nan", "None"))]
    md = md[md["familia"].astype(str).str.strip() == FAMILIA_DC]
    md[COL_LINEA] = md[COL_LINEA].astype(str).str.strip()
    md = md.sort_values(COL_SKU).reset_index(drop=True)

    md_por_linea = {}
    for linea, marcas in MARCAS_LINEA.items():
        bloque = md[md[COL_LINEA].isin(marcas)].copy()
        md_por_linea[linea] = bloque.reset_index(drop=True)

    compartidos = set(md_por_linea["L2"][COL_SKU]) & set(md_por_linea["L4"][COL_SKU])
    log(f"MD: {len(md)} SKU de familia {FAMILIA_DC} | "
        f"L2: {len(md_por_linea['L2'])} | L4: {len(md_por_linea['L4'])} | "
        f"compartidos: {len(compartidos)}")
    return md, md_por_linea


def calcular_relaciones(md_por_linea):
    """Matriz de relacion por linea: (atributos de formato distintos) * 10.

    Se compara cada par de SKU de la misma linea en las cuatro columnas de
    formato. Entre mas alta la relation, mas tarda el cambio entre esos dos
    productos. Esta funcion solo necesita correr una vez por catalogo.
    Regresa {'L2': df, 'L4': df} con columnas sku_1, sku_2, relation.
    """
    pesos = {}
    for linea, tabla in md_por_linea.items():
        atributos = tabla.set_index(COL_SKU)[COLS_FORMATO]
        atributos = atributos[~atributos.index.duplicated(keep="first")]
        filas = []
        for sku_1, sku_2 in combinations(atributos.index, 2):
            diffs = int((atributos.loc[sku_1] != atributos.loc[sku_2]).sum())
            filas.append((sku_1, sku_2, diffs * 10))
        pesos[linea] = pd.DataFrame(filas,
                                    columns=["sku_1", "sku_2", "relation"])
        log(f"Relaciones {linea}: {len(filas)} pares")
    return pesos


def mapa_relaciones(tabla_pesos):
    """Convierte la tabla de pesos en un diccionario simetrico de consulta.

    La tabla solo trae la mitad superior de la matriz, asi que aqui se
    registran las dos direcciones. Un SKU contra si mismo es relation 0.

    NOTA DE RENDIMIENTO (pendiente de optimizar):
    con ~60 SKU por linea salen ~1,800 pares por tabla y cerca de 3,000 en
    total entre L2 y L4, asi que el diccionario queda en unos 6,000 pares
    contando las dos direcciones. Para el tamano actual eso son unos cientos
    de KB y la consulta es O(1), que es justo lo que necesitan el greedy, la
    reinsercion y el 2-opt: todos piden la misma relacion miles de veces.
    Lo que si conviene revisar cuando se haga la pasada de desempeno:
        - `mapa_relaciones` se vuelve a construir en cada llamada a
          `secuenciar` y a `generar_cronograma`, asi que en una corrida de 12
          meses x 2 lineas x 2 esquemas se arma decenas de veces. Conviene
          cachearlo una sola vez por linea (es constante en todo el
          horizonte) y pasarlo ya armado.
        - si el catalogo crece mucho, cambiar el dict por una matriz de
          numpy indexada por posicion de SKU: menos memoria y acceso mas
          rapido que las tuplas de strings.
    """
    mapa = {}
    for sku_1, sku_2, relation in tabla_pesos.itertuples(index=False):
        valor = float(relation)
        mapa[(sku_1, sku_2)] = valor
        mapa[(sku_2, sku_1)] = valor
    return mapa


def relacion(mapa, sku_1, sku_2, avisos=None):
    """Relation entre dos SKU. Un par que no esta en MD se castiga.

    Si el par no aparece en la tabla se asume el peor caso para que un hueco
    en MD no produzca un plan optimista falso, y se deja constancia.
    """
    if sku_1 == sku_2:
        return 0.0
    clave = (sku_1, sku_2)
    if clave in mapa:
        return mapa[clave]
    if avisos is not None:
        avisos.add(clave)
    return float(RELATION_DESCONOCIDA)


def pide_cambio(valor_relation):
    """True si la transicion cruza el umbral y hay que parar a cambiar formato."""
    return valor_relation > UMBRAL_CAMBIO


# ==============================================================================
# 2. Demanda de acondicionamiento: un renglon por SKU
# ==============================================================================
def preparar_demanda(forecast):
    """Colapsa el forecast balanceado a un renglon por SKU.

    El acondicionamiento no distingue DC1 de DC2: las dos lineas de llenado
    alimentan al mismo empaque, asi que la demanda se suma por SKU. Se
    conservan la descripcion y el granel solo para que el reporte se lea.
    """
    tabla = forecast.copy()
    meses = pp.meses_disponibles(tabla)
    if not meses:
        raise ValueError("El forecast balanceado no trae columnas de mes.")

    tabla[COL_SKU] = tabla[COL_SKU].astype(str).str.strip()
    for mes in meses:
        tabla[mes] = pd.to_numeric(tabla[mes], errors="coerce").fillna(0.0)

    demanda = tabla.groupby(COL_SKU, as_index=False)[meses].sum()

    descriptivas = [c for c in (COL_DESC, COL_GRANEL) if c in tabla.columns]
    if descriptivas:
        etiquetas = (tabla[[COL_SKU] + descriptivas]
                     .drop_duplicates(subset=[COL_SKU]))
        demanda = demanda.merge(etiquetas, on=COL_SKU, how="left")

    demanda = demanda[[COL_SKU] + descriptivas + meses]
    demanda = demanda.sort_values(COL_SKU).reset_index(drop=True)
    log(f"Demanda de acondicionamiento: {len(demanda)} SKU | "
        f"{len(meses)} meses ({meses[0]} a {meses[-1]})")
    return demanda, meses


# ==============================================================================
# 3. Reparto entre L2 y L4
# ==============================================================================
def repartir_lineas(demanda, mes, md_por_linea, lote=TAMANO_LOTE):
    """Reparte la demanda del mes entre L2 y L4 en bloques de un lote.

    Primero se asignan los SKU exclusivos, porque no hay decision que tomar:
    van a la unica linea que los puede correr. Con esa carga ya en el marcador
    se reparten los compartidos, de mayor a menor demanda y lote por lote,
    mandando cada lote a la linea que vaya mas descargada. Asi las dos lineas
    terminan lo mas parejas posible sin partir un lote a la mitad.

    El residuo que no completa un lote se manda tambien a la linea mas
    descargada. Esa regla es provisional: aqui es donde entrara la logica de
    residuos cuando se defina.
    """
    skus_l2 = set(md_por_linea["L2"][COL_SKU])
    skus_l4 = set(md_por_linea["L4"][COL_SKU])

    bloque = demanda[[COL_SKU, mes]].copy()
    bloque = bloque[bloque[mes] > 0]
    if bloque.empty:
        return pd.DataFrame(columns=[COL_LINEA, COL_SKU, COL_DESC,
                                     "jeringas", "lotes"])

    # La descripcion se arrastra solo como etiqueta para el reporte: todo el
    # reparto, la secuencia y las horas se calculan contra el SKU.
    if COL_DESC in demanda.columns:
        descripciones = (demanda.drop_duplicates(subset=[COL_SKU])
                         .set_index(COL_SKU)[COL_DESC].to_dict())
    else:
        descripciones = {}

    sin_md = sorted(set(bloque[COL_SKU]) - (skus_l2 | skus_l4))
    if sin_md:
        log(f"Aviso {mes}: {len(sin_md)} SKU con demanda no estan en MD y "
            f"quedan fuera del plan de acondicionamiento ({sin_md[:5]}...)")
        bloque = bloque[~bloque[COL_SKU].isin(sin_md)]

    carga = {linea: 0.0 for linea in LINEAS_ACONDI}
    asignado = {linea: {} for linea in LINEAS_ACONDI}

    def anotar(linea, sku, piezas):
        asignado[linea][sku] = asignado[linea].get(sku, 0.0) + piezas
        carga[linea] += piezas

    # --- SKU exclusivos: no hay nada que decidir ---
    compartidos = []
    for sku, piezas in bloque.itertuples(index=False):
        piezas = float(piezas)
        en_l2, en_l4 = sku in skus_l2, sku in skus_l4
        if en_l2 and en_l4:
            compartidos.append((sku, piezas))
        elif en_l2:
            anotar("L2", sku, piezas)
        else:
            anotar("L4", sku, piezas)

    # --- SKU compartidos: lote por lote hacia la linea mas descargada ---
    for sku, piezas in sorted(compartidos, key=lambda par: -par[1]):
        lotes_completos = int(piezas // lote)
        residuo = piezas - lotes_completos * lote
        for _ in range(lotes_completos):
            destino = min(LINEAS_ACONDI, key=lambda ln: carga[ln])
            anotar(destino, sku, float(lote))
        if residuo > 1e-9:
            destino = min(LINEAS_ACONDI, key=lambda ln: carga[ln])
            anotar(destino, sku, residuo)

    filas = [{COL_LINEA: linea, COL_SKU: sku,
              COL_DESC: descripciones.get(sku, ""),
              "jeringas": piezas, "lotes": piezas / lote}
             for linea in LINEAS_ACONDI
             for sku, piezas in asignado[linea].items()]
    reparto = pd.DataFrame(filas).sort_values([COL_LINEA, COL_SKU])

    log(f"Reparto {mes}: " + " | ".join(
        f"{ln} {carga[ln]:,.0f} jeringas ({carga[ln] / lote:,.2f} lotes)"
        for ln in LINEAS_ACONDI))
    return reparto.reset_index(drop=True)


# ==============================================================================
# 4. Secuencia: minimo numero de cambios de formato
# ==============================================================================
def costo_secuencia(orden, mapa, avisos=None):
    """Costo de un orden: (cambios de formato, relation acumulada).

    La tupla se compara en ese orden de prioridad, asi que entre dos secuencias
    gana la que pague menos cambios de formato y, solo si empatan, la de menor
    relation total. Esto importa porque la relation no es proporcional al
    tiempo: 20, 30 y 40 cuestan los mismos 9 h, y 0 y 10 cuestan cero.
    """
    cambios = 0
    suma = 0.0
    for previo, actual in zip(orden, orden[1:]):
        valor = relacion(mapa, previo, actual, avisos)
        suma += valor
        if pide_cambio(valor):
            cambios += 1
    return cambios, suma


def _mejor_greedy(skus, mapa):
    """Construye una secuencia arrancando en cada SKU y se queda con la mejor.

    El greedy salta siempre al vecino mas barato que quede libre. Por si solo
    se pinta en una esquina: al final deja un SKU huerfano caro. Probar todos
    los arranques es barato y da un punto de partida mucho mejor.
    """
    mejor_orden, mejor_costo = None, None
    for inicio in skus:
        orden = [inicio]
        libres = set(skus) - {inicio}
        while libres:
            actual = orden[-1]
            siguiente = min(
                libres,
                key=lambda sku: (pide_cambio(relacion(mapa, actual, sku)),
                                 relacion(mapa, actual, sku), sku))
            orden.append(siguiente)
            libres.discard(siguiente)
        costo = costo_secuencia(orden, mapa)
        if mejor_costo is None or costo < mejor_costo:
            mejor_orden, mejor_costo = orden, costo
    return mejor_orden, mejor_costo


def _reinsertar(orden, mapa):
    """Saca el SKU de cada transicion cara y lo reacomoda donde estorbe menos.

    Esta es la pasada que pidio el negocio: localizar las transiciones de 20,
    30 y 40, y ver si ese producto cabe entre otros con los que tenga 0 o 10
    sin encarecer el resto del plan. Se mueve un SKU a la vez y solo se acepta
    el movimiento si el costo total baja.
    """
    mejor = list(orden)
    mejor_costo = costo_secuencia(mejor, mapa)
    mejoro = True
    while mejoro:
        mejoro = False
        # Posiciones involucradas en una transicion que paga cambio de formato.
        sospechosos = set()
        for i in range(len(mejor) - 1):
            if pide_cambio(relacion(mapa, mejor[i], mejor[i + 1])):
                sospechosos.update((i, i + 1))
        for i in sorted(sospechosos, reverse=True):
            resto = mejor[:i] + mejor[i + 1:]
            sku = mejor[i]
            for j in range(len(resto) + 1):
                candidato = resto[:j] + [sku] + resto[j:]
                costo = costo_secuencia(candidato, mapa)
                if costo < mejor_costo:
                    mejor, mejor_costo = candidato, costo
                    mejoro = True
                    break
            if mejoro:
                break
    return mejor, mejor_costo


def _dos_opt(orden, mapa):
    """Invierte segmentos internos y conserva la inversion si abarata el plan.

    Arregla los cruces que deja el greedy y que la reinsercion de un solo SKU
    no alcanza a deshacer. Se repite hasta que ninguna inversion mejore.
    """
    mejor = list(orden)
    mejor_costo = costo_secuencia(mejor, mapa)
    mejoro = True
    while mejoro:
        mejoro = False
        for i in range(len(mejor) - 1):
            for j in range(i + 1, len(mejor)):
                candidato = mejor[:i] + mejor[i:j + 1][::-1] + mejor[j + 1:]
                costo = costo_secuencia(candidato, mapa)
                if costo < mejor_costo:
                    mejor, mejor_costo = candidato, costo
                    mejoro = True
                    break
            if mejoro:
                break
    return mejor, mejor_costo


def _exacto(skus, mapa):
    """Orden optimo por busqueda exhaustiva con poda. Solo para pocos SKU."""
    mejor = {"orden": None, "costo": None}

    def avanzar(parcial, libres, cambios, suma):
        if mejor["costo"] is not None and (cambios, suma) > mejor["costo"]:
            return                                  # ya es peor: se poda
        if not libres:
            mejor["orden"], mejor["costo"] = list(parcial), (cambios, suma)
            return
        actual = parcial[-1]
        siguientes = sorted(
            libres, key=lambda sku: (relacion(mapa, actual, sku), sku))
        for sku in siguientes:
            valor = relacion(mapa, actual, sku)
            parcial.append(sku)
            avanzar(parcial, libres - {sku},
                    cambios + (1 if pide_cambio(valor) else 0), suma + valor)
            parcial.pop()

    for inicio in skus:
        avanzar([inicio], set(skus) - {inicio}, 0, 0.0)
    return mejor["orden"], mejor["costo"]


def secuenciar(skus, tabla_pesos, n_exacto=N_EXACTO):
    """Ordena los SKU de una linea para pagar el menor numero de cambios.

    Es un problema de ruta abierta: no hay formula cerrada y el numero de
    ordenes posibles crece como n factorial. Por eso:
        - hasta `n_exacto` SKU se busca el optimo con poda;
        - arriba de eso se construye un greedy desde cada arranque, se
          reacomodan los SKU de las transiciones caras y se cierra con 2-opt.
    Regresa (orden, cambios, relation_total, avisos) donde `avisos` son los
    pares que no aparecen en MD.
    """
    skus = list(dict.fromkeys(skus))
    mapa = mapa_relaciones(tabla_pesos)
    avisos = set()

    if len(skus) <= 1:
        return skus, 0, 0.0, avisos

    if len(skus) <= n_exacto:
        orden, (cambios, suma) = _exacto(skus, mapa)
        metodo = "exacto"
    else:
        orden, _ = _mejor_greedy(skus, mapa)
        orden, _ = _reinsertar(orden, mapa)
        orden, (cambios, suma) = _dos_opt(orden, mapa)
        metodo = "greedy + reinsercion + 2-opt"

    cambios, suma = costo_secuencia(orden, mapa, avisos)
    log(f"Secuencia {metodo}: {len(orden)} SKU | {cambios} cambios de formato "
        f"({cambios * HORAS_CAMBIO_FORMATO:,.1f} h) | relation total {suma:,.0f}")
    return orden, cambios, suma, avisos


# ==============================================================================
# 5. Cronograma de acondicionamiento
# ==============================================================================
def _fila_vacia(nombre, horas=None):
    """Renglon separador del cronograma (fin de semana o bloque sin reloj)."""
    return {"Nombre": nombre, COL_SKU: "", "Secuencia": None, "Jeringas": None,
            "Lotes": None, "Horas": horas, "Relation previa": None,
            "Fecha y hora de inicio": None,
            "Fecha y hora de finalizacion": None}


def generar_cronograma(reparto, linea, orden, tabla_pesos, fecha_inicio,
                       esquema, lote=TAMANO_LOTE):
    """Una fila por lote y una fila 'Cambio de formato' cuando toca pararse.

    Se respeta el orden que entrego el secuenciador. Igual que en produccion,
    en 24/5 un lote no se puede pausar: si no cabe antes del sabado a las
    00:00 se inserta un renglon de fin de semana y el lote abre el lunes a las
    07:00. El reloj es el de planning_produccion, para que los dos planes
    midan el calendario igual.
    """
    bloque = reparto[reparto[COL_LINEA] == linea].set_index(COL_SKU)
    mapa = mapa_relaciones(tabla_pesos)
    filas = []
    tiempo = pp._avanzar_tiempo(fecha_inicio, 0, esquema)
    previo = None

    def colocar(nombre, sku, horas, jeringas=None, lotes=None,
                secuencia=None, relation_previa=None):
        """Agenda un bloque de trabajo cuidando el corte de fin de semana."""
        nonlocal tiempo
        cabe, _, corte = pp._cabe_sin_pausa(tiempo, horas, esquema)
        if not cabe:
            filas.append(_fila_vacia(ETIQUETA_FIN_SEMANA))
            tiempo = pp._avanzar_tiempo(corte, 0, esquema)
        fin = pp._avanzar_tiempo(tiempo, horas, esquema)
        filas.append({"Nombre": nombre, COL_SKU: sku, "Secuencia": secuencia,
                      "Jeringas": jeringas, "Lotes": lotes, "Horas": horas,
                      "Relation previa": relation_previa,
                      "Fecha y hora de inicio": tiempo,
                      "Fecha y hora de finalizacion": fin})
        tiempo = fin

    for posicion, sku in enumerate(orden, start=1):
        if sku not in bloque.index:
            continue
        jeringas_totales = float(bloque.at[sku, "jeringas"])
        if jeringas_totales <= 0:
            continue
        descripcion = (str(bloque.at[sku, COL_DESC])
                       if COL_DESC in bloque.columns
                       and pd.notna(bloque.at[sku, COL_DESC])
                       and str(bloque.at[sku, COL_DESC]).strip() != ""
                       else sku)

        valor_relation = (None if previo is None
                          else relacion(mapa, previo, sku))
        if valor_relation is not None and pide_cambio(valor_relation):
            colocar(ETIQUETA_CAMBIO, "", HORAS_CAMBIO_FORMATO,
                    relation_previa=valor_relation)

        lotes_completos = int(jeringas_totales // lote)
        residuo = jeringas_totales - lotes_completos * lote
        tramos = [(float(lote), HORAS_POR_LOTE)] * lotes_completos
        if residuo > 1e-9:
            tramos.append((residuo, (residuo / lote) * HORAS_POR_LOTE))

        for jeringas_tramo, horas_tramo in tramos:
            # En el cronograma se muestra la descripcion para que el plan se
            # lea; el SKU viaja en su propia columna y es el que manda.
            colocar(descripcion, sku, horas_tramo, jeringas=jeringas_tramo,
                    lotes=jeringas_tramo / lote, secuencia=posicion,
                    relation_previa=valor_relation)
            valor_relation = None       # la relation solo se reporta una vez

        previo = sku

    return pd.DataFrame(filas, columns=COLS_CRONOGRAMA)


def pintar_cronograma(hoja, tabla):
    """Colorea el cronograma: alterna por bloque y resalta las paradas.

    Mismo criterio visual que produccion, pero aqui el separador es el cambio
    de formato en lugar de la limpieza.
    """
    from openpyxl.styles import PatternFill

    if "Nombre" not in tabla.columns or tabla.empty:
        return
    colores = (COLOR_BLOQUE_A, COLOR_BLOQUE_B)
    turno = 0
    for i, nombre in enumerate(tabla["Nombre"], start=2):
        if nombre == ETIQUETA_CAMBIO:
            color = COLOR_CAMBIO
        elif nombre == ETIQUETA_FIN_SEMANA:
            color = COLOR_FIN_SEMANA
        else:
            color = colores[turno % 2]
        relleno = PatternFill("solid", fgColor=color)
        for celda in hoja[i]:
            celda.fill = relleno
        if nombre == ETIQUETA_CAMBIO:
            turno += 1


# ==============================================================================
# 6. Planes de todos los meses y resumen
# ==============================================================================
def planear_mes(demanda, mes, md_por_linea, pesos, inicio=None,
                lote=TAMANO_LOTE, esquemas=pp.ESQUEMAS):
    """Reparto, secuencia y cronogramas de un mes, para cada esquema.

    Regresa {'reparto': df, 'secuencias': {...}, 'cronogramas': {clave: df}}
    donde la clave de cronograma es (esquema, linea), igual que en produccion.
    """
    reparto = repartir_lineas(demanda, mes, md_por_linea, lote)
    if reparto.empty:
        log(f"Mes {mes}: sin demanda de acondicionamiento.")
        return None

    anio, numero = pp.mes_a_anio_mes(mes)
    inicio_mes = inicio or datetime(anio, numero, 1, pp.HORA_INICIO_DEFECTO, 0)

    secuencias, cronogramas = {}, {}
    for linea in LINEAS_ACONDI:
        skus = list(reparto.loc[reparto[COL_LINEA] == linea, COL_SKU])
        if not skus:
            continue
        orden, cambios, suma, avisos = secuenciar(skus, pesos[linea])
        secuencias[linea] = {"orden": orden, "cambios": cambios,
                             "relation_total": suma, "pares_sin_md": avisos}
        if avisos:
            log(f"Aviso {mes} / {linea}: {len(avisos)} pares sin relacion en "
                f"MD, se tomaron como {RELATION_DESCONOCIDA}.")
        for esquema in esquemas:
            arranque = pp._avanzar_tiempo(inicio_mes, 0, esquema)
            cronogramas[(esquema, linea)] = generar_cronograma(
                reparto, linea, orden, pesos[linea], arranque, esquema, lote)

    return {"reparto": reparto, "secuencias": secuencias,
            "cronogramas": cronogramas}


def planear_todos(demanda, meses, md_por_linea, pesos, lote=TAMANO_LOTE,
                  esquemas=pp.ESQUEMAS):
    """Corre `planear_mes` en todo el horizonte. Regresa {mes: plan}."""
    planes = {}
    for mes in meses:
        plan = planear_mes(demanda, mes, md_por_linea, pesos, lote=lote,
                           esquemas=esquemas)
        if plan is not None:
            planes[mes] = plan
    log(f"Planes de acondicionamiento generados: {len(planes)} meses")
    return planes


def resumen_acondi(planes, esquemas=pp.ESQUEMAS):
    """Tabla resumen por mes, linea y esquema.

    Misma idea que `resumen_esquemas` del plan de produccion: una fila por
    combinacion, con la ventana de fechas, las horas de trabajo, las horas
    perdidas en cambios de formato y la holgura contra el calendario.
    """
    filas = []
    for mes, plan in planes.items():
        anio, numero = pp.mes_a_anio_mes(mes)
        for linea in LINEAS_ACONDI:
            secuencia = plan["secuencias"].get(linea)
            if secuencia is None:
                continue
            for esquema in esquemas:
                cro = plan["cronogramas"].get((esquema, linea))
                if cro is None or cro.empty:
                    continue
                inicios = cro["Fecha y hora de inicio"].dropna()
                fines = cro["Fecha y hora de finalizacion"].dropna()
                es_cambio = cro["Nombre"] == ETIQUETA_CAMBIO
                horas_cambio = float(cro.loc[es_cambio, "Horas"].sum())
                horas_totales = float(cro["Horas"].sum())
                disp, _ = pp.horas_disponibles(anio, numero, esquema, 0)
                filas.append({
                    "Mes": mes,
                    COL_LINEA: linea,
                    "Esquema": esquema,
                    "Fecha inicio": inicios.iloc[0] if len(inicios) else None,
                    "Fecha fin": fines.iloc[-1] if len(fines) else None,
                    "SKU": int(cro[COL_SKU].replace("", pd.NA).nunique()),
                    "Jeringas": float(cro["Jeringas"].sum()),
                    "Lotes": float(cro["Lotes"].sum()),
                    "Horas acondicionamiento": horas_totales - horas_cambio,
                    "Cambios de formato": int(es_cambio.sum()),
                    "Horas cambio de formato": horas_cambio,
                    "Relation total": secuencia["relation_total"],
                    "Horas requeridas": horas_totales,
                    "Horas disponibles": disp,
                    "Holgura horas": disp - horas_totales,
                    "Factible": disp - horas_totales >= 0})
    columnas = ["Mes", COL_LINEA, "Esquema", "Fecha inicio", "Fecha fin",
                "SKU", "Jeringas", "Lotes", "Horas acondicionamiento",
                "Cambios de formato", "Horas cambio de formato",
                "Relation total", "Horas requeridas", "Horas disponibles",
                "Holgura horas", "Factible"]
    return pd.DataFrame(filas, columns=columnas)


def tabla_reparto(planes):
    """Reparto de todos los meses en una sola tabla, para auditar la division."""
    filas = []
    for mes, plan in planes.items():
        bloque = plan["reparto"].copy()
        bloque.insert(0, "Mes", mes)
        filas.append(bloque)
    if not filas:
        return pd.DataFrame(columns=["Mes", COL_LINEA, COL_SKU, "jeringas",
                                     "lotes"])
    return pd.concat(filas, ignore_index=True)


def tabla_secuencias(planes):
    """Orden final de cada linea y mes, con el costo que pago la secuencia."""
    filas = []
    for mes, plan in planes.items():
        for linea, secuencia in plan["secuencias"].items():
            for posicion, sku in enumerate(secuencia["orden"], start=1):
                filas.append({"Mes": mes, COL_LINEA: linea,
                              "Posicion": posicion, COL_SKU: sku,
                              "Cambios de la linea": secuencia["cambios"],
                              "Relation total de la linea":
                                  secuencia["relation_total"]})
    return pd.DataFrame(filas, columns=["Mes", COL_LINEA, "Posicion", COL_SKU,
                                        "Cambios de la linea",
                                        "Relation total de la linea"])


# ==============================================================================
# 7. Exportacion
# ==============================================================================
def armar_resumen(mes, origen_forecast, origen_md, lote, planes):
    """Tabla de trazabilidad que encabeza el libro."""
    filas = [
        ("Fecha de ejecucion", pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("Mes planeado", str(mes)),
        ("Jeringas por lote", f"{lote:,}"),
        ("Horas por lote", f"{HORAS_POR_LOTE} h"),
        ("Cambio de formato", f"{HORAS_CAMBIO_FORMATO} h cuando relation > "
                              f"{UMBRAL_CAMBIO}"),
        ("Lineas de acondicionamiento", " / ".join(LINEAS_ACONDI)),
        ("Origen del forecast", origen_forecast),
        ("Origen de relaciones", origen_md),
        ("Meses con plan", str(len(planes))),
    ]
    plan = planes.get(mes)
    if plan is not None:
        for linea, secuencia in plan["secuencias"].items():
            filas.append((
                f"Secuencia {linea}",
                f"{len(secuencia['orden'])} SKU | {secuencia['cambios']} "
                f"cambios ({secuencia['cambios'] * HORAS_CAMBIO_FORMATO:,.1f} h)"
                f" | relation total {secuencia['relation_total']:,.0f}"))
    return pd.DataFrame(filas, columns=["Concepto", "Valor"])


def exportar_excel(ruta_salida, resumen, planes_resumen, reparto, secuencias,
                   plan_mes, mes):
    """Libro del mes: Resumen, planes_resumen, Reparto, Secuencia y cronogramas.

    `planes_resumen` sustituye a la vieja hoja de horas disponibles: ya trae
    el calendario, las horas requeridas y la holgura en la misma tabla.
    """
    ruta_salida = Path(ruta_salida).with_suffix(".xlsx")
    ruta_salida.parent.mkdir(parents=True, exist_ok=True)

    hojas = {
        "Resumen": (resumen, "@"),
        "Planes resumen": (planes_resumen, FORMATO_HORAS),
        "Reparto L2 L4": (reparto, FORMATO_PIEZAS),
        "Secuencia": (secuencias, FORMATO_PIEZAS),
    }
    if plan_mes is not None:
        for (esquema, linea), cro in plan_mes["cronogramas"].items():
            etiqueta = esquema.replace("/", "")
            hojas[f"Acondi {linea} {etiqueta}"] = (cro, FORMATO_PIEZAS)

    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        for nombre, (tabla, formato) in hojas.items():
            tabla = pd.DataFrame() if tabla is None else tabla
            if not len(tabla.columns):
                tabla = pd.DataFrame({"Sin registros": []})
            tabla.to_excel(writer, sheet_name=nombre, index=False)
            pp.dar_formato(writer.book[nombre], tabla, formato)
            if nombre.startswith("Acondi"):
                pintar_cronograma(writer.book[nombre], tabla)
        writer.book["Resumen"].column_dimensions["B"].width = 70

    log(f"Excel de acondicionamiento generado: {ruta_salida}")
    return ruta_salida


# La creacion de carpetas vive en produccion: los dos modulos escriben en el
# mismo arbol, asi que no tiene sentido tener dos versiones de lo mismo.
make_or_break = pp.make_or_break


def gen_all_plans_xlsx(carpeta_raiz, planes, tipo="Acondi",
                       esquemas=pp.ESQUEMAS, lineas=LINEAS_ACONDI,
                       mes_omitido=None):
    """Guarda SOLO el cronograma de cada plan a futuro, un Excel por archivo.

    Estos archivos son el plan pelon: nada de resumen ni de reparto, solo el
    cronograma. El libro completo con todas las hojas es el principal, el del
    mes elegido en la app, y ese se escribe aparte.

    El arbol es el mismo que usa el plan de produccion, para que los dos tipos
    de plan del mismo mes y esquema caigan en la misma carpeta:
        <raiz>/<mes>/<esquema>/<mes>_<esquema>_<linea>_Acondi.xlsx
    Asi en '2027/01-2027/245' conviven los cronogramas de DC1, DC2, L2 y L4.

    `mes_omitido` salta el mes que ya se entrego en el libro principal, para no
    duplicarlo.
    """
    raiz = Path(carpeta_raiz)
    make_or_break(raiz)
    generados = []
    for mes, plan in planes.items():
        if mes_omitido is not None and mes == mes_omitido:
            continue
        for esquema in esquemas:
            etiqueta = esquema.replace("/", "")           # 24/5 -> 245
            carpeta = raiz / mes / etiqueta
            make_or_break(carpeta)
            for linea in lineas:
                cro = plan["cronogramas"].get((esquema, linea))
                if cro is None or cro.empty:
                    continue
                archivo = carpeta / f"{mes}_{etiqueta}_{linea}_{tipo}.xlsx"
                nombre_hoja = f"{tipo}_{linea}_{etiqueta}"[:31]
                with pd.ExcelWriter(archivo, engine="openpyxl") as writer:
                    cro.to_excel(writer, sheet_name=nombre_hoja, index=False)
                    hoja = writer.book[nombre_hoja]
                    pp.dar_formato(hoja, cro, FORMATO_PIEZAS)
                    pintar_cronograma(hoja, cro)
                generados.append(archivo)
    log(f"Cronogramas de acondicionamiento a futuro: {len(generados)} archivos "
        f"en {raiz}")
    return generados


# ==============================================================================
# 8. Proceso completo
# ==============================================================================
def ejecutar(origen, ruta_catalogo, ruta_salida, hoja_md=HOJA_MD, mes=None,
             lote=TAMANO_LOTE, hoja_forecast=pp.HOJA_FORECAST,
             generar_futuros=False, carpeta_futuros=None,
             esquemas=pp.ESQUEMAS):
    """Corre el plan de acondicionamiento y escribe el Excel.

    `origen` es lo mismo que acepta el plan de produccion: el dict de
    planning_balanceo.ejecutar(), un DataFrame del forecast balanceado o la
    ruta del .xlsx del balanceo.
    `mes` es la etiqueta MM-AAAA del libro principal; si no se manda se usa el
    mes en curso, y si ese no tiene demanda se toma el primero del horizonte.
    `generar_futuros` escribe ademas un Excel por mes/linea/esquema en
    `carpeta_futuros`, sin tocar el libro principal.
    """
    forecast = pp.obtener_forecast(origen, hoja_forecast)
    etiqueta_origen = ("resultado en memoria del balanceo"
                       if isinstance(origen, (dict, pd.DataFrame))
                       else f"{Path(origen).name} (hoja '{hoja_forecast}')")
    etiqueta_md = f"{Path(ruta_catalogo).name} (hoja '{hoja_md}')"
    log(f"Inicio del acondicionamiento | origen: {etiqueta_origen} | "
        f"relaciones: {etiqueta_md}")

    _, md_por_linea = leer_md(ruta_catalogo, hoja_md)
    pesos = calcular_relaciones(md_por_linea)
    demanda, meses = preparar_demanda(forecast)

    planes = planear_todos(demanda, meses, md_por_linea, pesos, lote, esquemas)
    if not planes:
        raise ValueError("Ningun mes del forecast tiene demanda de "
                         "acondicionamiento.")

    mes = mes or pd.Timestamp.today().strftime(FORMATO_MES)
    if mes not in planes:
        mes_anterior, mes = mes, sorted(planes)[0]
        log(f"El mes {mes_anterior} no tiene plan de acondicionamiento; "
            f"el libro principal se arma con {mes}.")

    planes_resumen = resumen_acondi(planes, esquemas)
    reparto = tabla_reparto(planes)
    secuencias = tabla_secuencias(planes)
    resumen = armar_resumen(mes, etiqueta_origen, etiqueta_md, lote, planes)

    ruta_excel = exportar_excel(ruta_salida, resumen, planes_resumen, reparto,
                                secuencias, planes[mes], mes)

    archivos_futuros = []
    if generar_futuros:
        raiz = carpeta_futuros or Path(ruta_salida).parent
        # El mes del libro principal ya se entrego completo, no se duplica.
        archivos_futuros = gen_all_plans_xlsx(raiz, planes, esquemas=esquemas,
                                              mes_omitido=mes)

    log(f"Fin del acondicionamiento | mes del libro: {mes} | "
        f"meses planeados: {len(planes)}")

    return {
        "mes": mes,
        "meses": list(planes),
        "demanda": demanda,
        "planes": planes,
        "planes_resumen": planes_resumen,
        "reparto": reparto,
        "secuencias": secuencias,
        "pesos": pesos,
        "resumen": resumen,
        "ruta_excel": ruta_excel,
        "archivos_futuros": archivos_futuros,
    }
