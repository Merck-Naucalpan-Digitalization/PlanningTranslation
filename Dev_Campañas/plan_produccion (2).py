"""
Planeador de producción para líneas DC1 / DC2.

Reglas de negocio modeladas:
- Esquema de producción: 24/7 o 24/5.
- Paros programados (calibración, mantenimiento, etc.) definidos antes de iniciar el mes.
    * 24/5  -> se intentan acomodar en fin de semana.
    * 24/7  -> se descuentan directamente del tiempo total disponible.
- Tiempo de fabricación por lote: DC1 = 27 h, DC2 = 24 h (sin importar el producto).
- Limpieza entre fabricaciones: 8 h.
- Campañas: hasta 3 lotes del MISMO producto, sin limpieza entre lotes de la
  campaña; solo se limpia (8 h) al cerrar la campaña (se trata como una sola
  "corrida" para efectos de limpieza).
- Productos de DC2 se pueden trasladar a DC1 para completar una campaña en
  DC1 (nunca al revés).
"""

import pandas as pd
from calendar import monthrange
from datetime import datetime, timedelta

# ---------------------------------------------------------------------------
# CONFIGURACIÓN GLOBAL (ajustable)
# ---------------------------------------------------------------------------
PROD_TIME = {'DC1': 27, 'DC2': 24}     # horas por lote
CLEAN_TIME = 8                          # horas de limpieza por corrida
MAX_CAMPAIGN = 3                        # lotes máx. por campaña (mismo producto)
# Si True, se limpia también después de la última corrida del mes (deja la
# línea lista para el mes siguiente). Si False, solo se limpia ENTRE corridas.
CLEANING_COUNTS_LAST_RUN = False


# ---------------------------------------------------------------------------
# 1) HORAS DISPONIBLES SEGÚN ESQUEMA Y PAROS PROGRAMADOS
# ---------------------------------------------------------------------------
def horas_disponibles(year, month, esquema='24/7', horas_paro=0):
    """
    Calcula las horas de producción disponibles en el mes para una línea.

    esquema: '24/7' o '24/5'
    horas_paro: horas de paro programado (mantenimiento, calibración, etc.)
    """
    dias_en_mes = monthrange(year, month)[1]
    primer_dia = datetime(year, month, 1)

    dias_habiles = sum(
        1 for d in range(dias_en_mes)
        if (primer_dia + timedelta(days=d)).weekday() < 5
    )
    dias_finde = dias_en_mes - dias_habiles

    if esquema == '24/7':
        horas_totales = dias_en_mes * 24
        disp = horas_totales - horas_paro
        detalle = {
            'esquema': esquema,
            'horas_totales_calendario': horas_totales,
            'horas_paro_programado': horas_paro,
            'horas_disponibles_produccion': disp,
        }

    elif esquema == '24/5':
        horas_habiles = dias_habiles * 24
        horas_finde = dias_finde * 24
        if horas_paro <= horas_finde:
            # el paro cabe completo en el fin de semana, no toca tiempo hábil
            disp = horas_habiles
        else:
            excedente = horas_paro - horas_finde
            disp = horas_habiles - excedente
        detalle = {
            'esquema': esquema,
            'dias_habiles': dias_habiles,
            'horas_habiles_calendario': horas_habiles,
            'horas_fin_de_semana_disponibles': horas_finde,
            'horas_paro_programado': horas_paro,
            'horas_disponibles_produccion': disp,
        }
    else:
        raise ValueError("esquema debe ser '24/7' o '24/5'")

    return disp, detalle


# ---------------------------------------------------------------------------
# 2) OPTIMIZACIÓN DE TRASLADOS DC2 -> DC1 (para completar campañas)
# ---------------------------------------------------------------------------
def optimizar_transferencias(df, col_lotes):
    """
    Para cada producto con lotes en DC1 y DC2, evalúa si conviene trasladar
    lotes de DC2 a DC1 para completar una campaña (hasta 3 lotes).
    Solo se mueve lo necesario para cerrar el múltiplo de 3 (nunca se manda
    más de lo que se necesita, y nunca se traslada de DC1 a DC2).
    """
    piv = df.pivot_table(index='Nombre granel', columns='DC',
                          values=col_lotes, aggfunc='sum', fill_value=0)
    for c in ['DC1', 'DC2']:
        if c not in piv.columns:
            piv[c] = 0

    movimientos = []
    for producto in piv.index:
        n1, n2 = int(piv.loc[producto, 'DC1']), int(piv.loc[producto, 'DC2'])
        if n1 == 0 or n2 == 0:
            continue
        residuo = n1 % MAX_CAMPAIGN
        if residuo == 0:
            continue
        faltante = MAX_CAMPAIGN - residuo
        a_mover = min(faltante, n2)
        if a_mover > 0:
            piv.loc[producto, 'DC1'] += a_mover
            piv.loc[producto, 'DC2'] -= a_mover
            movimientos.append({
                'producto': producto,
                'lotes_movidos_DC2_a_DC1': a_mover,
                'DC1_final': int(piv.loc[producto, 'DC1']),
                'DC2_final': int(piv.loc[producto, 'DC2']),
            })

    df_ajustado = (
        piv.reset_index()
           .melt(id_vars='Nombre granel', value_vars=['DC1', 'DC2'],
                 var_name='DC', value_name=col_lotes)
    )
    df_ajustado = df_ajustado[df_ajustado[col_lotes] > 0].reset_index(drop=True)
    return df_ajustado, pd.DataFrame(movimientos)


# ---------------------------------------------------------------------------
# 3) ARMADO DE CAMPAÑAS / CORRIDAS
# ---------------------------------------------------------------------------
def construir_campanas(df, col_lotes):
    """Agrupa los lotes de cada producto/DC en corridas de máx. 3 lotes."""
    corridas = []
    for _, row in df.iterrows():
        producto, dc, lotes = row['Nombre granel'], row['DC'], int(row[col_lotes])
        restante = lotes
        while restante > 0:
            tam = min(MAX_CAMPAIGN, restante)
            corridas.append({'producto': producto, 'DC': dc, 'lotes_en_corrida': tam})
            restante -= tam
    return pd.DataFrame(corridas)


# ---------------------------------------------------------------------------
# 4) HORAS REQUERIDAS POR DC
# ---------------------------------------------------------------------------
def calcular_horas_por_dc(corridas):
    resumen = []
    for dc, grupo in corridas.groupby('DC'):
        num_corridas = len(grupo)
        lotes_totales = int(grupo['lotes_en_corrida'].sum())
        horas_produccion = lotes_totales * PROD_TIME[dc]
        num_limpiezas = num_corridas if CLEANING_COUNTS_LAST_RUN else max(num_corridas - 1, 0)
        horas_limpieza = num_limpiezas * CLEAN_TIME
        resumen.append({
            'DC': dc,
            'lotes_totales': lotes_totales,
            'num_corridas': num_corridas,
            'horas_produccion': horas_produccion,
            'num_limpiezas': num_limpiezas,
            'horas_limpieza': horas_limpieza,
            'horas_totales_requeridas': horas_produccion + horas_limpieza,
        })
    return pd.DataFrame(resumen)


# ---------------------------------------------------------------------------
# 5) ORQUESTADOR PRINCIPAL
# ---------------------------------------------------------------------------
def plan_produccion(df, col_lotes, year, month, esquema='24/7', horas_paro=0):
    """
    horas_paro puede ser:
      - un número -> se aplica igual a DC1 y DC2
      - un dict {'DC1': x, 'DC2': y} -> paro específico por línea
    """
    if not isinstance(horas_paro, dict):
        horas_paro = {'DC1': horas_paro, 'DC2': horas_paro}

    df_opt, movimientos = optimizar_transferencias(df, col_lotes)
    corridas = construir_campanas(df_opt, col_lotes)
    resumen_horas = calcular_horas_por_dc(corridas)

    filas = []
    for _, row in resumen_horas.iterrows():
        dc = row['DC']
        disp, detalle = horas_disponibles(year, month, esquema, horas_paro.get(dc, 0))
        fila = row.to_dict()
        fila['horas_disponibles'] = disp
        fila['holgura_horas'] = disp - row['horas_totales_requeridas']
        fila['factible'] = fila['holgura_horas'] >= 0
        filas.append(fila)

    resultado = pd.DataFrame(filas)
    return resultado, corridas, movimientos


# ---------------------------------------------------------------------------
# 6) CRONOGRAMA (fecha/hora de inicio y fin de cada lote y limpieza)
# ---------------------------------------------------------------------------
def _avanzar_tiempo(inicio, horas, esquema):
    """
    Suma `horas` a `inicio`. En 24/7 es una suma directa; en 24/5 el reloj
    de producción se detiene el fin de semana (sáb 00:00 - lun 00:00) y
    continúa el lunes donde se quedó.
    """
    if esquema == '24/7':
        return inicio + timedelta(hours=horas)

    actual = inicio
    restante = horas
    while restante > 1e-9:
        if actual.weekday() >= 5:  # sábado(5) o domingo(6) -> saltar a lunes 00:00
            dias_a_saltar = 7 - actual.weekday()
            actual = (actual.replace(hour=0, minute=0, second=0, microsecond=0)
                      + timedelta(days=dias_a_saltar))
            continue
        dias_hasta_finde = 5 - actual.weekday()  # lun(0)->5 ... vie(4)->1
        limite_semana = (actual.replace(hour=0, minute=0, second=0, microsecond=0)
                          + timedelta(days=dias_hasta_finde))
        horas_bloque = (limite_semana - actual).total_seconds() / 3600
        if restante <= horas_bloque:
            actual = actual + timedelta(hours=restante)
            restante = 0
        else:
            actual = limite_semana
            restante -= horas_bloque
    return actual


def generar_cronograma(corridas, dc, fecha_inicio, esquema='24/7'):
    """
    Expande las corridas de una línea (DC1 o DC2) en un cronograma detallado:
    una fila por lote individual + una fila 'Limpieza' al cerrar cada corrida
    (campaña o lote suelto). Respeta el orden en que aparecen en `corridas`.
    """
    filas = []
    tiempo_actual = fecha_inicio
    grupo = corridas[corridas['DC'] == dc]

    for _, row in grupo.iterrows():
        producto, lotes = row['producto'], int(row['lotes_en_corrida'])
        for _ in range(lotes):
            fin = _avanzar_tiempo(tiempo_actual, PROD_TIME[dc], esquema)
            filas.append({'Nombre': producto,
                           'Fecha y hora de inicio': tiempo_actual,
                           'Fecha y hora de finalización': fin})
            tiempo_actual = fin
        fin_limpieza = _avanzar_tiempo(tiempo_actual, CLEAN_TIME, esquema)
        filas.append({'Nombre': 'Limpieza',
                       'Fecha y hora de inicio': tiempo_actual,
                       'Fecha y hora de finalización': fin_limpieza})
        tiempo_actual = fin_limpieza

    return pd.DataFrame(filas)


def exportar_cronograma_csv(corridas, fecha_inicio_dc1, fecha_inicio_dc2,
                             esquema='24/7', ruta='plan_produccion.csv'):
    """
    Genera UN solo CSV con dos bloques (DC1 y DC2), cada uno con su propio
    encabezado, en el formato:
        DC1
        Nombre,Fecha y hora de inicio,Fecha y hora de finalización
        ...filas...
        (línea en blanco)
        DC2
        Nombre,Fecha y hora de inicio,Fecha y hora de finalización
        ...filas...
    """
    fmt = '%d/%m/%Y %H:%M'
    cro_dc1 = generar_cronograma(corridas, 'DC1', fecha_inicio_dc1, esquema)
    cro_dc2 = generar_cronograma(corridas, 'DC2', fecha_inicio_dc2, esquema)

    for cro in (cro_dc1, cro_dc2):
        cro['Fecha y hora de inicio'] = cro['Fecha y hora de inicio'].dt.strftime(fmt)
        cro['Fecha y hora de finalización'] = cro['Fecha y hora de finalización'].dt.strftime(fmt)

    with open(ruta, 'w', encoding='utf-8-sig', newline='') as f:
        f.write('DC1\n')
        cro_dc1.to_csv(f, index=False, lineterminator='\n')
        f.write('\n')
        f.write('DC2\n')
        cro_dc2.to_csv(f, index=False, lineterminator='\n')

    return ruta


# ---------------------------------------------------------------------------
# 7) CAPTURA DE PARÁMETROS POR TERMINAL
# ---------------------------------------------------------------------------
def solicitar_parametros_usuario():
    """
    Pide por terminal:
      - Tipo de turno (24/7 o 24/5)
      - Horas de paro programado para DC1 y DC2
    Valida la entrada y regresa (esquema, horas_paro_dict).
    """
    print("=== Parámetros del plan de producción ===")

    esquema = None
    while esquema not in ('24/7', '24/5'):
        esquema = input("Tipo de turno [24/7 o 24/5]: ").strip()
        if esquema not in ('24/7', '24/5'):
            print("  -> Entrada no válida, escribe exactamente '24/7' o '24/5'.")

    horas_paro = {}
    for dc in ('DC1', 'DC2'):
        while True:
            valor = input(f"Horas de paro programado en {dc} (Enter = 0): ").strip()
            if valor == '':
                horas_paro[dc] = 0
                break
            try:
                horas_paro[dc] = float(valor)
                break
            except ValueError:
                print("  -> Ingresa un número (ej. 24, 48, 0).")

    return esquema, horas_paro


# ---------------------------------------------------------------------------
# EJEMPLO DE USO CON LOS DATOS PROPORCIONADOS
# ---------------------------------------------------------------------------
if __name__ == '__main__':
    data = [
        ('DCS NEUROBION 10,000', 'DC1', 2),
        ('DCS NEUROBION 10,000', 'DC2', 1),
        ('DCS NEUROBION 10,000 DOUBLE FILTRATION', 'DC1', 0),
        ('DCS NEUROBION 25,000', 'DC1', 1),
        ('DEXABION DC BULK - MEX', 'DC1', 2),
        ('DEXABION DC BULK - MEX', 'DC2', 0),
        ('DOLO NEUROBION DC BULK - MEX', 'DC1', 2),
        ('DOLO NEUROBION DC BULK - MEX', 'DC2', 4),
        ('DOLO NEUROBION FORTE DC - MEX', 'DC1', 1),
        ('DOLO NEUROBION FORTE DC - MEX', 'DC2', 0),
        ('DOLO NEUROBION FORTE DC BULK - MEX', 'DC1', 0),
    ]
    df = pd.DataFrame(data, columns=['Nombre granel', 'DC', 'ago-26'])
    df = df[df['ago-26'] > 0].reset_index(drop=True)  # quitamos ceros

    esquema, horas_paro = solicitar_parametros_usuario()

    resultado, corridas, movimientos = plan_produccion(
        df, col_lotes='ago-26', year=2026, month=8,
        esquema=esquema, horas_paro=horas_paro
    )

    print("=== Traslados DC2 -> DC1 sugeridos ===")
    print(movimientos.to_string(index=False) if not movimientos.empty else "(ninguno)")

    print("\n=== Corridas / campañas resultantes ===")
    print(corridas.to_string(index=False))

    print("\n=== Resumen de factibilidad por DC ===")
    print(resultado.to_string(index=False))

    ruta = exportar_cronograma_csv(
        corridas,
        fecha_inicio_dc1=datetime(2026, 8, 1, 7, 0),
        fecha_inicio_dc2=datetime(2026, 8, 1, 7, 0),
        esquema=esquema,
        ruta='./plan_produccion.csv',
    )
    print(f"\n=== CSV generado: {ruta} ===")
