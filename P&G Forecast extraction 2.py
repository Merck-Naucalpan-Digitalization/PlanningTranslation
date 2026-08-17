"""""
Module: P&G Forecast extraction 2
Purpose: Extraer ## de jeringas para todos los meses por producto
Date: 30/07/2026
Author: J.Gonzalez
"""


def run():
    pass
if __name__ == "__main__":
    import pandas as pd
    import numpy as np
    run()

id_cols = ['SKUMERCK']

#Calculo de mes y año actual
Month_today=pd.to_datetime("today").month-1
Year_today=pd.to_datetime("today").year
demand_today=pd.to_datetime(str(Year_today) + "-01-" + str(Month_today), format="%Y-%m-%d")
demand_today=str(demand_today)[0:10]

#seleccion de meses con demanda futura en el FCST
matching_cols = []
for i in range(len(FCST.columns) - 1, -1, -1):  # from [-1] backwards
    col = FCST.columns[i]
    parsed = pd.to_datetime(col, errors='coerce')
    if pd.notna(parsed):
        matching_cols.append(col)
        if parsed.month <= Month_today and parsed.year <= Year_today:
            break  #
matching_cols = list(reversed(matching_cols))

#Union de columnas id con columna demanda de cada mes
FCST_month = FCST[id_cols + matching_cols].copy()
FCST_month = FCST_month.dropna(subset=['SKUMERCK']).reset_index(drop=True)