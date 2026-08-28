"""""
Module: Produccion por linea 2
Purpose: Separa los productos por familia, y despues por linea que utiliza dentro de cada mes esa familia especifica
Date: 30/07/2026
Author: J.Gonzalez
"""

def run():
    pass
if __name__ == "__main__":
    import pandas as pd
    import numpy as np
    run()

#Merge de datos y eliminacion de SKU
PROD_DC_2['SKUMERCK'] = PROD_DC_2['SKUMERCK'].astype(str).str.strip()
FCST_month['SKUMERCK'] = FCST_month['SKUMERCK'].astype(str).str.strip()
PROD_DC_2 = PROD_DC_2.merge(FCST_month, on='SKUMERCK', how='left')
PROD_DC_2.drop(['SKUMERCK'], axis=1, inplace=True)

#dicccionario para la suma del group
PROD_DC_2.columns = [pd.to_datetime(col, errors='coerce').strftime('%m-%Y') if pd.notna(pd.to_datetime(col, errors='coerce')) else col for col in PROD_DC_2.columns]
agg_dict = {col: 'sum' for col in PROD_DC_2.columns if col not in ['Nombre granel', 'Linea Granel 1', 'Linea Granel 2']}

#Group by de familias de granel
PROD_DC_2 = PROD_DC_2.groupby(['Nombre granel', 'Linea Granel 1', 'Linea Granel 2'], as_index=False).agg(agg_dict)

#Se reduce de dos columnas a una para identificar DC
PROD_DC_2['Linea Granel 1'] = np.where(
    (PROD_DC_2['Linea Granel 1'] == 1) & (PROD_DC_2['Linea Granel 2'] == 1),
    'DC2',
    'DC1')
PROD_DC_2.drop(['Linea Granel 2'], axis=1, inplace=True)
PROD_DC_2 = PROD_DC_2.rename(columns={'Linea Granel 1': 'DC'})