"""""
Module: Balanceo mensual
Purpose: Balancea DC2 mandando su residuo a DC1 para cada familia en cada mes 
Date: 30/07/2026
Author: J.Gonzalez
"""

def run():
    pass

if __name__ == "__main__":
    import pandas as pd
    import numpy as np
    run()
    
# Numero de jeringas por lote y redondeo de residuo
jeringas=115200
decimales_req=2

# Identifica todas las familias disponibles y registra cuales tienen capacidad en DC2
familias= PROD_DC_2['Nombre granel'].unique()
familias_con_dc2=PROD_DC_2.loc[PROD_DC_2['DC'] == 'DC2', 'Nombre granel'].unique()

#Base de datos para residuos en jeringas
registro_residuos = PROD_DC_2.copy()

# Balanceo, en casos que detecte que la familia con DC2 tiene residuo mando todo ese residuo a DC1 de esa misma familia                      
for familia in familias_con_dc2:
    for col in agg_dict.keys():
        # Filtra las filas para esta familia específica a DC2 y DC1
        mask_dc2 = (PROD_DC_2['Nombre granel'] == familia) & (PROD_DC_2['DC'] == 'DC2')
        mask_dc1 = (PROD_DC_2['Nombre granel'] == familia) & (PROD_DC_2['DC'] == 'DC1')
        # Valor de DC2 para la familia y mes especificado
        dc2_val = PROD_DC_2.loc[mask_dc2, col].values[0]
        
        # Balanceo de residuos DC2
        if dc2_val % jeringas != 0:
            residue = dc2_val % jeringas
            # suma el residuo a DC1
            PROD_DC_2.loc[mask_dc1, col] = (PROD_DC_2.loc[mask_dc1, col] + residue)
            # resta el residuo de DC2
            PROD_DC_2.loc[mask_dc2, col] =(PROD_DC_2.loc[mask_dc2, col] - residue)

for familia in familias:
    for col in agg_dict.keys():
        mask = (PROD_DC_2['Nombre granel'] == familia)
        #calcula el residuo de la familia en dc1
        registro_residuos.loc[mask, col]=jeringas-PROD_DC_2.loc[mask, col]%jeringas
        registro_residuos.loc[mask, col] = np.where(
            (registro_residuos.loc[mask, col] <= jeringas * 0.01) |
            (registro_residuos.loc[mask, col] >= jeringas * (1 - 0.01)),
            0,
            registro_residuos.loc[mask, col])
        #cambia toda la demanda de la tabla de jeringas a lotes
        PROD_DC_2.loc[mask, col] = (PROD_DC_2.loc[mask, col]/jeringas)
        
# Exportación de resultados en archivos CSV
DC1=PROD_DC_2[PROD_DC_2['DC'] == 'DC1'].copy()
DC2=PROD_DC_2[PROD_DC_2['DC'] == 'DC2'].copy()
DC1.to_csv("DC1_FCST.csv", index=False)
DC2.to_csv("DC2_FCST.csv", index=False)
PROD_DC_2.to_csv("PROD_DC_2_balanceado.csv", index=False)
pre_balanceo.to_csv("pre_balanceo.csv", index=False)
registro_residuos.to_csv("registro_residuos.csv", index=False)