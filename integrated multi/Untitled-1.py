# el propodito de este nuevo archivo es de empezar a consolidar las formulas compartidas entre los diferentes scripts de planificación
# por ejemplo planning balance y planning_produccion tienen un archivo para leer sus archivos previos pero podrian ser el mismo dentro de 
# un mejor archivo mas grande. 


####            log                      ####
#version de planning_balanceo
def log(mensaje):
    """Unico punto de salida de mensajes. La app de escritorio reemplaza esta
    funcion por la suya para mostrarlos en la bitacora de la ventana."""
    print(mensaje)
#version de planning_produccion
def log(mensaje):
    """Unico punto de salida de mensajes. La app de escritorio la reemplaza."""
    print(mensaje)


#version de planning_balanceo
def leer_entradas(ruta_fcst, hoja_fcst, ruta_catalogo, hoja_catalogo):
    """Lee los dos Excel y revisa que traigan lo minimo necesario."""
    fcst = pd.read_excel(ruta_fcst, sheet_name=hoja_fcst)
    fcst[id_cols[0]]=fcst[id_cols[0]].astype(str).str.strip()
    catalogo = pd.read_excel(ruta_catalogo, sheet_name=hoja_catalogo)
    catalogo[id_cols[0]]=catalogo[id_cols[0]].astype(str).str.strip()
    log(f"Forecast: {len(fcst)} filas | Catalogo: {len(catalogo)} filas")
    return fcst, catalogo

#version de planning_produccion
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
    """Lee la hoja 'Forecast balanceado' del libro que genero el balanceo."""
    tabla = pd.read_excel(Path(origen), sheet_name=hoja)
    log(f"Forecast balanceado leido de {origen.name} (hoja '{hoja}'): "
        f"{len(tabla)} renglones")
    return tabla
#version de planning_balanceo

#version de planning_produccion


#version de planning_balanceo

#version de planning_produccion


#version de planning_balanceo

#version de planning_produccion

'''
lista de funciones y variables que si son invocadas en la interfaz directamente. 
planning produccion
    pp.log
    pp.HOJA_FORECAST
    pp.mes_actual()
        sospecho que mes actual tiene una contraprte de planning balanceo 
    pp.meses_disponibles()
        aparece 2 veces
    pp.ejecutar()
    pp.ESQUEMAS
planning_balanceo
    pb.log
    pb.TAMANO_LOTE
    pb.listar_meses()
    pb.ejecutar()

podremos eliminar la funcino de log, si solo mandamos el print. 
'''