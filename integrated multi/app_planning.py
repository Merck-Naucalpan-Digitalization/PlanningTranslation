"""
app_planning.py
Ventana de escritorio para el proceso de Planning en tres pasos:

    Pestana 1 - Balanceo          -> planning_balanceo.py
    Pestana 2 - Plan de produccion -> planning_produccion.py
    Pestana 3 - Acondicionamiento  -> planning_acondi.py

Aqui no hay reglas de negocio: todo el calculo vive en los modulos. Esta capa
solo pide los archivos, arma los parametros, lanza el proceso en un hilo
aparte para que la ventana no se congele y muestra la bitacora.

Los pasos 2 y 3 toman el "Forecast balanceado" del paso 1. Si el balanceo ya
corrio en esta sesion usan su resultado en memoria; si no, se puede elegir el
.xlsx de balanceo que ya exista en el disco.

Las horas de paro programado ya no se escriben a mano: se leen de una hoja del
Excel de catalogo, que se elige igual que el forecast (ruta + hoja).

Los dos planes pueden ademas exportar TODO el horizonte a futuro, un Excel por
mes / linea / esquema dentro de su propio arbol de carpetas. Eso se activa con
un checkbox y es independiente del libro principal, que siempre se genera.

Requisitos: pip install customtkinter pandas openpyxl
"""

import calendar
import os
import queue
import subprocess
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk
import pandas as pd

# Importacion de funciones y variables de la logica.
import planning_acondi as pa
import planning_balanceo as pb
import planning_produccion as pp

# Colores de la interfaz
ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

TIPOS_EXCEL = [("Archivos de Excel", "*.xlsx *.xlsm *.xls")]
SIN_ARCHIVO = "Ningun archivo seleccionado"
SIN_HOJA = "Selecciona una hoja"
SIN_SALIDA = "Ningun archivo definido"
SIN_CARPETA = "Ninguna carpeta definida"
TODOS_LOS_MESES = "Horizonte completo"
SIN_MES = "Selecciona un mes"
EN_MEMORIA = "Resultado del balanceo de esta sesion"

MESES_ES = ["01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril",
            "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto",
            "09 - Septiembre", "10 - Octubre", "11 - Noviembre",
            "12 - Diciembre"]


class SelectorExcel(ctk.CTkFrame):
    """Bloque para elegir un archivo de Excel y una de sus hojas."""

    def __init__(self, master, titulo, al_cambiar, hoja_sugerida=None,
                 ayuda=None):
        super().__init__(master)
        self.al_cambiar = al_cambiar
        self.hoja_sugerida = hoja_sugerida
        self.ruta = None
        self.hoja = None

        self.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(self, text=titulo, font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(10, 4))
        # Joe: configuracion del boton para seleccionar archivo
        self.btn = ctk.CTkButton(self, text="Examinar...", width=110,
                                 command=self.elegir_archivo)
        self.btn.grid(row=1, column=0, padx=(12, 8), pady=(0, 12))
        self.lbl = ctk.CTkLabel(self, text=SIN_ARCHIVO, anchor="w",
                                text_color=("gray40", "gray60"))
        self.lbl.grid(row=1, column=1, sticky="ew", padx=(0, 8), pady=(0, 12))

        self.combo = ctk.CTkComboBox(self, values=[], width=190, state="disabled",
                                     command=self.elegir_hoja)
        self.combo.set(SIN_HOJA)
        self.combo.grid(row=1, column=2, padx=(0, 12), pady=(0, 12))

        if ayuda:
            ctk.CTkLabel(self, text=ayuda, justify="left",
                         font=ctk.CTkFont(size=11),
                         text_color=("gray40", "gray60")).grid(
                row=2, column=0, columnspan=3, sticky="w", padx=12, pady=(0, 10))

    def elegir_archivo(self):
        # Joe: esta es la funcion que selecciona el archivo en una segunda ventana
        ruta = filedialog.askopenfilename(filetypes=TIPOS_EXCEL)
        if not ruta:
            return
        try:
            with pd.ExcelFile(ruta) as libro:
                # Joe: la variable hojas guarda todas las hojas posibles para analizar.
                hojas = list(libro.sheet_names)
        except Exception as error:
            messagebox.showerror("No se pudo leer el archivo",
                                 f"{Path(ruta).name}\n\n{error}")
            return

        # Joe: al cambiar de archivo, la hoja anterior deja de valer.
        self.ruta = Path(ruta)
        self.hoja = None
        self.lbl.configure(text=self.ruta.name, text_color=("gray10", "gray90"))
        self.combo.configure(values=hojas, state="readonly")

        if self.hoja_sugerida in hojas:
            self.hoja = self.hoja_sugerida
        elif len(hojas) == 1:
            self.hoja = hojas[0]
        self.combo.set(self.hoja or SIN_HOJA)

        self.al_cambiar()

    def elegir_hoja(self, hoja):
        self.hoja = hoja
        self.al_cambiar()

    def fijar_archivo(self, ruta, hoja_sugerida=None):
        """Precarga una ruta ya elegida en otra pestana, sin volver a pedirla.

        Sirve para que el catalogo del balanceo se reutilice en los pasos 2 y 3
        sin que el usuario lo busque dos veces. La hoja se deja a su eleccion.
        """
        if ruta is None or self.ruta is not None:
            return
        try:
            with pd.ExcelFile(ruta) as libro:
                hojas = list(libro.sheet_names)
        except Exception:
            return
        self.ruta = Path(ruta)
        self.lbl.configure(text=self.ruta.name, text_color=("gray10", "gray90"))
        self.combo.configure(values=hojas, state="readonly")
        sugerida = hoja_sugerida or self.hoja_sugerida
        self.hoja = sugerida if sugerida in hojas else None
        self.combo.set(self.hoja or SIN_HOJA)
        self.al_cambiar()

    def esta_listo(self):
        return self.ruta is not None and self.hoja is not None

    def habilitar(self, activo):
        self.btn.configure(state="normal" if activo else "disabled")
        self.combo.configure(state="readonly" if activo and self.ruta else "disabled")


class SelectorSalida(ctk.CTkFrame):
    """Bloque para definir la ruta del archivo de salida."""

    def __init__(self, master, titulo, al_cambiar, nombre_sugerido):
        super().__init__(master)
        self.al_cambiar = al_cambiar
        self.nombre_sugerido = nombre_sugerido
        self.ruta = None

        self.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(self, text=titulo, font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=12, pady=(10, 4))

        self.btn = ctk.CTkButton(self, text="Guardar como...", width=110,
                                 command=self.elegir)
        self.btn.grid(row=1, column=0, padx=(12, 8), pady=(0, 12))
        self.lbl = ctk.CTkLabel(self, text=SIN_SALIDA, anchor="w",
                                text_color=("gray40", "gray60"))
        self.lbl.grid(row=1, column=1, sticky="ew", padx=(0, 12), pady=(0, 12))

    def elegir(self):
        ruta = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=self.nombre_sugerido(),
            filetypes=[("Libro de Excel", "*.xlsx")])
        if not ruta:
            return
        self.ruta = Path(ruta)
        self.lbl.configure(text=str(self.ruta), text_color=("gray10", "gray90"))
        self.al_cambiar()

    def esta_listo(self):
        return self.ruta is not None

    def habilitar(self, activo):
        self.btn.configure(state="normal" if activo else "disabled")


class SelectorFecha(ctk.CTkFrame):
    """Fecha y hora por listas, sin escribir texto.

    Se arma con cuatro combos (dia, mes, anio, hora) para que no haya forma de
    teclear una fecha invalida. Los dias se recalculan cuando cambia el mes o
    el anio, asi que nunca ofrece un 31 de febrero.
    """

    def __init__(self, master, titulo, al_cambiar=None,
                 hora_defecto=pp.HORA_INICIO_DEFECTO):
        super().__init__(master, fg_color="transparent")
        self.al_cambiar = al_cambiar or (lambda: None)
        hoy = datetime.today()

        ctk.CTkLabel(self, text=titulo).grid(row=0, column=0, columnspan=4,
                                             sticky="w", padx=(0, 6))

        anios = [str(a) for a in range(hoy.year - 1, hoy.year + 4)]
        self.cb_anio = ctk.CTkComboBox(self, values=anios, width=85,
                                       command=self._al_cambiar_mes)
        self.cb_anio.set(str(hoy.year))
        self.cb_mes = ctk.CTkComboBox(self, values=MESES_ES, width=135,
                                      command=self._al_cambiar_mes)
        self.cb_mes.set(MESES_ES[hoy.month - 1])
        self.cb_dia = ctk.CTkComboBox(self, values=[], width=65,
                                      command=lambda _=None: self.al_cambiar())
        self.cb_hora = ctk.CTkComboBox(
            self, values=[f"{h:02d}:00" for h in range(24)], width=85,
            command=lambda _=None: self.al_cambiar())
        self.cb_hora.set(f"{hora_defecto:02d}:00")

        for col, widget in enumerate((self.cb_dia, self.cb_mes, self.cb_anio,
                                      self.cb_hora)):
            widget.grid(row=1, column=col, padx=(0, 6), pady=(2, 0), sticky="w")

        self._refrescar_dias(preferido=1)

    def _al_cambiar_mes(self, _=None):
        self._refrescar_dias(preferido=int(self.cb_dia.get() or 1))
        self.al_cambiar()

    def _refrescar_dias(self, preferido=1):
        """Ajusta la lista de dias al mes y anio elegidos."""
        anio = int(self.cb_anio.get())
        mes = MESES_ES.index(self.cb_mes.get()) + 1
        ultimo = calendar.monthrange(anio, mes)[1]
        dias = [f"{d:02d}" for d in range(1, ultimo + 1)]
        self.cb_dia.configure(values=dias)
        self.cb_dia.set(f"{min(preferido, ultimo):02d}")

    def fijar_mes(self, anio, mes):
        """Mueve el selector al mes que se va a planear, dia 1."""
        self.cb_anio.set(str(anio))
        self.cb_mes.set(MESES_ES[mes - 1])
        self._refrescar_dias(preferido=1)

    def valor(self):
        anio = int(self.cb_anio.get())
        mes = MESES_ES.index(self.cb_mes.get()) + 1
        dia = int(self.cb_dia.get())
        hora = int(self.cb_hora.get().split(":")[0])
        return datetime(anio, mes, dia, hora, 0)

    def habilitar(self, activo):
        estado = "readonly" if activo else "disabled"
        for widget in (self.cb_dia, self.cb_mes, self.cb_anio, self.cb_hora):
            widget.configure(state=estado)


class SelectorCarpeta(ctk.CTkFrame):
    """Bloque para elegir la carpeta raiz de los planes a futuro.

    Produccion y acondicionamiento comparten carpeta: los dos cronogramas del
    mismo mes y esquema tienen que caer juntos. Por eso el widget avisa al
    elegir y la ventana replica la ruta en el otro selector.
    """

    def __init__(self, master, titulo, al_cambiar):
        super().__init__(master, fg_color="transparent")
        self.al_cambiar = al_cambiar
        self.ruta = None

        self.grid_columnconfigure(1, weight=1)
        self.btn = ctk.CTkButton(self, text=titulo, width=170,
                                 command=self.elegir)
        self.btn.grid(row=0, column=0, padx=(0, 8), sticky="w")
        self.lbl = ctk.CTkLabel(self, text=SIN_CARPETA, anchor="w",
                                font=ctk.CTkFont(size=11),
                                text_color=("gray40", "gray60"))
        self.lbl.grid(row=0, column=1, sticky="ew")

    def elegir(self):
        ruta = filedialog.askdirectory()
        if not ruta:
            return
        self.fijar(ruta)
        self.al_cambiar()

    def fijar(self, ruta):
        """Pone la ruta sin disparar el aviso, para replicarla entre pestanas."""
        self.ruta = Path(ruta)
        self.lbl.configure(text=str(self.ruta), text_color=("gray10", "gray90"))

    def habilitar(self, activo):
        self.btn.configure(state="normal" if activo else "disabled")


class AppPlanning(ctk.CTk):
    """Ventana principal con las tres etapas del proceso."""

    def __init__(self):
        super().__init__()
        self.title("Planning DC - Balanceo, produccion y acondicionamiento")
        self.geometry("980x900")
        self.minsize(900, 760)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)

        # Estado compartido entre las etapas
        self.meses = []
        self.resultado_balanceo = None     # dict que devuelve pb.ejecutar
        self.ruta_balanceo = None          # .xlsx generado o elegido a mano
        self.ruta_plan = None
        self.ruta_acondi = None
        self.ejecutando = False
        self.mensajes = queue.Queue()

        self.tabs = ctk.CTkTabview(self)
        self.tabs.grid(row=0, column=0, sticky="nsew", padx=16, pady=(16, 8))
        self.tab_bal = self.tabs.add("1. Balanceo")
        self.tab_prod = self.tabs.add("2. Plan de produccion")
        self.tab_acondi = self.tabs.add("3. Acondicionamiento")
        for tab in (self.tab_bal, self.tab_prod, self.tab_acondi):
            tab.grid_columnconfigure(0, weight=1)

        self._construir_balanceo()
        self._construir_produccion()
        self._construir_acondicionamiento()
        self._construir_bitacora()

        # Joe: los modulos de calculo mandan sus mensajes a la cola en vez de la consola.
        pb.log = self.mensajes.put
        pp.log = self.mensajes.put
        pa.log = self.mensajes.put
        self._vaciar_cola()
        self._revisar_estados()

    # -- Pestana 1: balanceo ---------------------------------------------------

    def _construir_balanceo(self):
        marco = self.tab_bal

        self.sel_fcst = SelectorExcel(marco, "1. Forecast P&G (demanda por SKU)",
                                      self._al_cambiar_forecast)
        self.sel_fcst.grid(row=0, column=0, sticky="ew", pady=(8, 8))

        self.sel_catalogo = SelectorExcel(marco, "2. Catalogo de productos DC",
                                          self._al_cambiar_catalogo,
                                          hoja_sugerida="Catalogo")
        self.sel_catalogo.grid(row=1, column=0, sticky="ew", pady=8)

        params = ctk.CTkFrame(marco)
        params.grid(row=2, column=0, sticky="ew", pady=8)
        params.grid_columnconfigure(5, weight=1)
        ctk.CTkLabel(params, text="3. Mes del ejercicio y parametros",
                     font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=0, column=0, columnspan=6, sticky="w", padx=12, pady=(10, 4))
        # Joe: seccion que selecciona el mes a analizar; si se deja, se queda
        # Horizonte completo. La variable TODOS_LOS_MESES dicta que se esta haciendo.
        ctk.CTkLabel(params, text="Mes a analizar").grid(
            row=1, column=0, padx=(12, 6), pady=(0, 10), sticky="e")
        self.combo_mes = ctk.CTkComboBox(params, values=[TODOS_LOS_MESES], width=180,
                                         state="disabled")
        self.combo_mes.set(TODOS_LOS_MESES)
        self.combo_mes.grid(row=1, column=1, columnspan=2, pady=(0, 10), sticky="w")

        ctk.CTkLabel(params, text="Deja 'Horizonte completo' para que los planes de\n"
                                  "produccion y acondicionamiento puedan elegir "
                                  "cualquier mes despues.",
                     justify="left", font=ctk.CTkFont(size=11),
                     text_color=("gray40", "gray60")).grid(
            row=1, column=3, columnspan=3, padx=(4, 12), pady=(0, 10), sticky="w")
        # Joe: aqui es donde importa el tamano de lote al codigo
        self.ent_lote = self._campo(params, "Jeringas por lote", 0, pb.TAMANO_LOTE, 110)

        self.sel_salida_bal = SelectorSalida(
            marco, "4. Archivo de salida del balanceo (.xlsx)", self._revisar_estados,
            lambda: f"Balanceo_DC_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx")
        self.sel_salida_bal.grid(row=3, column=0, sticky="ew", pady=8)

        acciones = ctk.CTkFrame(marco, fg_color="transparent")
        acciones.grid(row=4, column=0, sticky="ew", pady=(4, 8))
        acciones.grid_columnconfigure(1, weight=1)
        # Joe: establece la funcion del boton de ejecutar
        self.btn_balanceo = ctk.CTkButton(acciones, text="Ejecutar balanceo", height=40,
                                          font=ctk.CTkFont(size=14, weight="bold"),
                                          state="disabled", command=self.ejecutar_balanceo)
        self.btn_balanceo.grid(row=0, column=0, padx=(0, 10), sticky="w")
        self.barra_bal = ctk.CTkProgressBar(acciones, mode="indeterminate")
        self.barra_bal.grid(row=0, column=1, sticky="ew", padx=(0, 10))
        self.barra_bal.set(0)
        self.btn_abrir_bal = ctk.CTkButton(
            acciones, text="Abrir Excel", width=120, state="disabled",
            command=lambda: self._abrir(self.ruta_balanceo))
        self.btn_abrir_bal.grid(row=0, column=2, sticky="e")

    def _campo(self, marco, texto, columna, valor, ancho, fila=2):
        """Etiqueta + caja de texto en una fila de parametros."""
        ctk.CTkLabel(marco, text=texto).grid(row=fila, column=columna, padx=(12, 6),
                                             pady=(0, 12), sticky="e")
        entrada = ctk.CTkEntry(marco, width=ancho)
        entrada.insert(0, str(valor))
        entrada.grid(row=fila, column=columna + 1, pady=(0, 12), sticky="w")
        return entrada

    # -- Pestana 2: plan de produccion ----------------------------------------

    def _construir_produccion(self):
        marco = self.tab_prod

        origen = ctk.CTkFrame(marco)
        origen.grid(row=0, column=0, sticky="ew", pady=(8, 8))
        origen.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(origen, text="1. Forecast balanceado (salida del balanceo)",
                     font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(10, 4))
        self.btn_lotes = ctk.CTkButton(origen, text="Elegir Excel...", width=120,
                                       command=self.elegir_balanceo_existente)
        self.btn_lotes.grid(row=1, column=0, padx=(12, 8), pady=(0, 4))
        self.lbl_lotes = ctk.CTkLabel(origen, text="Corre el balanceo o elige un libro "
                                                   "de balanceo ya generado.",
                                      anchor="w", text_color=("gray40", "gray60"))
        self.lbl_lotes.grid(row=1, column=1, columnspan=2, sticky="ew",
                            padx=(0, 12), pady=(0, 4))
        
        '''
        ctk.CTkLabel(origen, text=f"Se lee la hoja '{pp.HOJA_FORECAST}'. Las cantidades "
                                  "no se modifican: solo cambia de linea un producto\n"
                                  "cuando tiene equivalente en DC1 y DC2.",
                     justify="left", font=ctk.CTkFont(size=11),
                     text_color=("gray40", "gray60")).grid(
            row=2, column=0, columnspan=3, sticky="w", padx=12, pady=(0, 12))
        '''

        # --- Paros programados: ya no se teclean, se leen del catalogo ---
        self.sel_paros = SelectorExcel(
            marco, "2. Paros programados (hoja del Excel de catalogo)",
            self._revisar_estados, hoja_sugerida="Paros Programados")
        self.sel_paros.grid(row=1, column=0, sticky="ew", pady=6)

        params = ctk.CTkFrame(marco)
        params.grid(row=2, column=0, sticky="ew", pady=6)
        params.grid_columnconfigure(4, weight=1)
        ctk.CTkLabel(params, text="3. Mes a planear y arranque de cada linea",
                     font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=0, column=0, columnspan=5, sticky="w", padx=12, pady=(10, 4))

        ctk.CTkLabel(params, text="Mes a planear").grid(
            row=1, column=0, padx=(12, 6), pady=(0, 10), sticky="e")
        self.combo_mes_plan = ctk.CTkComboBox(params, values=[SIN_MES], width=180,
                                              state="disabled",
                                              command=self._al_cambiar_mes_plan)
        self.combo_mes_plan.set(SIN_MES)
        self.combo_mes_plan.grid(row=1, column=1, pady=(0, 10), sticky="w")

        ctk.CTkLabel(params, text="Se planea un solo mes en el libro principal.\n"
                                  "Las horas de paro salen de la hoja de paros.",
                     justify="left", font=ctk.CTkFont(size=11),
                     text_color=("gray40", "gray60")).grid(
            row=1, column=2, columnspan=3, padx=(12, 12), pady=(0, 10), sticky="w")

        self.fecha_dc1 = SelectorFecha(params, "Inicio DC1")
        self.fecha_dc1.grid(row=2, column=0, columnspan=2, padx=(12, 12),
                            pady=(0, 10), sticky="w")
        self.fecha_dc2 = SelectorFecha(params, "Inicio DC2")
        self.fecha_dc2.grid(row=2, column=2, columnspan=3, padx=(12, 12),
                            pady=(0, 10), sticky="w")

        # --- Planes a futuro: gen_all_plans_xlsx ---
        futuros = ctk.CTkFrame(marco)
        futuros.grid(row=3, column=0, sticky="ew", pady=6)
        futuros.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(futuros, text="4. Planes a futuro (opcional)",
                     font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=12, pady=(10, 4))
        self.var_futuros_prod = ctk.BooleanVar(value=False)
        self.chk_futuros_prod = ctk.CTkCheckBox(
            futuros, text="Generar todos los planes del horizonte en su propia "
                          "carpeta por mes", variable=self.var_futuros_prod,
            command=self._revisar_estados)
        self.chk_futuros_prod.grid(row=1, column=0, sticky="w", padx=12, pady=(0, 6))
        self.carpeta_futuros_prod = SelectorCarpeta(
            futuros, "Carpeta de los planes...", self._al_cambiar_carpeta_prod)
        self.carpeta_futuros_prod.grid(row=2, column=0, sticky="ew",
                                       padx=12, pady=(0, 6))
        '''
        ctk.CTkLabel(futuros, text="Aqui solo se guardan los cronogramas, uno por "
                                   "archivo, en <carpeta>/<mes>/<esquema>.\nEl libro "
                                   "completo del mes elegido se exporta siempre, "
                                   "marques o no esta casilla.\nEs la misma carpeta "
                                   "que usa el acondicionamiento.",
                     justify="left", font=ctk.CTkFont(size=11),
                     text_color=("gray40", "gray60")).grid(
            row=3, column=0, sticky="w", padx=12, pady=(0, 10))
        '''
        self.sel_salida_plan = SelectorSalida(
            marco, "5. Archivo de salida del plan (.xlsx)", self._revisar_estados,
            lambda: f"Plan_Produccion_DC_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx")
        self.sel_salida_plan.grid(row=4, column=0, sticky="ew", pady=6)

        acciones = ctk.CTkFrame(marco, fg_color="transparent")
        acciones.grid(row=5, column=0, sticky="ew", pady=(4, 8))
        acciones.grid_columnconfigure(1, weight=1)
        self.btn_plan = ctk.CTkButton(acciones, text="Generar plan de produccion",
                                      height=40,
                                      font=ctk.CTkFont(size=14, weight="bold"),
                                      state="disabled", command=self.ejecutar_plan)
        self.btn_plan.grid(row=0, column=0, padx=(0, 10), sticky="w")
        self.barra_plan = ctk.CTkProgressBar(acciones, mode="indeterminate")
        self.barra_plan.grid(row=0, column=1, sticky="ew", padx=(0, 10))
        self.barra_plan.set(0)
        self.btn_abrir_plan = ctk.CTkButton(
            acciones, text="Abrir Excel", width=120, state="disabled",
            command=lambda: self._abrir(self.ruta_plan))
        self.btn_abrir_plan.grid(row=0, column=2, sticky="e")
        self.btn_carpeta_prod = ctk.CTkButton(
            acciones, text="Abrir carpeta", width=120, state="disabled",
            command=lambda: self._abrir(self.carpeta_futuros_prod.ruta))
        self.btn_carpeta_prod.grid(row=0, column=3, padx=(10, 0), sticky="e")

    # -- Pestana 3: acondicionamiento -----------------------------------------

    def _construir_acondicionamiento(self):
        marco = self.tab_acondi

        origen = ctk.CTkFrame(marco)
        origen.grid(row=0, column=0, sticky="ew", pady=(4, 4))
        origen.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(origen, text="1. Forecast balanceado",
                     font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=12, pady=(8, 4))
        self.lbl_origen_acondi = ctk.CTkLabel(
            origen, text="Se usa el mismo forecast balanceado del paso 2.",
            anchor="w", text_color=("gray40", "gray60"))
        self.lbl_origen_acondi.grid(row=1, column=0, sticky="ew", padx=12,
                                    pady=(0, 2))
        self.sel_md = SelectorExcel(
            marco, "2. Relaciones de formato (hoja 'MD' del catalogo)",
            self._revisar_estados, hoja_sugerida=pa.HOJA_MD,
            ayuda="De aqui salen las tablas de relacion L2 y L4. Se esperan las "
                  "columnas SKU, los cuatro atributos de formato, familia y Linea.")
        self.sel_md.grid(row=1, column=0, sticky="ew", pady=6)

        params = ctk.CTkFrame(marco)
        params.grid(row=2, column=0, sticky="ew", pady=8)
        params.grid_columnconfigure(4, weight=1)
        ctk.CTkLabel(params, text="3. Mes a planear",
                     font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=0, column=0, columnspan=5, sticky="w", padx=12, pady=(10, 4))
        ctk.CTkLabel(params, text="Mes a planear").grid(
            row=1, column=0, padx=(12, 6), pady=(0, 10), sticky="e")
        self.combo_mes_acondi = ctk.CTkComboBox(params, values=[SIN_MES], width=180,
                                                state="disabled",
                                                command=self._al_cambiar_mes_acondi)
        self.combo_mes_acondi.set(SIN_MES)
        self.combo_mes_acondi.grid(row=1, column=1, pady=(0, 10), sticky="w")
        ctk.CTkLabel(params, text="La secuencia busca el minimo de cambios de formato "
                                  f"({pa.HORAS_CAMBIO_FORMATO} h cada uno)\ny "
                                  "desempata por relation total.",
                     justify="left", font=ctk.CTkFont(size=11),
                     text_color=("gray40", "gray60")).grid(
            row=1, column=2, columnspan=3, padx=(12, 12), pady=(0, 10), sticky="w")

        futuros = ctk.CTkFrame(marco)
        futuros.grid(row=3, column=0, sticky="ew", pady=8)
        futuros.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(futuros, text="4. Planes a futuro (opcional)",
                     font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=12, pady=(10, 4))
        self.var_futuros_acondi = ctk.BooleanVar(value=False)
        self.chk_futuros_acondi = ctk.CTkCheckBox(
            futuros, text="Generar todos los planes de acondicionamiento del "
                          "horizonte", variable=self.var_futuros_acondi,
            command=self._revisar_estados)
        self.chk_futuros_acondi.grid(row=1, column=0, sticky="w", padx=12,
                                     pady=(0, 6))
        self.carpeta_futuros_acondi = SelectorCarpeta(
            futuros, "Carpeta de los planes...", self._al_cambiar_carpeta_acondi)
        self.carpeta_futuros_acondi.grid(row=2, column=0, sticky="ew",
                                         padx=12, pady=(0, 6))


        self.sel_salida_acondi = SelectorSalida(
            marco, "5. Archivo de salida del acondicionamiento (.xlsx)",
            self._revisar_estados,
            lambda: f"Plan_Acondi_DC_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx")
        self.sel_salida_acondi.grid(row=4, column=0, sticky="ew", pady=8)

        acciones = ctk.CTkFrame(marco, fg_color="transparent")
        acciones.grid(row=5, column=0, sticky="ew", pady=(4, 8))
        acciones.grid_columnconfigure(1, weight=1)
        self.btn_acondi = ctk.CTkButton(
            acciones, text="Generar plan de acondicionamiento", height=40,
            font=ctk.CTkFont(size=14, weight="bold"), state="disabled",
            command=self.ejecutar_acondi)
        self.btn_acondi.grid(row=0, column=0, padx=(0, 10), sticky="w")
        self.barra_acondi = ctk.CTkProgressBar(acciones, mode="indeterminate")
        self.barra_acondi.grid(row=0, column=1, sticky="ew", padx=(0, 10))
        self.barra_acondi.set(0)
        self.btn_abrir_acondi = ctk.CTkButton(
            acciones, text="Abrir Excel", width=120, state="disabled",
            command=lambda: self._abrir(self.ruta_acondi))
        self.btn_abrir_acondi.grid(row=0, column=2, sticky="e")
        self.btn_carpeta_acondi = ctk.CTkButton(
            acciones, text="Abrir carpeta", width=120, state="disabled",
            command=lambda: self._abrir(self.carpeta_futuros_acondi.ruta))
        self.btn_carpeta_acondi.grid(row=0, column=3, padx=(10, 0), sticky="e")

    # -- Bitacora --------------------------------------------------------------

    def _construir_bitacora(self):
        marco = ctk.CTkFrame(self)
        marco.grid(row=1, column=0, sticky="nsew", padx=16, pady=(0, 12))
        marco.grid_columnconfigure(0, weight=1)
        marco.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(marco, text="Bitacora de ejecucion",
                     font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=12, pady=(10, 4))

        self.txt_log = ctk.CTkTextbox(marco, wrap="none", height=70,
                                      font=ctk.CTkFont(family="Consolas", size=11))
        self.txt_log.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 8))
        self.txt_log.configure(state="disabled")

    def _vaciar_cola(self):
        """Pasa a la pantalla lo que dejo el hilo de calculo en la cola.

        Tkinter solo se puede tocar desde el hilo principal, por eso el hilo
        escribe en la cola y la ventana la revisa cada 200 ms.
        """
        while not self.mensajes.empty():
            self._escribir(self.mensajes.get())
        self.after(200, self._vaciar_cola)

    def _escribir(self, texto):
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", f"{pd.Timestamp.now():%H:%M:%S} | {texto}\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    # -- Acciones del usuario --------------------------------------------------

    def _al_cambiar_forecast(self):
        """Refresca la lista de meses cuando cambia el forecast o su hoja."""
        self.meses = []
        if self.sel_fcst.esta_listo():
            try:
                self.meses = pb.listar_meses(self.sel_fcst.ruta, self.sel_fcst.hoja)
            except Exception as error:
                messagebox.showwarning("No se pudieron leer los meses", str(error))
                self.mensajes.put(traceback.format_exc())

        if self.meses:
            self.combo_mes.configure(values=[TODOS_LOS_MESES] + self.meses,
                                     state="readonly")
            self._escribir(f"Meses en el forecast: {len(self.meses)} "
                           f"({self.meses[0]} a {self.meses[-1]})")
        else:
            self.combo_mes.configure(values=[TODOS_LOS_MESES], state="disabled")
        self.combo_mes.set(TODOS_LOS_MESES)

        self._revisar_estados()

    def _al_cambiar_catalogo(self):
        """El catalogo del paso 1 precarga las hojas de paros y de relaciones.

        Es el mismo libro, asi que no tiene sentido pedirlo tres veces: solo se
        deja que el usuario elija que hoja usar en cada paso.
        """
        if self.sel_catalogo.ruta is not None:
            self.sel_paros.fijar_archivo(self.sel_catalogo.ruta)
            self.sel_md.fijar_archivo(self.sel_catalogo.ruta)
        self._revisar_estados()

    def _al_cambiar_mes_plan(self, _mes=None):
        """Mueve los selectores de arranque al mes que se va a planear."""
        mes = self.combo_mes_plan.get()
        if mes not in ("", SIN_MES):
            try:
                anio, numero = pp.mes_a_anio_mes(mes)
                self.fecha_dc1.fijar_mes(anio, numero)
                self.fecha_dc2.fijar_mes(anio, numero)
            except ValueError:
                pass
        self._revisar_estados()

    def _al_cambiar_mes_acondi(self, _mes=None):
        self._revisar_estados()

    # -- Carpeta de planes a futuro: una sola para los dos pasos --------------
    #
    # El usuario pidio que los planes de produccion y de acondicionamiento del
    # mismo mes y esquema queden juntos. Para no obligarlo a elegir dos veces la
    # misma ruta, el selector que se toque replica su valor en el otro.

    def _al_cambiar_carpeta_prod(self):
        if self.carpeta_futuros_prod.ruta is not None:
            self.carpeta_futuros_acondi.fijar(self.carpeta_futuros_prod.ruta)
        self._revisar_estados()

    def _al_cambiar_carpeta_acondi(self):
        if self.carpeta_futuros_acondi.ruta is not None:
            self.carpeta_futuros_prod.fijar(self.carpeta_futuros_acondi.ruta)
        self._revisar_estados()

    def _llenar_meses_plan(self, meses):
        """Carga los meses del forecast balanceado y preselecciona el actual."""
        meses = list(meses)
        actual = pd.Timestamp.today().strftime("%m-%Y")
        for combo in (self.combo_mes_plan, self.combo_mes_acondi):
            if not meses:
                combo.configure(values=[SIN_MES], state="disabled")
                combo.set(SIN_MES)
            else:
                # Por omision se planea el mes en curso; si no esta, el primero.
                combo.configure(values=meses, state="readonly")
                combo.set(actual if actual in meses else meses[0])
        if meses:
            self._al_cambiar_mes_plan()
        self._revisar_estados()

    def elegir_balanceo_existente(self):
        """Toma el forecast balanceado de un libro ya generado en el disco."""
        ruta = filedialog.askopenfilename(filetypes=TIPOS_EXCEL)
        if not ruta:
            return
        ruta = Path(ruta)
        try:
            tabla = pd.read_excel(ruta, sheet_name=pp.HOJA_FORECAST)
            self.mensajes.put(f"Forecast balanceado leido de {ruta.name} "
                              f"(hoja '{pp.HOJA_FORECAST}'): {len(tabla)} renglones")
            meses = pp.meses_disponibles(tabla)
        except Exception as error:
            messagebox.showerror("No se pudo leer el forecast balanceado",
                                 f"{ruta.name}\n\n{error}")
            self.mensajes.put(traceback.format_exc())
            return

        # El archivo manda sobre el resultado en memoria.
        self.resultado_balanceo = None
        self.ruta_balanceo = ruta
        self.lbl_lotes.configure(text=str(self.ruta_balanceo),
                                 text_color=("gray10", "gray90"))
        self.lbl_origen_acondi.configure(text=str(self.ruta_balanceo),
                                         text_color=("gray10", "gray90"))
        self._llenar_meses_plan(meses)

    def _abrir(self, ruta):
        """Abre el libro o la carpeta generados con la aplicacion del sistema."""
        if not ruta or not Path(ruta).exists():
            return
        ruta = str(ruta)
        if sys.platform.startswith("win"):
            os.startfile(ruta)
        elif sys.platform == "darwin":
            subprocess.run(["open", ruta], check=False)
        else:
            subprocess.run(["xdg-open", ruta], check=False)

    def _hay_forecast(self):
        return self.resultado_balanceo is not None or self.ruta_balanceo is not None

    def _origen_forecast(self):
        """El resultado en memoria manda; si no hay, el libro del disco."""
        return self.resultado_balanceo or self.ruta_balanceo

    def _revisar_estados(self):
        listo_bal = (self.sel_fcst.esta_listo() and self.sel_catalogo.esta_listo()
                     and self.sel_salida_bal.esta_listo() and bool(self.meses)
                     and not self.ejecutando)
        self.btn_balanceo.configure(state="normal" if listo_bal else "disabled")

        # El checkbox de planes a futuro exige su carpeta antes de correr.
        futuros_prod_ok = (not self.var_futuros_prod.get()
                           or self.carpeta_futuros_prod.ruta is not None)
        listo_plan = (self._hay_forecast() and self.sel_salida_plan.esta_listo()
                      and self.sel_paros.esta_listo()
                      and self.combo_mes_plan.get() not in ("", SIN_MES)
                      and futuros_prod_ok and not self.ejecutando)
        self.btn_plan.configure(state="normal" if listo_plan else "disabled")

        futuros_acondi_ok = (not self.var_futuros_acondi.get()
                             or self.carpeta_futuros_acondi.ruta is not None)
        listo_acondi = (self._hay_forecast() and self.sel_md.esta_listo()
                        and self.sel_salida_acondi.esta_listo()
                        and self.combo_mes_acondi.get() not in ("", SIN_MES)
                        and futuros_acondi_ok and not self.ejecutando)
        self.btn_acondi.configure(state="normal" if listo_acondi else "disabled")

        self.carpeta_futuros_prod.habilitar(
            self.var_futuros_prod.get() and not self.ejecutando)
        self.carpeta_futuros_acondi.habilitar(
            self.var_futuros_acondi.get() and not self.ejecutando)

    # -- Lectura de parametros -------------------------------------------------

    def _leer_parametros_balanceo(self):
        """Valida las cajas del paso 1. Regresa un dict o None si hay error."""
        try:
            lote = int(float(self.ent_lote.get().replace(",", "")))
        except ValueError:
            messagebox.showerror("Parametros invalidos",
                                 "Jeringas por lote debe ser un numero.")
            return None
        if lote <= 0:
            messagebox.showerror("Parametros invalidos",
                                 "El lote debe ser mayor a cero.")
            return None
        return {"lote": lote}

    def _leer_parametros_plan(self):
        """Arma los parametros del paso 2. Regresa un dict o None si hay error.

        Las horas de paro ya no se teclean: el modulo las lee de la hoja que se
        eligio aqui. Las fechas de arranque vienen de listas, asi que no hay
        nada que validar: no existe forma de elegir una fecha invalida.
        """
        mes = self.combo_mes_plan.get()
        if mes in ("", SIN_MES):
            messagebox.showerror("Falta el mes", "Elige el mes a planear.")
            return None

        parametros = {
            "mes": mes,
            "ruta_paros": self.sel_paros.ruta,
            "hoja_paros": self.sel_paros.hoja,
            "fecha_inicio_dc1": self.fecha_dc1.valor(),
            "fecha_inicio_dc2": self.fecha_dc2.valor(),
            "generar_futuros": bool(self.var_futuros_prod.get()),
            "carpeta_futuros": self.carpeta_futuros_prod.ruta,
        }
        return parametros

    def _leer_parametros_acondi(self):
        """Arma los parametros del paso 3."""
        mes = self.combo_mes_acondi.get()
        if mes in ("", SIN_MES):
            messagebox.showerror("Falta el mes", "Elige el mes a planear.")
            return None
        return {
            "ruta_catalogo": self.sel_md.ruta,
            "hoja_md": self.sel_md.hoja,
            "mes": mes,
            "generar_futuros": bool(self.var_futuros_acondi.get()),
            "carpeta_futuros": self.carpeta_futuros_acondi.ruta,
        }

    # -- Ejecucion en segundo plano --------------------------------------------

    def ejecutar_balanceo(self):
        parametros = self._leer_parametros_balanceo()
        if parametros is None:
            return

        mes = self.combo_mes.get()
        parametros["mes"] = mes if mes in self.meses else None

        self._bloquear_interfaz(True)
        self._escribir("=" * 70)
        threading.Thread(target=self._trabajo_balanceo, args=(parametros,),
                         daemon=True).start()

    def _trabajo_balanceo(self, parametros):
        """Cuerpo del hilo: calcula y exporta. No toca widgets."""
        try:
            resultado = pb.ejecutar(
                self.sel_fcst.ruta, self.sel_fcst.hoja,
                self.sel_catalogo.ruta, self.sel_catalogo.hoja,
                self.sel_salida_bal.ruta, **parametros)
        except Exception as error:
            self.mensajes.put(f"ERROR: {error}")
            self.mensajes.put(traceback.format_exc())
            self.after(0, self._al_fallar, error)
            return
        self.after(0, self._al_terminar_balanceo, resultado)

    def _al_terminar_balanceo(self, resultado):
        self._bloquear_interfaz(False)
        self.resultado_balanceo = resultado
        self.ruta_balanceo = Path(self.sel_salida_bal.ruta).with_suffix(".xlsx")
        self.btn_abrir_bal.configure(state="normal")

        # Los pasos 2 y 3 ya pueden trabajar con el forecast de esta corrida.
        etiqueta = f"{EN_MEMORIA} ({self.ruta_balanceo.name})"
        self.lbl_lotes.configure(text=etiqueta, text_color=("gray10", "gray90"))
        self.lbl_origen_acondi.configure(text=etiqueta,
                                         text_color=("gray10", "gray90"))
        self._llenar_meses_plan(pp.meses_disponibles(resultado["detalle"]))
        if not self.sel_salida_plan.esta_listo():
            self.sel_salida_plan.lbl.configure(
                text="Define el archivo de salida para generar el plan.")

        meses = resultado["meses"]
        messagebox.showinfo(
            "Balanceo terminado",
            f"Meses procesados: {len(meses)} ({meses[0]} a {meses[-1]})\n"
            f"Archivo generado:\n{self.ruta_balanceo}\n\n"
            "Ya puedes pasar a la pestana '2. Plan de produccion'.")

    def ejecutar_plan(self):
        parametros = self._leer_parametros_plan()
        if parametros is None:
            return

        self._bloquear_interfaz(True)
        self._escribir("=" * 70)
        threading.Thread(target=self._trabajo_plan,
                         args=(self._origen_forecast(), parametros),
                         daemon=True).start()

    def _trabajo_plan(self, origen, parametros):
        try:
            resultado = pp.ejecutar(origen, ruta_salida=self.sel_salida_plan.ruta,
                                    **parametros)
        except Exception as error:
            self.mensajes.put(f"ERROR: {error}")
            self.mensajes.put(traceback.format_exc())
            self.after(0, self._al_fallar, error)
            return
        self.after(0, self._al_terminar_plan, resultado)

    def _al_terminar_plan(self, resultado):
        self._bloquear_interfaz(False)
        self.ruta_plan = resultado["ruta_excel"]
        self.btn_abrir_plan.configure(state="normal")

        bloques = []
        for esquema in pp.ESQUEMAS:
            plan = resultado["planes"].get(esquema)
            if plan is None:
                bloques.append(f"{esquema}: no generado, no alcanzan las horas")
                continue
            detalle = " | ".join(
                f"{f[pp.COL_DC]} {f['horas_requeridas']:,.0f}/"
                f"{f['horas_disponibles']:,.0f} h"
                for _, f in plan["estado"].iterrows())
            bloques.append(f"{esquema}: {detalle}")

        futuros = resultado.get("archivos_futuros") or []
        if futuros:
            self.btn_carpeta_prod.configure(state="normal")
            bloques.append(f"Planes a futuro: {len(futuros)} archivos en "
                           f"{self.carpeta_futuros_prod.ruta}")

        messagebox.showinfo(
            "Plan de produccion terminado",
            f"Mes planeado: {resultado['mes']}\n\n" + "\n".join(bloques) +
            f"\n\nArchivo generado:\n{self.ruta_plan}")

    def ejecutar_acondi(self):
        parametros = self._leer_parametros_acondi()
        if parametros is None:
            return

        self._bloquear_interfaz(True)
        self._escribir("=" * 70)
        threading.Thread(target=self._trabajo_acondi,
                         args=(self._origen_forecast(), parametros),
                         daemon=True).start()

    def _trabajo_acondi(self, origen, parametros):
        try:
            resultado = pa.ejecutar(origen,
                                    ruta_salida=self.sel_salida_acondi.ruta,
                                    **parametros)
        except Exception as error:
            self.mensajes.put(f"ERROR: {error}")
            self.mensajes.put(traceback.format_exc())
            self.after(0, self._al_fallar, error)
            return
        self.after(0, self._al_terminar_acondi, resultado)

    def _al_terminar_acondi(self, resultado):
        self._bloquear_interfaz(False)
        self.ruta_acondi = resultado["ruta_excel"]
        self.btn_abrir_acondi.configure(state="normal")

        bloques = []
        plan = resultado["planes"].get(resultado["mes"])
        if plan is not None:
            for linea, secuencia in plan["secuencias"].items():
                horas = secuencia["cambios"] * pa.HORAS_CAMBIO_FORMATO
                bloques.append(
                    f"{linea}: {len(secuencia['orden'])} SKU | "
                    f"{secuencia['cambios']} cambios de formato ({horas:,.1f} h)")

        futuros = resultado.get("archivos_futuros") or []
        if futuros:
            self.btn_carpeta_acondi.configure(state="normal")
            bloques.append(f"Planes a futuro: {len(futuros)} archivos en "
                           f"{self.carpeta_futuros_acondi.ruta}")

        messagebox.showinfo(
            "Plan de acondicionamiento terminado",
            f"Mes planeado: {resultado['mes']}\n\n" + "\n".join(bloques) +
            f"\n\nArchivo generado:\n{self.ruta_acondi}")

    def _al_fallar(self, error):
        self._bloquear_interfaz(False)
        messagebox.showerror("El proceso no se completo", str(error))

    def _bloquear_interfaz(self, activo):
        """Bloquea o libera los controles segun si hay un calculo en curso."""
        self.ejecutando = activo
        estado = "disabled" if activo else "normal"

        for selector in (self.sel_fcst, self.sel_catalogo, self.sel_paros,
                         self.sel_md, self.sel_salida_bal, self.sel_salida_plan,
                         self.sel_salida_acondi):
            selector.habilitar(not activo)
        for fecha in (self.fecha_dc1, self.fecha_dc2):
            fecha.habilitar(not activo)
        self.btn_lotes.configure(state=estado)
        self.ent_lote.configure(state=estado)
        self.chk_futuros_prod.configure(state=estado)
        self.chk_futuros_acondi.configure(state=estado)

        self.combo_mes.configure(
            state="disabled" if activo or not self.meses else "readonly")
        for combo in (self.combo_mes_plan, self.combo_mes_acondi):
            combo.configure(
                state="disabled" if activo or combo.get() == SIN_MES
                else "readonly")

        if activo:
            self.btn_balanceo.configure(state="disabled", text="Procesando...")
            self.btn_plan.configure(state="disabled", text="Procesando...")
            self.btn_acondi.configure(state="disabled", text="Procesando...")
            for boton in (self.btn_abrir_bal, self.btn_abrir_plan,
                          self.btn_abrir_acondi, self.btn_carpeta_prod,
                          self.btn_carpeta_acondi):
                boton.configure(state="disabled")
            for barra in (self.barra_bal, self.barra_plan, self.barra_acondi):
                barra.start()
        else:
            self.btn_balanceo.configure(text="Ejecutar balanceo")
            self.btn_plan.configure(text="Generar plan de produccion")
            self.btn_acondi.configure(text="Generar plan de acondicionamiento")
            for barra in (self.barra_bal, self.barra_plan, self.barra_acondi):
                barra.stop()
                barra.set(0)
            if self.ruta_balanceo:
                self.btn_abrir_bal.configure(state="normal")
            if self.ruta_plan:
                self.btn_abrir_plan.configure(state="normal")
            if self.ruta_acondi:
                self.btn_abrir_acondi.configure(state="normal")
            self._revisar_estados()


if __name__ == "__main__":
    AppPlanning().mainloop()
