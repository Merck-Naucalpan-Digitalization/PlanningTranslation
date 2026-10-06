"""
app_planning.py
Ventana de escritorio para el proceso de Planning en dos pasos:

    Pestana 1 - Balanceo    -> planning_balanceo.py
    Pestana 2 - Produccion  -> planning_produccion.py

Aqui no hay reglas de negocio: todo el calculo vive en los modulos. Esta capa
solo pide los archivos, arma los parametros, lanza el proceso en un hilo
aparte para que la ventana no se congele y muestra la bitacora.

El paso 2 toma el "Forecast balanceado" del paso 1. Si el balanceo ya corrio
en esta sesion usa su resultado en memoria; si no, se puede elegir el .xlsx de
balanceo que ya exista en el disco. Se planea un solo mes: el que se elija en
el combo, y por omision el mes en curso.

Requisitos: pip install customtkinter pandas openpyxl
"""

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
import planning_balanceo as pb
import planning_produccion as pp

# Colores de la interfaz
ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

TIPOS_EXCEL = [("Archivos de Excel", "*.xlsx *.xlsm *.xls")]
SIN_ARCHIVO = "Ningun archivo seleccionado"
SIN_HOJA = "Selecciona una hoja"
SIN_SALIDA = "Ningun archivo definido"
TODOS_LOS_MESES = "Horizonte completo"
SIN_MES = "Selecciona un mes"
EN_MEMORIA = "Resultado del balanceo de esta sesion"


class SelectorExcel(ctk.CTkFrame):
    """Bloque para elegir un archivo de Excel y una de sus hojas."""

    def __init__(self, master, titulo, al_cambiar, hoja_sugerida=None):
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


class AppPlanning(ctk.CTk):
    """Ventana principal con las dos etapas del proceso."""

    def __init__(self):
        super().__init__()
        self.title("Planning DC1 / DC2 - Balanceo y plan de produccion")
        self.geometry("900x820")
        self.minsize(820, 700)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # Estado compartido entre las dos etapas
        self.meses = []
        self.resultado_balanceo = None     # dict que devuelve pb.ejecutar
        self.ruta_balanceo = None          # .xlsx generado o elegido a mano
        self.ruta_plan = None
        self.ejecutando = False
        self.mensajes = queue.Queue()

        self.tabs = ctk.CTkTabview(self)
        self.tabs.grid(row=0, column=0, sticky="nsew", padx=16, pady=(16, 8))
        self.tab_bal = self.tabs.add("1. Balanceo")
        self.tab_prod = self.tabs.add("2. Plan de produccion")
        for tab in (self.tab_bal, self.tab_prod):
            tab.grid_columnconfigure(0, weight=1)

        self._construir_balanceo()
        self._construir_produccion()
        self._construir_bitacora()



        # Joe: los modulos de calculo mandan sus mensajes a la cola en vez de la consola.
        pb.log = self.mensajes.put
        pp.log = self.mensajes.put
        self._vaciar_cola()
        self._revisar_estados()

    # -- Pestana 1: balanceo ---------------------------------------------------

    def _construir_balanceo(self):
        marco = self.tab_bal

        self.sel_fcst = SelectorExcel(marco, "1. Forecast P&G (demanda por SKU)",
                                      self._al_cambiar_forecast)
        self.sel_fcst.grid(row=0, column=0, sticky="ew", pady=(8, 8))

        self.sel_catalogo = SelectorExcel(marco, "2. Catalogo de productos DC",
                                          self._revisar_estados,
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

        ctk.CTkLabel(params, text="Deja 'Horizonte completo' para que el plan de\n"
                                  "produccion pueda elegir cualquier mes despues.",
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
        ctk.CTkLabel(origen, text=f"Se lee la hoja '{pp.HOJA_FORECAST}'. Las cantidades "
                                  "no se modifican: solo cambia de linea un producto\n"
                                  "cuando tiene equivalente en DC1 y DC2.",
                     justify="left", font=ctk.CTkFont(size=11),
                     text_color=("gray40", "gray60")).grid(
            row=2, column=0, columnspan=3, sticky="w", padx=12, pady=(0, 12))

        params = ctk.CTkFrame(marco)
        params.grid(row=1, column=0, sticky="ew", pady=8)
        params.grid_columnconfigure(5, weight=1)
        ctk.CTkLabel(params, text="2. Mes a planear y paros programados",
                     font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=0, column=0, columnspan=6, sticky="w", padx=12, pady=(10, 4))

        ctk.CTkLabel(params, text="Mes a planear").grid(
            row=1, column=0, padx=(12, 6), pady=(0, 10), sticky="e")
        self.combo_mes_plan = ctk.CTkComboBox(params, values=[SIN_MES], width=180,
                                              state="disabled",
                                              command=self._al_cambiar_mes_plan)
        self.combo_mes_plan.set(SIN_MES)
        self.combo_mes_plan.grid(row=1, column=1, pady=(0, 10), sticky="w")

        ctk.CTkLabel(params, text="Se planea un solo mes. Siempre se genera el plan "
                                  "24/7;\nel 24/5 solo si alcanzan las horas del mes.",
                     justify="left", font=ctk.CTkFont(size=11),
                     text_color=("gray40", "gray60")).grid(
            row=1, column=2, columnspan=4, padx=(12, 12), pady=(0, 10), sticky="w")

        self.ent_paro_dc1 = self._campo(params, "Horas de paro DC1", 0, 0, 110)
        self.ent_paro_dc2 = self._campo(params, "Horas de paro DC2", 2, 0, 110)

        self.ent_inicio_dc1 = self._campo(params, "Inicio DC1 (dd/mm/aaaa hh:mm)",
                                          0, "", 170, fila=3)
        self.ent_inicio_dc2 = self._campo(params, "Inicio DC2 (dd/mm/aaaa hh:mm)",
                                          2, "", 170, fila=3)
        ctk.CTkLabel(params, text="Vacio = dia 1 del mes a las 07:00",
                     font=ctk.CTkFont(size=11),
                     text_color=("gray40", "gray60")).grid(
            row=3, column=4, columnspan=2, padx=(12, 12), pady=(0, 12), sticky="w")

        self.sel_salida_plan = SelectorSalida(
            marco, "3. Archivo de salida del plan (.xlsx)", self._revisar_estados,
            lambda: f"Plan_Produccion_DC_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx")
        self.sel_salida_plan.grid(row=2, column=0, sticky="ew", pady=8)

        acciones = ctk.CTkFrame(marco, fg_color="transparent")
        acciones.grid(row=3, column=0, sticky="ew", pady=(4, 8))
        acciones.grid_columnconfigure(1, weight=1)
        self.btn_plan = ctk.CTkButton(acciones, text="Generar plan de produccion",
                                      height=40,
                                      font=ctk.CTkFont(size=14, weight="bold"),
                                      state="enabled", command=self.ejecutar_plan)
        self.btn_plan.grid(row=0, column=0, padx=(0, 10), sticky="w")
        self.barra_plan = ctk.CTkProgressBar(acciones, mode="indeterminate")
        self.barra_plan.grid(row=0, column=1, sticky="ew", padx=(0, 10))
        self.barra_plan.set(0)
        self.btn_abrir_plan = ctk.CTkButton(
            acciones, text="Abrir Excel", width=120, state="disabled",
            command=lambda: self._abrir(self.ruta_plan))
        self.btn_abrir_plan.grid(row=0, column=2, sticky="e")

    # -- Bitacora --------------------------------------------------------------

    def _construir_bitacora(self):
        marco = ctk.CTkFrame(self)
        marco.grid(row=1, column=0, sticky="nsew", padx=16, pady=(0, 16))
        marco.grid_columnconfigure(0, weight=1)
        marco.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(marco, text="Bitacora de ejecucion",
                     font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=12, pady=(10, 4))

        self.txt_log = ctk.CTkTextbox(marco, wrap="none", height=180,
                                      font=ctk.CTkFont(family="Consolas", size=11))
        self.txt_log.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))
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

    def _al_cambiar_mes_plan(self, _mes=None):
        self._revisar_estados()

    def _llenar_meses_plan(self, meses):
        """Carga los meses del forecast balanceado y preselecciona el actual."""
        meses = list(meses)
        if not meses:
            self.combo_mes_plan.configure(values=[SIN_MES], state="disabled")
            self.combo_mes_plan.set(SIN_MES)
        else:
            # Por omision se planea el mes en curso; si no esta, el primero.
            actual = pd.Timestamp.today().strftime("%m-%Y")
            self.combo_mes_plan.configure(values=meses, state="readonly")
            self.combo_mes_plan.set(actual if actual in meses else meses[0])
        self._revisar_estados()

    def elegir_balanceo_existente(self):
        """Toma el forecast balanceado de un libro ya generado en el disco."""
        ruta = filedialog.askopenfilename(filetypes=TIPOS_EXCEL)
        if not ruta:
            return
        try:
            tabla = pd.read_excel(Path(ruta), sheet_name=pp.HOJA_FORECAST)
            hoja=pp.HOJA_FORECAST
            pp.log(f"Forecast balanceado leido de {ruta.name} (hoja '{hoja}'): "
                    f"{len(tabla)} renglones")    
            meses = pp.meses_disponibles(tabla)
        except Exception as error:
            messagebox.showerror("No se pudo leer el forecast balanceado",
                                 f"{Path(ruta).name}\n\n{error}")
            self.mensajes.put(traceback.format_exc())
            return

        # El archivo manda sobre el resultado en memoria.
        self.resultado_balanceo = None
        self.ruta_balanceo = Path(ruta)
        self.lbl_lotes.configure(text=str(self.ruta_balanceo),
                                 text_color=("gray10", "gray90"))
        self._llenar_meses_plan(meses)

    def _abrir(self, ruta):
        """Abre el libro generado con la aplicacion del sistema."""
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

    def _revisar_estados(self):
        listo_bal = (self.sel_fcst.esta_listo() and self.sel_catalogo.esta_listo()
                     and self.sel_salida_bal.esta_listo() and bool(self.meses)
                     and not self.ejecutando)
        self.btn_balanceo.configure(state="normal" if listo_bal else "disabled")

        listo_plan = (self._hay_forecast() and self.sel_salida_plan.esta_listo()
                      and self.combo_mes_plan.get() not in ("", SIN_MES)
                      and not self.ejecutando)
        self.btn_plan.configure(state="normal" if listo_plan else "disabled")

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

    def _leer_hora(self, entrada, etiqueta):
        """Lee una fecha/hora opcional. Regresa (ok, valor)."""
        texto = entrada.get().strip()
        if not texto:
            return True, None
        for formato in ("%d/%m/%Y %H:%M", "%d/%m/%Y", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
            try:
                return True, datetime.strptime(texto, formato)
            except ValueError:
                continue
        messagebox.showerror("Parametros invalidos",
                             f"{etiqueta} no se pudo leer. Use el formato "
                             f"dd/mm/aaaa hh:mm o deje la caja vacia.")
        return False, None

    def _leer_parametros_plan(self):
        """Valida las cajas del paso 2. Regresa un dict o None si hay error."""
        paros = {}
        for dc, entrada in (("DC1", self.ent_paro_dc1), ("DC2", self.ent_paro_dc2)):
            texto = entrada.get().strip().replace(",", "")
            try:
                valor = 0.0 if texto == "" else float(texto)
            except ValueError:
                messagebox.showerror("Parametros invalidos",
                                     f"Las horas de paro de {dc} deben ser un numero.")
                return None
            if valor < 0:
                messagebox.showerror("Parametros invalidos",
                                     f"Las horas de paro de {dc} no pueden ser "
                                     "negativas.")
                return None
            paros[dc] = valor

        ok1, inicio_dc1 = self._leer_hora(self.ent_inicio_dc1, "Inicio DC1")
        if not ok1:
            return None
        ok2, inicio_dc2 = self._leer_hora(self.ent_inicio_dc2, "Inicio DC2")
        if not ok2:
            return None

        return {"mes": self.combo_mes_plan.get(),
                "horas_paro": paros,
                "fecha_inicio_dc1": inicio_dc1,
                "fecha_inicio_dc2": inicio_dc2}

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

        # El paso 2 ya puede trabajar con el forecast balanceado de esta corrida.
        self.lbl_lotes.configure(text=f"{EN_MEMORIA} ({self.ruta_balanceo.name})",
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

        # Si el balanceo corrio en esta sesion se usa su resultado en memoria;
        # si no, se lee el forecast balanceado del libro elegido.
        origen = self.resultado_balanceo or self.ruta_balanceo

        self._bloquear_interfaz(True)
        self._escribir("=" * 70)
        threading.Thread(target=self._trabajo_plan, args=(origen, parametros),
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

        messagebox.showinfo(
            "Plan de produccion terminado",
            f"Mes planeado: {resultado['mes']}\n\n" + "\n".join(bloques) +
            f"\n\nArchivo generado:\n{self.ruta_plan}")

    def _al_fallar(self, error):
        self._bloquear_interfaz(False)
        messagebox.showerror("El proceso no se completo", str(error))

    def _bloquear_interfaz(self, activo):
        """Bloquea o libera los controles segun si hay un calculo en curso."""
        self.ejecutando = activo
        estado = "disabled" if activo else "normal"

        self.sel_fcst.habilitar(not activo)
        self.sel_catalogo.habilitar(not activo)
        self.sel_salida_bal.habilitar(not activo)
        self.sel_salida_plan.habilitar(not activo)
        self.btn_lotes.configure(state=estado)
        self.combo_mes.configure(
            state="disabled" if activo or not self.meses else "readonly")
        self.combo_mes_plan.configure(
            state="disabled" if activo or self.combo_mes_plan.get() == SIN_MES
            else "readonly")
        for widget in (self.ent_lote, self.ent_paro_dc1, self.ent_paro_dc2,
                       self.ent_inicio_dc1, self.ent_inicio_dc2):
            widget.configure(state=estado)

        if activo:
            self.btn_balanceo.configure(state="disabled", text="Procesando...")
            self.btn_plan.configure(state="disabled", text="Procesando...")
            self.btn_abrir_bal.configure(state="disabled")
            self.btn_abrir_plan.configure(state="disabled")
            self.barra_bal.start()
            self.barra_plan.start()
        else:
            self.btn_balanceo.configure(text="Ejecutar balanceo")
            self.btn_plan.configure(text="Generar plan de produccion")
            for barra in (self.barra_bal, self.barra_plan):
                barra.stop()
                barra.set(0)
            if self.ruta_balanceo:
                self.btn_abrir_bal.configure(state="normal")
            if self.ruta_plan:
                self.btn_abrir_plan.configure(state="normal")
            self._revisar_estados()


if __name__ == "__main__":
    AppPlanning().mainloop()
