"""
app_planning.py
Ventana de escritorio para el balanceo DC1 / DC2.

Aqui no hay reglas de negocio: todo el calculo vive en planning_balanceo.py.
Esta capa solo pide los archivos, arma los parametros, lanza el proceso en un
hilo aparte para que la ventana no se congele y muestra la bitacora.

Requisitos: pip install customtkinter pandas openpyxl
"""

import os
import queue
import subprocess
import sys
import threading
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk
import pandas as pd

import planning_balanceo as pb

ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

TIPOS_EXCEL = [("Archivos de Excel", "*.xlsx *.xlsm *.xls")]
SIN_ARCHIVO = "Ningun archivo seleccionado"
SIN_HOJA = "Selecciona una hoja"
TODOS_LOS_MESES = "Horizonte completo"


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
        ruta = filedialog.askopenfilename(filetypes=TIPOS_EXCEL)
        if not ruta:
            return

        try:
            with pd.ExcelFile(ruta) as libro:
                hojas = list(libro.sheet_names)
        except Exception as error:
            messagebox.showerror("No se pudo leer el archivo",
                                 f"{Path(ruta).name}\n\n{error}")
            return

        # Al cambiar de archivo, la hoja anterior deja de valer.
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


class AppPlanning(ctk.CTk):
    """Ventana principal."""

    def __init__(self):
        super().__init__()
        self.title("Balanceo de lineas DC1 / DC2 - Planning")
        self.geometry("760x720")
        self.minsize(700, 640)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)

        self.ruta_salida = None
        self.meses = []
        self.ejecutando = False
        self.mensajes = queue.Queue()

        self._construir_entradas()
        self._construir_parametros()
        self._construir_salida()
        self._construir_bitacora()
        self._construir_acciones()

        # El modulo de calculo manda sus mensajes a la cola en vez de la consola.
        pb.log = self.mensajes.put
        self._vaciar_cola()
        self._revisar_si_puede_ejecutar()

    # -- Construccion de la interfaz -------------------------------------------

    def _construir_entradas(self):
        self.sel_fcst = SelectorExcel(self, "1. Forecast P&G (demanda por SKU)",
                                      self._al_cambiar_forecast)
        self.sel_fcst.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 8))

        self.sel_catalogo = SelectorExcel(self, "2. Catalogo de productos DC",
                                          self._revisar_si_puede_ejecutar,
                                          hoja_sugerida="Catalogo")
        self.sel_catalogo.grid(row=1, column=0, sticky="ew", padx=16, pady=8)

    def _construir_parametros(self):
        marco = ctk.CTkFrame(self)
        marco.grid(row=2, column=0, sticky="ew", padx=16, pady=8)
        marco.grid_columnconfigure(5, weight=1)

        ctk.CTkLabel(marco, text="3. Mes del ejercicio y parametros",
                     font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=0, column=0, columnspan=6, sticky="w", padx=12, pady=(10, 4))

        ctk.CTkLabel(marco, text="Mes a analizar").grid(row=1, column=0, padx=(12, 6),
                                                        pady=(0, 10), sticky="e")
        self.combo_mes = ctk.CTkComboBox(marco, values=[TODOS_LOS_MESES], width=180,
                                         state="disabled")
        self.combo_mes.set(TODOS_LOS_MESES)
        self.combo_mes.grid(row=1, column=1, columnspan=2, pady=(0, 10), sticky="w")

        ctk.CTkLabel(marco, text="Limita las hojas 'Detalle SKU' y 'Residuos'; los "
                                 "lotes\nse calculan siempre sobre todo el horizonte.",
                     justify="left", font=ctk.CTkFont(size=11),
                     text_color=("gray40", "gray60")).grid(
            row=1, column=3, columnspan=3, padx=(4, 12), pady=(0, 10), sticky="w")

        self.ent_lote = self._campo(marco, "Jeringas por lote", 0, pb.TAMANO_LOTE, 110)
        self.ent_tolerancia = self._campo(marco, "Tolerancia (%)", 2,
                                          pb.TOLERANCIA * 100, 80)
        self.ent_decimales = self._campo(marco, "Decimales", 4, pb.DECIMALES, 70)

    def _campo(self, marco, texto, columna, valor, ancho):
        """Etiqueta + caja de texto en la fila de parametros."""
        ctk.CTkLabel(marco, text=texto).grid(row=2, column=columna, padx=(12, 6),
                                             pady=(0, 12), sticky="e")
        entrada = ctk.CTkEntry(marco, width=ancho)
        entrada.insert(0, str(valor))
        entrada.grid(row=2, column=columna + 1, pady=(0, 12), sticky="w")
        return entrada

    def _construir_salida(self):
        marco = ctk.CTkFrame(self)
        marco.grid(row=3, column=0, sticky="ew", padx=16, pady=8)
        marco.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(marco, text="4. Archivo de salida (.xlsx)",
                     font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=12, pady=(10, 4))

        self.btn_guardar = ctk.CTkButton(marco, text="Guardar como...", width=110,
                                         command=self.elegir_salida)
        self.btn_guardar.grid(row=1, column=0, padx=(12, 8), pady=(0, 12))

        self.lbl_salida = ctk.CTkLabel(marco, text="Ningun archivo definido",
                                       anchor="w", text_color=("gray40", "gray60"))
        self.lbl_salida.grid(row=1, column=1, sticky="ew", padx=(0, 12), pady=(0, 12))

    def _construir_bitacora(self):
        marco = ctk.CTkFrame(self)
        marco.grid(row=4, column=0, sticky="nsew", padx=16, pady=8)
        marco.grid_columnconfigure(0, weight=1)
        marco.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(marco, text="Bitacora de ejecucion",
                     font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=12, pady=(10, 4))

        self.txt_log = ctk.CTkTextbox(marco, wrap="none",
                                      font=ctk.CTkFont(family="Consolas", size=11))
        self.txt_log.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))
        self.txt_log.configure(state="disabled")

    def _construir_acciones(self):
        marco = ctk.CTkFrame(self, fg_color="transparent")
        marco.grid(row=5, column=0, sticky="ew", padx=16, pady=(0, 8))
        marco.grid_columnconfigure(1, weight=1)

        self.btn_ejecutar = ctk.CTkButton(marco, text="Ejecutar balanceo", height=40,
                                          font=ctk.CTkFont(size=14, weight="bold"),
                                          state="disabled", command=self.ejecutar)
        self.btn_ejecutar.grid(row=0, column=0, padx=(0, 10), sticky="w")

        self.barra = ctk.CTkProgressBar(marco, mode="indeterminate")
        self.barra.grid(row=0, column=1, sticky="ew", padx=(0, 10))
        self.barra.set(0)

        self.btn_abrir = ctk.CTkButton(marco, text="Abrir Excel", width=120,
                                       state="disabled", command=self.abrir_salida)
        self.btn_abrir.grid(row=0, column=2, sticky="e")

    # -- Bitacora --------------------------------------------------------------

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

        if self.meses:
            self.combo_mes.configure(values=[TODOS_LOS_MESES] + self.meses,
                                     state="readonly")
            self._escribir(f"Meses en el forecast: {len(self.meses)} "
                           f"({self.meses[0]} a {self.meses[-1]})")
        else:
            self.combo_mes.configure(values=[TODOS_LOS_MESES], state="disabled")
        self.combo_mes.set(TODOS_LOS_MESES)

        self._revisar_si_puede_ejecutar()

    def elegir_salida(self):
        ruta = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=f"Balanceo_DC_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx",
            filetypes=[("Libro de Excel", "*.xlsx")])
        if not ruta:
            return
        self.ruta_salida = Path(ruta)
        self.lbl_salida.configure(text=str(self.ruta_salida),
                                  text_color=("gray10", "gray90"))
        self._revisar_si_puede_ejecutar()

    def abrir_salida(self):
        """Abre el libro generado con la aplicacion del sistema."""
        if not self.ruta_salida or not self.ruta_salida.exists():
            return
        if sys.platform.startswith("win"):
            os.startfile(self.ruta_salida)
        elif sys.platform == "darwin":
            subprocess.run(["open", str(self.ruta_salida)], check=False)
        else:
            subprocess.run(["xdg-open", str(self.ruta_salida)], check=False)

    def _revisar_si_puede_ejecutar(self):
        listo = (self.sel_fcst.esta_listo() and self.sel_catalogo.esta_listo()
                 and self.ruta_salida is not None and bool(self.meses)
                 and not self.ejecutando)
        self.btn_ejecutar.configure(state="normal" if listo else "disabled")

    def _leer_parametros(self):
        """Valida las cajas de parametros. Regresa un dict o None si hay error."""
        try:
            lote = int(float(self.ent_lote.get().replace(",", "")))
            tolerancia = float(self.ent_tolerancia.get().replace(",", ".")) / 100
            decimales = int(self.ent_decimales.get())
        except ValueError:
            messagebox.showerror("Parametros invalidos",
                                 "Lote, tolerancia y decimales deben ser numericos.")
            return None

        if lote <= 0 or decimales < 0 or not 0 <= tolerancia < 0.5:
            messagebox.showerror(
                "Parametros invalidos",
                "El lote debe ser mayor a cero, los decimales no pueden ser "
                "negativos y la tolerancia debe estar entre 0 y 50 %.")
            return None

        return {"lote": lote, "tolerancia": tolerancia, "decimales": decimales}

    # -- Ejecucion en segundo plano --------------------------------------------

    def ejecutar(self):
        parametros = self._leer_parametros()
        if parametros is None:
            return

        mes = self.combo_mes.get()
        parametros["mes"] = mes if mes in self.meses else None

        self._bloquear_interfaz(True)
        self._escribir("=" * 70)
        threading.Thread(target=self._trabajo, args=(parametros,), daemon=True).start()

    def _trabajo(self, parametros):
        """Cuerpo del hilo: calcula y exporta. No toca widgets."""
        try:
            resultado = pb.ejecutar(
                self.sel_fcst.ruta, self.sel_fcst.hoja,
                self.sel_catalogo.ruta, self.sel_catalogo.hoja,
                self.ruta_salida, **parametros)
        except Exception as error:
            self.mensajes.put(f"ERROR: {error}")
            self.after(0, self._al_fallar, error)
            return
        self.after(0, self._al_terminar, resultado)

    def _al_terminar(self, resultado):
        self._bloquear_interfaz(False)
        self.btn_abrir.configure(state="normal")
        meses = resultado["meses"]
        messagebox.showinfo(
            "Balanceo terminado",
            f"Meses procesados: {len(meses)} ({meses[0]} a {meses[-1]})\n"
            f"Lotes en DC1: {resultado['lotes_dc1']}\n"
            f"Lotes en DC2: {resultado['lotes_dc2']}\n\n"
            f"Archivo generado:\n{self.ruta_salida}")

    def _al_fallar(self, error):
        self._bloquear_interfaz(False)
        messagebox.showerror("El proceso no se completo", str(error))

    def _bloquear_interfaz(self, activo):
        """Bloquea o libera los controles segun si hay un calculo en curso."""
        self.ejecutando = activo
        estado = "disabled" if activo else "normal"

        self.sel_fcst.habilitar(not activo)
        self.sel_catalogo.habilitar(not activo)
        self.btn_guardar.configure(state=estado)
        self.combo_mes.configure(
            state="disabled" if activo or not self.meses else "readonly")
        for entrada in (self.ent_lote, self.ent_tolerancia, self.ent_decimales):
            entrada.configure(state=estado)

        if activo:
            self.btn_ejecutar.configure(state="disabled", text="Procesando...")
            self.btn_abrir.configure(state="disabled")
            self.barra.start()
        else:
            self.btn_ejecutar.configure(text="Ejecutar balanceo")
            self.barra.stop()
            self.barra.set(0)
            self._revisar_si_puede_ejecutar()


if __name__ == "__main__":
    AppPlanning().mainloop()
