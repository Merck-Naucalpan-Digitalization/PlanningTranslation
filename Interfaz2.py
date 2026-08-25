import customtkinter as ctk
import pandas as pd
from tkinter import filedialog

class ExcelAnalyzer(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Analizador de Excel")
        self.geometry("400x250")

        self.excel_file = None
        self.hoja_actual = None

        self.btn_cargar = ctk.CTkButton(self, text="Cargar archivo", command=self.cargar_archivo)
        self.btn_cargar.pack(pady=10)

        self.combo_hojas = ctk.CTkComboBox(self, values=[], command=self.hoja_seleccionada)
        self.combo_hojas.pack(pady=10)
        self.combo_hojas.set("Selecciona una hoja")

        self.btn_siguiente = ctk.CTkButton(self, text="Siguiente", command=self.siguiente, state="disabled")
        self.btn_siguiente.pack(pady=10)

    def cargar_archivo(self):
        ruta = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not ruta:
            return

        self.excel_file = pd.ExcelFile(ruta)
        self.combo_hojas.configure(values=self.excel_file.sheet_names)
        self.combo_hojas.set("Selecciona una hoja")

        # Reinicia estado al cargar un nuevo archivo
        self.hoja_actual = None
        self.btn_siguiente.configure(state="disabled")

    def hoja_seleccionada(self, hoja):
        self.hoja_actual = hoja
        self.btn_siguiente.configure(state="normal")

    def siguiente(self):
        print(f"Hoja seleccionada: {self.hoja_actual}")

if __name__ == "__main__":
    app = ExcelAnalyzer()
    app.mainloop()