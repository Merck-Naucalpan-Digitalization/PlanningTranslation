import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

class ExcelApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Planning")
        self.geometry("400x250")

        self.input_path = None

        self.label_status = ctk.CTkLabel(self, text="No hay archivo seleccionado")
        self.label_status.pack(pady=20)

        self.btn_select = ctk.CTkButton(self, text="Seleccionar Excel", command=self.select_file)
        self.btn_select.pack(pady=10)

        self.btn_process = ctk.CTkButton(self, text="Procesar y Guardar", command=self.process_file, state="disabled")
        self.btn_process.pack(pady=10)

    def select_file(self):
        path = filedialog.askopenfilename(
            title="Selecciona el archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )
        if path:
            self.input_path = path
            self.label_status.configure(text=f"Archivo: {path.split('/')[-1]}")
            self.btn_process.configure(state="normal")

    def process_file(self):
        try:
            df = pd.read_excel(self.input_path)

            # --- aquí va lógica de procesamiento -------
            df_procesado = df  # ejemplo: sin transformar
            # -------------------------------------------

            save_path = filedialog.asksaveasfilename(
                title="Guardar archivo procesado",
                defaultextension=".xlsx",
                filetypes=[("Archivos Excel", "*.xlsx")]
            )
            if save_path:
                df_procesado.to_excel(save_path, index=False)
                messagebox.showinfo("Éxito", f"Archivo guardado en:\n{save_path}")

        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un problema:\n{e}")

if __name__ == "__main__":
    app = ExcelApp()
    app.mainloop()