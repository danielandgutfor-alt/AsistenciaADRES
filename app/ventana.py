from pathlib import Path
import threading
import traceback
import subprocess
import os
import sys

import customtkinter as ctk
from tkinter import filedialog, messagebox

from core.parser import Parser
from core.excel import ExcelExporter
from core.zip_manager import ZipManager


ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class VentanaPrincipal(ctk.CTk):

    def __init__(self):

        super().__init__()

        self.title("Asistencia ADRES")
        self.geometry("850x620")
        self.minsize(820, 600)

        self.archivo = ctk.StringVar()
        self.carpeta = ctk.StringVar(value="salida")

        self._crear_interfaz()

    # ---------------------------------------------------------

    def _crear_interfaz(self):

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        contenedor = ctk.CTkFrame(self)
        contenedor.grid(
            row=0,
            column=0,
            padx=20,
            pady=20,
            sticky="nsew"
        )

        contenedor.grid_columnconfigure(1, weight=1)

        titulo = ctk.CTkLabel(
            contenedor,
            text="Generador de Consolidado de Asistencia",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        titulo.grid(
            row=0,
            column=0,
            columnspan=3,
            pady=(15,25)
        )

        # ---------------- Archivo ----------------

        ctk.CTkLabel(
            contenedor,
            text="Archivo Excel:"
        ).grid(
            row=1,
            column=0,
            padx=10,
            pady=8,
            sticky="w"
        )

        self.entry_archivo = ctk.CTkEntry(
            contenedor,
            textvariable=self.archivo
        )

        self.entry_archivo.grid(
            row=1,
            column=1,
            sticky="ew",
            padx=10
        )

        self.btn_archivo = ctk.CTkButton(
            contenedor,
            text="Examinar",
            width=120,
            command=self.seleccionar_archivo
        )

        self.btn_archivo.grid(
            row=1,
            column=2,
            padx=10
        )

        # ---------------- Carpeta ----------------

        ctk.CTkLabel(
            contenedor,
            text="Carpeta salida:"
        ).grid(
            row=2,
            column=0,
            padx=10,
            pady=8,
            sticky="w"
        )

        self.entry_carpeta = ctk.CTkEntry(
            contenedor,
            textvariable=self.carpeta
        )

        self.entry_carpeta.grid(
            row=2,
            column=1,
            sticky="ew",
            padx=10
        )

        self.btn_carpeta = ctk.CTkButton(
            contenedor,
            text="Examinar",
            width=120,
            command=self.seleccionar_carpeta
        )

        self.btn_carpeta.grid(
            row=2,
            column=2,
            padx=10
        )

        # ---------------- Botones ----------------

        botones = ctk.CTkFrame(
            contenedor,
            fg_color="transparent"
        )

        botones.grid(
            row=3,
            column=0,
            columnspan=3,
            pady=20
        )

        self.btn_procesar = ctk.CTkButton(
            botones,
            text="Procesar",
            width=180,
            command=self.procesar
        )

        self.btn_procesar.pack(
            side="left",
            padx=8
        )

        self.btn_abrir = ctk.CTkButton(
            botones,
            text="Abrir carpeta",
            width=180,
            state="disabled",
            command=self.abrir_carpeta
        )

        self.btn_abrir.pack(
            side="left",
            padx=8
        )

        # ---------------- Barra ----------------

        self.progress = ctk.CTkProgressBar(contenedor)
        self.progress.grid(
            row=4,
            column=0,
            columnspan=3,
            padx=10,
            sticky="ew"
        )

        self.progress.set(0)

        self.lbl_estado = ctk.CTkLabel(
            contenedor,
            text="Esperando..."
        )

        self.lbl_estado.grid(
            row=5,
            column=0,
            columnspan=3,
            pady=8
        )

        self.log = ctk.CTkTextbox(
            contenedor,
            height=260
        )

        self.log.grid(
            row=6,
            column=0,
            columnspan=3,
            padx=10,
            pady=15,
            sticky="nsew"
        )

        contenedor.grid_rowconfigure(6, weight=1)

            # ---------------------------------------------------------

    def seleccionar_archivo(self):

        archivo = filedialog.askopenfilename(
            title="Seleccione el archivo Excel",
            filetypes=[
                ("Archivos Excel", "*.xlsx *.xls"),
                ("Todos los archivos", "*.*")
            ]
        )

        if archivo:
            self.archivo.set(archivo)

    # ---------------------------------------------------------

    def seleccionar_carpeta(self):

        carpeta = filedialog.askdirectory(
            title="Seleccione la carpeta de salida"
        )

        if carpeta:
            self.carpeta.set(carpeta)

    # ---------------------------------------------------------

    def escribir_log(self, texto):

        self.log.insert("end", texto + "\n")
        self.log.see("end")
        self.update_idletasks()

    # ---------------------------------------------------------

    def cambiar_estado(self, texto):

        self.lbl_estado.configure(text=texto)
        self.update_idletasks()

    # ---------------------------------------------------------

    def bloquear(self):

        self.btn_procesar.configure(state="disabled")
        self.btn_archivo.configure(state="disabled")
        self.btn_carpeta.configure(state="disabled")
        self.btn_abrir.configure(state="disabled")

    # ---------------------------------------------------------

    def desbloquear(self):

        self.btn_procesar.configure(state="normal")
        self.btn_archivo.configure(state="normal")
        self.btn_carpeta.configure(state="normal")

    # ---------------------------------------------------------

    def procesar(self):

        if not self.archivo.get():

            messagebox.showwarning(
                "Archivo",
                "Seleccione el archivo Excel."
            )
            return

        hilo = threading.Thread(
            target=self._procesar,
            daemon=True
        )

        hilo.start()

    # ---------------------------------------------------------

    def _procesar(self):

        self.bloquear()

        self.progress.set(0)

        self.log.delete("1.0", "end")

        try:

            archivo = Path(self.archivo.get())
            carpeta = Path(self.carpeta.get())

            carpeta.mkdir(
                parents=True,
                exist_ok=True
            )

            self.cambiar_estado("Leyendo archivo...")

            self.escribir_log(
                f"Archivo: {archivo.name}"
            )

            self.progress.set(0.15)

            parser = Parser(archivo)

            dataframe = parser.obtener_dataframe()

            self.progress.set(0.55)

            self.escribir_log(
                f"Registros encontrados: {len(dataframe)}"
            )

            self.cambiar_estado(
                "Generando archivos Excel..."
            )

            exportador = ExcelExporter(carpeta)

            exportador.exportar(dataframe)

            self.progress.set(0.85)

            self.cambiar_estado(
                "Comprimiendo dependencias..."
            )

            zip_manager = ZipManager(carpeta)

            archivo_zip = zip_manager.comprimir_dependencias()

            self.progress.set(1)

            self.escribir_log("")
            self.escribir_log("Proceso finalizado correctamente.")
            self.escribir_log("")
            self.escribir_log(
                f"Consolidado: {carpeta / 'Consolidado_General.xlsx'}"
            )
            self.escribir_log(
                f"ZIP: {archivo_zip}"
            )

            self.cambiar_estado("Proceso terminado.")

            self._carpeta_resultado = carpeta

            self.btn_abrir.configure(
                state="normal"
            )

        except Exception:

            self.progress.set(0)

            self.cambiar_estado("Ocurrió un error.")

            self.escribir_log("")
            self.escribir_log(traceback.format_exc())

            messagebox.showerror(
                "Error",
                "Se produjo un error durante el procesamiento.\n\nRevise el log."
            )

        finally:

            self.desbloquear()

                # ---------------------------------------------------------

    def abrir_carpeta(self):

        if not hasattr(self, "_carpeta_resultado"):
            return

        carpeta = self._carpeta_resultado

        try:

            if sys.platform.startswith("win"):

                os.startfile(carpeta)

            elif sys.platform == "darwin":

                subprocess.Popen(["open", str(carpeta)])

            else:

                subprocess.Popen(["xdg-open", str(carpeta)])

        except Exception:

            messagebox.showerror(
                "Error",
                "No fue posible abrir la carpeta de salida."
            )

    # ---------------------------------------------------------

    def limpiar(self):

        self.archivo.set("")
        self.progress.set(0)
        self.lbl_estado.configure(
            text="Esperando..."
        )
        self.log.delete(
            "1.0",
            "end"
        )

    # ---------------------------------------------------------

    def ejecutar(self):

        self.mainloop()