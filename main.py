import os
import time
import datetime
import threading
import pandas as pd
from openpyxl import load_workbook
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC

import tkinter as tk
from tkinter import filedialog, ttk, messagebox

from config import URL
from login import iniciar_sesion

# =========================
# COLORES CORPORATIVOS SENA
# =========================
VERDE_PRINCIPAL = "#39A900"
VERDE_OSCURO    = "#007832"
AMARILLO        = "#FFD100"
AZUL_OSCURO     = "#003DA5"
BLANCO          = "#FFFFFF"
GRIS_FONDO      = "#F5F5F5"
GRIS_TEXTO      = "#333333"
ROJO_CANCELAR   = "#C0392B"
FUENTE          = "Calibri"

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Consulta Aprendices - SENA Sofia Plus")
        self.root.geometry("540x600")
        self.root.resizable(False, False)
        self.root.configure(bg=GRIS_FONDO)

        self.cancelado = False
        self.proceso_activo = False

        # =========================
        # HEADER SENA
        # =========================
        header = tk.Frame(root, bg=VERDE_PRINCIPAL, height=70)
        header.pack(fill="x")
        header.pack_propagate(False)

        tk.Label(
            header,
            text="SENA",
            font=(FUENTE, 22, "bold"),
            bg=VERDE_PRINCIPAL,
            fg=BLANCO
        ).pack(side="left", padx=(20, 4), pady=15)

        tk.Label(
            header,
            text="Consulta de Aprendices · Sofia Plus",
            font=(FUENTE, 11),
            bg=VERDE_PRINCIPAL,
            fg=BLANCO
        ).pack(side="left", pady=15)

        # FRANJA AMARILLA
        tk.Frame(root, bg=AMARILLO, height=4).pack(fill="x")

        # =========================
        # CONTENIDO
        # =========================
        contenido = tk.Frame(root, bg=GRIS_FONDO)
        contenido.pack(fill="both", expand=True, padx=35, pady=20)

        def label(texto):
            tk.Label(
                contenido, text=texto,
                font=(FUENTE, 9, "bold"),
                bg=GRIS_FONDO, fg=VERDE_OSCURO,
                anchor="w"
            ).pack(fill="x", pady=(8, 2))

        def entry_field(**kwargs):
            e = tk.Entry(
                contenido,
                font=(FUENTE, 10),
                relief="solid", bd=1,
                bg=BLANCO,
                **kwargs
            )
            e.pack(fill="x", ipady=4, pady=(0, 2))
            return e

        # TIPO DE IDENTIFICACIÓN
        label("Tipo de Identificación")
        self.combo_tipo = ttk.Combobox(contenido, state="readonly", font=(FUENTE, 10))
        self.combo_tipo["values"] = [
            "CC - Cédula de Ciudadanía",
            "CE - Cédula de Extranjería",
            "TI - Tarjeta de Identidad",
            "PA - Pasaporte",
            "RC - Registro Civil",
        ]
        self.combo_tipo.current(0)
        self.combo_tipo.pack(fill="x", pady=(0, 2))

        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "TCombobox",
            fieldbackground=BLANCO,
            background=VERDE_PRINCIPAL,
            foreground=GRIS_TEXTO,
            selectbackground=VERDE_PRINCIPAL,
            selectforeground=BLANCO,
        )

        # USUARIO
        label("Usuario (Número de documento)")
        self.entry_usuario = entry_field(width=40)

        # CONTRASEÑA
        label("Contraseña")
        self.entry_password = entry_field(width=40, show="*")

        # ARCHIVO EXCEL
        label("Archivo Excel")
        frame_archivo = tk.Frame(contenido, bg=GRIS_FONDO)
        frame_archivo.pack(fill="x", pady=(0, 2))

        self.entry_archivo = tk.Entry(
            frame_archivo,
            font=(FUENTE, 10),
            relief="solid", bd=1,
            bg=BLANCO
        )
        self.entry_archivo.pack(side="left", fill="x", expand=True, ipady=4)

        tk.Button(
            frame_archivo, text="📂 Buscar",
            font=(FUENTE, 9, "bold"),
            bg=AMARILLO, fg=GRIS_TEXTO,
            relief="flat", cursor="hand2",
            padx=8,
            command=self.seleccionar_archivo
        ).pack(side="left", padx=(6, 0), ipady=4)

        # BARRA DE PROGRESO
        label("Progreso")

        style.configure(
            "SENA.Horizontal.TProgressbar",
            troughcolor="#E0E0E0",
            background=VERDE_PRINCIPAL,
            thickness=18
        )

        self.barra = ttk.Progressbar(
            contenido,
            orient="horizontal",
            mode="determinate",
            style="SENA.Horizontal.TProgressbar"
        )
        self.barra.pack(fill="x", pady=(0, 4))

        self.label_progreso = tk.Label(
            contenido, text="",
            font=(FUENTE, 9),
            bg=GRIS_FONDO, fg="#555555",
            anchor="w"
        )
        self.label_progreso.pack(fill="x")

        # BOTONES
        frame_botones = tk.Frame(contenido, bg=GRIS_FONDO)
        frame_botones.pack(fill="x", pady=(20, 0))

        self.boton_iniciar = tk.Button(
            frame_botones, text="▶  Iniciar Consulta",
            font=(FUENTE, 11, "bold"),
            bg=VERDE_PRINCIPAL, fg=BLANCO,
            relief="flat", cursor="hand2",
            height=2,
            command=self.iniciar_en_hilo,
            activebackground=VERDE_OSCURO,
            activeforeground=BLANCO
        )
        self.boton_iniciar.pack(side="left", expand=True, fill="x", padx=(0, 8))

        self.boton_cancelar = tk.Button(
            frame_botones, text="✖  Cancelar",
            font=(FUENTE, 11, "bold"),
            bg=ROJO_CANCELAR, fg=BLANCO,
            relief="flat", cursor="hand2",
            height=2, state="disabled",
            command=self.cancelar_proceso,
            activebackground="#922B21",
            activeforeground=BLANCO
        )
        self.boton_cancelar.pack(side="left", expand=True, fill="x")

        # FOOTER
        tk.Frame(root, bg=VERDE_OSCURO, height=3).pack(fill="x", side="bottom")
        tk.Label(
            root,
            text="Servicio Nacional de Aprendizaje · SENA © 2024",
            font=(FUENTE, 8),
            bg=GRIS_FONDO, fg="#888888"
        ).pack(side="bottom", pady=4)

    # -------------------------
    def seleccionar_archivo(self):
        archivo = filedialog.askopenfilename(
            title="Selecciona el archivo de aprendices",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )
        if archivo:
            self.entry_archivo.delete(0, tk.END)
            self.entry_archivo.insert(0, archivo)

    # -------------------------
    def cancelar_proceso(self):
        if self.proceso_activo:
            self.cancelado = True
            self.label_progreso.config(text="⛔ Cancelando proceso...")

    # -------------------------
    def iniciar_en_hilo(self):
        hilo = threading.Thread(target=self.ejecutar)
        hilo.daemon = True
        hilo.start()

    # -------------------------
    def ejecutar(self):
        tipo_id  = self.combo_tipo.get().split(" - ")[0].strip()
        usuario  = self.entry_usuario.get().strip()
        password = self.entry_password.get().strip()
        archivo  = self.entry_archivo.get().strip()

        if not usuario or not password:
            messagebox.showerror("Error", "Ingresa usuario y contraseña.")
            return
        if not archivo:
            messagebox.showerror("Error", "Selecciona un archivo Excel.")
            return

        self.cancelado = False
        self.proceso_activo = True
        self.boton_iniciar.config(state="disabled")
        self.boton_cancelar.config(state="normal")
        self.barra["value"] = 0
        self.label_progreso.config(text="Iniciando sesión...")

        try:
            driver, wait = iniciar_sesion(URL, tipo_id, usuario, password)

            df = pd.read_excel(archivo, dtype=str, header=3)
            df = df.dropna(how='all').reset_index(drop=True)
            wb = load_workbook(archivo)
            ws = wb.active

            total = len(df)
            self.barra["maximum"] = total

            def ir_a_formulario():
                driver.switch_to.default_content()
                opcion = wait.until(EC.presence_of_element_located((By.ID, "4011Opcion")))
                driver.execute_script("arguments[0].click();", opcion)
                time.sleep(5)
                for frame in driver.find_elements(By.TAG_NAME, "iframe"):
                    try:
                        driver.switch_to.default_content()
                        driver.switch_to.frame(frame)
                        driver.find_element(By.ID, "inputNumeroIdentificacion")
                        return
                    except:
                        continue
                raise Exception("No se encontró el iframe del formulario")

            for i in range(total):

                if self.cancelado:
                    self.label_progreso.config(text=f"⛔ Proceso cancelado en fila {i+1}")
                    break

                try:
                    tipo_id_fila = str(df["Tipo de Identificación"].iloc[i]).strip()
                    numero_id    = str(df["Numero de Identificación"].iloc[i]).strip()

                    if numero_id.lower() == "nan" or numero_id == "":
                        self.actualizar_progreso(i + 1, total, f"⚠️ Fila {i+1} vacía")
                        continue

                    self.actualizar_progreso(i + 1, total, f"🔎 Consultando {numero_id}...")

                    ir_a_formulario()

                    Select(wait.until(EC.element_to_be_clickable(
                        (By.ID, "seleccionTipoIdentificacion")
                    ))).select_by_value(tipo_id_fila)

                    campo = wait.until(EC.element_to_be_clickable((By.ID, "inputNumeroIdentificacion")))
                    campo.clear()
                    campo.send_keys(numero_id)

                    boton = wait.until(EC.presence_of_element_located((By.ID, "formPrincipal:btnAplicar")))
                    driver.execute_script("arguments[0].click();", boton)

                    time.sleep(3)

                    # =========================
                    # VALIDACIÓN DE RESULTADO
                    # =========================
                    # Verificar si hay filas en la tabla de resultados
                    filas_tabla = driver.find_elements(
                        By.XPATH,
                        "//table//tr[contains(@class,'ui-widget-content')]"
                    )

                    mensajes = driver.find_elements(By.ID, "messages")
                    texto_mensaje = mensajes[0].text.strip() if mensajes else ""

                    if filas_tabla:
                        # Hay registros en la tabla = está matriculado
                        resultado = "MATRICULADO"
                    elif "exitosamente" in texto_mensaje.lower():
                        resultado = "MATRICULADO"
                    elif "no se encontr" in texto_mensaje.lower() or "no existe" in texto_mensaje.lower():
                        resultado = "NO REGISTRADO"
                    else:
                        resultado = "NO REGISTRADO"

                    ws[f"E{i + 5}"] = resultado

                except Exception as e:
                    error_msg = str(e)[:100]
                    ws[f"F{i + 5}"] = f"ERROR: {error_msg}"
                    self.actualizar_progreso(i + 1, total, f"⚠️ Error fila {i+1}: {error_msg}")
                    continue

            # GUARDAR EN LA MISMA CARPETA DEL EXCEL
            carpeta    = os.path.dirname(os.path.abspath(archivo))
            nombre     = f"resultado_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            ruta_final = os.path.join(carpeta, nombre)
            wb.save(ruta_final)

            if not self.cancelado:
                self.label_progreso.config(text=f"✅ Guardado en: {ruta_final}")
                messagebox.showinfo("Proceso completado", f"Archivo guardado en:\n{ruta_final}")
            else:
                messagebox.showwarning("Cancelado", f"Proceso cancelado.\nResultados parciales guardados en:\n{ruta_final}")

        except Exception as e:
            messagebox.showerror("Error general", str(e))

        finally:
            self.proceso_activo = False
            self.cancelado = False
            self.boton_iniciar.config(state="normal")
            self.boton_cancelar.config(state="disabled")

    # -------------------------
    def actualizar_progreso(self, actual, total, mensaje):
        self.barra["value"] = actual
        porcentaje = int((actual / total) * 100)
        self.label_progreso.config(text=f"{mensaje}  ({actual}/{total} — {porcentaje}%)")
        self.root.update_idletasks()


# =========================
# EJECUTAR
# =========================
if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()