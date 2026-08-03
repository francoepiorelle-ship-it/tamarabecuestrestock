import os
import threading
from pathlib import Path
import sqlite3
import datetime
import subprocess
import pandas as pd
import customtkinter as ctk
from tkinter import messagebox
from playwright.sync_api import sync_playwright

# Configuración inicial de apariencia de CustomTkinter
ctk.set_appearance_mode("System")  # Options: "System", "Dark", "Light"
ctk.set_default_color_theme("blue")

EMAIL = "admtamara.b.ecuestre@hotmail.com"
PASSWORD = "HelloKitty1912"

# Ruta absoluta principal de tu proyecto
PROJECT_DIR = Path(r"C:\Users\franc\OneDrive\Escritorio\tamarabecuestrestock-main")
DOWNLOAD_FOLDER = PROJECT_DIR / "excel"
DOWNLOAD_FOLDER.mkdir(exist_ok=True)


class AppGestionTamara(ctk.CTk):

    def __init__(self):
        super().__init__()

        self.title("Gestión Tamara B - Control de Stock")
        self.geometry("500x450")
        self.resizable(False, False)

        # Título principal
        self.label_titulo = ctk.CTkLabel(
            self,
            text="Sistema de Gestión Tamara B",
            font=ctk.CTkFont(size=20, weight="bold"),
        )
        self.label_titulo.pack(pady=20)

        # Estado del proceso
        self.label_estado = ctk.CTkLabel(
            self,
            text="Estado: Esperando acción...",
            font=ctk.CTkFont(size=14),
            text_color="gray",
        )
        self.label_estado.pack(pady=10)

        # Cuadro de texto / consola interna para ver mensajes
        self.textbox_log = ctk.CTkTextbox(self, width=420, height=180)
        self.textbox_log.pack(pady=10)
        self.textbox_log.insert(
            "0.0", "Bienvenido. Haz clic en el botón para sincronizar stock.\n"
        )
        self.textbox_log.configure(state="disabled")

        # Botón para ejecutar la automatización
        self.btn_sincronizar = ctk.CTkButton(
            self,
            text="Sincronizar Stock desde Contabilium",
            command=self.iniciar_hilo_descarga,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#2fa572",
            hover_color="#248258",
            height=40,
        )
        self.btn_sincronizar.pack(pady=15)

    def log(self, mensaje):
        """Agrega texto al cuadro de logs de la interfaz de forma segura"""
        self.textbox_log.configure(state="normal")
        self.textbox_log.insert("end", mensaje + "\n")
        self.textbox_log.see("end")
        self.textbox_log.configure(state="disabled")

    def iniciar_hilo_descarga(self):
        """Ejecuta la automatización en un hilo secundario para que la ventana no se congele"""
        self.btn_sincronizar.configure(state="disabled")
        self.label_estado.configure(
            text="Estado: Sincronizando...", text_color="#1f538d"
        )
        hilo = threading.Thread(target=self.ejecutar_playwright)
        hilo.start()

    def ejecutar_playwright(self):
        try:
            self.log("Abriendo Contabilium...")
            with sync_playwright() as p:
                # Usamos Google Chrome del sistema para evitar errores al compilar
                browser = p.chromium.launch(
                    headless=True,
                    channel="chrome",
                    args=["--no-sandbox", "--disable-setuid-sandbox"],
                )
                context = browser.new_context(accept_downloads=True)
                page = context.new_page()

                page.goto(
                    "https://app.contabilium.com/v3/login?logOut=true",
                    wait_until="commit",
                    timeout=60000,
                )

                campo_email = page.get_by_role("textbox", name="Email")
                campo_email.wait_for(state="visible", timeout=45000)
                campo_email.fill(EMAIL)

                page.get_by_role("textbox", name="Contraseña").fill(PASSWORD)
                page.get_by_role("button", name="Ingresar").click()

                self.log("Esperando ingreso al sistema...")
                page.wait_for_url("**/dashboard**", timeout=45000)
                page.wait_for_load_state("domcontentloaded")

                self.log("Navegando a Productos y Servicios...")
                page.goto(
                    "https://app.contabilium.com/conceptos.aspx",
                    wait_until="commit",
                    timeout=60000,
                )
                page.wait_for_load_state("networkidle")
                page.wait_for_timeout(3000)

                self.log("Buscando boton de exportar...")
                page.locator("button.btn.btn-white").first.wait_for(
                    timeout=45000
                )

                self.log("Desplegando opciones de exportacion...")
                page.locator("span.fa.fa-caret-down").evaluate(
                    "element => element.parentElement.click()"
                )

                self.log("Seleccionando formato Simple...")
                opcion_simple = page.get_by_text("Simple", exact=True)
                opcion_simple.wait_for(state="visible", timeout=15000)
                opcion_simple.click()

                self.log("Esperando generacion y descarga del archivo...")
                link_descargar = page.get_by_role("link", name="Descargar")
                link_descargar.wait_for(state="visible", timeout=60000)
                page.wait_for_timeout(2000)

                with page.expect_download(timeout=60000) as download_info:
                    link_descargar.click()

                download = download_info.value
                archivo_destino = DOWNLOAD_FOLDER / "stock.xlsx"

                if archivo_destino.exists():
                    archivo_destino.unlink()

                download.save_as(archivo_destino)
                browser.close()

            self.log("¡Stock descargado con éxito!")

            # --- ACTUALIZACIÓN DE LA BASE DE DATOS LOCAL Y HORA ---
            self.log("Actualizando base de datos local y hora...")
            df_subido = pd.read_excel(archivo_destino)

            if not df_subido.empty:
                ahora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                for col in ["Última Actualización", "Ultima Actualizacion", "Fecha de Actualización"]:
                    if col in df_subido.columns:
                        df_subido[col] = ahora

                db_path = PROJECT_DIR / "database.db"
                conn = sqlite3.connect(db_path)
                df_subido.to_sql("stock", conn, if_exists="replace", index=False)
                conn.close()
                self.log("¡Base de datos local actualizada correctamente!")
            else:
                self.log("Aviso: El archivo descargado está vacío.")

            # --- SUBIDA AUTOMÁTICA A GITHUB ---
            self.log("Subiendo cambios a GitHub (actualizando la nube)...")
            os.chdir(PROJECT_DIR)

            # Ejecuta los comandos git para guardar y subir
            subprocess.run(["git", "add", "database.db"], check=True)
            subprocess.run(["git", "commit", "-m", f"Actualizacion automatica stock {ahora}"], check=False)
            resultado_push = subprocess.run(["git", "push"], capture_output=True, text=True)

            if resultado_push.returncode == 0:
                self.log("¡Cambios subidos a GitHub con éxito!")
            else:
                self.log("Aviso en Git push (verificar conexión o credenciales si es necesario).")

            self.label_estado.configure(
                text="Estado: Sincronización Completa", text_color="green"
            )
            messagebox.showinfo(
                "Éxito", "El stock se sincronizó, actualizó en la base de datos y se subió a GitHub correctamente."
            )

        except Exception as e:
            self.log(f"Error: {str(e)}")
            self.label_estado.configure(
                text="Estado: Error en la sincronización", text_color="red"
            )
            messagebox.showerror("Error", f"Ocurrió un error:\n{str(e)}")

        finally:
            self.btn_sincronizar.configure(state="normal")


if __name__ == "__main__":
    app = AppGestionTamara()
    app.mainloop()