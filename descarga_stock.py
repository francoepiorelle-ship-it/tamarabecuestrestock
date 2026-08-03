import os
from pathlib import Path
import sqlite3
import datetime
import subprocess
import pandas as pd
from playwright.sync_api import sync_playwright

EMAIL = "admtamara.b.ecuestre@hotmail.com"
PASSWORD = "HelloKitty1912"

# Ruta absoluta principal de tu proyecto
PROJECT_DIR = Path(r"C:\Users\franc\OneDrive\Escritorio\tamarabecuestrestock-main")
DOWNLOAD_FOLDER = PROJECT_DIR / "excel"
DOWNLOAD_FOLDER.mkdir(exist_ok=True)


def log(mensaje):
    """Muestra los mensajes directamente en la consola con marca de tiempo"""
    timestamp = datetime.datetime.now().strftime("%H:%M:%S")
    print(f"[{timestamp}] {mensaje}")


def sincronizar_stock():
    try:
        log("🔍 Abriendo Contabilium mediante Playwright...")
        with sync_playwright() as p:
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

            log("⏳ Esperando ingreso al sistema...")
            page.wait_for_url("**/dashboard**", timeout=45000)
            page.wait_for_load_state("domcontentloaded")

            log("📂 Navegando a Productos y Servicios...")
            page.goto(
                "https://app.contabilium.com/conceptos.aspx",
                wait_until="commit",
                timeout=60000,
            )
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(3000)

            log("🔘 Buscando botón de exportar...")
            page.locator("button.btn.btn-white").first.wait_for(timeout=45000)

            log("📋 Desplegando opciones de exportación...")
            page.locator("span.fa.fa-caret-down").evaluate(
                "element => element.parentElement.click()"
            )

            log("📥 Seleccionando formato Simple...")
            opcion_simple = page.get_by_text("Simple", exact=True)
            opcion_simple.wait_for(state="visible", timeout=15000)
            opcion_simple.click()

            log("⏳ Esperando generación y descarga del archivo...")
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

        log("✅ ¡Stock descargado con éxito!")

        # --- ACTUALIZACIÓN DE LA BASE DE DATOS LOCAL Y HORA ---
        log("🔄 Actualizando base de datos local y fecha...")
        df_subido = pd.read_excel(archivo_destino)

        ahora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        if not df_subido.empty:
            for col in ["Última Actualización", "Ultima Actualizacion", "Fecha de Actualización"]:
                if col in df_subido.columns:
                    df_subido[col] = ahora

            db_path = PROJECT_DIR / "database.db"
            conn = sqlite3.connect(db_path)
            df_subido.to_sql("stock", conn, if_exists="replace", index=False)
            conn.close()
            log("✅ ¡Base de datos local actualizada correctamente!")
        else:
            log("⚠️ Aviso: El archivo descargado está vacío.")

        # --- SUBIDA AUTOMÁTICA A GITHUB ---
        log("🚀 Subiendo cambios automáticamente a GitHub...")
        os.chdir(PROJECT_DIR)

        subprocess.run(["git", "add", "database.db"], check=True)
        subprocess.run(["git", "commit", "-m", f"Actualizacion automatica stock {ahora}"], check=False)
        
        resultado_push = subprocess.run(["git", "push", "origin", "main"], capture_output=True, text=True)

        if resultado_push.returncode == 0:
            log("🎉 ¡Base de datos subida a GitHub con éxito!")
        else:
            log("⚠️ Aviso: No se pudo hacer push automáticamente (verificar conexión).")

        log("✨ ¡Proceso de sincronización completado por completo!")

    except Exception as e:
        log(f"❌ Ocurrió un error crítico: {str(e)}")


if __name__ == "__main__":
    sincronizar_stock()