import os
import shutil
import subprocess
from pathlib import Path
import datetime
import sqlite3
import pandas as pd
from playwright.sync_api import sync_playwright

# --- CONFIGURACIÓN ---
EMAIL = "admtamara.b.ecuestre@hotmail.com"
PASSWORD = "HelloKitty1912"

# Rutas del proyecto
CARPETA_REPO_GITHUB = Path(r"C:\Users\tamar\OneDrive\Escritorio\StockControl")
CARPETA_ORIGEN_EXCEL = CARPETA_REPO_GITHUB / "excel"
CARPETA_ORIGEN_EXCEL.mkdir(exist_ok=True)
NOMBRE_ARCHIVO_DESTINO = "stock_actualizado.xlsx"


def log(mensaje):
    """Muestra los mensajes en la consola con marca de tiempo"""
    timestamp = datetime.datetime.now().strftime("%H:%M:%S")
    print(f"[{timestamp}] {mensaje}")


def ejecutar_sincronizacion_completa():
    try:
        # ==========================================
        # PASO 1: Descargar stock desde Contabilium
        # ==========================================
        log("🌐 Abriendo Contabilium mediante Playwright...")
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

            log("📦 Navegando a Productos y Servicios...")
            page.goto(
                "https://app.contabilium.com/conceptos.aspx",
                wait_until="commit",
                timeout=60000,
            )
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(3000)

            log("🔍 Buscando botón de exportar...")
            page.locator("button.btn.btn-white").first.wait_for(timeout=45000)

            log("📥 Desplegando opciones de exportación...")
            page.locator("span.fa.fa-caret-down").evaluate(
                "element => element.parentElement.click()"
            )

            log("📋 Seleccionando formato Simple...")
            opcion_simple = page.get_by_text("Simple", exact=True)
            opcion_simple.wait_for(state="visible", timeout=15000)
            opcion_simple.click()

            log("⬇️ Esperando generación del archivo y descargando...")
            link_descargar = page.get_by_role("link", name="Descargar")
            link_descargar.wait_for(state="visible", timeout=60000)
            page.wait_for_timeout(2000)

            with page.expect_download(timeout=60000) as download_info:
                link_descargar.click()

            download = download_info.value
            archivo_temporal = CARPETA_ORIGEN_EXCEL / "stock.xlsx"
            
            if archivo_temporal.exists():
                archivo_temporal.unlink()

            download.save_as(archivo_temporal)
            browser.close()

        log(f"✅ ¡Stock descargado con éxito en: {archivo_temporal}!")

   # ==========================================
        # PASO 2: Actualizar Base de Datos local
        # ==========================================
        log("🔄 Actualizando base de datos local...")
        df_subido = pd.read_excel(archivo_temporal)
        ahora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        if not df_subido.empty:
            for col in ["Última Actualización", "Ultima Actualizacion", "Fecha de Actualización"]:
                if col in df_subido.columns:
                    df_subido[col] = ahora

            db_path = CARPETA_REPO_GITHUB / "database.db"
            
            # 1. Rescatar los subrubros anteriores de forma segura
            subrubros_guardados = {}
            if db_path.exists():
                try:
                    conn_temp = sqlite3.connect(db_path)
                    cursor_temp = conn_temp.cursor()
                    # Verificar si la tabla stock existe y si tiene la columna Subrubro
                    cursor_temp.execute("PRAGMA table_info(stock)")
                    columnas_existentes = [col[1] for col in cursor_temp.fetchall()]
                    
                    if "Subrubro" in columnas_existentes and "SKU" in columnas_existentes:
                        query_check = "SELECT SKU, Subrubro FROM stock"
                        df_antiguo = pd.read_sql(query_check, conn_temp)
                        for _, row in df_antiguo.iterrows():
                            sku = row["SKU"]
                            subrubro = row["Subrubro"]
                            if pd.notna(subrubro) and str(subrubro).lower() != "none":
                                subrubros_guardados[sku] = subrubro
                        log(f"📋 Se recuperaron {len(subrubros_guardados)} subrubros de la base anterior.")
                    else:
                        log("ℹ️ La tabla anterior no tenía la columna Subrubro todavía.")
                    conn_temp.close()
                except Exception as e:
                    log(f"⚠️ No se pudo leer el subrubro anterior: {e}")

            # 2. Rellenar la columna Subrubro si viene vacía de Contabilium
            if "Subrubro" in df_subido.columns and subrubros_guardados:
                df_subido["Subrubro"] = df_subido.apply(
                    lambda row: subrubros_guardados.get(row["SKU"], row["Subrubro"]) 
                    if pd.isna(row["Subrubro"]) or str(row["Subrubro"]).lower() == "none" 
                    else row["Subrubro"], 
                    axis=1
                )

            # 3. Guardar en la base de datos
            conn = sqlite3.connect(db_path)
            df_subido.to_sql("stock", conn, if_exists="replace", index=False)
            conn.close()
            log("✅ ¡Base de datos local actualizada correctamente con sus subrubros!")
        # ==========================================
        # PASO 3: Copiar archivo y subir a GitHub
        # ==========================================
        ruta_destino = CARPETA_REPO_GITHUB / NOMBRE_ARCHIVO_DESTINO
        shutil.copy(archivo_temporal, ruta_destino)
        log(f"📋 Archivo copiado al repositorio local en: {ruta_destino}")

        os.chdir(CARPETA_REPO_GITHUB)
        
        log("🌿 Asegurando rama principal (main)...")
        subprocess.run(["git", "checkout", "main"], capture_output=True, text=True)
        
        log("➕ Añadiendo archivos a Git...")
        subprocess.run(["git", "add", NOMBRE_ARCHIVO_DESTINO, "database.db"], check=True)
        
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        mensaje_commit = f"Actualización automática de stock: {timestamp}"
        
        resultado_commit = subprocess.run(["git", "commit", "-m", mensaje_commit], capture_output=True, text=True)
        
        if "nothing to commit" in resultado_commit.stdout or "no hay nada para confirmar" in resultado_commit.stdout:
            log("ℹ️ El archivo de stock no sufrió cambios respecto a la última vez. No es necesario subirlo.")
            return
        
        # Sincronizar cambios remotos antes de hacer push para evitar rechazos
        log("📥 Sincronizando cambios remotos (git pull)...")
        subprocess.run(["git", "pull", "origin", "main", "--rebase"], capture_output=True, text=True)

        log("🚀 Subiendo cambios a GitHub (rama main)...")
        subprocess.run(["git", "push", "origin", "main"], check=True)
        
        log("🎉 ¡Sincronización completa con GitHub completada con éxito!")

    except subprocess.CalledProcessError as e:
        log(f"❌ Error al ejecutar comandos de Git: {e}")
    except Exception as e:
        log(f"❌ Ocurrió un error inesperado: {e}")


if __name__ == "__main__":
    ejecutar_sincronizacion_completa()