from pathlib import Path
from playwright.sync_api import sync_playwright

EMAIL = "admtamara.b.ecuestre@hotmail.com"
PASSWORD = "HelloKitty1912"

DOWNLOAD_FOLDER = Path("excel")
DOWNLOAD_FOLDER.mkdir(exist_ok=True)

with sync_playwright() as p:
    # IMPORTANTE: headless=True para que corra invisible desde el programador de tareas sin bloquearse
    browser = p.chromium.launch(headless=True)
    context = browser.new_context(accept_downloads=True)
    page = context.new_page()

    print("Abriendo Contabilium...")
    page.goto("https://app.contabilium.com/v3/login?logOut=true", wait_until="commit", timeout=60000)

    campo_email = page.get_by_role("textbox", name="Email")
    campo_email.wait_for(state="visible", timeout=45000)
    campo_email.fill(EMAIL)
    
    page.get_by_role("textbox", name="Contraseña").fill(PASSWORD)
    page.get_by_role("button", name="Ingresar").click()

    print("Esperando ingreso al sistema...")
    page.wait_for_url("**/dashboard**", timeout=45000)
    page.wait_for_load_state("domcontentloaded")

    print("Navegando a Productos y Servicios...")
    page.goto("https://app.contabilium.com/conceptos.aspx", wait_until="commit", timeout=60000)
    
    # Espera adicional para asegurar que la tabla cargue los datos más recientes de la web
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(3000)

    print("Buscando boton de exportar...")
    page.locator("button.btn.btn-white").first.wait_for(timeout=45000)

    print("Desplegando opciones de exportacion...")
    page.locator("span.fa.fa-caret-down").evaluate(
        "element => element.parentElement.click()"
    )

    print("Seleccionando formato Simple...")
    opcion_simple = page.get_by_text("Simple", exact=True)
    opcion_simple.wait_for(state="visible", timeout=15000)
    opcion_simple.click()

    print("Esperando generacion del archivo y descargando...")
    link_descargar = page.get_by_role("link", name="Descargar")
    link_descargar.wait_for(state="visible", timeout=60000)

    # Pequeña pausa preventiva para que el servidor termine de empaquetar el reporte actualizado
    page.wait_for_timeout(2000)

    with page.expect_download(timeout=60000) as download_info:
        link_descargar.click()

    download = download_info.value
    archivo_destino = DOWNLOAD_FOLDER / "stock.xlsx"
    
    # Si ya existía un archivo previo, lo sobrescribe de forma limpia
    if archivo_destino.exists():
        archivo_destino.unlink()

    download.save_as(archivo_destino)

    print(f"Stock descargado y guardado con exito en: {archivo_destino}")
    browser.close()