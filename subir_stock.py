import os
import subprocess
from pathlib import Path

CARPETA_EXCEL = Path(r"C:\Users\franc\OneDrive\Escritorio\StockControl\dist\GestionStock\excel")
REPODIR = Path(r"C:\Users\franc\OneDrive\Escritorio\StockControl")
EXCEL_DESTINO = REPODIR / "stock.xlsx"

def subir_stock_automatico():
    archivos_excel = list(CARPETA_EXCEL.glob("*.xlsx"))
    
    if not archivos_excel:
        print("⚠️ No se encontraron archivos Excel en la carpeta especificada.")
        return

    archivo_reciente = max(archivos_excel, key=os.path.getmtime)
    print(f"📄 Archivo detectado: {archivo_reciente.name}")

    EXCEL_DESTINO.write_bytes(archivo_reciente.read_bytes())
    print("📁 Archivo copiado al repositorio local.")

    try:
        os.chdir(REPODIR)
        subprocess.run(["git", "add", "stock.xlsx"], check=True)
        
        # Intenta hacer commit, si no hay cambios reales no rompe la ejecución
        resultado_commit = subprocess.run(
            ["git", "commit", "-m", f"Actualización automática de stock: {archivo_reciente.name}"],
            capture_output=True, text=True
        )
        
        if "nothing to commit" in resultado_commit.stdout or "no changes added to commit" in resultado_commit.stdout:
            print("ℹ️ El archivo de stock no tiene modificaciones respecto al anterior. No se requiere nuevo envío.")
            return

        subprocess.run(["git", "push"], check=True)
        print("🚀 ¡Stock actualizado y subido a GitHub con éxito!")
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Error al ejecutar los comandos de Git: {e}")

if __name__ == "__main__":
    subir_stock_automatico()