import sqlite3
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "database.db"

print(f"Buscando base de datos en: {DB_PATH}")
print(f"¿Existe el archivo?: {DB_PATH.exists()}")

if DB_PATH.exists():
    conexion = sqlite3.connect(DB_PATH)
    cursor = conexion.cursor()
    try:
        cursor.execute("SELECT COUNT(*) FROM productos")
        cantidad = cursor.fetchone()[0]
        print(f"¡Cantidad de productos encontrados en la base de datos: {cantidad}!")
    except Exception as e:
        print(f"La tabla 'productos' no existe o dio error: {e}")
    conexion.close()