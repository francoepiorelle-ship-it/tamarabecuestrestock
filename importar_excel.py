from datetime import datetime
from pathlib import Path
import sqlite3
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR / "excel" / "stock.xlsx"
DB_PATH = BASE_DIR / "database.db"

def inicializar_base_datos():
    conexion = sqlite3.connect(DB_PATH)
    cursor = conexion.cursor()
    
    # Creamos la tabla si no existe
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS productos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT UNIQUE,
            descripcion TEXT,
            stock REAL,
            stock_reservado REAL,
            stock_disponible REAL,
            ultima_actualizacion TEXT
        )
    """)
    
    # Por si la tabla ya existía de antes sin estas columnas, intentamos agregarlas de forma segura
    try:
        cursor.execute("ALTER TABLE productos ADD COLUMN stock_reservado REAL")
    except sqlite3.OperationalError:
        pass # La columna ya existe
        
    try:
        cursor.execute("ALTER TABLE productos ADD COLUMN stock_disponible REAL")
    except sqlite3.OperationalError:
        pass # La columna ya existe

    conexion.commit()
    conexion.close()

def importar_stock():
    if not EXCEL_PATH.exists():
        print(f"No se encontró el archivo en {EXCEL_PATH}")
        return

    print("Leyendo archivo Excel...")
    df = pd.read_excel(EXCEL_PATH)
    
    inicializar_base_datos()
    
    conexion = sqlite3.connect(DB_PATH)
    cursor = conexion.cursor()
    
    ahora_local = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    actualizados = 0
    nuevos = 0
    
    for _, row in df.iterrows():
        codigo = str(row.get("SKU", row.get("Código", ""))).strip()
        
        nombre_raw = row.get("Nombre", row.get("Descripción", ""))
        descripcion = str(nombre_raw).strip() if pd.notna(nombre_raw) else ""
        
        def limpiar_stock(valor):
            try:
                if pd.isna(valor):
                    return 0.0
                if isinstance(valor, str):
                    valor = valor.replace(".", "").replace(",", ".")
                return float(valor)
            except (ValueError, TypeError):
                return 0.0

        stock = limpiar_stock(row.get("Stock", 0))
        stock_reservado = limpiar_stock(row.get("Stock Reservado", 0))
        stock_disponible = limpiar_stock(row.get("Stock Disponible", 0))
        
        if not codigo or codigo.lower() == "nan" or codigo == "":
            continue

        cursor.execute("SELECT id FROM productos WHERE codigo = ?", (codigo,))
        resultado = cursor.fetchone()
        
        if resultado:
            cursor.execute("""
                UPDATE productos 
                SET descripcion = ?, stock = ?, stock_reservado = ?, stock_disponible = ?, ultima_actualizacion = ? 
                WHERE codigo = ?
            """, (descripcion, stock, stock_reservado, stock_disponible, ahora_local, codigo))
            actualizados += 1
        else:
            cursor.execute("""
                INSERT INTO productos (codigo, descripcion, stock, stock_reservado, stock_disponible, ultima_actualizacion) 
                VALUES (?, ?, ?, ?, ?, ?)
            """, (descripcion, stock, stock_reservado, stock_disponible, ahora_local))
            nuevos += 1
            
    conexion.commit()
    conexion.close()
    
    print(f"Importacion completada: {nuevos} nuevos, {actualizados} actualizados a las {ahora_local}.")

if __name__ == "__main__":
    importar_stock()