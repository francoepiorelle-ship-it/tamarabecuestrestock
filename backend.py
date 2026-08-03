from fastapi import FastAPI, UploadFile, File
import pandas as pd
import sqlite3
import io
from datetime import datetime

app = FastAPI()

@app.post("/api/actualizar-stock")
async def actualizar_stock(file: UploadFile = File(...)):
    # 1. Recibir el archivo en memoria
    contenido = await file.read()
    
    try:
        # 2. Leer el Excel con Pandas
        df = pd.read_excel(io.BytesIO(contenido))
    except Exception as e:
        return {"error": f"Error al leer el archivo Excel: {str(e)}"}
    
    # 3. Filtrar las 4 columnas clave del Excel
    columnas_necesarias = ["SKU", "Nombre", "Stock", "Stock Reservado"]
    
    faltantes = [col for col in columnas_necesarias if col not in df.columns]
    if faltantes:
        return {"error": f"Faltan estas columnas en el Excel: {faltantes}"}
        
    df_filtrado = df[columnas_necesarias].copy()
    
    # 4. Limpiar y convertir stock a números para evitar errores operativos
    df_filtrado["Stock"] = pd.to_numeric(df_filtrado["Stock"], errors="coerce").fillna(0)
    df_filtrado["Stock Reservado"] = pd.to_numeric(df_filtrado["Stock Reservado"], errors="coerce").fillna(0)
    
    # 5. Renombrar "Nombre" a "Descripción" y calcular Stock Disponible
    df_filtrado.rename(columns={"Nombre": "Descripción"}, inplace=True)
    df_filtrado["Stock Disponible"] = df_filtrado["Stock"] - df_filtrado["Stock Reservado"]
    df_filtrado["Última Actualización"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # 6. Guardar en la Base de Datos SQLite
    try:
        conn = sqlite3.connect("inventario.db")
        df_filtrado.to_sql("productos", conn, if_exists="replace", index=False)
        conn.close()
    except Exception as e:
        return {"error": f"Error al guardar en la base de datos: {str(e)}"}
    
    return {
        "status": "success",
        "mensaje": "Stock procesado y actualizado correctamente en la base de datos.",
        "total_productos_actualizados": len(df_filtrado)
    }

@app.get("/api/productos")
def obtener_productos():
    try:
        conn = sqlite3.connect("inventario.db")
        df = pd.read_sql("SELECT * FROM productos", conn)
        conn.close()
        return df.to_dict(orient="records")
    except Exception as e:
        return {"error": f"Error al leer la base de datos: {str(e)}"}
