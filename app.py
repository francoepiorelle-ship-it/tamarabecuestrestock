from pathlib import Path
import sqlite3
import pandas as pd
import streamlit as st
import datetime
from zoneinfo import ZoneInfo
import io
import base64

# Configuración inicial de la página
st.set_page_config(
    page_title="GestíonTamaraB - Control de Stock",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed"
)

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "inventario.db"
LOGO_PATH = BASE_DIR / "Diseño Sin Título - 2_2.jpg"
REMITOS_DIR = BASE_DIR / "remitos"
REMITOS_DIR.mkdir(exist_ok=True)
ETIQUETAS_DIR = BASE_DIR / "etiquetas"
ETIQUETAS_DIR.mkdir(exist_ok=True)

def asegurar_base_datos():
    conexion = sqlite3.connect(DB_PATH)
    cursor = conexion.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS productos (
            SKU TEXT UNIQUE,
            Descripción TEXT,
            Rubro TEXT,
            Subrubro TEXT,
            Proveedor TEXT,
            Stock REAL,
            "Stock Reservado" REAL,
            "Stock Disponible" REAL,
            "Última Actualización" TEXT
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS controles_fisicos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT,
            hora TEXT,
            responsable TEXT,
            sku TEXT,
            producto TEXT,
            stock_fisico REAL,
            stock_sistema REAL,
            diferencia REAL,
            fecha_registro TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS movimientos_stock (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT,
            hora TEXT,
            sku TEXT,
            producto TEXT,
            rubro TEXT,
            subrubro TEXT,
            proveedor TEXT,
            cantidad REAL,
            resp_conteo TEXT,
            resp_calidad TEXT,
            resp_remito TEXT,
            resp_etiquetado TEXT,
            resp_ubicacion TEXT,
            observacion TEXT,
            remito_archivo TEXT,
            fecha_registro TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    # Tabla borrador para evitar pérdida de datos por recargas o doble clic en Recepción
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS recepcion_borrador (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sku TEXT,
            producto TEXT,
            rubro TEXT,
            subrubro TEXT,
            cantidad REAL,
            observacion TEXT,
            precio_unitario REAL
        )
    """)
    
    cursor.execute("PRAGMA table_info(productos)")
    cols_prod = [col[1] for col in cursor.fetchall()]
    if "Rubro" not in cols_prod:
        cursor.execute("ALTER TABLE productos ADD COLUMN Rubro TEXT")
    if "Subrubro" not in cols_prod:
        cursor.execute("ALTER TABLE productos ADD COLUMN Subrubro TEXT")

    cursor.execute("PRAGMA table_info(movimientos_stock)")
    cols_mov = [col[1] for col in cursor.fetchall()]
    if "remito_archivo" not in cols_mov:
        cursor.execute("ALTER TABLE movimientos_stock ADD COLUMN remito_archivo TEXT")
    if "rubro" not in cols_mov:
        cursor.execute("ALTER TABLE movimientos_stock ADD COLUMN rubro TEXT")
    if "subrubro" not in cols_mov:
        cursor.execute("ALTER TABLE movimientos_stock ADD COLUMN subrubro TEXT")
    if "resp_conteo" not in cols_mov:
        cursor.execute("ALTER TABLE movimientos_stock ADD COLUMN resp_conteo TEXT")
    if "resp_calidad" not in cols_mov:
        cursor.execute("ALTER TABLE movimientos_stock ADD COLUMN resp_calidad TEXT")
    if "resp_remito" not in cols_mov:
        cursor.execute("ALTER TABLE movimientos_stock ADD COLUMN resp_remito TEXT")
    if "resp_etiquetado" not in cols_mov:
        cursor.execute("ALTER TABLE movimientos_stock ADD COLUMN resp_etiquetado TEXT")
    if "resp_ubicacion" not in cols_mov:
        cursor.execute("ALTER TABLE movimientos_stock ADD COLUMN resp_ubicacion TEXT")
        
    conexion.commit()
    conexion.close()

def sincronizar_excel_automatico():
    ruta_excel = BASE_DIR / "stock_actualizado.xlsx"
    if ruta_excel.exists():
        try:
            df = pd.read_excel(ruta_excel)
            df.columns = df.columns.str.strip()
            
            if "Sub Rubro" in df.columns and "Subrubro" not in df.columns:
                df.rename(columns={"Sub Rubro": "Subrubro"}, inplace=True)

            columnas_necesarias = ["SKU", "Nombre", "Stock", "Stock Reservado"]
            
            if all(col in df.columns for col in columnas_necesarias):
                df_filtrado = df.copy()
                
                if "Proveedor" not in df_filtrado.columns:
                    df_filtrado["Proveedor"] = None
                if "Rubro" not in df_filtrado.columns:
                    df_filtrado["Rubro"] = None
                if "Subrubro" not in df_filtrado.columns:
                    df_filtrado["Subrubro"] = None
                
                cols_finales = ["SKU", "Nombre", "Rubro", "Subrubro", "Proveedor", "Stock", "Stock Reservado"]
                df_filtrado = df_filtrado[[c for c in cols_finales if c in df_filtrado.columns]].copy()
                
                for col in ["Stock", "Stock Reservado"]:
                    if col in df_filtrado.columns:
                        if df_filtrado[col].dtype == object or pd.api.types.is_string_dtype(df_filtrado[col]):
                            df_filtrado[col] = (
                                df_filtrado[col]
                                .astype(str)
                                .str.replace(' ', '', regex=False)
                                .str.replace('.', '', regex=False)
                                .str.replace(',', '.', regex=False)
                            )
                            df_filtrado[col] = pd.to_numeric(df_filtrado[col], errors="coerce").fillna(0)
                
                df_filtrado.rename(columns={"Nombre": "Descripción"}, inplace=True)
                df_filtrado["Stock Disponible"] = df_filtrado["Stock"] - df_filtrado["Stock Reservado"]
                
                hora_arg = datetime.datetime.now(ZoneInfo("America/Argentina/Buenos_Aires"))
                df_filtrado["Última Actualización"] = hora_arg.strftime("%Y-%m-%d %H:%M:%S")

                asegurar_base_datos()
                conn = sqlite3.connect(DB_PATH)
                df_filtrado.to_sql("productos", conn, if_exists="replace", index=False)
                conn.close()
        except Exception as e:
            print(f"Error sincronizando stock automático: {e}")

def cargar_datos():
    sincronizar_excel_automatico()
    asegurar_base_datos()
    conexion = sqlite3.connect(DB_PATH)
    try:
        df = pd.read_sql('SELECT SKU, Descripción, Rubro, Subrubro, Proveedor, Stock, "Stock Reservado", "Stock Disponible", "Última Actualización" FROM productos', conexion)
    except Exception:
        df = pd.DataFrame(columns=["SKU", "Descripción", "Rubro", "Subrubro", "Proveedor", "Stock", "Stock Reservado", "Stock Disponible", "Última Actualización"])
    conexion.close()
    return df

def obtener_lista_nombres_fantasia():
    ruta_prov_excel = BASE_DIR / "proveedores.xlsx"
    nombres = []
    if ruta_prov_excel.exists():
        try:
            df_p = pd.read_excel(ruta_prov_excel)
            df_p.columns = df_p.columns.str.strip()
            if "Nombre Fantasia" in df_p.columns:
                nombres = df_p["Nombre Fantasia"].dropna().astype(str).str.strip().unique().tolist()
                nombres = [n for n in nombres if n and n.lower() != 'nan']
        except Exception as e:
            print(f"Error leyendo proveedores.xlsx: {e}")
    if not nombres:
        nombres = ["Sin Proveedor / General"]
    return sorted(nombres)

def cargar_borrador_recepcion():
    asegurar_base_datos()
    conexion = sqlite3.connect(DB_PATH)
    try:
        cursor = conexion.cursor()
        cursor.execute("SELECT sku, producto, rubro, subrubro, cantidad, observacion, precio_unitario FROM recepcion_borrador")
        filas = cursor.fetchall()
        items = []
        for f in filas:
            items.append({
                "SKU": f[0],
                "Producto": f[1],
                "Rubro": f[2],
                "Subrubro": f[3],
                "Cantidad": f[4],
                "Observación": f[5] if f[5] else "",
                "Precio Unitario": f[6] if f[6] is not None else 0.0
            })
    except Exception:
        items = []
    conexion.close()
    return items

def guardar_item_borrador(item):
    asegurar_base_datos()
    conexion = sqlite3.connect(DB_PATH)
    cursor = conexion.cursor()
    cursor.execute("""
        INSERT INTO recepcion_borrador (sku, producto, rubro, subrubro, cantidad, observacion, precio_unitario)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (item['SKU'], item['Producto'], item['Rubro'], item['Subrubro'], item['Cantidad'], item['Observación'], item.get('Precio Unitario', 0.0)))
    conexion.commit()
    conexion.close()

def actualizar_borrador_en_db(items):
    asegurar_base_datos()
    conexion = sqlite3.connect(DB_PATH)
    cursor = conexion.cursor()
    cursor.execute("DELETE FROM recepcion_borrador")
    for item in items:
        cursor.execute("""
            INSERT INTO recepcion_borrador (sku, producto, rubro, subrubro, cantidad, observacion, precio_unitario)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (item['SKU'], item['Producto'], item['Rubro'], item['Subrubro'], item['Cantidad'], item['Observación'], item.get('Precio Unitario', 0.0)))
    conexion.commit()
    conexion.close()

def vaciar_borrador_db():
    asegurar_base_datos()
    conexion = sqlite3.connect(DB_PATH)
    cursor = conexion.cursor()
    cursor.execute("DELETE FROM recepcion_borrador")
    conexion.commit()
    conexion.close()

def cargar_historial():
    conexion = sqlite3.connect(DB_PATH)
    try:
        df = pd.read_sql("SELECT id AS ID, fecha AS Fecha, hora AS Hora, responsable AS Responsable, sku AS SKU, producto AS Producto, stock_fisico AS 'Stock Físico', stock_sistema AS 'Stock Sistema', diferencia AS Diferencia FROM controles_fisicos ORDER BY fecha DESC, hora DESC, id DESC", conexion)
        if not df.empty and 'Stock Físico' in df.columns:
            df['Stock Físico'] = df['Stock Físico'].astype(int)
        if not df.empty and 'Stock Sistema' in df.columns:
            df['Stock Sistema'] = df['Stock Sistema'].astype(int)
        if not df.empty and 'Diferencia' in df.columns:
            df['Diferencia'] = df['Diferencia'].astype(int)
    except Exception:
        df = pd.DataFrame(columns=["ID", "Fecha", "Hora", "Responsable", "SKU", "Producto", "Stock Físico", "Stock Sistema", "Diferencia"])
    conexion.close()
    return df

def cargar_historial_movimientos():
    conexion = sqlite3.connect(DB_PATH)
    try:
        df = pd.read_sql("""
            SELECT id AS ID, fecha AS Fecha, hora AS Hora, sku AS SKU, producto AS Producto, 
                   rubro AS Rubro, subrubro AS Subrubro, proveedor AS Proveedor, cantidad AS Cantidad, 
                   resp_conteo AS 'Conteo Inicial', resp_calidad AS 'Control de Calidad', 
                   resp_remito AS 'Cotejo Remito', resp_etiquetado AS 'Etiquetado SKU', 
                   resp_ubicacion AS 'Ubicación Depósito', observacion AS Observación, remito_archivo AS Remito 
            FROM movimientos_stock ORDER BY fecha DESC, hora DESC, id DESC
        """, conexion)
        if not df.empty and 'Cantidad' in df.columns:
            df['Cantidad'] = df['Cantidad'].astype(int)
    except Exception:
        df = pd.DataFrame(columns=["ID", "Fecha", "Hora", "SKU", "Producto", "Rubro", "Subrubro", "Proveedor", "Cantidad", "Conteo Inicial", "Control de Calidad", "Cotejo Remito", "Etiquetado SKU", "Ubicación Depósito", "Observación", "Remito"])
    conexion.close()
    return df

def convertir_df_a_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Planilla')
    processed_data = output.getvalue()
    return processed_data

def convertir_multiples_tandas_a_excel(lista_de_dfs_con_meta):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for i, item in enumerate(lista_de_dfs_con_meta):
            nombre_hoja = f"Tanda_{i+1}_{item['fecha'].replace('-', '')}_{item['hora'].replace(':', '')[:4]}"
            nombre_hoja = "".join(c for c in nombre_hoja if c.isalnum() or c == '_')[:31]
            item['df'].to_excel(writer, index=False, sheet_name=nombre_hoja)
    return output.getvalue()

def convertir_multiples_tandas_a_html_impresion(lista_de_tandas_info, titulo_reporte):
    bloques_html = ""
    mes_actual = None
    dia_actual = None

    for idx, item in enumerate(lista_de_tandas_info):
        meta_data = item['meta_data']
        df = item['df']
        t_tanda = item['titulo_tanda']
        fecha_str = item.get('fecha', '')
        
        mes_str = fecha_str[:7] if len(fecha_str) >= 7 else "General"
        separador_jerarquico_html = ""
        
        if mes_str != mes_actual:
            mes_actual = mes_str
            dia_actual = None
            separador_jerarquico_html += f"""
            <div class="mes-separator">
                <h2>Mes: {mes_actual}</h2>
            </div>
            """
            
        if fecha_str != dia_actual:
            dia_actual = fecha_str
            separador_jerarquico_html += f"""
            <div class="dia-separator">
                <h3>Día: {dia_actual}</h3>
            </div>
            """

        meta_html = ""
        if meta_data:
            meta_html = f"""
            <div class="meta-box">
                <table class="meta-table">
                    <tr>
                        <td><b>Fecha:</b> {meta_data.get('fecha', '-')}</td>
                        <td><b>Proveedor (Nombre Fantasía):</b> {meta_data.get('proveedor', '-')}</td>
                    </tr>
                    <tr>
                        <td><b>1) Conteo inicial:</b> {meta_data.get('c1', '-')}</td>
                        <td><b>2) Control de calidad:</b> {meta_data.get('c2', '-')}</td>
                    </tr>
                    <tr>
                        <td><b>3) Cotejo con remito:</b> {meta_data.get('c3', '-')}</td>
                        <td><b>4) Etiquetado SKU:</b> {meta_data.get('c4', '-')}</td>
                    </tr>
                    <tr>
                        <td colspan="2"><b>5) Ubicación depósito:</b> {meta_data.get('c5', '-')}</td>
                    </tr>
                </table>
            </div>
            """

        remito_html_seccion = ""
        if meta_data and meta_data.get('remito_path'):
            ruta_r = Path(meta_data.get('remito_path'))
            if ruta_r.exists():
                ext = ruta_r.suffix.lower()
                if ext in ['.png', '.jpg', '.jpeg']:
                    with open(ruta_r, "rb") as img_file:
                        encoded_string = base64.b64encode(img_file.read()).decode('utf-8')
                    mime_type = "image/jpeg" if ext in ['.jpg', '.jpeg'] else "image/png"
                    remito_html_seccion = f"""
                    <div class="remito-box">
                        <b>Comprobante / Remito Adjunto de esta Tanda:</b><br><br>
                        <img src="data:{mime_type};base64,{encoded_string}" class="remito-img"/>
                    </div>
                    """
                else:
                    remito_html_seccion = f"""
                    <div class="remito-box">
                        <b>Comprobante / Remito Adjunto de esta Tanda:</b> {ruta_r.name} (Archivo PDF)
                    </div>
                    """

        page_break_style = "page-break-before: always;" if idx > 0 else ""
        bloques_html += f"""
        {separador_jerarquico_html}
        <div class="tanda-container" style="{page_break_style}">
            <h4 style="color: #00b89f; border-bottom: 2px solid #00b89f; padding-bottom: 4px; margin-top: 15px;">{t_tanda}</h4>
            {meta_html}
            {df.to_html(index=False, classes='data-table')}
            {remito_html_seccion}
        </div>
        """

    html = f"""
    <html>
        <head>
            <title>{titulo_reporte}</title>
            <style>
                @page {{ size: portrait; margin: 15mm; }}
                body {{ font-family: Arial, sans-serif; margin: 0; color: #0f172a; }}
                h2.main-title {{ text-align: center; color: #00b89f; margin-bottom: 5px; }}
                .mes-separator {{ background-color: #0f172a; color: #ffffff; padding: 10px 15px; border-radius: 6px; margin-top: 25px; margin-bottom: 10px; page-break-after: avoid; }}
                .mes-separator h2 {{ color: #ffffff; margin: 0; font-size: 18px; }}
                .dia-separator {{ background-color: #e2e8f0; color: #0f172a; padding: 8px 12px; border-radius: 4px; margin-top: 15px; margin-bottom: 10px; page-break-after: avoid; border-left: 4px solid #00b89f; }}
                .dia-separator h3 {{ margin: 0; font-size: 15px; }}
                .meta-box {{ border: 1px solid #cbd5e1; background-color: #f8fafc; padding: 10px; border-radius: 6px; margin-bottom: 12px; }}
                .meta-table {{ width: 100%; border-collapse: collapse; }}
                .meta-table td {{ padding: 4px 8px; font-size: 13px; border: none; }}
                .remito-box {{ margin-top: 12px; padding: 10px; border: 1px dashed #00b89f; background-color: #f0fdf4; font-size: 13px; border-radius: 6px; page-break-inside: avoid; text-align: center; }}
                .remito-img {{ max-width: 100%; height: auto; border: 1px solid #cbd5e1; border-radius: 4px; margin-top: 10px; }}
                table.data-table {{ width: 100%; border-collapse: collapse; margin-top: 8px; }}
                table.data-table th, table.data-table td {{ border: 1px solid #cbd5e1; padding: 6px 10px; text-align: left; font-size: 12px; }}
                table.data-table th {{ background-color: #f1f5f9; color: #0f172a; }}
                table.data-table tr:nth-child(even) {{ background-color: #f8fafc; }}
                .tanda-container {{ margin-bottom: 25px; }}
            </style>
        </head>
        <body>
            <h2 class="main-title">GestionTamaraB - {titulo_reporte}</h2>
            <p style="text-align: center; font-size: 12px; color: #64748b;"><b>Emitido:</b> {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
            {bloques_html}
            <script>
                window.onload = function() {{ window.print(); }}
            </script>
        </body>
    </html>
    """
    return html

def generar_html_remito_digital(proveedor, fecha, items_con_precios, total_general):
    filas_html = ""
    for item in items_con_precios:
        subtotal = item['Cantidad'] * item['Precio Unitario']
        filas_html += f"""
        <tr>
            <td>{item['SKU']}</td>
            <td>{item['Producto']}</td>
            <td style="text-align: center;">{int(item['Cantidad'])}</td>
            <td style="text-align: right;">$ {item['Precio Unitario']:,.2f}</td>
            <td style="text-align: right;">$ {subtotal:,.2f}</td>
        </tr>
        """
        
    html = f"""
    <html>
        <head>
            <title>Remito Digital - GestionTamaraB</title>
            <style>
                @page {{ size: portrait; margin: 20mm; }}
                body {{ font-family: Arial, sans-serif; margin: 0; color: #0f172a; }}
                .header-box {{ border-bottom: 2px solid #00b89f; padding-bottom: 10px; margin-bottom: 20px; display: flex; justify-content: space-between; }}
                .title {{ color: #00b89f; font-size: 24px; font-weight: bold; margin: 0; }}
                .subtitle {{ color: #64748b; font-size: 14px; margin: 5px 0 0 0; }}
                .info-box {{ background-color: #f8fafc; border: 1px solid #cbd5e1; padding: 12px; border-radius: 6px; margin-bottom: 20px; font-size: 14px; }}
                table.remito-table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
                table.remito-table th, table.remito-table td {{ border: 1px solid #cbd5e1; padding: 8px 12px; font-size: 13px; }}
                table.remito-table th {{ background-color: #f1f5f9; color: #0f172a; text-align: left; }}
                .total-box {{ text-align: right; margin-top: 20px; font-size: 18px; font-weight: bold; color: #0f172a; background-color: #f0fdf4; padding: 12px; border: 1px solid #00b89f; border-radius: 6px; }}
            </style>
        </head>
        <body>
            <div class="header-box">
                <div>
                    <h2 class="title">REMITO DIGITAL</h2>
                    <p class="subtitle">GestionTamaraB - Control de Stock</p>
                </div>
                <div style="text-align: right;">
                    <p style="margin: 0; font-size: 14px;"><b>Fecha:</b> {fecha}</p>
                    <p style="margin: 5px 0 0 0; font-size: 14px;"><b>Emisión:</b> {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}</p>
                </div>
            </div>
            
            <div class="info-box">
                <p style="margin: 0;"><b>Proveedor / Destino:</b> {proveedor}</p>
            </div>
            
            <table class="remito-table">
                <thead>
                    <tr>
                        <th>SKU</th>
                        <th>Descripción del Producto</th>
                        <th style="text-align: center;">Cantidad</th>
                        <th style="text-align: right;">Precio Unitario</th>
                        <th style="text-align: right;">Subtotal</th>
                    </tr>
                </thead>
                <tbody>
                    {filas_html}
                </tbody>
            </table>
            
            <div class="total-box">
                TOTAL GENERAL: $ {total_general:,.2f}
            </div>
            
            <script>
                window.onload = function() {{ window.print(); }}
            </script>
        </body>
    </html>
    """
    return html

# Inicialización de estados
if 'lista_control' not in st.session_state:
    st.session_state.lista_control = []

# Sincronizamos la tabla de recepción directamente desde el borrador persistente en SQLite
st.session_state.tabla_recepcion_items = cargar_borrador_recepcion()

if 'ultimo_movimiento_guardado' not in st.session_state:
    st.session_state.ultimo_movimiento_guardado = []

# --- SISTEMA DE LOGIN PERSISTENTE POR URL ---
query_params = st.query_params
if "auth" in query_params and query_params["auth"] == "ok":
    st.session_state['logueado'] = True

if 'logueado' not in st.session_state:
    st.session_state['logueado'] = False

if not st.session_state['logueado']:
    st.markdown("""
        <style>
            .stApp { background-color: #f8fafc; }
            .stButton > button { background-color: #00b89f !important; color: #ffffff !important; border: none !important; border-radius: 8px !important; font-weight: 600 !important; width: 100%; padding: 12px; }
            .stButton > button:hover { background-color: #009984 !important; }
        </style>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        if LOGO_PATH.exists():
            st.image(str(LOGO_PATH), width=None)
        st.markdown("<h2 style='text-align: center; color: #0f172a; font-weight: 700;'>GestionTamaraB</h2>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; color: #64748b;'>Panel de Control de Stock</p>", unsafe_allow_html=True)
        
        usuario_ingresado = st.text_input("Usuario")
        password_ingresada = st.text_input("Contraseña", type="password")
        
        if st.button("Iniciar Sesión"):
            try:
                usuarios_validos = st.secrets["passwords"]
                if usuario_ingresado in usuarios_validos and usuarios_validos[usuario_ingresado] == password_ingresada:
                    st.session_state['logueado'] = True
                    st.query_params["auth"] = "ok"
                    st.rerun()
                else:
                    st.error("Usuario o contraseña incorrectos.")
            except Exception as e:
                st.error("Configure los Secrets en Streamlit Cloud.")
                
        st.stop()

# --- CSS PROFESIONAL ---
st.markdown("""
    <style>
        .stApp { background-color: #f8fafc !important; }
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header[data-testid="stHeader"] { display: none !important; }
        [data-testid="stSidebar"] { display: none !important; }

        div[data-testid="metric-container"] { 
            background-color: #ffffff; 
            padding: 20px; 
            border-radius: 12px; 
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); 
            border: 1px solid #e2e8f0;
            border-top: 4px solid #00b89f;
        }
        div[data-testid="metric-container"] label { color: #64748b !important; font-size: 0.9rem !important; font-weight: 500 !important; }
        div[data-testid="metric-container"] [data-testid="stMetricValue"] { color: #0f172a !important; font-weight: 700 !important; font-size: 1.8rem !important; }

        .stButton > button { 
            background-color: #00b89f !important; 
            color: #ffffff !important; 
            border: none !important; 
            border-radius: 8px !important; 
            font-weight: 600 !important; 
            padding: 10px 20px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
            transition: all 0.2s;
        }
        .stButton > button:hover { background-color: #009984 !important; transform: translateY(-1px); }

        .stTabs [data-baseweb="tab-list"] {
            gap: 8px;
            background-color: #ffffff;
            padding: 10px 14px;
            border-radius: 12px;
            box-shadow: 0 2px 6px rgba(0,0,0,0.04);
            border: 1px solid #e2e8f0;
        }
        .stTabs [data-baseweb="tab"] {
            height: 44px;
            background-color: #f1f5f9;
            border-radius: 8px;
            padding: 0 24px;
            font-weight: 600;
            color: #475569;
            border: 1px solid #cbd5e1;
        }
        .stTabs [aria-selected="true"] {
            background-color: #00b89f !important;
            color: #ffffff !important;
            border-color: #009984 !important;
        }

        div[data-testid="stExpander"] {
            background-color: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 12px;
        }

        .stDataFrame {
            border: 1px solid #e2e8f0;
            border-radius: 12px;
            background-color: #ffffff;
            overflow: hidden;
        }

        .main h1, .main h2, .main h3, .main p, .main span, .main label { color: #0f172a !important; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif !important; }
    </style>
""", unsafe_allow_html=True)

df_productos = cargar_datos()
lista_nombres_fantasia = obtener_lista_nombres_fantasia()

# --- CABECERA SUPERIOR ---
col_head1, col_head2 = st.columns([5, 1])
with col_head1:
    st.title("📦 GestíonTamaraB - Panel de Control")
with col_head2:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚪 Cerrar Sesión"):
        st.session_state['logueado'] = False
        st.query_params.clear()
        st.rerun()

st.markdown("<br>", unsafe_allow_html=True)

tab_dash, tab_control, tab_movimientos, tab_historial = st.tabs([
    "📊 Dashboard General", 
    "📋 Control de Stock", 
    "📦 Recepción de Mercadería", 
    "📂 Historial y Auditorías"
])

# ==========================================
# 1. SOLAPA: DASHBOARD GENERAL
# ==========================================
with tab_dash:
    st.caption("ℹ️ Vista general del estado actual de la mercadería sincronizada.")
    st.markdown("<br>", unsafe_allow_html=True)

    if df_productos.empty:
        st.warning("⚠️ Esperando a que el sistema automatizado sincronice el primer archivo de stock en GitHub.")
    else:
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total de Productos", len(df_productos))
        with col2:
            st.metric("Unidades en Sistema", int(df_productos['Stock'].sum()) if 'Stock' in df_productos else 0)
        with col3:
            ultima_act = df_productos['Última Actualización'].max() if 'Última Actualización' in df_productos else "N/A"
            st.metric("Última Sincronización", str(ultima_act).split()[0] if isinstance(ultima_act, str) else "N/A")

        st.markdown("<br><br>", unsafe_allow_html=True)
        
        with st.expander("🔍 Ventana de Búsqueda y Filtros Avanzados"):
            f_col1, f_col2 = st.columns(2)
            with f_col1:
                busqueda = st.text_input("Filtrar por SKU o Descripción:", placeholder="Escribí aquí para buscar...")
            with f_col2:
                rubros_disponibles = ["Todos"] + sorted(df_productos['Rubro'].dropna().unique().tolist()) if 'Rubro' in df_productos.columns else ["Todos"]
                filtro_rubro_dash = st.selectbox("Filtrar por Rubro:", rubros_disponibles)

        df_filtrado = df_productos.copy()
        if 'busqueda' in locals() and busqueda:
            df_filtrado = df_filtrado[
                df_filtrado['SKU'].astype(str).str.contains(busqueda, case=False, na=False) |
                df_filtrado['Descripción'].astype(str).str.contains(busqueda, case=False, na=False)
            ]
        if 'filtro_rubro_dash' in locals() and filtro_rubro_dash != "Todos":
            df_filtrado = df_filtrado[df_filtrado['Rubro'] == filtro_rubro_dash]

        st.markdown("<br>", unsafe_allow_html=True)
        df_dash_mostrar = df_filtrado.drop(columns=['Proveedor'], errors='ignore')
        st.dataframe(df_dash_mostrar, use_container_width=True, hide_index=True)

# ==========================================
# 2. SOLAPA: CONTROL DE STOCK
# ==========================================
with tab_control:
    st.markdown("### Control de Stock")
    st.caption("Podés realizar varios conteos en el mismo día (ej. turno mañana y turno tarde) guardándolos de forma independiente.")
    st.markdown("<br>", unsafe_allow_html=True)
    
    col_fecha, col_resp = st.columns(2)
    with col_fecha:
        fecha_control = st.date_input("📅 Fecha del Conteo", datetime.date.today(), key="f_ctrl_dia")
    with col_resp:
        responsable = st.text_input("👤 Responsable del Conteo", placeholder="Ej: Tamara", key="resp_ctrl_dia")

    if not df_productos.empty:
        opciones_buscador = df_productos.apply(lambda x: f"{x['Descripción']} | SKU: {x['SKU']} | Rubro: {x.get('Rubro', 'N/A')}", axis=1).tolist()
        opciones_buscador.insert(0, "Seleccione un producto...")
    else:
        opciones_buscador = ["No hay productos disponibles"]

    st.markdown("<br>", unsafe_allow_html=True)
    with st.container():
        st.markdown("#### ➕ Agregar Producto a la Tanda Actual de Conteo")
        with st.form("form_conteo_dia", clear_on_submit=True):
            col1, col2 = st.columns([3, 1])
            with col1:
                producto_seleccionado = st.selectbox("Buscar Producto", opciones_buscador)
            with col2:
                stock_fisico_input = st.number_input("Stock Físico", min_value=0, step=1, format="%d")
                
            st.markdown("<br>", unsafe_allow_html=True)
            agregar_btn = st.form_submit_button("Agregar a la tanda de hoy")
            
            if agregar_btn and producto_seleccionado != "Seleccione un producto..." and producto_seleccionado != "No hay productos disponibles":
                partes = producto_seleccionado.split(" | SKU: ")
                nombre_prod = partes[0]
                sku_extraido = partes[1].split(" | Rubro: ")[0] if len(partes) > 1 else "Sin SKU"
                
                prod_info = df_productos[df_productos['SKU'].astype(str) == str(sku_extraido).strip()]
                stock_sis = int(prod_info.iloc[0]['Stock']) if not prod_info.empty else 0
                diferencia = int(stock_fisico_input) - stock_sis
                
                st.session_state.lista_control.append({
                    "SKU": sku_extraido,
                    "Producto": nombre_prod,
                    "Stock Físico": int(stock_fisico_input),
                    "Stock Sistema": stock_sis,
                    "Diferencia": diferencia
                })
                st.success(f"¡Agregado a la tanda! ({nombre_prod})")

    if st.session_state.lista_control:
        st.markdown("<br><hr><br>", unsafe_allow_html=True)
        st.markdown(f"### 📋 Tanda Actual de Conteo ({fecha_control})")
        df_control_actual = pd.DataFrame(st.session_state.lista_control)
        
        def pintar_diferencia(val):
            color = 'green' if val == 0 else 'red'
            return f'color: {color}; font-weight: bold;'
        
        st.dataframe(df_control_actual.style.map(pintar_diferencia, subset=['Diferencia']), use_container_width=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        excel_bytes_ctrl = convertir_df_a_excel(df_control_actual)
        st.download_button(
            label="📥 Descargar Tanda Actual en Excel (.xlsx)",
            data=excel_bytes_ctrl,
            file_name=f"control_stock_{fecha_control}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_excel_ctrl"
        )
        
        with st.expander("⚙️ Opciones de edición y corrección de la tanda actual"):
            col_err1, col_err2, col_err3 = st.columns([2, 1, 1])
            with col_err1:
                opciones_borrar = [f"{i} - {item['Producto']} (Stock Físico: {item['Stock Físico']})" for i, item in enumerate(st.session_state.lista_control)]
                item_a_borrar = st.selectbox("Elegí cuál borrar:", opciones_borrar, label_visibility="collapsed")
            with col_err2:
                if st.button("❌ Borrar ítem"):
                    if item_a_borrar:
                        indice = int(item_a_borrar.split(" - ")[0])
                        st.session_state.lista_control.pop(indice)
                        st.rerun()
            with col_err3:
                if st.button("🧹 Vaciar tanda"):
                    st.session_state.lista_control = []
                    st.rerun()
                
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("💾 Guardar Tanda en el Historial"):
            if not responsable:
                st.error("Por favor, ingresá el nombre del responsable antes de guardar.")
            else:
                hora_arg = datetime.datetime.now(ZoneInfo("America/Argentina/Buenos_Aires")).strftime("%H:%M:%S")
                conexion = sqlite3.connect(DB_PATH)
                cursor = conexion.cursor()
                for item in st.session_state.lista_control:
                    cursor.execute("""
                        INSERT INTO controles_fisicos (fecha, hora, responsable, sku, producto, stock_fisico, stock_sistema, diferencia)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """, (str(fecha_control), hora_arg, responsable, item['SKU'], item['Producto'], item['Stock Físico'], item['Stock Sistema'], item['Diferencia']))
                conexion.commit()
                conexion.close()
                
                st.session_state.lista_control = []
                st.success(f"¡Tanda de control guardada exitosamente a las {hora_arg}!")
                st.rerun()

# ==========================================
# 3. SOLAPA: RECEPCIÓN DE MERCADERÍA
# ==========================================
with tab_movimientos:
    st.markdown("### Recepción de Mercadería")
    st.caption("Completá obligatoriamente la fecha, el proveedor y todos los responsables del proceso antes de guardar. Todo lo que agregues se guarda automáticamente en la base de datos para evitar pérdidas por recargas o doble clic.")
    st.markdown("<br>", unsafe_allow_html=True)
    
    col_mfech, col_mprov = st.columns(2)
    with col_mfech:
        fecha_recepcion_dia = st.date_input("📅 Fecha de Recepción", datetime.date.today(), key="f_rec_dia")
    with col_mprov:
        proveedor_recepcion_global = st.selectbox("📦 Proveedor (Nombre Fantasía)", lista_nombres_fantasia, key="prov_rec_global")

    st.markdown("#### 👥 Responsables del Proceso *(Obligatorios)*")
    r_col1, r_col2, r_col3, r_col4, r_col5 = st.columns(5)
    with r_col1:
        resp_conteo = st.text_input("1) Conteo inicial *", placeholder="Nombre")
    with r_col2:
        resp_calidad = st.text_input("2) Control de calidad *", placeholder="Nombre")
    with r_col3:
        resp_remito = st.text_input("3) Cotejo con remito *", placeholder="Nombre")
    with r_col4:
        resp_etiquetado = st.text_input("4) Etiquetado SKU *", placeholder="Nombre")
    with r_col5:
        resp_ubicacion = st.text_input("5) Ubicación depósito *", placeholder="Nombre")

    archivo_remito_subido = st.file_uploader("📎 Adjuntar Remito General (Foto o PDF)", type=["png", "jpg", "jpeg", "pdf"], key="remito_masivo_subida")

    st.markdown("<br>", unsafe_allow_html=True)

    # --- SECCIÓN 1: BUSCADOR HABITUAL ---
    with st.container():
        st.markdown("#### 🔍 Buscador de Productos Existentes")
        
        if not df_productos.empty:
            opciones_skus_dict = {f"{row['Descripción']} | SKU: {row['SKU']} | Rubro: {row.get('Rubro', 'N/A')} | Subrubro: {row.get('Subrubro', 'N/A')}": row for _, row in df_productos.iterrows()}
            lista_opciones_prod = ["Seleccione un producto para agregar..."] + list(opciones_skus_dict.keys())
        else:
            lista_opciones_prod = ["No hay productos disponibles"]

        col_b_sel, col_b_btn = st.columns([3, 1])
        with col_b_sel:
            producto_individual_elegido = st.selectbox("Buscar Producto", lista_opciones_prod, label_visibility="collapsed")
        with col_b_btn:
            if st.button("➕ Agregar de la lista"):
                if producto_individual_elegido != "Seleccione un producto para agregar..." and producto_individual_elegido != "No hay productos disponibles":
                    datos_prod = opciones_skus_dict[producto_individual_elegido]
                    sku_val = str(datos_prod['SKU'])
                    desc_val = str(datos_prod['Descripción'])
                    rubro_val = str(datos_prod.get('Rubro', ''))
                    subrubro_val = str(datos_prod.get('Subrubro', ''))

                    items_actuales = cargar_borrador_recepcion()
                    ya_existe = any(item['SKU'] == sku_val for item in items_actuales)
                    
                    if not ya_existe:
                        nuevo_item = {
                            "SKU": sku_val,
                            "Producto": desc_val,
                            "Rubro": rubro_val,
                            "Subrubro": subrubro_val,
                            "Cantidad": 1,
                            "Observación": "",
                            "Precio Unitario": 0.0
                        }
                        guardar_item_borrador(nuevo_item)
                        st.success(f"¡Agregado: {desc_val}!")
                        st.rerun()
                    else:
                        st.warning("El producto ya se encuentra en la lista.")

    st.markdown("<br>", unsafe_allow_html=True)

    # --- SECCIÓN 2: AGREGAR PRODUCTO NUEVO (SOLO PARA LA PLANILLA) ---
    with st.expander("➕ Agregar Producto Nuevo (Solo para esta Planilla / Remito)"):
        st.caption("Usá esta opción para añadir mercadería nueva que no esté registrada en la base de datos principal.")
        with st.form("form_producto_nuevo_planilla", clear_on_submit=True):
            f_col_n1, f_col_n2 = st.columns(2)
            with f_col_n1:
                nuevo_sku = st.text_input("Código / SKU Nuevo", placeholder="Ej: NUEVO-001")
                nuevo_rubro = st.text_input("Rubro (Opcional)", placeholder="Ej: Varios")
            with f_col_n2:
                nuevo_desc = st.text_input("Descripción / Nombre del Producto *", placeholder="Ej: Producto Importado Nuevo")
                nuevo_subrubro = st.text_input("Subrubro (Opcional)", placeholder="Ej: General")
            
            f_col_n3, f_col_n4 = st.columns(2)
            with f_col_n3:
                nueva_cantidad = st.number_input("Cantidad", min_value=1, step=1, value=1)
            with f_col_n4:
                nuevo_precio_u = st.number_input("Precio Unitario ($)", min_value=0.0, step=100.0, format="%.2f")

            st.markdown("<br>", unsafe_allow_html=True)
            btn_add_nuevo = st.form_submit_button("Añadir producto nuevo a la planilla")

            if btn_add_nuevo:
                if not nuevo_desc.strip():
                    st.error("Por favor, ingresá al menos la descripción del producto.")
                else:
                    items_actuales = cargar_borrador_recepcion()
                    sku_final_nuevo = nuevo_sku.strip() if nuevo_sku.strip() else f"NUEVO-{len(items_actuales)+1}"
                    
                    nuevo_item = {
                        "SKU": sku_final_nuevo,
                        "Producto": nuevo_desc.strip(),
                        "Rubro": nuevo_rubro.strip() if nuevo_rubro else "Nuevo / Remito",
                        "Subrubro": nuevo_subrubro.strip() if nuevo_subrubro else "",
                        "Cantidad": int(nueva_cantidad),
                        "Observación": "",
                        "Precio Unitario": float(nuevo_precio_u)
                    }
                    guardar_item_borrador(nuevo_item)
                    st.success(f"¡Producto nuevo '{nuevo_desc.strip()}' agregado a la planilla correctamente!")
                    st.rerun()

    st.markdown("<br><hr><br>", unsafe_allow_html=True)
    st.markdown("#### 📋 Lista de Productos a Recibir (Control de Cantidades)")

    st.session_state.tabla_recepcion_items = cargar_borrador_recepcion()

    if not st.session_state.tabla_recepcion_items:
        st.info("ℹ️ La lista está vacía. Buscá productos existentes o agregá nuevos utilizando los paneles superiores.")
    else:
        indices_a_borrar = []
        
        for idx, item in enumerate(st.session_state.tabla_recepcion_items):
            with st.container():
                cols_item = st.columns([3, 1.5, 1.5, 0.8])
                with cols_item[0]:
                    st.markdown(f"<b>{item['Producto']}</b><br><span style='font-size:11px; color:#64748b;'>SKU: {item['SKU']} | {item['Rubro']}</span>", unsafe_allow_html=True)
                with cols_item[1]:
                    c_menos, c_cant, c_mas = st.columns([1, 1.5, 1])
                    with c_menos:
                        if st.button("➖", key=f"btn_menos_{idx}"):
                            if st.session_state.tabla_recepcion_items[idx]['Cantidad'] > 1:
                                st.session_state.tabla_recepcion_items[idx]['Cantidad'] -= 1
                                actualizar_borrador_en_db(st.session_state.tabla_recepcion_items)
                                st.rerun()
                    with c_cant:
                        st.markdown(f"<p style='text-align: center; font-weight: bold; margin-top: 5px;'>{int(item['Cantidad'])}</p>", unsafe_allow_html=True)
                    with c_mas:
                        if st.button("➕", key=f"btn_mas_{idx}"):
                            st.session_state.tabla_recepcion_items[idx]['Cantidad'] += 1
                            actualizar_borrador_en_db(st.session_state.tabla_recepcion_items)
                            st.rerun()
                with cols_item[2]:
                    obs_val = st.text_input("Obs", value=item['Observación'], placeholder="Observación...", key=f"obs_{idx}", label_visibility="collapsed")
                    if obs_val != item['Observación']:
                        st.session_state.tabla_recepcion_items[idx]['Observación'] = obs_val
                        actualizar_borrador_en_db(st.session_state.tabla_recepcion_items)
                with cols_item[3]:
                    if st.button("🗑️", key=f"del_item_{idx}"):
                        indices_a_borrar.append(idx)
                st.markdown("<hr style='margin: 5px 0; border-color: #f1f5f9;'>", unsafe_allow_html=True)

        if indices_a_borrar:
            for i in sorted(indices_a_borrar, reverse=True):
                st.session_state.tabla_recepcion_items.pop(i)
            actualizar_borrador_en_db(st.session_state.tabla_recepcion_items)
            st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)
        
        # --- SECCIÓN: REMITO DIGITAL CON PRECIOS ---
        with st.expander("🧾 Generación de Remito Digital (Precios y Totales)", expanded=True):
            st.caption("Verificá y ajustá el precio unitario de cada producto para calcular subtotales, el total general y generar el remito digital.")
            
            items_con_precios_temp = []
            total_remito_digital = 0.0
            
            for idx, item in enumerate(st.session_state.tabla_recepcion_items):
                col_p1, col_p2, col_p3 = st.columns([3, 1.5, 1.5])
                with col_p1:
                    st.markdown(f"<b>{item['Producto']}</b><br><span style='font-size:11px; color:#64748b;'>SKU: {item['SKU']} | Cant: {int(item['Cantidad'])}</span>", unsafe_allow_html=True)
                with col_p2:
                    precio_actual = float(item.get('Precio Unitario', 0.0))
                    precio_u = st.number_input("Precio Unitario ($)", min_value=0.0, step=100.0, value=precio_actual, format="%.2f", key=f"precio_u_{idx}")
                    if precio_u != precio_actual:
                        st.session_state.tabla_recepcion_items[idx]['Precio Unitario'] = precio_u
                        actualizar_borrador_en_db(st.session_state.tabla_recepcion_items)
                with col_p3:
                    subt = item['Cantidad'] * precio_u
                    total_remito_digital += subt
                    st.markdown(f"<p style='margin-top: 28px; font-weight: bold; color: #0f172a;'>Subtotal: $ {subt:,.2f}</p>", unsafe_allow_html=True)
                
                items_con_precios_temp.append({
                    "SKU": item['SKU'],
                    "Producto": item['Producto'],
                    "Cantidad": item['Cantidad'],
                    "Precio Unitario": precio_u
                })
                st.markdown("<hr style='margin: 5px 0; border-color: #f1f5f9;'>", unsafe_allow_html=True)
            
            st.markdown(f"<h3 style='text-align: right; color: #00b89f;'>TOTAL GENERAL: $ {total_remito_digital:,.2f}</h3>", unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)
            
            html_remito_generado = generar_html_remito_digital(proveedor_recepcion_global, str(fecha_recepcion_dia), items_con_precios_temp, total_remito_digital)
            st.download_button(
                label="🖨️ Descargar / Imprimir Remito Digital (PDF / HTML)",
                data=html_remito_generado,
                file_name=f"remito_digital_{proveedor_recepcion_global}_{fecha_recepcion_dia}.html",
                mime="text/html",
                key="btn_descargar_remito_digital"
            )

        st.markdown("<br>", unsafe_allow_html=True)
        
        df_para_excel_recepcion = pd.DataFrame(st.session_state.tabla_recepcion_items)
        excel_bytes_rec = convertir_df_a_excel(df_para_excel_recepcion)
        st.download_button(
            label="📥 Descargar Planilla de Recepción en Excel (.xlsx)",
            data=excel_bytes_rec,
            file_name=f"recepcion_{fecha_recepcion_dia}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_excel_recepcion"
        )
        st.markdown("<br>", unsafe_allow_html=True)

        col_acc1, col_acc2 = st.columns([2, 1])
        with col_acc1:
            if st.button("💾 Guardar Recepción Completa en el Historial"):
                if not resp_conteo or not resp_calidad or not resp_remito or not resp_etiquetado or not resp_ubicacion:
                    st.error("⚠️ No se puede continuar: Debes completar todos los casilleros de responsables del proceso (1 al 5).")
                else:
                    ruta_guardada = ""
                    if archivo_remito_subido is not None:
                        timestamp_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                        nombre_archivo_seguro = f"remito_{timestamp_str}_{archivo_remito_subido.name}"
                        ruta_completa = REMITOS_DIR / nombre_archivo_seguro
                        with open(ruta_completa, "wb") as f:
                            f.write(archivo_remito_subido.getbuffer())
                        ruta_guardada = str(ruta_completa)

                    hora_arg = datetime.datetime.now(ZoneInfo("America/Argentina/Buenos_Aires")).strftime("%H:%M:%S")
                    conexion = sqlite3.connect(DB_PATH)
                    cursor = conexion.cursor()
                    
                    lista_para_etiquetas = []

                    for item in st.session_state.tabla_recepcion_items:
                        cursor.execute("""
                            INSERT INTO movimientos_stock (fecha, hora, sku, producto, rubro, subrubro, proveedor, cantidad, resp_conteo, resp_calidad, resp_remito, resp_etiquetado, resp_ubicacion, observacion, remito_archivo)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, (str(fecha_recepcion_dia), hora_arg, item['SKU'], item['Producto'], item['Rubro'], item['Subrubro'], proveedor_recepcion_global, int(item['Cantidad']), resp_conteo, resp_calidad, resp_remito, resp_etiquetado, resp_ubicacion, item['Observación'], ruta_guardada))

                        lista_para_etiquetas.append({
                            "Producto": item['Producto'],
                            "SKU": item['SKU'],
                            "Cantidad": int(item['Cantidad'])
                        })

                    conexion.commit()
                    conexion.close()

                    vaciar_borrador_db()

                    st.session_state.ultimo_movimiento_guardado = lista_para_etiquetas
                    st.session_state.tabla_recepcion_items = []
                    st.success(f"¡Recepción completa guardada exitosamente a las {hora_arg}!")
                    st.rerun()
        with col_acc2:
            if st.button("🧹 Vaciar Lista"):
                vaciar_borrador_db()
                st.session_state.tabla_recepcion_items = []
                st.rerun()

    if st.session_state.ultimo_movimiento_guardado:
        st.markdown("<br><hr><br>", unsafe_allow_html=True)
        st.markdown("### 🏷️ Generar Archivo CSV para Etiquetas de esta Recepción")
        st.caption("Podés descargar el archivo con formato exacto (delimitado con punto y coma) listo para tu impresora de etiquetas.")
        
        tipo_cant_etiquetas = st.radio("Cantidad de etiquetas por producto:", ["Imprimir 1 etiqueta por ítem", "Imprimir tantas etiquetas como la cantidad recibida"], horizontal=True, key="radio_etiquetas_masivo")
        
        filas_etiquetas = []
        for item in st.session_state.ultimo_movimiento_guardado:
            repeticiones = int(item["Cantidad"]) if tipo_cant_etiquetas == "Imprimir tantas etiquetas como la cantidad recibida" else 1
            for _ in range(repeticiones):
                filas_etiquetas.append({
                    "Producto": item["Producto"],
                    "SKU": str(int(float(item["SKU"]))) if str(item["SKU"]).replace('.','',1).isdigit() else str(item["SKU"])
                })
        
        if filas_etiquetas:
            df_etiquetas = pd.DataFrame(filas_etiquetas)
            csv_buffer = df_etiquetas.to_csv(index=False, header=False, encoding="utf-8-sig", sep=";")
            
            timestamp_etiq = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            nombre_csv_etiquetas = f"etiquetas_{timestamp_etiq}.csv"
            
            st.download_button(
                label="📥 Descargar CSV para Etiquetas",
                data=csv_buffer,
                file_name=nombre_csv_etiquetas,
                mime="text/csv",
                key="btn_descarga_csv_masivo"
            )

# ==========================================
# 4. SOLAPA: HISTORIAL Y AUDITORÍAS
# ==========================================
with tab_historial:
    st.markdown("### Historial y Auditorías Avanzadas")
    st.caption("Seleccioná tandas específicas, por día completo o por mes completo. Los reportes separan limpiamente cada planilla por meses y días.")
    st.markdown("<br>", unsafe_allow_html=True)
    
    tipo_historial_seleccionado = st.radio("Seleccioná el tipo de historial a visualizar:", ["Control de Stock", "Recepción de Mercadería"], horizontal=True)
    st.markdown("<br>", unsafe_allow_html=True)

    if tipo_historial_seleccionado == "Control de Stock":
        st.markdown("#### 📋 Historial de Controles de Stock")
        df_historial = cargar_historial()
        
        if df_historial.empty:
            st.info("No hay controles de stock registrados todavía.")
        else:
            df_historial['Tanda_Label'] = df_historial.apply(lambda r: f"Fecha: {r['Fecha']} | Hora: {r['Hora']} | Resp: {r['Responsable']}", axis=1)
            
            modo_filtro_c = st.radio("Modo de selección:", ["Seleccionar por Tanda(s) específica(s)", "Seleccionar por Día completo", "Seleccionar por Mes completo"], horizontal=True, key="modo_c_stock")
            
            tandas_seleccionadas_f = []
            if modo_filtro_c == "Seleccionar por Tanda(s) específica(s)":
                tandas_fisicas_disponibles = list(df_historial['Tanda_Label'].unique())
                seleccion_checks = st.multiselect("🔍 Seleccioná una o varias tandas:", tandas_fisicas_disponibles)
                tandas_seleccionadas_f = seleccion_checks
            elif modo_filtro_c == "Seleccionar por Día completo":
                dias_disp = sorted(df_historial['Fecha'].unique().tolist(), reverse=True)
                dia_elegido = st.selectbox("📅 Seleccioná el Día:", dias_disp)
                if dia_elegido:
                    tandas_seleccionadas_f = df_historial[df_historial['Fecha'] == dia_elegido]['Tanda_Label'].unique().tolist()
            else:
                df_historial['Mes'] = df_historial['Fecha'].astype(str).str[:7]
                meses_disp = sorted(df_historial['Mes'].unique().tolist(), reverse=True)
                mes_elegido = st.selectbox("🗓️ Seleccioná el Mes (YYYY-MM):", meses_disp)
                if mes_elegido:
                    tandas_seleccionadas_f = df_historial[df_historial['Mes'] == mes_elegido]['Tanda_Label'].unique().tolist()

            if not tandas_seleccionadas_f:
                st.info("ℹ️ Seleccioná al menos una tanda, día o mes para ver las planillas.")
            else:
                df_hist_mostrar = df_historial[df_historial['Tanda_Label'].isin(tandas_seleccionadas_f)].sort_values(by=['Fecha', 'Hora'], ascending=[True, True])
                
                def pintar_historial(val):
                    if pd.isna(val): return ''
                    color = 'green' if val == 0 else 'red'
                    return f'color: {color}; font-weight: bold;'
                    
                df_final_hist_fisico = df_hist_mostrar.drop(columns=["ID", "Tanda_Label", "Mes"] if "Mes" in df_hist_mostrar.columns else ["ID", "Tanda_Label"])
                st.dataframe(df_final_hist_fisico.style.map(pintar_historial, subset=['Diferencia']), use_container_width=True, hide_index=True)
                
                st.markdown("<br>", unsafe_allow_html=True)
                
                lista_tandas_para_exportar = []
                for tanda_lbl in tandas_seleccionadas_f:
                    sub_df = df_historial[df_historial['Tanda_Label'] == tanda_lbl]
                    fila_m = sub_df.iloc[0]
                    lista_tandas_para_exportar.append({
                        "df": sub_df.drop(columns=["ID", "Tanda_Label", "Mes"] if "Mes" in sub_df.columns else ["ID", "Tanda_Label"]),
                        "fecha": fila_m['Fecha'],
                        "hora": fila_m['Hora'],
                        "meta_data": {
                            "fecha": f"{fila_m['Fecha']} | Hora: {fila_m['Hora']}",
                            "proveedor": "Control Físico Interno",
                            "c1": fila_m['Responsable'],
                            "c2": "-", "c3": "-", "c4": "-", "c5": "-"
                        },
                        "titulo_tanda": f"Control de Stock - {tanda_lbl}"
                    })
                
                lista_tandas_para_exportar = sorted(lista_tandas_para_exportar, key=lambda x: (x['fecha'], x['hora']))

                col_exp1, col_exp2 = st.columns(2)
                with col_exp1:
                    excel_bytes_multi = convertir_multiples_tandas_a_excel(lista_tandas_para_exportar)
                    st.download_button(
                        label="📥 Descargar Planillas Seleccionadas en Excel (.xlsx)",
                        data=excel_bytes_multi,
                        file_name=f"control_stock_seleccion_{datetime.date.today()}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_excel_hist_fisico_multi"
                    )
                with col_exp2:
                    html_impresion_multi = convertir_multiples_tandas_a_html_impresion(lista_tandas_para_exportar, "Reporte Consolidado - Control de Stock")
                    st.download_button(
                        label="🖨️ Imprimir / Descargar Reporte PDF con Separación de Meses y Días",
                        data=html_impresion_multi,
                        file_name=f"control_stock_reporte_{datetime.date.today()}.html",
                        mime="text/html",
                        key="download_pdf_hist_fisico_multi"
                    )

            st.markdown("<br>", unsafe_allow_html=True)
            with st.expander("🗑️ Zona de administración: Eliminar tandas o registros específicos"):
                opcion_eliminacion = st.radio("¿Qué deseas eliminar?", ["Un ítem específico", "Una tanda entera (por hora)", "Un día entero completo"], key="del_fisico")
                
                if opcion_eliminacion == "Un ítem específico":
                    opciones_borrar_historial = [f"{row['ID']} - [{row['Fecha']} {row['Hora']}] {row['Producto']}" for _, row in df_historial.iterrows()]
                    opciones_borrar_historial.insert(0, "Seleccione un registro...")
                    
                    col_del1, col_del2 = st.columns([3, 1])
                    with col_del1:
                        registro_a_borrar = st.selectbox("Registro:", opciones_borrar_historial, label_visibility="collapsed")
                    with col_del2:
                        if st.button("❌ Eliminar ítem"):
                            if registro_a_borrar != "Seleccione un registro...":
                                id_borrar = int(registro_a_borrar.split(" - ")[0])
                                conexion = sqlite3.connect(DB_PATH)
                                cursor = conexion.cursor()
                                cursor.execute("DELETE FROM controles_fisicos WHERE id = ?", (id_borrar,))
                                conexion.commit()
                                conexion.close()
                                st.success("¡Registro eliminado!")
                                st.rerun()
                elif opcion_eliminacion == "Una tanda entera (por hora)":
                    sesiones_disponibles = df_historial.apply(lambda r: f"Fecha: {r['Fecha']} - Hora: {r['Hora']} - Resp: {r['Responsable']}", axis=1).unique().tolist()
                    sesiones_disponibles.insert(0, "Seleccione una tanda...")
                    tanda_a_borrar = st.selectbox("Tanda:", sesiones_disponibles)
                    if st.button("🗑️ Eliminar Tanda Entera"):
                        if tanda_a_borrar != "Seleccione una tanda...":
                            partes = tanda_a_borrar.split(" - ")
                            f_val = partes[0].replace("Fecha: ", "")
                            h_val = partes[1].replace("Hora: ", "")
                            conexion = sqlite3.connect(DB_PATH)
                            cursor = conexion.cursor()
                            cursor.execute("DELETE FROM controles_fisicos WHERE fecha = ? AND hora = ?", (f_val, h_val))
                            conexion.commit()
                            conexion.close()
                            st.success("¡Tanda eliminada exitosamente!")
                            st.rerun()
                elif opcion_eliminacion == "Un día entero completo":
                    dias_disponibles = sorted(df_historial['Fecha'].unique().tolist(), reverse=True)
                    dias_disponibles.insert(0, "Seleccione un día...")
                    dia_a_borrar = st.selectbox("Día a eliminar:", dias_disponibles)
                    if st.button("💥 Eliminar Día Completo"):
                        if dia_a_borrar != "Seleccione un día...":
                            conexion = sqlite3.connect(DB_PATH)
                            cursor = conexion.cursor()
                            cursor.execute("DELETE FROM controles_fisicos WHERE fecha = ?", (dia_a_borrar,))
                            conexion.commit()
                            conexion.close()
                            st.success(f"¡Todos los controles del día {dia_a_borrar} fueron eliminados!")
                            st.rerun()

    else:
        st.markdown("#### 📦 Historial de Recepciones de Mercadería")
        df_mov = cargar_historial_movimientos()
        
        if df_mov.empty:
            st.info("No hay recepciones registradas todavía.")
        else:
            df_mov['Mov_Label'] = df_mov.apply(lambda r: f"Fecha: {r['Fecha']} | Hora: {r['Hora']} | Proveedor: {r['Proveedor']}", axis=1)
            
            modo_filtro_m = st.radio("Modo de selección:", ["Seleccionar por Recepción(es) específica(s)", "Seleccionar por Día completo", "Seleccionar por Mes completo"], horizontal=True, key="modo_m_stock")
            
            tandas_seleccionadas_m = []
            if modo_filtro_m == "Seleccionar por Recepción(es) específica(s)":
                movs_disponibles = list(df_mov['Mov_Label'].unique())
                seleccion_checks_m = st.multiselect("🔍 Seleccioná una o varias recepciones:", movs_disponibles)
                tandas_seleccionadas_m = seleccion_checks_m
            elif modo_filtro_m == "Seleccionar por Día completo":
                dias_disp_m = sorted(df_mov['Fecha'].unique().tolist(), reverse=True)
                dia_elegido_m = st.selectbox("📅 Seleccioná el Día:", dias_disp_m)
                if dia_elegido_m:
                    tandas_seleccionadas_m = df_mov[df_mov['Fecha'] == dia_elegido_m]['Mov_Label'].unique().tolist()
            else:
                df_mov['Mes'] = df_mov['Fecha'].astype(str).str[:7]
                meses_disp_m = sorted(df_mov['Mes'].unique().tolist(), reverse=True)
                mes_elegido_m = st.selectbox("🗓️ Seleccioná el Mes (YYYY-MM):", meses_disp_m)
                if mes_elegido_m:
                    tandas_seleccionadas_m = df_mov[df_mov['Mes'] == mes_elegido_m]['Mov_Label'].unique().tolist()

            if not tandas_seleccionadas_m:
                st.info("ℹ️ Seleccioná al menos una recepción, día o mes para ver las planillas.")
            else:
                df_mov_mostrar = df_mov[df_mov['Mov_Label'].isin(tandas_seleccionadas_m)].sort_values(by=['Fecha', 'Hora'], ascending=[True, True])

                columnas_a_quitar = ["ID", "Mov_Label", "Fecha", "Hora", "Conteo Inicial", "Control de Calidad", "Cotejo Remito", "Etiquetado SKU", "Ubicación Depósito", "Remito", "Mes"]
                df_final_hist_mov = df_mov_mostrar.drop(columns=[c for c in columnas_a_quitar if c in df_mov_mostrar.columns])
                
                st.dataframe(df_final_hist_mov, use_container_width=True, hide_index=True)

                lista_recepciones_para_exportar = []
                conexion = sqlite3.connect(DB_PATH)
                cursor = conexion.cursor()

                for tanda_lbl in tandas_seleccionadas_m:
                    sub_df = df_mov[df_mov['Mov_Label'] == tanda_lbl]
                    fila_m = sub_df.iloc[0]
                    
                    cursor.execute("SELECT remito_archivo FROM movimientos_stock WHERE fecha = ? AND hora = ? AND proveedor = ? LIMIT 1", (fila_m['Fecha'], fila_m['Hora'], fila_m['Proveedor']))
                    res_rem = cursor.fetchone()
                    ruta_rem_val = res_rem[0] if res_rem else None

                    cols_q_sub = ["ID", "Mov_Label", "Fecha", "Hora", "Conteo Inicial", "Control de Calidad", "Cotejo Remito", "Etiquetado SKU", "Ubicación Depósito", "Remito", "Mes"]
                    df_sub_limpio = sub_df.drop(columns=[c for c in cols_q_sub if c in sub_df.columns])

                    lista_recepciones_para_exportar.append({
                        "df": df_sub_limpio,
                        "fecha": fila_m['Fecha'],
                        "hora": fila_m['Hora'],
                        "meta_data": {
                            "fecha": f"{fila_m['Fecha']} | Hora: {fila_m['Hora']}",
                            "proveedor": fila_m['Proveedor'],
                            "c1": fila_m['Conteo Inicial'],
                            "c2": fila_m['Control de Calidad'],
                            "c3": fila_m['Cotejo Remito'],
                            "c4": fila_m['Etiquetado SKU'],
                            "c5": fila_m['Ubicación Depósito'],
                            "remito_path": ruta_rem_val
                        },
                        "titulo_tanda": f"Recepción - {tanda_lbl}"
                    })
                conexion.close()
                
                lista_recepciones_para_exportar = sorted(lista_recepciones_para_exportar, key=lambda x: (x['fecha'], x['hora']))

                st.markdown("<br>", unsafe_allow_html=True)
                
                col_mexp1, col_mexp2 = st.columns(2)
                with col_mexp1:
                    excel_bytes_multi_m = convertir_multiples_tandas_a_excel(lista_recepciones_para_exportar)
                    st.download_button(
                        label="📥 Descargar Planillas Seleccionadas en Excel (.xlsx)",
                        data=excel_bytes_multi_m,
                        file_name=f"recepciones_seleccion_{datetime.date.today()}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_excel_hist_mov_multi"
                    )
                with col_mexp2:
                    html_impresion_multi_m = convertir_multiples_tandas_a_html_impresion(lista_recepciones_para_exportar, "Reporte Consolidado - Recepción de Mercadería")
                    st.download_button(
                        label="🖨️ Imprimir / Descargar Reporte PDF con Separación de Meses y Días",
                        data=html_impresion_multi_m,
                        file_name=f"recepciones_reporte_{datetime.date.today()}.html",
                        mime="text/html",
                        key="download_pdf_hist_mov_multi"
                    )

            st.markdown("<br>", unsafe_allow_html=True)
            with st.expander("📎 Ver o descargar remitos históricos"):
                df_con_remito = df_mov[df_mov['Remito'].notna() & (df_mov['Remito'] != "")]
                if df_con_remito.empty:
                    st.info("No hay remitos adjuntos en los registros actuales.")
                else:
                    opciones_remitos = [f"Fecha: {r['Fecha']} | Hora: {r['Hora']} | Prov: {r['Proveedor']} | Archivo: {Path(r['Remito']).name}" for _, r in df_con_remito.iterrows()]
                    remito_elegido = st.selectbox("Seleccione el remito a ver:", opciones_remitos)
                    if remito_elegido:
                        idx_sel = opciones_remitos.index(remito_elegido)
                        ruta_archivo_remito = df_con_remito.iloc[idx_sel]['Remito']
                        if Path(ruta_archivo_remito).exists():
                            if Path(ruta_archivo_remito).suffix.lower() in ['.png', '.jpg', '.jpeg']:
                                st.image(ruta_archivo_remito, caption="Vista previa del remito", use_container_width=True)
                            with open(ruta_archivo_remito, "rb") as file_in:
                                st.download_button(
                                    label="📥 Descargar Archivo de Remito",
                                    data=file_in,
                                    file_name=Path(ruta_archivo_remito).name,
                                    mime="application/octet-stream"
                                )
                        else:
                            st.error("El archivo físico del remito ya no se encuentra en el servidor.")
