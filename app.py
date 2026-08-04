from pathlib import Path
import sqlite3
import pandas as pd
import streamlit as st
import datetime
from zoneinfo import ZoneInfo

# Configuración inicial de la página
st.set_page_config(
    page_title="GestionTamaraB - Control de Stock",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed"
)

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "inventario.db"
LOGO_PATH = BASE_DIR / "Diseño Sin Título - 2_2.jpg"

def asegurar_base_datos():
    conexion = sqlite3.connect(DB_PATH)
    cursor = conexion.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS productos (
            SKU TEXT UNIQUE,
            Descripción TEXT,
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
            responsable TEXT,
            sku TEXT,
            producto TEXT,
            talle TEXT,
            color TEXT,
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
            tipo TEXT,
            sku TEXT,
            producto TEXT,
            talle TEXT,
            color TEXT,
            cantidad REAL,
            responsable TEXT,
            observacion TEXT,
            fecha_registro TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conexion.commit()
    conexion.close()

def sincronizar_excel_automatico():
    """Lee el Excel actualizado que sube automáticamente GitHub y refresca la base SQLite"""
    ruta_excel = BASE_DIR / "stock_actualizado.xlsx"
    if ruta_excel.exists():
        try:
            df = pd.read_excel(ruta_excel)
            columnas_necesarias = ["SKU", "Nombre", "Stock", "Stock Reservado"]
            
            if all(col in df.columns for col in columnas_necesarias):
                df_filtrado = df[columnas_necesarias].copy()
                
                for col in ["Stock", "Stock Reservado"]:
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
        df = pd.read_sql("SELECT SKU, Descripción, Stock, \"Stock Reservado\", \"Stock Disponible\", \"Última Actualización\" FROM productos", conexion)
    except Exception:
        df = pd.DataFrame(columns=["SKU", "Descripción", "Stock", "Stock Reservado", "Stock Disponible", "Última Actualización"])
    conexion.close()
    return df

def cargar_historial():
    conexion = sqlite3.connect(DB_PATH)
    try:
        df = pd.read_sql("SELECT id AS ID, fecha AS Fecha, responsable AS Responsable, sku AS SKU, producto AS Producto, talle AS Talle, color AS Color, stock_fisico AS 'Stock Físico', stock_sistema AS 'Stock Sistema', diferencia AS Diferencia FROM controles_fisicos ORDER BY fecha DESC, id DESC", conexion)
    except Exception:
        df = pd.DataFrame(columns=["ID", "Fecha", "Responsable", "SKU", "Producto", "Talle", "Color", "Stock Físico", "Stock Sistema", "Diferencia"])
    conexion.close()
    return df

def cargar_historial_movimientos():
    conexion = sqlite3.connect(DB_PATH)
    try:
        df = pd.read_sql("SELECT id AS ID, fecha AS Fecha, tipo AS Tipo, sku AS SKU, producto AS Producto, talle AS Talle, color AS Color, cantidad AS Cantidad, responsable AS Responsable, observacion AS Observación FROM movimientos_stock ORDER BY fecha DESC, id DESC", conexion)
    except Exception:
        df = pd.DataFrame(columns=["ID", "Fecha", "Tipo", "SKU", "Producto", "Talle", "Color", "Cantidad", "Responsable", "Observación"])
    conexion.close()
    return df

if 'lista_control' not in st.session_state:
    st.session_state.lista_control = []

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

# --- CABECERA SUPERIOR ---
col_head1, col_head2 = st.columns([5, 1])
with col_head1:
    st.title("📦 GestionTamaraB - Panel de Control")
with col_head2:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚪 Cerrar Sesión"):
        st.session_state['logueado'] = False
        st.query_params.clear()
        st.rerun()

st.markdown("<br>", unsafe_allow_html=True)

tab_dash, tab_control, tab_historial, tab_movimientos = st.tabs(["📊 Dashboard General", "📋 Control Físico por Día", "📂 Historial de Auditorías", "📦 Recepcion de mercaderia"])

# ==========================================
# 1. SOLAPA: DASHBOARD GENERAL
# ==========================================
with tab_dash:
    st.caption("ℹ️ Vista general del estado actual de la mercadería sincronizada de forma automática.")
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
            busqueda = st.text_input("Filtrar por SKU o Descripción:", placeholder="Escribí aquí para buscar...")

        if 'busqueda' in locals() and busqueda:
            df_filtrado = df_productos[
                df_productos['SKU'].astype(str).str.contains(busqueda, case=False, na=False) |
                df_productos['Descripción'].astype(str).str.contains(busqueda, case=False, na=False)
            ]
        else:
            df_filtrado = df_productos

        st.markdown("<br>", unsafe_allow_html=True)
        st.dataframe(df_filtrado, width='stretch', hide_index=True)

# ==========================================
# 2. SOLAPA: CONTROL FÍSICO POR DÍA
# ==========================================
with tab_control:
    st.markdown("### Auditoría de Stock Físico por Jornada")
    st.caption("Agrupá el conteo de múltiples productos realizados en un día específico y guardalos en conjunto.")
    st.markdown("<br>", unsafe_allow_html=True)
    
    col_fecha, col_resp = st.columns(2)
    with col_fecha:
        fecha_control = st.date_input("📅 Fecha del Conteo Diario", datetime.date.today())
    with col_resp:
        responsable = st.text_input("👤 Responsable del Conteo", placeholder="Ej: Tamara")

    if not df_productos.empty:
        opciones_buscador = df_productos.apply(lambda x: f"{x['Descripción']} | SKU: {x['SKU']}", axis=1).tolist()
        opciones_buscador.insert(0, "Seleccione un producto...")
    else:
        opciones_buscador = ["No hay productos disponibles"]

    st.markdown("<br>", unsafe_allow_html=True)
    with st.container():
        st.markdown("#### ➕ Agregar Producto al Conteo del Día")
        with st.form("form_conteo_dia", clear_on_submit=True):
            col1, col2, col3, col4 = st.columns([3, 1, 1, 1])
            with col1:
                producto_seleccionado = st.selectbox("Buscar Producto", opciones_buscador)
            with col2:
                talle_input = st.text_input("Talle")
            with col3:
                color_input = st.text_input("Color")
            with col4:
                stock_fisico_input = st.number_input("Stock Físico", min_value=0, step=1)
                
            st.markdown("<br>", unsafe_allow_html=True)
            agregar_btn = st.form_submit_button("Agregar a la sesión de hoy")
            
            if agregar_btn and producto_seleccionado != "Seleccione un producto..." and producto_seleccionado != "No hay productos disponibles":
                partes = producto_seleccionado.split(" | SKU: ")
                nombre_prod = partes[0]
                sku_extraido = partes[1] if len(partes) > 1 else "Sin SKU"
                
                prod_info = df_productos[df_productos['SKU'].astype(str) == str(sku_extraido).strip()]
                stock_sis = prod_info.iloc[0]['Stock'] if not prod_info.empty else 0.0
                diferencia = stock_fisico_input - stock_sis
                
                st.session_state.lista_control.append({
                    "SKU": sku_extraido,
                    "Producto": nombre_prod,
                    "Talle": talle_input,
                    "Color": color_input,
                    "Stock Físico": stock_fisico_input,
                    "Stock Sistema": stock_sis,
                    "Diferencia": diferencia
                })
                st.success(f"¡Agregado a la lista del día! ({nombre_prod})")

    if st.session_state.lista_control:
        st.markdown("<br><hr><br>", unsafe_allow_html=True)
        st.markdown(f"### 📋 Resumen de Conteo para el día: {fecha_control}")
        df_control_actual = pd.DataFrame(st.session_state.lista_control)
        
        def pintar_diferencia(val):
            color = 'green' if val == 0 else 'red'
            return f'color: {color}; font-weight: bold;'
        
        st.dataframe(df_control_actual.style.map(pintar_diferencia, subset=['Diferencia']), width='stretch')
        
        with st.expander("⚙️ Opciones de edición y corrección de la lista actual"):
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
                if st.button("🧹 Vaciar todo"):
                    st.session_state.lista_control = []
                    st.rerun()
                
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("💾 Guardar Conteo del Día en el Historial"):
            if not responsable:
                st.error("Por favor, ingresá el nombre del responsable antes de guardar.")
            else:
                conexion = sqlite3.connect(DB_PATH)
                cursor = conexion.cursor()
                for item in st.session_state.lista_control:
                    cursor.execute("""
                        INSERT INTO controles_fisicos (fecha, responsable, sku, producto, talle, color, stock_fisico, stock_sistema, diferencia)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (str(fecha_control), responsable, item['SKU'], item['Producto'], item['Talle'], item['Color'], item['Stock Físico'], item['Stock Sistema'], item['Diferencia']))
                conexion.commit()
                conexion.close()
                
                st.session_state.lista_control = []
                st.success(f"¡Conteo del día {fecha_control} guardado exitosamente en el historial!")
                st.rerun()

# ==========================================
# 3. SOLAPA: HISTORIAL
# ==========================================
with tab_historial:
    st.markdown("### Historial de Auditorías por Día")
    st.caption("Registro completo de los controles físicos agrupados y ordenados por fecha.")
    st.markdown("<br>", unsafe_allow_html=True)
    
    df_historial = cargar_historial()
    
    if df_historial.empty:
        st.info("No hay controles registrados todavía.")
    else:
        # Filtro opcional por Fecha en el historial
        fechas_disponibles = ["Todas"] + list(df_historial['Fecha'].unique())
        filtro_fecha = st.selectbox("Filtrar historial por Fecha de Conteo:", fechas_disponibles)
        
        if filtro_fecha != "Todas":
            df_hist_mostrar = df_historial[df_historial['Fecha'] == filtro_fecha]
        else:
            df_hist_mostrar = df_historial

        def pintar_historial(val):
            if pd.isna(val): return ''
            color = 'green' if val == 0 else 'red'
            return f'color: {color}; font-weight: bold;'
            
        st.dataframe(df_hist_mostrar.drop(columns=["ID"]).style.map(pintar_historial, subset=['Diferencia']), width='stretch', hide_index=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        with st.expander("🗑️ Zona de administración: Eliminar registros o días enteros"):
            opcion_eliminacion = st.radio("¿Qué deseas eliminar?", ["Un ítem específico", "Un día entero de conteo"])
            
            if opcion_eliminacion == "Un ítem específico":
                opciones_borrar_historial = [f"{row['ID']} - Fecha: {row['Fecha']} | Producto: {row['Producto']}" for _, row in df_hist_mostrar.iterrows()]
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
                            
                            st.success("¡Registro eliminado correctamente!")
                            st.rerun()
            else:
                fechas_para_borrar = list(df_historial['Fecha'].unique())
                fechas_para_borrar.insert(0, "Seleccione una fecha...")
                
                col_dia1, col_dia2 = st.columns([3, 1])
                with col_dia1:
                    fecha_a_borrar = st.selectbox("Fecha a eliminar:", fechas_para_borrar, label_visibility="collapsed")
                with col_dia2:
                    if st.button("🗑️ Borrar día completo"):
                        if fecha_a_borrar != "Seleccione una fecha...":
                            conexion = sqlite3.connect(DB_PATH)
                            cursor = conexion.cursor()
                            cursor.execute("DELETE FROM controles_fisicos WHERE fecha = ?", (fecha_a_borrar,))
                            conexion.commit()
                            conexion.close()
                            
                            st.success(f"¡Todos los registros del día {fecha_a_borrar} fueron eliminados exitosamente!")
                            st.rerun()

# ==========================================
# 4. SOLAPA: RECEPCION DE MERCADERIA
# ==========================================
with tab_movimientos:
    st.markdown("### Recepción de Mercadería")
    st.caption("Registra la entrada o recepción de nueva mercadería al stock.")
    st.markdown("<br>", unsafe_allow_html=True)

    with st.form("form_movimiento", clear_on_submit=True):
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            fecha_mov = st.date_input("📅 Fecha de Recepción", datetime.date.today())
            tipo_mov = st.selectbox("Tipo de Movimiento", ["Ingreso (+)", "Egreso (-)"])
            responsable_mov = st.text_input("👤 Responsable", placeholder="Ej: Tamara")
        with col_m2:
            if not df_productos.empty:
                opciones_mov = df_productos.apply(lambda x: f"{x['Descripción']} | SKU: {x['SKU']}", axis=1).tolist()
                opciones_mov.insert(0, "Seleccione un producto...")
            else:
                opciones_mov = ["No hay productos disponibles"]
            
            producto_mov = st.selectbox("Producto", opciones_mov)
            cantidad_mov = st.number_input("Cantidad", min_value=1, step=1, value=1)

        col_m3, col_m4 = st.columns(2)
        with col_m3:
            talle_mov = st.text_input("Talle (Opcional)")
        with col_m4:
            color_mov = st.text_input("Color (Opcional)")

        observacion_mov = st.text_area("Observación / Motivo", placeholder="Ej: Ingreso de proveedor o reposición")
        
        st.markdown("<br>", unsafe_allow_html=True)
        btn_guardar_mov = st.form_submit_button("💾 Registrar Recepción")

        if btn_guardar_mov:
            if producto_mov in ["Seleccione un producto...", "No hay productos disponibles"]:
                st.error("Por favor, selecciona un producto válido.")
            elif not responsable_mov:
                st.error("Por favor, ingresa el nombre del responsable.")
            else:
                partes = producto_mov.split(" | SKU: ")
                nombre_p = partes[0]
                sku_p = partes[1] if len(partes) > 1 else "Sin SKU"

                conexion = sqlite3.connect(DB_PATH)
                cursor = conexion.cursor()
                cursor.execute("""
                    INSERT INTO movimientos_stock (fecha, tipo, sku, producto, talle, color, cantidad, responsable, observacion)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (str(fecha_mov), tipo_mov, sku_p, nombre_p, talle_mov, color_mov, float(cantidad_mov), responsable_mov, observacion_mov))
                conexion.commit()
                conexion.close()

                st.success(f"¡Recepción registrada correctamente para {nombre_p}!")

    st.markdown("<br><hr><br>", unsafe_allow_html=True)
    st.markdown("### Historial de Recepciones")
    df_movs = cargar_historial_movimientos()
    if df_movs.empty:
        st.info("No hay recepciones registradas todavía.")
    else:
        def pintar_tipo(val):
            color = 'green' if 'Ingreso' in str(val) else 'red'
            return f'color: {color}; font-weight: bold;'
        st.dataframe(df_movs.drop(columns=["ID"]).style.map(pintar_tipo, subset=['Tipo']), width='stretch', hide_index=True)
