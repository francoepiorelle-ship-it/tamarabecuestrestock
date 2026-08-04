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
            proveedor TEXT,
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
                df_filtrado = df.copy()
                
                if "Proveedor" not in df_filtrado.columns:
                    df_filtrado["Proveedor"] = "General"
                
                cols_finales = ["SKU", "Nombre", "Proveedor", "Stock", "Stock Reservado"]
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
        df = pd.read_sql("SELECT SKU, Descripción, Proveedor, Stock, \"Stock Reservado\", \"Stock Disponible\", \"Última Actualización\" FROM productos", conexion)
    except Exception:
        df = pd.DataFrame(columns=["SKU", "Descripción", "Proveedor", "Stock", "Stock Reservado", "Stock Disponible", "Última Actualización"])
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
        df = pd.read_sql("SELECT id AS ID, fecha AS Fecha, tipo AS Tipo, sku AS SKU, producto AS Producto, proveedor AS Proveedor, talle AS Talle, color AS Color, cantidad AS Cantidad, responsable AS Responsable, observacion AS Observación FROM movimientos_stock ORDER BY fecha DESC, id DESC", conexion)
    except Exception:
        df = pd.DataFrame(columns=["ID", "Fecha", "Tipo", "SKU", "Producto", "Proveedor", "Talle", "Color", "Cantidad", "Responsable", "Observación"])
    conexion.close()
    return df

if 'lista_control' not in st.session_state:
    st.session_state.lista_control = []

if 'lista_recepcion' not in st.session_state:
    st.session_state.lista_recepcion = []

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

tab_dash, tab_control, tab_movimientos, tab_historial = st.tabs(["📊 Dashboard General", "📋 Control Físico por Día", "📦 Recepción de Mercadería por Día", "📂 Historial y Auditorías"])

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
            f_col1, f_col2 = st.columns(2)
            with f_col1:
                busqueda = st.text_input("Filtrar por SKU o Descripción:", placeholder="Escribí aquí para buscar...")
            with f_col2:
                proveedores_disponibles = ["Todos"] + sorted(df_productos['Proveedor'].dropna().unique().tolist()) if 'Proveedor' in df_productos.columns else ["Todos"]
                filtro_proveedor_dash = st.selectbox("Filtrar por Proveedor:", proveedores_disponibles)

        df_filtrado = df_productos.copy()
        if 'busqueda' in locals() and busqueda:
            df_filtrado = df_filtrado[
                df_filtrado['SKU'].astype(str).str.contains(busqueda, case=False, na=False) |
                df_filtrado['Descripción'].astype(str).str.contains(busqueda, case=False, na=False)
            ]
        if 'filtro_proveedor_dash' in locals() and filtro_proveedor_dash != "Todos":
            df_filtrado = df_filtrado[df_filtrado['Proveedor'] == filtro_proveedor_dash]

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
        opciones_buscador = df_productos.apply(lambda x: f"{x['Descripción']} | SKU: {x['SKU']} | Prov: {x.get('Proveedor', 'N/A')}", axis=1).tolist()
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
                sku_extraido = partes[1].split(" | Prov: ")[0] if len(partes) > 1 else "Sin SKU"
                
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
# 3. SOLAPA: RECEPCION DE MERCADERIA POR DÍA
# ==========================================
with tab_movimientos:
    st.markdown("### Recepción de Mercadería por Jornada")
    st.caption("Agrupá la entrada o salida de múltiples productos en una misma fecha y guárdalos en conjunto.")
    st.markdown("<br>", unsafe_allow_html=True)
    
    col_mfech, col_mresp = st.columns(2)
    with col_mfech:
        fecha_recepcion_dia = st.date_input("📅 Fecha de Recepción / Movimiento", datetime.date.today(), key="f_rec_dia")
    with col_mresp:
        responsable_recepcion = st.text_input("👤 Responsable de la Recepción", placeholder="Ej: Tamara", key="resp_rec_dia")

    # Filtro aparte por Proveedor dentro del formulario
    lista_prov_form = sorted(df_productos['Proveedor'].dropna().unique().tolist()) if 'Proveedor' in df_productos.columns else ["General"]
    proveedor_seleccionado_form = st.selectbox("Filtrar por Proveedor (opcional):", ["Todos"] + lista_prov_form, key="prov_rec_filtro")

    if not df_productos.empty:
        df_prod_form = df_productos.copy()
        if proveedor_seleccionado_form != "Todos":
            df_prod_form = df_prod_form[df_prod_form['Proveedor'] == proveedor_seleccionado_form]
        
        opciones_mov = df_prod_form.apply(lambda x: f"{x['Descripción']} | SKU: {x['SKU']} | Prov: {x.get('Proveedor', 'N/A')}", axis=1).tolist()
        opciones_mov.insert(0, "Seleccione un producto...")
    else:
        opciones_mov = ["No hay productos disponibles"]

    st.markdown("<br>", unsafe_allow_html=True)
    with st.container():
        st.markdown("#### ➕ Agregar Producto a la Recepción del Día")
        with st.form("form_recepcion_dia", clear_on_submit=True):
            col_d1, col_d2, col_d3, col_d4, col_d5 = st.columns([2.5, 1, 1, 1, 1])
            with col_d1:
                producto_mov = st.selectbox("Producto", opciones_mov)
            with col_d2:
                tipo_mov = st.selectbox("Tipo", ["Ingreso (+)", "Egreso (-)"])
            with col_d3:
                cantidad_mov = st.number_input("Cantidad", min_value=1, step=1, value=1)
            with col_d4:
                talle_mov = st.text_input("Talle")
            with col_d5:
                color_mov = st.text_input("Color")
            
            observacion_mov = st.text_input("Observación / Motivo", placeholder="Ej: Reposición de stock o compra")
            
            st.markdown("<br>", unsafe_allow_html=True)
            agregar_rec_btn = st.form_submit_button("Agregar a la sesión de recepción de hoy")
            
            if agregar_rec_btn and producto_mov not in ["Seleccione un producto...", "No hay productos disponibles"]:
                partes = producto_mov.split(" | SKU: ")
                nombre_p = partes[0]
                sku_p = partes[1].split(" | Prov: ")[0] if len(partes) > 1 else "Sin SKU"
                
                prod_match = df_productos[df_productos['SKU'].astype(str) == str(sku_p).strip()]
                prov_p = prod_match.iloc[0]['Proveedor'] if not prod_match.empty and 'Proveedor' in prod_match.columns else "General"

                st.session_state.lista_recepcion.append({
                    "Tipo": tipo_mov,
                    "SKU": sku_p,
                    "Producto": nombre_p,
                    "Proveedor": prov_p,
                    "Talle": talle_mov,
                    "Color": color_mov,
                    "Cantidad": float(cantidad_mov),
                    "Observación": observacion_mov
                })
                st.success(f"¡Agregado a la recepción del día! ({nombre_p})")

    if st.session_state.lista_recepcion:
        st.markdown("<br><hr><br>", unsafe_allow_html=True)
        st.markdown(f"### 📋 Resumen de Recepción para el día: {fecha_recepcion_dia}")
        df_rec_actual = pd.DataFrame(st.session_state.lista_recepcion)
        
        def pintar_tipo_sesion(val):
            color = 'green' if 'Ingreso' in str(val) else 'red'
            return f'color: {color}; font-weight: bold;'
        
        st.dataframe(df_rec_actual.style.map(pintar_tipo_sesion, subset=['Tipo']), width='stretch')
        
        with st.expander("⚙️ Opciones de edición y corrección de la lista actual de recepción"):
            col_err1, col_err2, col_err3 = st.columns([2, 1, 1])
            with col_err1:
                opciones_borrar_rec = [f"{i} - {item['Producto']} (Cant: {item['Cantidad']})" for i, item in enumerate(st.session_state.lista_recepcion)]
                item_rec_a_borrar = st.selectbox("Elegí cuál borrar:", opciones_borrar_rec, label_visibility="collapsed")
            with col_err2:
                if st.button("❌ Borrar ítem de recepción"):
                    if item_rec_a_borrar:
                        indice = int(item_rec_a_borrar.split(" - ")[0])
                        st.session_state.lista_recepcion.pop(indice)
                        st.rerun()
            with col_err3:
                if st.button("🧹 Vaciar lista de recepción"):
                    st.session_state.lista_recepcion = []
                    st.rerun()
                
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("💾 Guardar Recepción del Día en el Historial"):
            if not responsable_recepcion:
                st.error("Por favor, ingresá el nombre del responsable antes de guardar.")
            else:
                conexion = sqlite3.connect(DB_PATH)
                cursor = conexion.cursor()
                for item in st.session_state.lista_recepcion:
                    cursor.execute("""
                        INSERT INTO movimientos_stock (fecha, tipo, sku, producto, proveedor, talle, color, cantidad, responsable, observacion)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (str(fecha_recepcion_dia), item['Tipo'], item['SKU'], item['Producto'], item['Proveedor'], item['Talle'], item['Color'], item['Cantidad'], responsable_recepcion, item['Observación']))
                conexion.commit()
                conexion.close()
                
                st.session_state.lista_recepcion = []
                st.success(f"¡Recepción del día {fecha_recepcion_dia} guardada exitosamente en el historial unificado!")
                st.rerun()

# ==========================================
# 4. SOLAPA: HISTORIAL Y AUDITORÍAS (EN CONJUNTO)
# ==========================================
with tab_historial:
    st.markdown("### Historial y Auditorías en Conjunto")
    st.caption("Consulta unificada de todos los controles físicos de stock y las recepciones de mercadería registradas.")
    st.markdown("<br>", unsafe_allow_html=True)
    
    tipo_historial_seleccionado = st.radio("Seleccione el tipo de historial a visualizar:", ["Control Físico de Stock", "Recepción / Movimientos de Mercadería"], horizontal=True)
    st.markdown("<br>", unsafe_allow_html=True)

    if tipo_historial_seleccionado == "Control Físico de Stock":
        st.markdown("#### 📋 Historial de Controles Físicos")
        df_historial = cargar_historial()
        
        if df_historial.empty:
            st.info("No hay controles físicos registrados todavía.")
        else:
            fechas_disponibles = ["Todas"] + list(df_historial['Fecha'].unique())
            filtro_fecha = st.selectbox("Filtrar historial físico por Fecha:", fechas_disponibles)
            
            df_hist_mostrar = df_historial if filtro_fecha == "Todas" else df_historial[df_historial['Fecha'] == filtro_fecha]

            def pintar_historial(val):
                if pd.isna(val): return ''
                color = 'green' if val == 0 else 'red'
                return f'color: {color}; font-weight: bold;'
                
            st.dataframe(df_hist_mostrar.drop(columns=["ID"]).style.map(pintar_historial, subset=['Diferencia']), width='stretch', hide_index=True)
            
            st.markdown("<br>", unsafe_allow_html=True)
            with st.expander("🗑️ Zona de administración: Eliminar registros físicos"):
                opcion_eliminacion = st.radio("¿Qué deseas eliminar?", ["Un ítem específico", "Un día entero de conteo"], key="del_fisico")
                
                if opcion_eliminacion == "Un ítem específico":
                    opciones_borrar_historial = [f"{row['ID']} - Fecha: {row['Fecha']} | Producto: {row['Producto']}" for _, row in df_hist_mostrar.iterrows()]
                    opciones_borrar_historial.insert(0, "Seleccione un registro...")
                    
                    col_del1, col_del2 = st.columns([3, 1])
                    with col_del1:
                        registro_a_borrar = st.selectbox("Registro:", opciones_borrar_historial, label_visibility="collapsed")
                    with col_del2:
                        if st.button("❌ Eliminar ítem físico"):
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
                        fecha_a_borrar = st.selectbox("Fecha a eliminar:", fechas_para_borrar, label_visibility="collapsed", key="f_del_dia_fis")
                    with col_dia2:
                        if st.button("🗑️ Borrar día físico completo"):
                            if fecha_a_borrar != "Seleccione una fecha...":
                                conexion = sqlite3.connect(DB_PATH)
                                cursor = conexion.cursor()
                                cursor.execute("DELETE FROM controles_fisicos WHERE fecha = ?", (fecha_a_borrar,))
                                conexion.commit()
                                conexion.close()
                                st.success(f"¡Todos los registros del día {fecha_a_borrar} fueron eliminados!")
                                st.rerun()

    else:
        st.markdown("#### 📦 Historial de Recepciones y Movimientos de Mercadería")
        df_movs = cargar_historial_movimientos()
        
        if df_movs.empty:
            st.info("No hay recepciones registradas todavía.")
        else:
            col_hf1, col_hf2 = st.columns(2)
            with col_hf1:
                fechas_mov_opts = ["Todas"] + list(df_movs['Fecha'].unique())
                filtro_fecha_mov = st.selectbox("Filtrar historial por Fecha:", fechas_mov_opts)
            with col_hf2:
                prov_hist_opts = ["Todos"] + sorted(df_movs['Proveedor'].dropna().unique().tolist()) if 'Proveedor' in df_movs.columns else ["Todos"]
                filtro_prov_hist = st.selectbox("Filtrar historial por Proveedor:", prov_hist_opts)

            df_movs_mostrar = df_movs.copy()
            if filtro_fecha_mov != "Todas":
                df_movs_mostrar = df_movs_mostrar[df_movs_mostrar['Fecha'] == filtro_fecha_mov]
            if filtro_prov_hist != "Todos":
                df_movs_mostrar = df_movs_mostrar[df_movs_mostrar['Proveedor'] == filtro_prov_hist]

            def pintar_tipo(val):
                color = 'green' if 'Ingreso' in str(val) else 'red'
                return f'color: {color}; font-weight: bold;'
                
            st.dataframe(df_movs_mostrar.drop(columns=["ID"]).style.map(pintar_tipo, subset=['Tipo']), width='stretch', hide_index=True)
            
            st.markdown("<br>", unsafe_allow_html=True)
            with st.expander("🗑️ Zona de administración: Eliminar registros de recepción"):
                opcion_eliminacion_rec = st.radio("¿Qué deseas eliminar?", ["Un ítem específico", "Un día entero de recepción"], key="del_rec")
                
                if opcion_eliminacion_rec == "Un ítem específico":
                    opciones_borrar_mov = [f"{row['ID']} - Fecha: {row['Fecha']} | Producto: {row['Producto']}" for _, row in df_movs_mostrar.iterrows()]
                    opciones_borrar_mov.insert(0, "Seleccione un registro...")
                    
                    col_del1, col_del2 = st.columns([3, 1])
                    with col_del1:
                        registro_mov_a_borrar = st.selectbox("Registro:", opciones_borrar_mov, label_visibility="collapsed")
                    with col_del2:
                        if st.button("❌ Eliminar ítem de recepción"):
                            if registro_mov_a_borrar != "Seleccione un registro...":
                                id_borrar = int(registro_mov_a_borrar.split(" - ")[0])
                                conexion = sqlite3.connect(DB_PATH)
                                cursor = conexion.cursor()
                                cursor.execute("DELETE FROM movimientos_stock WHERE id = ?", (id_borrar,))
                                conexion.commit()
                                conexion.close()
                                st.success("¡Registro de recepción eliminado correctamente!")
                                st.rerun()
                else:
                    fechas_mov_para_borrar = list(df_movs['Fecha'].unique())
                    fechas_mov_para_borrar.insert(0, "Seleccione una fecha...")
                    
                    col_dia1, col_dia2 = st.columns([3, 1])
                    with col_dia1:
                        fecha_mov_a_borrar = st.selectbox("Fecha a eliminar:", fechas_mov_para_borrar, label_visibility="collapsed", key="f_del_dia_rec")
                    with col_dia2:
                        if st.button("🗑️ Borrar día de recepción completo"):
                            if fecha_mov_a_borrar != "Seleccione una fecha...":
                                conexion = sqlite3.connect(DB_PATH)
                                cursor = conexion.cursor()
                                cursor.execute("DELETE FROM movimientos_stock WHERE fecha = ?", (fecha_mov_a_borrar,))
                                conexion.commit()
                                conexion.close()
                                st.success(f"¡Todos los registros de recepción del día {fecha_mov_a_borrar} fueron eliminados!")
                                st.rerun()
