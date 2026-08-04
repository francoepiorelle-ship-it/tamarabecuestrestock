from pathlib import Path
import sqlite3
import pandas as pd
import streamlit as st
import datetime
from zoneinfo import ZoneInfo
import io

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
            Proveedor TEXT,
            Rubro TEXT,
            Subrubro TEXT,
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
            hora TEXT,
            tipo TEXT,
            sku TEXT,
            producto TEXT,
            proveedor TEXT,
            talle TEXT,
            color TEXT,
            cantidad REAL,
            responsable TEXT,
            observacion TEXT,
            remito_archivo TEXT,
            fecha_registro TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    
    cursor.execute("PRAGMA table_info(movimientos_stock)")
    columnas = [col[1] for col in cursor.fetchall()]
    if "remito_archivo" not in columnas:
        cursor.execute("ALTER TABLE movimientos_stock ADD COLUMN remito_archivo TEXT")
        
    cursor.execute("PRAGMA table_info(productos)")
    cols_prod = [col[1] for col in cursor.fetchall()]
    if "Rubro" not in cols_prod:
        cursor.execute("ALTER TABLE productos ADD COLUMN Rubro TEXT")
    if "Subrubro" not in cols_prod:
        cursor.execute("ALTER TABLE productos ADD COLUMN Subrubro TEXT")
        
    conexion.commit()
    conexion.close()

def sincronizar_excel_automatico():
    ruta_excel = BASE_DIR / "stock_actualizado.xlsx"
    if ruta_excel.exists():
        try:
            df = pd.read_excel(ruta_excel)
            df.columns = df.columns.str.strip()
            
            col_map = {}
            for col in df.columns:
                col_lower = col.lower().replace(" ", "").replace("_", "").replace("-", "")
                if col_lower in ["sku"]:
                    col_map[col] = "SKU"
                elif col_lower in ["nombre", "descripcion", "producto"]:
                    col_map[col] = "Descripción"
                elif col_lower in ["proveedor", "prov"]:
                    col_map[col] = "Proveedor"
                elif col_lower in ["rubro", "categoria"]:
                    col_map[col] = "Rubro"
                elif col_lower in ["subrubro", "subcategoria", "subrub"]:
                    col_map[col] = "Subrubro"
                elif col_lower in ["stock", "cantidad"]:
                    col_map[col] = "Stock"
                elif col_lower in ["stockreservado", "reservado"]:
                    col_map[col] = "Stock Reservado"
            
            df = df.rename(columns=col_map)
            
            columnas_necesarias = ["SKU", "Descripción", "Stock"]
            if all(c in df.columns for c in columnas_necesarias):
                df_filtrado = df.copy()
                
                if "Proveedor" not in df_filtrado.columns:
                    df_filtrado["Proveedor"] = "General"
                if "Rubro" not in df_filtrado.columns:
                    df_filtrado["Rubro"] = "General"
                else:
                    df_filtrado["Rubro"] = df_filtrado["Rubro"].fillna("General")
                    
                if "Subrubro" not in df_filtrado.columns:
                    df_filtrado["Subrubro"] = "General"
                else:
                    df_filtrado["Subrubro"] = df_filtrado["Subrubro"].fillna("General")
                
                if "Stock Reservado" not in df_filtrado.columns:
                    df_filtrado["Stock Reservado"] = 0
                
                cols_finales = ["SKU", "Descripción", "Proveedor", "Rubro", "Subrubro", "Stock", "Stock Reservado"]
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
        df = pd.read_sql("SELECT SKU, Descripción, Proveedor, Rubro, Subrubro, Stock, "Stock Reservado", "Stock Disponible", "Última Actualización" FROM productos", conexion)
    except Exception:
        df = pd.DataFrame(columns=["SKU", "Descripción", "Proveedor", "Rubro", "Subrubro", "Stock", "Stock Reservado", "Stock Disponible", "Última Actualización"])
    conexion.close()
    return df

def cargar_historial():
    conexion = sqlite3.connect(DB_PATH)
    try:
        df = pd.read_sql("SELECT id AS ID, fecha AS Fecha, hora AS Hora, responsable AS Responsable, sku AS SKU, producto AS Producto, talle AS Talle, color AS Color, stock_fisico AS 'Stock Físico', stock_sistema AS 'Stock Sistema', diferencia AS Diferencia FROM controles_fisicos ORDER BY fecha DESC, hora DESC, id DESC", conexion)
    except Exception:
        df = pd.DataFrame(columns=["ID", "Fecha", "Hora", "Responsable", "SKU", "Producto", "Talle", "Color", "Stock Físico", "Stock Sistema", "Diferencia"])
    conexion.close()
    return df

def cargar_historial_movimientos():
    conexion = sqlite3.connect(DB_PATH)
    try:
        df = pd.read_sql("SELECT id AS ID, fecha AS Fecha, hora AS Hora, tipo AS Tipo, sku AS SKU, producto AS Producto, proveedor AS Proveedor, talle AS Talle, color AS Color, cantidad AS Cantidad, responsable AS Responsable, observacion AS Observación, remito_archivo AS Remito FROM movimientos_stock ORDER BY fecha DESC, hora DESC, id DESC", conexion)
    except Exception:
        df = pd.DataFrame(columns=["ID", "Fecha", "Hora", "Tipo", "SKU", "Producto", "Proveedor", "Talle", "Color", "Cantidad", "Responsable", "Observación", "Remito"])
    conexion.close()
    return df

def convertir_df_a_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Planilla')
    processed_data = output.getvalue()
    return processed_data

if 'lista_control' not in st.session_state:
    st.session_state.lista_control = []

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
            f_col1, f_col2, f_col3 = st.columns(3)
            with f_col1:
                busqueda = st.text_input("Filtrar por SKU o Descripción:", placeholder="Escribí aquí para buscar...")
            with f_col2:
                proveedores_disponibles = ["Todos"] + sorted(df_productos['Proveedor'].dropna().unique().tolist()) if 'Proveedor' in df_productos.columns else ["Todos"]
                filtro_proveedor_dash = st.selectbox("Filtrar por Proveedor:", proveedores_disponibles)
            with f_col3:
                rubros_disponibles = ["Todos"] + sorted(df_productos['Rubro'].dropna().unique().tolist()) if 'Rubro' in df_productos.columns else ["Todos"]
                filtro_rubro_dash = st.selectbox("Filtrar por Rubro:", rubros_disponibles)

        df_filtrado = df_productos.copy()
        if 'busqueda' in locals() and busqueda:
            df_filtrado = df_filtrado[
                df_filtrado['SKU'].astype(str).str.contains(busqueda, case=False, na=False) |
                df_filtrado['Descripción'].astype(str).str.contains(busqueda, case=False, na=False)
            ]
        if 'filtro_proveedor_dash' in locals() and filtro_proveedor_dash != "Todos":
            df_filtrado = df_filtrado[df_filtrado['Proveedor'] == filtro_proveedor_dash]
        if 'filtro_rubro_dash' in locals() and filtro_rubro_dash != "Todos":
            df_filtrado = df_filtrado[df_filtrado['Rubro'] == filtro_rubro_dash]

        st.markdown("<br>", unsafe_allow_html=True)
        st.dataframe(df_filtrado, use_container_width=True, hide_index=True)

# ==========================================
# 2. SOLAPA: CONTROL DE STOCK
# ==========================================
with tab_control:
    st.markdown("### Control de Stock")
    st.caption("Podes realizar varios conteos en el mismo día (ej. turno mañana y turno tarde) guardándolos de forma independiente.")
    st.markdown("<br>", unsafe_allow_html=True)
    
    col_fecha, col_resp = st.columns(2)
    with col_fecha:
        fecha_control = st.date_input("📅 Fecha del Conteo", datetime.date.today(), key="f_ctrl_dia")
    with col_resp:
        responsable = st.text_input("👤 Responsable del Conteo", placeholder="Ej: Tamara", key="resp_ctrl_dia")

    if not df_productos.empty:
        opciones_buscador = df_productos.apply(lambda x: f"{x['Descripción']} | SKU: {x['SKU']} | Rubro: {x.get('Rubro', 'N/A')} | Prov: {x.get('Proveedor', 'N/A')}", axis=1).tolist()
        opciones_buscador.insert(0, "Seleccione un producto...")
    else:
        opciones_buscador = ["No hay productos disponibles"]

    st.markdown("<br>", unsafe_allow_html=True)
    with st.container():
        st.markdown("#### ➕ Agregar Producto a la Tanda Actual de Conteo")
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
            agregar_btn = st.form_submit_button("Agregar a la tanda de hoy")
            
            if agregar_btn and producto_seleccionado != "Seleccione un producto..." and producto_seleccionado != "No hay productos disponibles":
                partes = producto_seleccionado.split(" | SKU: ")
                nombre_prod = partes[0]
                sku_extraido = partes[1].split(" | Rubro: ")[0] if len(partes) > 1 else "Sin SKU"
                
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
                        INSERT INTO controles_fisicos (fecha, hora, responsable, sku, producto, talle, color, stock_fisico, stock_sistema, diferencia)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (str(fecha_control), hora_arg, responsable, item['SKU'], item['Producto'], item['Talle'], item['Color'], item['Stock Físico'], item['Stock Sistema'], item['Diferencia']))
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
    st.caption("La planilla incluye el Rubro y Subrubro de cada producto. Utiliza el formulario de abajo para ir agregándolos uno a uno y adjunta un único remito para toda la planilla.")
    st.markdown("<br>", unsafe_allow_html=True)
    
    col_mfech, col_mresp, col_mprov = st.columns(3)
    with col_mfech:
        fecha_recepcion_dia = st.date_input("📅 Fecha de Recepción", datetime.date.today(), key="f_rec_dia")
    with col_mresp:
        responsable_recepcion = st.text_input("👤 Responsable", placeholder="Ej: Tamara", key="resp_rec_dia")
    with col_mprov:
        lista_prov_form = sorted(df_productos['Proveedor'].dropna().unique().tolist()) if 'Proveedor' in df_productos.columns else ["General"]
        proveedor_recepcion_global = st.selectbox("📦 Proveedor de la Recepción", ["General"] + lista_prov_form, key="prov_rec_global")

    archivo_remito_subido = st.file_uploader("📎 Adjuntar Remito General para Toda la Planilla (Foto o PDF)", type=["png", "jpg", "jpeg", "pdf"], key="remito_masivo_subida")

    st.markdown("<br>", unsafe_allow_html=True)

    with st.container():
        st.markdown("#### ➕ Agregar Producto a la Recepción")
        with st.form("form_agregar_item_recepcion", clear_on_submit=True):
            
            if not df_productos.empty:
                opciones_skus_dict = {f"{row['Descripción']} (SKU: {row['SKU']} | Rubro: {row.get('Rubro', 'N/A')} | Subrubro: {row.get('Subrubro', 'N/A')})": row for _, row in df_productos.iterrows()}
                lista_opciones_prod = ["Seleccione un producto..."] + list(opciones_skus_dict.keys())
            else:
                lista_opciones_prod = ["No hay productos disponibles"]

            f_col_p1, f_col_p2 = st.columns([2, 1])
            with f_col_p1:
                prod_elegido_form = st.selectbox("Seleccionar Producto", lista_opciones_prod)
            with f_col_p2:
                tipo_mov_form = st.selectbox("Tipo de Movimiento", ["Ingreso (+)", "Egreso (-)"])

            f_col_p3, f_col_p4, f_col_p5 = st.columns(3)
            with f_col_p3:
                cant_form = st.number_input("Cantidad", min_value=1, value=1, step=1)
            with f_col_p4:
                talle_form = st.text_input("Talle (Opcional)")
            with f_col_p5:
                color_form = st.text_input("Color (Opcional)")

            obs_form = st.text_input("Observación (Opcional)")

            st.markdown("<br>", unsafe_allow_html=True)
            btn_agregar_a_tabla = st.form_submit_button("➕ Añadir a la planilla de recepción")

            if btn_agregar_a_tabla:
                if prod_elegido_form == "Seleccione un producto..." or prod_elegido_form == "No hay productos disponibles":
                    st.error("Por favor, seleccione un producto válido.")
                else:
                    datos_prod = opciones_skus_dict[prod_elegido_form]
                    sku_val = str(datos_prod['SKU'])
                    desc_val = str(datos_prod['Descripción'])
                    rubro_val = str(datos_prod.get('Rubro', ''))
                    subrubro_val = str(datos_prod.get('Subrubro', ''))

                    nuevo_item = {
                        "SKU": sku_val,
                        "Producto": desc_val,
                        "Rubro": rubro_val,
                        "Subrubro": subrubro_val,
                        "Tipo": tipo_mov_form,
                        "Cantidad": float(cant_form),
                        "Talle": talle_input,
                        "Color": color_input,
                        "Observación": obs_form
                    }
                    
                    if 'tabla_recepcion_items' not in st.session_state:
                        st.session_state.tabla_recepcion_items = []
                    
                    st.session_state.tabla_recepcion_items.append(nuevo_item)
                    st.success(f"¡Producto agregado a la planilla correctamente!")

    st.markdown("<br><hr><br>", unsafe_allow_html=True)
    st.markdown("#### 📋 Planilla de Recepción Actual")

    if 'tabla_recepcion_items' not in st.session_state:
        st.session_state.tabla_recepcion_items = []

    if not st.session_state.tabla_recepcion_items:
        st.info("ℹ️ La planilla está vacía. Utiliza el formulario de arriba para agregar los productos que ingresan.")
        df_mostrar_recepcion = pd.DataFrame(columns=["SKU", "Producto", "Rubro", "Subrubro", "Tipo", "Cantidad", "Talle", "Color", "Observación"])
        st.dataframe(df_mostrar_recepcion, use_container_width=True, hide_index=True)
    else:
        df_mostrar_recepcion = pd.DataFrame(st.session_state.tabla_recepcion_items)
        st.dataframe(df_mostrar_recepcion, use_container_width=True, hide_index=True)

        st.markdown("<br>", unsafe_allow_html=True)
        excel_bytes_rec = convertir_df_a_excel(df_mostrar_recepcion)
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
                if not responsable_recepcion:
                    st.error("Por favor, ingresá el nombre del responsable antes de guardar.")
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
                            INSERT INTO movimientos_stock (fecha, hora, tipo, sku, producto, proveedor, talle, color, cantidad, responsable, observacion, remito_archivo)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, (str(fecha_recepcion_dia), hora_arg, item['Tipo'], item['SKU'], item['Producto'], proveedor_recepcion_global, item['Talle'], item['Color'], item['Cantidad'], responsable_recepcion, item['Observación'], ruta_guardada))

                        if "Ingreso" in item['Tipo']:
                            lista_para_etiquetas.append({
                                "Producto": item['Producto'],
                                "SKU": item['SKU'],
                                "Cantidad": item['Cantidad']
                            })

                    conexion.commit()
                    conexion.close()

                    st.session_state.ultimo_movimiento_guardado = lista_para_etiquetas
                    st.session_state.tabla_recepcion_items = []
                    st.success(f"¡Recepción completa guardada exitosamente a las {hora_arg}!")
                    st.rerun()
        with col_acc2:
            if st.button("🧹 Vaciar Planilla"):
                st.session_state.tabla_recepcion_items = []
                st.rerun()

    if st.session_state.ultimo_movimiento_guardado:
        st.markdown("<br><hr><br>", unsafe_allow_html=True)
        st.markdown("#### 🏷️ Última Recepción Guardada (Etiquetas / Resumen)")
        st.info("Aquí puedes ver los ítems ingresados recientemente para su etiquetado o control.")
        df_ult = pd.DataFrame(st.session_state.ultimo_movimiento_guardado)
        st.dataframe(df_ult, use_container_width=True, hide_index=True)

# ==========================================
# 4. SOLAPA: HISTORIAL Y AUDITORÍAS
# ==========================================
with tab_historial:
    st.markdown("### Historial y Auditorías")
    st.caption("Consulta los registros anteriores de controles físicos y movimientos de stock.")
    st.markdown("<br>", unsafe_allow_html=True)
    
    sub_tab1, sub_tab2 = st.tabs(["📋 Controles Físicos", "📦 Movimientos de Stock"])
    
    with sub_tab1:
        df_hist_ctrl = cargar_historial()
        if df_hist_ctrl.empty:
            st.info("No hay registros de controles físicos todavía.")
        else:
            st.dataframe(df_hist_ctrl, use_container_width=True, hide_index=True)
            st.markdown("<br>", unsafe_allow_html=True)
            st.download_button(
                label="📥 Descargar Historial de Controles (.xlsx)",
                data=convertir_df_a_excel(df_hist_ctrl),
                file_name="historial_controles_fisicos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    with sub_tab2:
        df_hist_mov = cargar_historial_movimientos()
        if df_hist_mov.empty:
            st.info("No hay registros de movimientos de stock todavía.")
        else:
            st.dataframe(df_hist_mov, use_container_width=True, hide_index=True)
            st.markdown("<br>", unsafe_allow_html=True)
            st.download_button(
                label="📥 Descargar Historial de Movimientos (.xlsx)",
                data=convertir_df_a_excel(df_hist_mov),
                file_name="historial_movimientos_stock.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
