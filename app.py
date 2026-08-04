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
        
    conexion.commit()
    conexion.close()

def sincronizar_excel_automatico():
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
        opciones_buscador = df_productos.apply(lambda x: f"{x['Descripción']} | SKU: {x['SKU']} | Prov: {x.get('Proveedor', 'N/A')}", axis=1).tolist()
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
                st.success(f"¡Agregado a la tanda! ({nombre_prod})")

    if st.session_state.lista_control:
        st.markdown("<br><hr><br>", unsafe_allow_html=True)
        st.markdown(f"### 📋 Tanda Actual de Conteo ({fecha_control})")
        df_control_actual = pd.DataFrame(st.session_state.lista_control)
        
        def pintar_diferencia(val):
            color = 'green' if val == 0 else 'red'
            return f'color: {color}; font-weight: bold;'
        
        st.dataframe(df_control_actual.style.map(pintar_diferencia, subset=['Diferencia']), width='stretch')
        
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
    st.markdown("### Recepción de Mercadería (Carga Masiva)")
    st.caption("Cargá todos los productos que vienen en esta recepción en una sola tabla, adjuntá el remito y guardá todo de un solo golpe.")
    st.markdown("<br>", unsafe_allow_html=True)
    
    col_mfech, col_mresp, col_mprov = st.columns(3)
    with col_mfech:
        fecha_recepcion_dia = st.date_input("📅 Fecha de Recepción", datetime.date.today(), key="f_rec_dia")
    with col_mresp:
        responsable_recepcion = st.text_input("👤 Responsable", placeholder="Ej: Tamara", key="resp_rec_dia")
    with col_mprov:
        lista_prov_form = sorted(df_productos['Proveedor'].dropna().unique().tolist()) if 'Proveedor' in df_productos.columns else ["General"]
        proveedor_recepcion_global = st.selectbox("📦 Proveedor de la Recepción", ["General"] + lista_prov_form, key="prov_rec_global")

    archivo_remito_subido = st.file_uploader("📎 Adjuntar Remito General (Foto o PDF)", type=["png", "jpg", "jpeg", "pdf"], key="remito_masivo_subida")

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("#### 📝 Detalle de Artículos Recibidos")
    st.caption("Agregá filas abajo de todo, buscá tu producto o completá la cantidad, talle y color de cada ítem que vino en el remito.")

    if 'df_recepcion_masiva' not in st.session_state:
        st.session_state.df_recepcion_masiva = pd.DataFrame(columns=[
            "Seleccionar", "SKU", "Producto", "Tipo", "Cantidad", "Talle", "Color", "Observación"
        ])

    if st.session_state.df_recepcion_masiva.empty and not df_productos.empty:
        filas_iniciales = []
        for _, prod in df_productos.head(5).iterrows():
            filas_iniciales.append({
                "Seleccionar": False,
                "SKU": str(prod["SKU"]),
                "Producto": prod["Descripción"],
                "Tipo": "Ingreso (+)",
                "Cantidad": 1.0,
                "Talle": "",
                "Color": "",
                "Observación": ""
            })
        st.session_state.df_recepcion_masiva = pd.DataFrame(filas_iniciales)

    opciones_skus_dict = {}
    if not df_productos.empty:
        opciones_skus_dict = {f"{row['Descripción']} (SKU: {row['SKU']})": row['SKU'] for _, row in df_productos.iterrows()}

    df_editado = st.data_editor(
        st.session_state.df_recepcion_masiva,
        num_rows="dynamic",
        column_config={
            "Seleccionar": st.column_config.CheckboxColumn("¿Incluir?", default=True),
            "Tipo": st.column_config.SelectboxColumn("Tipo", options=["Ingreso (+)", "Egreso (-)"], required=True),
            "Cantidad": st.column_config.NumberColumn("Cantidad", min_value=1, step=1, format="%d"),
        },
        use_container_width=True,
        key="editor_recepcion_masiva"
    )

    st.markdown("<br>", unsafe_allow_html=True)
    
    col_btn1, col_btn2 = st.columns([2, 1])
    with col_btn1:
        guardar_recepcion_masiva_btn = st.button("💾 Guardar Recepción Completa en el Historial")
    with col_btn2:
        if st.button("🧹 Limpiar Tabla Actual"):
            st.session_state.df_recepcion_masiva = pd.DataFrame(columns=[
                "Seleccionar", "SKU", "Producto", "Tipo", "Cantidad", "Talle", "Color", "Observación"
            ])
            st.rerun()

    if guardar_recepcion_masiva_btn:
        if not responsable_recepcion:
            st.error("Por favor, ingresá el nombre del responsable antes de guardar.")
        else:
            df_a_guardar = df_editado[df_editado["Seleccionar"] == True].copy()
            
            if df_a_guardar.empty:
                st.warning("⚠️ No hay ningún producto seleccionado para guardar en esta recepción.")
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

                for _, row in df_a_guardar.iterrows():
                    sku_val = str(row["SKU"]) if pd.notna(row["SKU"]) else "Sin SKU"
                    prod_val = str(row["Producto"]) if pd.notna(row["Producto"]) else "Sin Descripción"
                    tipo_val = str(row["Tipo"])
                    cant_val = float(row["Cantidad"]) if pd.notna(row["Cantidad"]) else 1.0
                    talle_val = str(row["Talle"]) if pd.notna(row["Talle"]) else ""
                    color_val = str(row["Color"]) if pd.notna(row["Color"]) else ""
                    obs_val = str(row["Observación"]) if pd.notna(row["Observación"]) else ""

                    cursor.execute("""
                        INSERT INTO movimientos_stock (fecha, hora, tipo, sku, producto, proveedor, talle, color, cantidad, responsable, observacion, remito_archivo)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (str(fecha_recepcion_dia), hora_arg, tipo_val, sku_val, prod_val, proveedor_recepcion_global, talle_val, color_val, cant_val, responsable_recepcion, obs_val, ruta_guardada))

                    if "Ingreso" in tipo_val:
                        lista_para_etiquetas.append({
                            "Producto": prod_val,
                            "SKU": sku_val,
                            "Cantidad": cant_val
                        })

                conexion.commit()
                conexion.close()

                st.session_state.ultimo_movimiento_guardado = lista_para_etiquetas
                st.success(f"¡Recepción completa guardada exitosamente a las {hora_arg}! Se registraron {len(df_a_guardar)} ítems juntos.")

    if st.session_state.ultimo_movimiento_guardado:
        st.markdown("<br><hr><br>", unsafe_allow_html=True)
        st.markdown("### 🏷️ Generar Archivo CSV para Etiquetas de esta Recepción")
        st.caption("Podes descargar el archivo con formato exacto (delimitado con comillas) listo para tu impresora de etiquetas.")
        
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
            csv_buffer = df_etiquetas.to_csv(index=False, header=False, encoding="utf-8-sig", quoting=1)
            
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
    st.markdown("### Historial y Auditorías Separadas por Tandas")
    st.caption("Visualiza cada control o recepción de forma totalmente independiente, pudiendo ver o descargar los remitos adjuntos.")
    st.markdown("<br>", unsafe_allow_html=True)
    
    tipo_historial_seleccionado = st.radio("Seleccione el tipo de historial a visualizar:", ["Control de Stock", "Recepción / Movimientos de Mercadería"], horizontal=True)
    st.markdown("<br>", unsafe_allow_html=True)

    if tipo_historial_seleccionado == "Control de Stock":
        st.markdown("#### 📋 Historial de Controles de Stock (por Tandas / Horas)")
        df_historial = cargar_historial()
        
        if df_historial.empty:
            st.info("No hay controles de stock registrados todavía.")
        else:
            df_historial['Tanda_Label'] = df_historial.apply(lambda r: f"Fecha: {r['Fecha']} | Hora: {r['Hora']} | Resp: {r['Responsable']}", axis=1)
            tandas_fisicas_disponibles = ["Todas las tandas"] + list(df_historial['Tanda_Label'].unique())
            
            filtro_tanda_fisica = st.selectbox("🔍 Filtrar por Tanda específica (Fecha y Hora):", tandas_fisicas_disponibles)
            df_hist_mostrar = df_historial if filtro_tanda_fisica == "Todas las tandas" else df_historial[df_historial['Tanda_Label'] == filtro_tanda_fisica]

            def pintar_historial(val):
                if pd.isna(val): return ''
                color = 'green' if val == 0 else 'red'
                return f'color: {color}; font-weight: bold;'
                
            st.dataframe(df_hist_mostrar.drop(columns=["ID", "Tanda_Label"]).style.map(pintar_historial, subset=['Diferencia']), width='stretch', hide_index=True)
            
            st.markdown("<br>", unsafe_allow_html=True)
            with st.expander("🗑️ Zona de administración: Eliminar tandas o registros específicos"):
                opcion_eliminacion = st.radio("¿Qué deseas eliminar?", ["Un ítem específico de una tanda", "Una tanda entera (por hora)", "Un día entero completo"], key="del_fisico")
                
                if opcion_eliminacion == "Un ítem específico de una tanda":
                    opciones_borrar_historial = [f"{row['ID']} - [{row['Fecha']} {row['Hora']}] {row['Producto']}" for _, row in df_hist_mostrar.iterrows()]
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
                else:
                    dias_disponibles = df_historial['Fecha'].unique().tolist()
                    dias_disponibles.insert(0, "Seleccione un día...")
                    dia_a_borrar = st.selectbox("Día:", dias_disponibles)
                    if st.button("🗑️ Eliminar Día Entero"):
                        if dia_a_borrar != "Seleccione un día...":
                            conexion = sqlite3.connect(DB_PATH)
                            cursor = conexion.cursor()
                            cursor.execute("DELETE FROM controles_fisicos WHERE fecha = ?", (dia_a_borrar,))
                            conexion.commit()
                            conexion.close()
                            st.success("¡Día completo eliminado exitosamente!")
                            st.rerun()

    else:
        st.markdown("#### 📦 Historial de Recepciones y Movimientos de Mercadería")
        df_movimientos = cargar_historial_movimientos()
        
        if df_movimientos.empty:
            st.info("No hay recepciones ni movimientos registrados todavía.")
        else:
            df_movimientos['Mov_Label'] = df_movimientos.apply(lambda r: f"Fecha: {r['Fecha']} | Hora: {r['Hora']} | Resp: {r['Responsable']} | Prov: {r['Proveedor']}", axis=1)
            movs_disponibles = ["Todos los movimientos"] + list(df_movimientos['Mov_Label'].unique())
            
            filtro_mov = st.selectbox("🔍 Filtrar por Recepción específica:", movs_disponibles)
            df_mov_mostrar = df_movimientos if filtro_mov == "Todos los movimientos" else df_movimientos[df_movimientos['Mov_Label'] == filtro_mov]

            st.dataframe(df_mov_mostrar.drop(columns=["ID", "Mov_Label"]), width='stretch', hide_index=True)

            # Sección para ver remitos adjuntos
            remitos_con_archivo = df_mov_mostrar[df_mov_mostrar['Remito'].notna() & (df_mov_mostrar['Remito'] != "")]
            if not remitos_con_archivo.empty:
                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown("#### 📎 Remitos Adjuntos en esta vista")
                for _, row in remitos_con_archivo.iterrows():
                    ruta_rem = Path(row['Remito'])
                    if ruta_rem.exists():
                        with st.expander(f"📄 Remito de {row['Fecha']} {row['Hora']} - Prov: {row['Proveedor']} ({row['Responsable']})"):
                            if ruta_rem.suffix.lower() in ['.png', '.jpg', '.jpeg']:
                                st.image(str(ruta_rem), caption=f"Remito {row['Fecha']}", use_column_width=True)
                            elif ruta_rem.suffix.lower() == '.pdf':
                                st.markdown(f"📄 Archivo PDF adjunto: `{ruta_rem.name}`")
                                with open(ruta_rem, "rb") as pdf_file:
                                    st.download_button(
                                        label="📥 Descargar PDF del Remito",
                                        data=pdf_file,
                                        file_name=ruta_rem.name,
                                        mime="application/pdf",
                                        key=f"dl_pdf_{row['ID']}"
                                    )

            st.markdown("<br>", unsafe_allow_html=True)
            with st.expander("🗑️ Zona de administración: Eliminar movimientos o recepciones"):
                opciones_borrar_mov = [f"{row['ID']} - [{row['Fecha']} {row['Hora']}] {row['Producto']} ({row['Tipo']})" for _, row in df_mov_mostrar.iterrows()]
                opciones_borrar_mov.insert(0, "Seleccione un movimiento...")
                
                col_mdel1, col_mdel2 = st.columns([3, 1])
                with col_mdel1:
                    mov_a_borrar = st.selectbox("Movimiento:", opciones_borrar_mov, label_visibility="collapsed")
                with col_mdel2:
                    if st.button("❌ Eliminar movimiento"):
                        if mov_a_borrar != "Seleccione un movimiento...":
                            id_mov_borrar = int(mov_a_borrar.split(" - ")[0])
                            conexion = sqlite3.connect(DB_PATH)
                            cursor = conexion.cursor()
                            cursor.execute("DELETE FROM movimientos_stock WHERE id = ?", (id_mov_borrar,))
                            conexion.commit()
                            conexion.close()
                            st.success("¡Movimiento eliminado!")
                            st.rerun()
