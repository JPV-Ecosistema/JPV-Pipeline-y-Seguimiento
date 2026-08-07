import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import io
import os
import json
import re
import gspread
from google.oauth2.service_account import Credentials

# --- CONTROL DE VERSIONES ---
# Incrementar APP_VERSION cada vez que se publique un cambio relevante en la app.
APP_VERSION = "1.12.0"

# --- ZONA HORARIA (Chile continental) ---
ZONA_HORARIA_CL = ZoneInfo("America/Santiago")

def ahora_cl():
    """Hora actual en Chile (naive, ya ajustada por CLT/CLST), sin depender de la zona horaria del servidor."""
    return datetime.now(ZONA_HORARIA_CL).replace(tzinfo=None)

def calcular_rangos_semana():
    """Devuelve (inicio, fin) de la semana pasada y de esta semana (lunes a domingo), según la fecha actual en Chile."""
    hoy = ahora_cl().date()
    inicio_semana_actual = hoy - timedelta(days=hoy.weekday())
    fin_semana_actual = inicio_semana_actual + timedelta(days=6)
    inicio_semana_pasada = inicio_semana_actual - timedelta(days=7)
    fin_semana_pasada = inicio_semana_actual - timedelta(days=1)
    return inicio_semana_pasada, fin_semana_pasada, inicio_semana_actual, fin_semana_actual

# --- CONFIGURACIÓN DE ETIQUETAS ---
PROB_MAP = {
    "0%": "Nula",
    "25%": "Remota",
    "50%": "Podría Ser",
    "75%": "Altamente probable",
    "100%": "Cierta"
}

# --- RESPALDO EN LA NUBE (Google Sheets vía cuenta de servicio) ---
PERSISTENCE_DIR = "persistence"
BACKUP_FILE = os.path.join(PERSISTENCE_DIR, "PIPELINE_BACKUP.json")

def get_google_client():
    try:
        if "gcp_service_account" in st.secrets and "google_sheet_url" in st.secrets:
            scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
            creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
            return gspread.authorize(creds)
    except Exception as e:
        st.session_state["_ultimo_respaldo_error"] = f"Error crítico de conexión a Google Cloud: {e}"
    return None

def get_backup_sheet():
    client = get_google_client()
    if client:
        try:
            doc = client.open_by_url(st.secrets["google_sheet_url"])
            try:
                return doc.worksheet("Pipeline_Backup")
            except Exception:
                return doc.add_worksheet(title="Pipeline_Backup", rows="2000", cols="30")
        except Exception as e:
            st.session_state["_ultimo_respaldo_error"] = f"Error al abrir la hoja de respaldo: {e}"
    return None

def guardar_local_pipeline(df):
    """Guarda una copia local (JSON) del pipeline activo en esta instancia del servidor."""
    os.makedirs(PERSISTENCE_DIR, exist_ok=True)
    fecha_hoy = ahora_cl().strftime("%Y-%m-%d %H:%M:%S")
    df_guardar = df.copy()
    df_guardar['Fecha probable de facturación'] = df_guardar['Fecha probable de facturación'].astype(str)
    datos_guardar = {"fecha": fecha_hoy, "data": df_guardar.fillna("").to_dict(orient="records")}
    try:
        with open(BACKUP_FILE, 'w', encoding='utf-8') as f:
            json.dump(datos_guardar, f, ensure_ascii=False)
    except Exception:
        pass

def guardar_respaldo_pipeline(df):
    """Deja una copia local (JSON) y sobrescribe el respaldo completo en la nube (Google Sheets).

    Usar solo para refrescos completos del pipeline (carga inicial o edición masiva en la
    tabla del Tab 1): sobrescribe TODA la hoja Pipeline_Backup. Si dos usuarios tienen sesiones
    abiertas al mismo tiempo, esto puede pisar cambios que otro usuario haya guardado en otros
    casos mientras tanto. Para actualizar un único caso sin ese riesgo, usar guardar_caso_en_nube().
    """
    guardar_local_pipeline(df)
    fecha_hoy = ahora_cl().strftime("%Y-%m-%d %H:%M:%S")
    df_guardar = df.copy()
    df_guardar['Fecha probable de facturación'] = df_guardar['Fecha probable de facturación'].astype(str)

    try:
        ws = get_backup_sheet()
        if ws:
            matriz = [["FECHA_ACTUALIZACION", fecha_hoy]]
            matriz.append(df_guardar.columns.astype(str).tolist())
            matriz.extend(df_guardar.fillna("").astype(str).values.tolist())
            ws.clear()
            ws.update("A1", matriz)
            st.session_state["_ultimo_respaldo_nube"] = fecha_hoy
            st.session_state["_ultimo_respaldo_error"] = None
            cargar_historial_desde_nube.clear()
    except Exception as cloud_error:
        if "429" in str(cloud_error):
            st.session_state["_ultimo_respaldo_error"] = "Google Cloud está en pausa (Límite 429 de peticiones). El respaldo local se guardó igualmente."
        else:
            st.session_state["_ultimo_respaldo_error"] = f"Hubo un detalle con el respaldo en nube: {cloud_error}"

def guardar_caso_en_nube(fila_caso_actualizada, col_llave='Número de caso'):
    """Actualiza en Pipeline_Backup solo la fila del caso editado (upsert por número de caso),
    sin sobrescribir el resto de la hoja. Así, si dos ajustadores guardan seguimientos de casos
    distintos al mismo tiempo, no se pisan los cambios entre sí."""
    fecha_hoy = ahora_cl().strftime("%Y-%m-%d %H:%M:%S")
    try:
        ws = get_backup_sheet()
        if not ws:
            return
        valores = ws.get_all_values()
        if len(valores) < 2:
            return
        headers = valores[1]
        filas_datos = valores[2:]

        if col_llave not in headers:
            return
        idx_col_llave = headers.index(col_llave)
        caso_valor = str(fila_caso_actualizada.get(col_llave, "")).strip()
        fila_nueva = [str(fila_caso_actualizada.get(h, "")) for h in headers]

        fila_encontrada = None
        for i, fila in enumerate(filas_datos):
            valor_actual = fila[idx_col_llave] if idx_col_llave < len(fila) else ""
            if str(valor_actual).strip() == caso_valor:
                fila_encontrada = i
                break

        if fila_encontrada is not None:
            num_fila_hoja = fila_encontrada + 3  # fila 1: metadata, fila 2: encabezados, datos desde fila 3
            ws.update(f"A{num_fila_hoja}", [fila_nueva])
        else:
            ws.append_row(fila_nueva)

        ws.update("B1", [[fecha_hoy]])
        st.session_state["_ultimo_respaldo_nube"] = fecha_hoy
        st.session_state["_ultimo_respaldo_error"] = None
        cargar_historial_desde_nube.clear()
    except Exception as cloud_error:
        if "429" in str(cloud_error):
            st.session_state["_ultimo_respaldo_error"] = "Google Cloud está en pausa (Límite 429 de peticiones). El cambio local se guardó igualmente."
        else:
            st.session_state["_ultimo_respaldo_error"] = f"Hubo un detalle al sincronizar el caso con la nube: {cloud_error}"

@st.cache_data(ttl=60, show_spinner=False)
def cargar_reporte_desde_nube():
    """Recupera el último 'Reporte de Acciones' (hoja Base_Maestra) subido desde el Planificador Semanal.

    Se cachea 60 segundos (compartido entre todas las sesiones de este servidor) para no
    golpear la cuota de la API de Google Sheets en cada rerun de Streamlit, que ocurre en
    cada clic, filtro o edición.
    """
    client = get_google_client()
    if client:
        try:
            doc = client.open_by_url(st.secrets["google_sheet_url"])
            ws = doc.worksheet("Base_Maestra")
            metadata = ws.row_values(1)
            if len(metadata) >= 2 and metadata[0] == "FECHA_ACTUALIZACION":
                fecha_str = metadata[1]
                registros = ws.get_all_records(head=2)
                df = pd.DataFrame(registros)
                if not df.empty:
                    return df, fecha_str
        except Exception as e:
            st.session_state["_ultimo_error_lectura_reporte"] = str(e)
            return None, None
    st.session_state["_ultimo_error_lectura_reporte"] = None
    return None, None

@st.cache_data(ttl=60, show_spinner=False)
def cargar_historial_desde_nube():
    """Recupera el último respaldo del pipeline (hoja Pipeline_Backup) para usarlo como Pipeline Anterior.

    Se cachea 60 segundos (compartido entre todas las sesiones de este servidor) por el mismo
    motivo que cargar_reporte_desde_nube(): evitar exceder la cuota de la API en cada rerun.
    """
    client = get_google_client()
    if client:
        try:
            doc = client.open_by_url(st.secrets["google_sheet_url"])
            ws = doc.worksheet("Pipeline_Backup")
            metadata = ws.row_values(1)
            if len(metadata) >= 2 and metadata[0] == "FECHA_ACTUALIZACION":
                fecha_str = metadata[1]
                registros = ws.get_all_records(head=2)
                df = pd.DataFrame(registros)
                if not df.empty:
                    return df, fecha_str
        except Exception as e:
            st.session_state["_ultimo_error_lectura_historial"] = str(e)
            return None, None
    st.session_state["_ultimo_error_lectura_historial"] = None
    return None, None

def detectar_historial_desde_excel(archivo_historial):
    """Detecta automáticamente la hoja más reciente del Pipeline Anterior (Excel maestro con varias hojas históricas)."""
    xl_historial = pd.ExcelFile(archivo_historial)
    hojas = xl_historial.sheet_names
    hoja_maestra = None
    fecha_reciente = datetime.min

    for h in hojas:
        match = re.search(r'(\d{2}-\d{2}-\d{2})', h)
        if match:
            try:
                fecha_hoja = datetime.strptime(match.group(1), "%d-%m-%y")
                if fecha_hoja > fecha_reciente:
                    fecha_reciente = fecha_hoja
                    hoja_maestra = h
            except:
                continue

    if not hoja_maestra:
        posibles_nombres = ['Número de caso', 'Numero de caso', 'N° caso', 'Caso']
        for h in reversed(hojas):
            df_check = pd.read_excel(xl_historial, sheet_name=h, nrows=10, header=None)
            if any(str(val).strip() in posibles_nombres for row in df_check.values for val in row):
                hoja_maestra = h
                break

    if not hoja_maestra:
        return None, None

    df_hist_raw = pd.read_excel(xl_historial, sheet_name=hoja_maestra, header=None)
    fila_h = 0
    for i, row in df_hist_raw.iterrows():
        if any(str(val).strip() in ['Número de caso', 'Numero de caso', 'N° caso', 'Caso'] for val in row.values):
            fila_h = i
            break

    df_hist = pd.read_excel(xl_historial, sheet_name=hoja_maestra, skiprows=fila_h)
    df_hist.columns = [str(c).strip() for c in df_hist.columns]
    return df_hist, hoja_maestra

def normalizar_para_base_maestra(df):
    """Replica la normalización que usa el Planificador Semanal antes de guardar en Base_Maestra."""
    df_norm = df.fillna("")
    for col in df_norm.columns:
        df_norm[col] = df_norm[col].astype(str)
        df_norm[col] = df_norm[col].apply(
            lambda x: "" if x.strip().lower() in ["nan", "nat", "none", "<na>", "inf", "-inf"] else x
        )
    return df_norm

def guardar_reporte_en_nube(df_nuevo_crudo):
    """Sube el Reporte de Acciones cargado manualmente en el Pipeline de vuelta a Base_Maestra,
    para que el Planificador Semanal también quede sincronizado (sincronización bidireccional)."""
    fecha_hoy = ahora_cl().strftime("%Y-%m-%d %H:%M:%S")
    df_norm = normalizar_para_base_maestra(df_nuevo_crudo)
    try:
        client = get_google_client()
        if client:
            doc = client.open_by_url(st.secrets["google_sheet_url"])
            try:
                ws = doc.worksheet("Base_Maestra")
            except Exception:
                ws = doc.add_worksheet(title="Base_Maestra", rows="100", cols="100")
            matriz = [["FECHA_ACTUALIZACION", fecha_hoy]]
            matriz.append(df_norm.columns.astype(str).tolist())
            matriz.extend(df_norm.values.tolist())
            ws.clear()
            ws.update("A1", matriz)
            st.session_state["_ultimo_envio_base_maestra"] = fecha_hoy
            st.session_state["_ultimo_envio_base_maestra_error"] = None
            cargar_reporte_desde_nube.clear()
    except Exception as cloud_error:
        if "429" in str(cloud_error):
            st.session_state["_ultimo_envio_base_maestra_error"] = "Google Cloud está en pausa (Límite 429 de peticiones). Se reintentará más adelante."
        else:
            st.session_state["_ultimo_envio_base_maestra_error"] = f"No se pudo sincronizar con Base_Maestra: {cloud_error}"

def render_sidebar_respaldo():
    st.sidebar.divider()
    st.sidebar.header("☁️ Respaldo en la Nube")
    if st.session_state.get("_ultimo_respaldo_nube"):
        st.sidebar.success(f"✅ Último respaldo: {st.session_state['_ultimo_respaldo_nube']}")
    if st.session_state.get("_ultimo_respaldo_error"):
        st.sidebar.warning(f"⚠️ {st.session_state['_ultimo_respaldo_error']}")
    if not st.session_state.get("_ultimo_respaldo_nube") and not st.session_state.get("_ultimo_respaldo_error"):
        st.sidebar.info("Aún no se ha generado un respaldo en esta sesión.")

def render_sidebar_version():
    try:
        fecha_revision = datetime.fromtimestamp(os.path.getmtime(__file__), tz=ZONA_HORARIA_CL).strftime("%d/%m/%Y %H:%M")
    except Exception:
        fecha_revision = "N/D"
    st.sidebar.divider()
    st.sidebar.caption(f"🧾 Versión {APP_VERSION} · Última revisión de la app: {fecha_revision}")

# Columnas definitivas para el reporte de salida
COLUMNAS_FINALES = [
    'Número de caso', 'Número de siniestro', 'Nickname', 'División', 
    'Compañía de seguros', 'Corredora', 'Ajustador senior', 'Asegurado', 
    'Creado en', 'Divisa', 'Perdida bruta (en moneda del caso)', 
    'Deducible (en moneda del caso)', 'Monto asegurado (en moneda del caso)', 
    'Honorarios (UF)', 'Facturado', 'Último movimiento', 
    'Contenido último movimiento', 'Probabilidad cierre 2026', 
    'Indicación Probabilidad', 'Hon Probables 2026', 'Observaciones', 
    'Fecha probable de facturación'
]

st.set_page_config(page_title="JPV Pipeline y Seguimiento", layout="wide")
st.title("🚀 JPV: Pipeline de Facturación Probable")

st.sidebar.header("Carga de Documentos")

df_nube, fecha_nube = cargar_reporte_desde_nube()

st.sidebar.subheader("1. Reporte Nuevo de Acciones")
if df_nube is not None:
    st.sidebar.success(f"☁️ Usando el Reporte de Acciones compartido con el Planificador (actualizado el {fecha_nube}).")
else:
    st.sidebar.info("No se encontró un Reporte de Acciones en la nube. Sube uno manualmente.")
    if st.session_state.get("_ultimo_error_lectura_reporte"):
        st.sidebar.caption(f"Detalle: {st.session_state['_ultimo_error_lectura_reporte']}")
archivo_nuevo = st.sidebar.file_uploader(
    "Cargar manualmente (opcional, reemplaza el de la nube)", type=["xlsx"]
)

df_nuevo_manual = None
if archivo_nuevo is not None:
    df_nuevo_manual = pd.read_excel(archivo_nuevo, skiprows=5)
    firma_archivo = f"{archivo_nuevo.name}_{archivo_nuevo.size}"
    if st.session_state.get("_ultimo_reporte_manual_enviado") != firma_archivo:
        guardar_reporte_en_nube(df_nuevo_manual)
        st.session_state["_ultimo_reporte_manual_enviado"] = firma_archivo

    if st.session_state.get("_ultimo_envio_base_maestra_error"):
        st.sidebar.warning(f"⚠️ {st.session_state['_ultimo_envio_base_maestra_error']}")
    elif st.session_state.get("_ultimo_envio_base_maestra"):
        st.sidebar.caption(f"🔁 Sincronizado con Base_Maestra: {st.session_state['_ultimo_envio_base_maestra']}")

df_hist_nube, fecha_hist_nube = cargar_historial_desde_nube()

st.sidebar.subheader("2. Pipeline Anterior (histórico)")
if df_hist_nube is not None:
    st.sidebar.success(f"☁️ Usando el último respaldo del Pipeline (actualizado el {fecha_hist_nube}).")
else:
    st.sidebar.info("No se encontró un respaldo del Pipeline en la nube. Sube el Excel maestro manualmente.")
    if st.session_state.get("_ultimo_error_lectura_historial"):
        st.sidebar.caption(f"Detalle: {st.session_state['_ultimo_error_lectura_historial']}")
archivo_historial = st.sidebar.file_uploader(
    "Cargar manualmente (opcional, reemplaza el de la nube)", type=["xlsx"], key="uploader_historial"
)

if archivo_historial is not None:
    st.sidebar.warning("⚠️ Esto reemplazará por completo el histórico en la nube (Pipeline_Backup). Úsalo solo en casos excepcionales.")
    password_historial = st.sidebar.text_input(
        "Contraseña para autorizar el reemplazo del histórico:",
        type="password",
        key="password_historial"
    )
    password_correcta = st.secrets.get("password_historial_manual")
    if not password_correcta:
        st.sidebar.error("⚠️ No hay contraseña configurada en los secrets de la app (password_historial_manual). Se ignorará el archivo.")
        archivo_historial = None
    elif not password_historial:
        st.sidebar.info("Ingresa la contraseña para habilitar este archivo. Mientras tanto se usará el respaldo en la nube.")
        archivo_historial = None
    elif password_historial != password_correcta:
        st.sidebar.error("❌ Contraseña incorrecta. Se ignorará el archivo subido y se usará el respaldo en la nube.")
        archivo_historial = None
    else:
        st.sidebar.success("✅ Contraseña correcta. Se usará el archivo subido en vez del respaldo en la nube.")

render_sidebar_respaldo()
render_sidebar_version()

if (df_nuevo_manual is not None or df_nube is not None) and (archivo_historial is not None or df_hist_nube is not None):
    # 1. Cargar Reporte Nuevo (Títulos en fila 6). Prioridad: archivo subido manualmente > reporte compartido en la nube.
    if df_nuevo_manual is not None:
        df_nuevo = df_nuevo_manual.copy()
    else:
        df_nuevo = df_nube.copy()
    df_nuevo.columns = [str(c).strip() for c in df_nuevo.columns]
    df_nuevo = df_nuevo.replace('', pd.NA).dropna(how='all', axis=0)

    # 2. Obtener el Pipeline Anterior. Prioridad: archivo subido manualmente > último respaldo en la nube.
    if archivo_historial is not None:
        df_hist, hoja_maestra = detectar_historial_desde_excel(archivo_historial)
    else:
        df_hist = df_hist_nube.copy()
        df_hist.columns = [str(c).strip() for c in df_hist.columns]
        if 'Probabilidad cierre 2026' in df_hist.columns:
            df_hist['Probabilidad cierre 2026'] = df_hist['Probabilidad cierre 2026'].astype(str).str.replace('%', '', regex=False)
        hoja_maestra = f"Respaldo en la nube ({fecha_hist_nube})"

    if not hoja_maestra:
        st.error("No se pudo identificar la hoja de datos en el historial.")
    else:
        st.info(f"Última actualización detectada: **{hoja_maestra}**")

        col_llave = next((c for c in df_nuevo.columns if c in ['Número de caso', 'Numero de caso', 'N° caso', 'Caso']), None)

        if col_llave:
            # Estandarización de llaves (Texto)
            df_nuevo[col_llave] = df_nuevo[col_llave].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            df_hist[col_llave] = df_hist[col_llave].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

            # Preparar columnas de persistencia
            cols_persistencia = [col_llave, 'Probabilidad cierre 2026', 'Observaciones', 'Fecha probable de facturación']
            for c in cols_persistencia:
                if c not in df_hist.columns: df_hist[c] = ""

            # --- CRUCE DE DATOS ---
            df_final = pd.merge(df_nuevo, df_hist[cols_persistencia], on=col_llave, how='left')

            # --- FORMATEO DE TIPOS PARA EL EDITOR ---
            def to_pct_str(val):
                try:
                    num = float(val)
                    if num <= 1.0: return f"{int(num * 100)}%"
                    return f"{int(num)}%"
                except: return "0%"

            df_final['Probabilidad cierre 2026'] = df_final['Probabilidad cierre 2026'].apply(to_pct_str)
            df_final['Observaciones'] = df_final['Observaciones'].astype(str).replace(['nan', 'None', '<NA>'], '')
            df_final['Fecha probable de facturación'] = pd.to_datetime(df_final['Fecha probable de facturación'], errors='coerce').dt.date

            for col in COLUMNAS_FINALES:
                if col not in df_final.columns: df_final[col] = ""
            
            # Cálculo inicial de Honorarios Probables
            if 'Honorarios (UF)' in df_final.columns:
                df_final['Honorarios (UF)'] = pd.to_numeric(df_final['Honorarios (UF)'], errors='coerce').fillna(0)
                prob_num = df_final['Probabilidad cierre 2026'].str.replace('%', '').astype(float) / 100
                df_final['Hon Probables 2026'] = df_final['Honorarios (UF)'] * prob_num

            df_final = df_final[COLUMNAS_FINALES]

            # --- INICIALIZAR ESTADO COMPARTIDO ENTRE PESTAÑAS ---
            if 'df_pipeline_activo' not in st.session_state:
                st.session_state['df_pipeline_activo'] = df_final.copy()
                guardar_respaldo_pipeline(df_final)

            # --- BLOQUE DE FILTROS PARA PESTAÑA 2 ---
            st.markdown("---")
            st.markdown("#### 🔎 Filtros para Seguimiento de Caso")
            col_f1, col_f2, col_f3, col_f4 = st.columns(4)

            with col_f1:
                divisiones_disponibles = sorted(
                    st.session_state['df_pipeline_activo']['División']
                    .dropna().astype(str).str.strip()
                    .replace('', pd.NA).dropna().unique().tolist()
                )
                filtro_division = st.selectbox(
                    "Filtrar por División:",
                    options=["Todas"] + divisiones_disponibles
                )

            with col_f2:
                ajustadores_disponibles = sorted(
                    st.session_state['df_pipeline_activo']['Ajustador senior']
                    .dropna().astype(str).str.strip()
                    .replace('', pd.NA).dropna().unique().tolist()
                )
                filtro_ajustadores = st.multiselect(
                    "Filtrar por Ajustador(es):",
                    options=ajustadores_disponibles,
                    placeholder="Todos los ajustadores"
                )

            with col_f3:
                filtro_probabilidad = st.multiselect(
                    "Filtrar por Probabilidad de cierre:",
                    options=["0%", "25%", "50%", "75%", "100%"],
                    placeholder="Todas las probabilidades"
                )

            with col_f4:
                filtro_sin_observaciones = st.checkbox(
                    "Solo casos sin observaciones (posibles nuevos)"
                )

            st.markdown("---")

            # --- PESTAÑAS PRINCIPALES ---
            tab_seguimiento, tab_pipeline = st.tabs(["🔍 Seguimiento de Caso", "📋 Pipeline General"])

            # ==========================================
            # PESTAÑA PIPELINE GENERAL (SIN CAMBIOS)
            # ==========================================
            with tab_pipeline:

                st.subheader("Panel de Gestión Semanal")
                
                casos_viejos = set(df_hist[col_llave].unique())
                casos_actuales = set(df_nuevo[col_llave].unique())
                
                nuevos_detectados = [c for c in casos_actuales if c not in casos_viejos]
                salientes_detectados = df_hist[~df_hist[col_llave].isin(casos_actuales)]
                
                col_res1, col_res2 = st.columns(2)
                with col_res1:
                    st.success(f"🆕 **Ingresos:** Se incorporaron **{len(nuevos_detectados)} casos nuevos**.")
                with col_res2:
                    st.warning(f"🔴 **Salidas:** **{len(salientes_detectados)} casos** del pipeline anterior ya no están en el reporte.")
                
                if not salientes_detectados.empty:
                    with st.expander("🔍 Ver listado de casos salientes"):
                        columnas_salientes = [col_llave, 'Nickname', 'Probabilidad cierre 2026', 'Observaciones']
                        cols_mostrar = [c for c in columnas_salientes if c in salientes_detectados.columns]
                        st.dataframe(salientes_detectados[cols_mostrar].fillna(''), hide_index=True)

                st.markdown("---")
                st.markdown("#### 🆕 Casos Nuevos por Fecha de Creación")

                inicio_sp, fin_sp, inicio_se, fin_se = calcular_rangos_semana()

                df_creacion = st.session_state['df_pipeline_activo'].copy()
                df_creacion['_fecha_creado'] = pd.to_datetime(df_creacion['Creado en'], errors='coerce').dt.date

                casos_semana_pasada = df_creacion[
                    (df_creacion['_fecha_creado'] >= inicio_sp) & (df_creacion['_fecha_creado'] <= fin_sp)
                ]
                casos_esta_semana = df_creacion[
                    (df_creacion['_fecha_creado'] >= inicio_se) & (df_creacion['_fecha_creado'] <= fin_se)
                ]

                columnas_nuevos = [c for c in [col_llave, 'Nickname', 'División', 'Ajustador senior', 'Creado en', 'Probabilidad cierre 2026'] if c in df_creacion.columns]

                col_n1, col_n2 = st.columns(2)
                with col_n1:
                    st.info(f"📅 **Casos Nuevos Semana Pasada** ({inicio_sp.strftime('%d/%m')} al {fin_sp.strftime('%d/%m')}): **{len(casos_semana_pasada)} casos**.")
                    with st.expander(f"Ver detalle ({len(casos_semana_pasada)} casos)"):
                        st.dataframe(casos_semana_pasada[columnas_nuevos].fillna(''), hide_index=True, use_container_width=True)
                with col_n2:
                    st.success(f"📅 **Casos Nuevos Esta Semana** ({inicio_se.strftime('%d/%m')} al {fin_se.strftime('%d/%m')}): **{len(casos_esta_semana)} casos**.")
                    with st.expander(f"Ver detalle ({len(casos_esta_semana)} casos)"):
                        st.dataframe(casos_esta_semana[columnas_nuevos].fillna(''), hide_index=True, use_container_width=True)

                def color_semaforo(val):
                    if val in ["75%", "100%"]:
                        return 'background-color: #c6efce; color: #006100;'
                    elif val == "50%":
                        return 'background-color: #ffeb9c; color: #9c5700;'
                    elif val in ["0%", "25%"]:
                        return 'background-color: #ffc7ce; color: #9c0006;'
                    return ''

                # Reconvertir fecha a tipo date antes de mostrar en el editor
                # (Tab2 la convierte a string al guardar, esto la restaura para Tab1)
                df_para_editor = st.session_state['df_pipeline_activo'].copy()
                df_para_editor['Fecha probable de facturación'] = pd.to_datetime(
                    df_para_editor['Fecha probable de facturación'], errors='coerce'
                ).dt.date
                df_styled = df_para_editor.style.map(color_semaforo, subset=['Probabilidad cierre 2026'])

                df_editado = st.data_editor(
                    df_styled,
                    column_config={
                        "Probabilidad cierre 2026": st.column_config.SelectboxColumn(
                            "Probabilidad (%)", 
                            options=["0%", "25%", "50%", "75%", "100%"]
                        ),
                        "Fecha probable de facturación": st.column_config.DateColumn("Fecha Fact."),
                        "Observaciones": st.column_config.TextColumn("Observaciones", width="large")
                    },
                    hide_index=True, 
                    use_container_width=True
                )

                prob_num_final = df_editado['Probabilidad cierre 2026'].str.replace('%', '').astype(float) / 100
                df_editado['Hon Probables 2026'] = df_editado['Honorarios (UF)'] * prob_num_final
                df_editado['Indicación Probabilidad'] = df_editado['Probabilidad cierre 2026'].map(PROB_MAP)

                if not df_editado.equals(st.session_state['df_pipeline_activo']):
                    st.session_state['df_pipeline_activo'] = df_editado.copy()
                    guardar_respaldo_pipeline(df_editado)

                st.metric("FACTURACIÓN PROBABLE TOTAL (UF)", f"{df_editado['Hon Probables 2026'].sum():,.2f}")

                fecha_desc = ahora_cl().strftime("%d-%m-%y")
                buffer = io.BytesIO()
                df_excel = df_editado.copy()
                df_excel['Probabilidad cierre 2026'] = df_excel['Probabilidad cierre 2026'].str.replace('%', '').astype(float) / 100

                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    nombre_hoja_descarga = f"Casos {fecha_desc}"
                    df_excel.to_excel(writer, sheet_name=nombre_hoja_descarga, index=False)
                    workbook = writer.book
                    worksheet = writer.sheets[nombre_hoja_descarga]
                    formato_pct = workbook.add_format({'num_format': '0%'})
                    formato_verde = workbook.add_format({'bg_color': '#c6efce', 'font_color': '#006100'})
                    formato_amarillo = workbook.add_format({'bg_color': '#ffeb9c', 'font_color': '#9c5700'})
                    formato_rojo = workbook.add_format({'bg_color': '#ffc7ce', 'font_color': '#9c0006'})
                    idx_prob = COLUMNAS_FINALES.index('Probabilidad cierre 2026')
                    worksheet.set_column(idx_prob, idx_prob, 15, formato_pct)
                    filas_totales = len(df_excel)
                    worksheet.conditional_format(1, idx_prob, filas_totales, idx_prob, 
                                                 {'type': 'cell', 'criteria': '>=', 'value': 0.75, 'format': formato_verde})
                    worksheet.conditional_format(1, idx_prob, filas_totales, idx_prob, 
                                                 {'type': 'cell', 'criteria': '==', 'value': 0.50, 'format': formato_amarillo})
                    worksheet.conditional_format(1, idx_prob, filas_totales, idx_prob, 
                                                 {'type': 'cell', 'criteria': '<=', 'value': 0.25, 'format': formato_rojo})
                
                st.sidebar.divider()
                st.sidebar.download_button(
                    label="📥 Descargar Pipeline Formateado",
                    data=buffer.getvalue(),
                    file_name=f"JPV_Pipeline_{fecha_desc}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            # ==========================================
            # PESTAÑA SEGUIMIENTO DE CASO (CON FILTROS)
            # ==========================================
            with tab_seguimiento:
                st.subheader("🔍 Seguimiento Individual de Caso")

                df_filtrado = st.session_state['df_pipeline_activo'].copy()

                if filtro_division != "Todas":
                    df_filtrado = df_filtrado[
                        df_filtrado['División'].astype(str).str.strip() == filtro_division
                    ]

                if filtro_ajustadores:
                    df_filtrado = df_filtrado[
                        df_filtrado['Ajustador senior'].astype(str).str.strip().isin(filtro_ajustadores)
                    ]

                if filtro_probabilidad:
                    df_filtrado = df_filtrado[
                        df_filtrado['Probabilidad cierre 2026'].isin(filtro_probabilidad)
                    ]

                if filtro_sin_observaciones:
                    df_filtrado = df_filtrado[
                        df_filtrado['Observaciones'].astype(str).str.strip().isin(['', 'nan', 'None'])
                    ]

                st.caption(f"Mostrando **{len(df_filtrado)}** casos según los filtros aplicados.")

                if filtro_probabilidad or filtro_sin_observaciones:
                    with st.expander(f"📋 Ver listado completo de los {len(df_filtrado)} casos filtrados", expanded=True):
                        columnas_resumen = [
                            'Número de caso', 'Nickname', 'División', 'Ajustador senior',
                            'Probabilidad cierre 2026', 'Observaciones'
                        ]
                        columnas_resumen = [c for c in columnas_resumen if c in df_filtrado.columns]
                        st.dataframe(df_filtrado[columnas_resumen].fillna(''), hide_index=True, use_container_width=True)

                df_filtrado['_num_caso_int'] = pd.to_numeric(
                    df_filtrado['Número de caso'].astype(str).str.replace(r'\.0$', '', regex=True),
                    errors='coerce'
                )
                df_filtrado_ordenado = df_filtrado.sort_values('_num_caso_int', ascending=True)
                df_filtrado_ordenado['_etiqueta'] = (
                    df_filtrado_ordenado['Número de caso'].astype(str) +
                    ' — ' +
                    df_filtrado_ordenado['Nickname'].astype(str).str.strip()
                )
                lista_etiquetas = df_filtrado_ordenado['_etiqueta'].unique().tolist()

                if not lista_etiquetas:
                    st.warning("No hay casos que coincidan con los filtros seleccionados.")
                else:
                    etiqueta_seleccionada = st.selectbox(
                        "Selecciona el Número de Caso a gestionar:",
                        options=["— Selecciona un caso —"] + lista_etiquetas
                    )

                    if etiqueta_seleccionada != "— Selecciona un caso —":
                        caso_seleccionado = etiqueta_seleccionada.split(' — ')[0].strip()

                        fila_caso = st.session_state['df_pipeline_activo'][
                            st.session_state['df_pipeline_activo']['Número de caso'].astype(str) == caso_seleccionado
                        ].iloc[0]

                        st.divider()

                        try:
                            fecha_creacion = pd.to_datetime(fila_caso.get('Creado en', ''), errors='coerce')
                            dias_activo = (ahora_cl() - fecha_creacion).days if pd.notna(fecha_creacion) else None
                        except:
                            dias_activo = None

                        col_dest1, col_dest2, col_dest3 = st.columns(3)
                        with col_dest1:
                            st.markdown(f"""
                                <div style='background-color:#1e3a5f; padding:16px; border-radius:10px; text-align:center;'>
                                    <div style='color:#a0b4c8; font-size:13px; margin-bottom:4px;'>NÚMERO DE CASO</div>
                                    <div style='color:#ffffff; font-size:28px; font-weight:bold;'>{fila_caso.get('Número de caso', '')}</div>
                                </div>
                            """, unsafe_allow_html=True)
                        with col_dest2:
                            st.markdown(f"""
                                <div style='background-color:#1e3a5f; padding:16px; border-radius:10px; text-align:center;'>
                                    <div style='color:#a0b4c8; font-size:13px; margin-bottom:4px;'>NICKNAME</div>
                                    <div style='color:#ffffff; font-size:22px; font-weight:bold;'>{fila_caso.get('Nickname', '')}</div>
                                </div>
                            """, unsafe_allow_html=True)
                        with col_dest3:
                            if dias_activo is not None:
                                color_dias = '#c0392b' if dias_activo > 365 else '#e67e22' if dias_activo > 180 else '#27ae60'
                                st.markdown(f"""
                                    <div style='background-color:{color_dias}; padding:16px; border-radius:10px; text-align:center;'>
                                        <div style='color:#ffffff; font-size:13px; margin-bottom:4px;'>DÍAS EN CARTERA</div>
                                        <div style='color:#ffffff; font-size:36px; font-weight:bold;'>{dias_activo}</div>
                                        <div style='color:#ffffff; font-size:11px;'>desde {fecha_creacion.strftime("%d/%m/%Y")}</div>
                                    </div>
                                """, unsafe_allow_html=True)
                            else:
                                st.metric("Días en cartera", "Sin fecha")

                        st.divider()

                        st.markdown("##### 📄 Datos del Caso")
                        col_a, col_b, col_c = st.columns(3)
                        with col_a:
                            st.text_input("Número de siniestro",    value=str(fila_caso.get('Número de siniestro', '')),                 disabled=True)
                            st.text_input("División",               value=str(fila_caso.get('División', '')),                            disabled=True)
                            st.text_input("Compañía de seguros",    value=str(fila_caso.get('Compañía de seguros', '')),                 disabled=True)
                            st.text_input("Corredora",              value=str(fila_caso.get('Corredora', '')),                           disabled=True)
                            st.text_input("Ajustador senior",       value=str(fila_caso.get('Ajustador senior', '')),                    disabled=True)
                        with col_b:
                            st.text_input("Asegurado",              value=str(fila_caso.get('Asegurado', '')),                           disabled=True)
                            st.text_input("Creado en",              value=str(fila_caso.get('Creado en', '')),                           disabled=True)
                            st.text_input("Divisa",                 value=str(fila_caso.get('Divisa', '')),                              disabled=True)
                            st.text_input("Pérdida bruta",          value=str(fila_caso.get('Perdida bruta (en moneda del caso)', '')),   disabled=True)
                            st.text_input("Monto asegurado",        value=str(fila_caso.get('Monto asegurado (en moneda del caso)', '')), disabled=True)
                            st.text_input("Honorarios (UF)",        value=str(fila_caso.get('Honorarios (UF)', '')),                     disabled=True)
                        with col_c:
                            st.text_input("Último movimiento",      value=str(fila_caso.get('Último movimiento', '')),                   disabled=True)
                            st.text_area("Contenido último mov.",   value=str(fila_caso.get('Contenido último movimiento', '')),         disabled=True, height=100)
                            st.text_input("Indicación probabilidad", value=str(fila_caso.get('Indicación Probabilidad', '')),            disabled=True)
                            st.text_input("Hon. Probables 2026 (UF)", value=str(round(float(fila_caso.get('Hon Probables 2026', 0) or 0), 2)), disabled=True)

                        st.divider()

                        obs_anterior = str(fila_caso.get('Observaciones', '') or '')
                        if obs_anterior.strip():
                            st.markdown("##### 📌 Última Observación Registrada")
                            st.info(obs_anterior)
                        else:
                            st.markdown("##### 📌 Última Observación Registrada")
                            st.warning("Este caso no tiene observaciones previas.")

                        st.divider()

                        st.markdown("##### ✏️ Actualizar Seguimiento")

                        prob_actual = str(fila_caso.get('Probabilidad cierre 2026', '0%'))
                        opciones_prob = ["0%", "25%", "50%", "75%", "100%"]
                        idx_prob_actual = opciones_prob.index(prob_actual) if prob_actual in opciones_prob else 0
                        nueva_prob = st.selectbox(
                            "Probabilidad de cierre 2026",
                            options=opciones_prob,
                            index=idx_prob_actual
                        )

                        nueva_obs = st.text_area(
                            "Nueva Observación (obligatorio) *",
                            placeholder="Escribe aquí el estado actual del caso...",
                            height=120
                        )

                        fecha_actual = fila_caso.get('Fecha probable de facturación', None)
                        nueva_fecha = st.date_input(
                            "Fecha probable de facturación",
                            value=fecha_actual if pd.notna(fecha_actual) and fecha_actual != '' else None
                        )

                        # --- MENSAJE DE ÉXITO PERSISTENTE ---
                        if st.session_state.get("_ultimo_guardado"):
                            st.success(st.session_state["_ultimo_guardado"])
                            st.session_state["_ultimo_guardado"] = None

                        # --- BOTÓN DE GUARDAR ---
                        if st.button("💾 Guardar Seguimiento", type="primary"):
                            if not nueva_obs.strip():
                                st.error("⚠️ La observación es obligatoria. Por favor completa el campo antes de guardar.")
                            else:
                                try:
                                    timestamp_ahora = ahora_cl().strftime("%d/%m/%Y %H:%M")
                                    obs_con_fecha = f"[{timestamp_ahora}] {nueva_obs.strip()}"

                                    hon_uf = float(fila_caso.get("Honorarios (UF)", 0) or 0)
                                    prob_decimal = float(nueva_prob.replace("%", "")) / 100
                                    hon_probables_nuevo = hon_uf * prob_decimal

                                    fecha_str = nueva_fecha.strftime("%Y-%m-%d") if nueva_fecha else ""

                                    df_temp = st.session_state["df_pipeline_activo"].copy()
                                    df_temp["Fecha probable de facturación"] = df_temp["Fecha probable de facturación"].astype(str).replace("NaT", "").replace("None", "")

                                    mask = df_temp["Número de caso"].astype(str) == caso_seleccionado
                                    if not mask.any():
                                        st.error(f"⚠️ No se encontró el caso {caso_seleccionado} en el pipeline activo. Intenta recargar la app.")
                                    else:
                                        df_temp.loc[mask, "Observaciones"] = obs_con_fecha
                                        df_temp.loc[mask, "Fecha probable de facturación"] = fecha_str
                                        df_temp.loc[mask, "Probabilidad cierre 2026"] = nueva_prob
                                        df_temp.loc[mask, "Indicación Probabilidad"] = PROB_MAP.get(nueva_prob, "")
                                        df_temp.loc[mask, "Hon Probables 2026"] = hon_probables_nuevo

                                        st.session_state["df_pipeline_activo"] = df_temp

                                        guardar_local_pipeline(df_temp)
                                        fila_actualizada = df_temp.loc[mask].fillna("").astype(str).iloc[0].to_dict()

                                        try:
                                            guardar_caso_en_nube(fila_actualizada)
                                        except Exception as error_nube:
                                            st.session_state["_ultimo_respaldo_error"] = f"El caso se guardó localmente pero no se pudo sincronizar con la nube: {error_nube}"

                                        st.session_state["_ultimo_guardado"] = f"✅ Caso **{caso_seleccionado}** actualizado correctamente el {timestamp_ahora}."
                                        st.rerun()
                                except Exception as error_guardado:
                                    st.error(f"❌ Ocurrió un error al guardar el seguimiento: {error_guardado}")

else:
    st.info("Sube los archivos para procesar el Pipeline. El sistema reportará ingresos, salidas y aplicará el formato al Excel.")
