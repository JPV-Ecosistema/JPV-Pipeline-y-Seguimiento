import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import io
import os
import json
import re
import time
import uuid
import gspread
from google.oauth2.service_account import Credentials
from forecast_pptx import generar_pptx_forecast

# --- CONTROL DE VERSIONES ---
# Incrementar APP_VERSION cada vez que se publique un cambio relevante en la app.
APP_VERSION = "1.23.1"

def con_reintento(func, intentos=3, espera_inicial=1.5):
    """Ejecuta func() reintentando con backoff exponencial si Google responde 429 (cuota excedida).
    Con varios ajustadores guardando casos a la vez es común chocar momentáneamente con la
    cuota de la API de Sheets; un par de reintentos cortos suele bastar para que pase solo."""
    espera = espera_inicial
    for intento in range(intentos):
        try:
            return func()
        except Exception as e:
            if "429" in str(e) and intento < intentos - 1:
                time.sleep(espera)
                espera *= 2
                continue
            raise

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
                return con_reintento(lambda: doc.worksheet("Pipeline_Backup"))
            except gspread.exceptions.WorksheetNotFound:
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
            con_reintento(lambda: ws.clear())
            con_reintento(lambda: ws.update("A1", matriz))
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
    distintos al mismo tiempo, no se pisan los cambios entre sí.

    Optimizado para consumir la menor cantidad posible de peticiones a la API de Google Sheets
    (cuota compartida entre todos los ajustadores, ya que usan la misma cuenta de servicio):
    la fila del caso y la fecha de actualización se escriben en una sola llamada cuando el caso
    ya existe, y cada llamada a Google reintenta sola si choca con la cuota (429)."""
    fecha_hoy = ahora_cl().strftime("%Y-%m-%d %H:%M:%S")
    try:
        ws = get_backup_sheet()
        if not ws:
            return
        valores = con_reintento(lambda: ws.get_all_values())
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
            # Una sola petición para la fila del caso y la fecha de actualización.
            con_reintento(lambda: ws.batch_update([
                {"range": f"A{num_fila_hoja}", "values": [fila_nueva]},
                {"range": "B1", "values": [[fecha_hoy]]},
            ]))
        else:
            con_reintento(lambda: ws.append_row(fila_nueva))
            con_reintento(lambda: ws.update("B1", [[fecha_hoy]]))

        st.session_state["_ultimo_respaldo_nube"] = fecha_hoy
        st.session_state["_ultimo_respaldo_error"] = None
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
            def _leer():
                doc = client.open_by_url(st.secrets["google_sheet_url"])
                ws = doc.worksheet("Base_Maestra")
                metadata = ws.row_values(1)
                registros = ws.get_all_records(head=2) if len(metadata) >= 2 and metadata[0] == "FECHA_ACTUALIZACION" else None
                return metadata, registros
            metadata, registros = con_reintento(_leer)
            if registros is not None:
                fecha_str = metadata[1]
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
            def _leer():
                doc = client.open_by_url(st.secrets["google_sheet_url"])
                ws = doc.worksheet("Pipeline_Backup")
                metadata = ws.row_values(1)
                registros = ws.get_all_records(head=2) if len(metadata) >= 2 and metadata[0] == "FECHA_ACTUALIZACION" else None
                return metadata, registros
            metadata, registros = con_reintento(_leer)
            if registros is not None:
                fecha_str = metadata[1]
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
    df_norm = df.copy()
    for col in df_norm.columns:
        df_norm[col] = df_norm[col].apply(
            lambda x: ""
            if pd.isna(x) or str(x).strip().lower() in ["nan", "nat", "none", "<na>", "inf", "-inf"]
            else str(x)
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
                ws = con_reintento(lambda: doc.worksheet("Base_Maestra"))
            except gspread.exceptions.WorksheetNotFound:
                ws = doc.add_worksheet(title="Base_Maestra", rows="100", cols="100")
            matriz = [["FECHA_ACTUALIZACION", fecha_hoy]]
            matriz.append(df_norm.columns.astype(str).tolist())
            matriz.extend(df_norm.values.tolist())
            con_reintento(lambda: ws.clear())
            con_reintento(lambda: ws.update("A1", matriz))
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

def generar_excel_pipeline(df):
    """Genera un Excel formateado (con semáforo de color según probabilidad) a partir de un
    DataFrame de casos. Reordena/selecciona siempre las columnas de COLUMNAS_FINALES, sin
    importar qué columnas auxiliares traiga el DataFrame de entrada. Devuelve (bytes, fecha)."""
    fecha_desc = ahora_cl().strftime("%d-%m-%y")
    buffer = io.BytesIO()
    df_excel = df[COLUMNAS_FINALES].copy()
    df_excel['Probabilidad cierre 2026'] = df_excel['Probabilidad cierre 2026'].astype(str).str.replace('%', '').astype(float) / 100

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
    return buffer.getvalue(), fecha_desc

# ---------------------------------------------------------
# LÍNEA BASE FORECAST — motor de cálculo (Fase 1)
# ---------------------------------------------------------
FORECAST_ANIO = 2026
FORECAST_META_DEFAULT = 39400
MESES_NOMBRE = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
                "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

def clasificar_division_forecast(division):
    return 'IE' if 'Ingeniería' in str(division) else 'EM'

def es_engie_ctm_forecast(nickname):
    keywords = ['CTM', 'Wally', 'Engie', 'Mejillones']
    texto = str(nickname).lower()
    return any(k.lower() in texto for k in keywords)

def calcular_corte_forecast(modo="oficial"):
    """Determina el mes de corte y arma la etiqueta N+M.
    modo='oficial': último mes completo cerrado (mes_actual - 1), el forecast normal.
    modo='avance': mes en curso, todavía sin cerrar facturación (mes_actual) — permite
        adelantar un AVANCE del próximo forecast (p.ej. 8+4 antes de que cierre agosto)
        con datos parciales.
    modo='anterior': un mes antes del oficial (mes_actual - 2)."""
    mes_actual = ahora_cl().month
    if modo == "avance":
        mes_corte = mes_actual
    elif modo == "anterior":
        mes_corte = mes_actual - 2
    else:
        mes_corte = mes_actual - 1

    bloqueado = False
    aviso = None
    if mes_corte <= 0:
        bloqueado = True
        aviso = "No hay meses completos disponibles. El análisis se habilitará desde febrero."
    elif mes_corte > 12:
        bloqueado = True
        aviso = "Mes de corte inválido."
    else:
        meses_proyectados_chk = 12 - mes_corte
        if meses_proyectados_chk <= 0:
            bloqueado = True
            aviso = "No quedan meses por proyectar este año."
        elif mes_corte == 11:
            aviso = "La proyección es de un solo mes (diciembre)."

    meses_reales = mes_corte if not bloqueado else 0
    meses_proyectados = (12 - mes_corte) if not bloqueado else 0
    label = f"{meses_reales}+{meses_proyectados}" if not bloqueado else "—"

    if not bloqueado:
        if modo == "avance":
            aviso = (
                f"⚠️ AVANCE {label} con datos PARCIALES de {MESES_NOMBRE[mes_corte]} "
                "(mes en curso, todavía sin cerrar facturación). No reemplaza al forecast oficial."
            )
        elif modo == "anterior":
            aviso = f"⚠️ Emitiendo forecast {label} con datos hasta {MESES_NOMBRE[mes_corte]}."

    return {
        'mes_corte': mes_corte,
        'meses_reales': meses_reales,
        'meses_proyectados': meses_proyectados,
        'label': label,
        'bloqueado': bloqueado,
        'aviso': aviso,
        'nombre_mes_corte': MESES_NOMBRE[mes_corte] if 1 <= mes_corte <= 12 else '—',
        'es_avance': modo == "avance",
    }

def cargar_reporte_produccion_forecast(archivo, anio):
    """Lee la hoja 'Casos' del Reporte de Producción (Faro), header en fila 6."""
    df = pd.read_excel(archivo, sheet_name='Casos', skiprows=5)
    df.columns = [str(c).strip() for c in df.columns]
    df['Fecha factura'] = pd.to_datetime(df['Fecha factura'], errors='coerce')
    df['Honorarios (UF)'] = pd.to_numeric(df['Honorarios (UF)'], errors='coerce')
    df = df[df['Fecha factura'].notna() & (df['Honorarios (UF)'] > 0)].copy()
    df['_anio'] = df['Fecha factura'].dt.year
    df['_mes'] = df['Fecha factura'].dt.month
    df['_div'] = df['División'].apply(clasificar_division_forecast)
    col_perdida = 'Perdida bruta (en moneda del caso)'
    if col_perdida in df.columns:
        df[col_perdida] = pd.to_numeric(df[col_perdida], errors='coerce').fillna(0)
    else:
        df[col_perdida] = 0
        st.warning(
            f"⚠️ El Reporte de Producción no tiene la columna '{col_perdida}': la "
            "clasificación de casos IE menores/mayores del forecast quedará en 0 hasta "
            "que se agregue esa columna al reporte."
        )
    return df

def calcular_ytd_forecast(df_reporte, anio, mes_corte):
    df_ytd = df_reporte[(df_reporte['_anio'] == anio) & (df_reporte['_mes'] <= mes_corte)].copy()
    em_ytd = df_ytd[df_ytd['_div'] == 'EM']['Honorarios (UF)'].sum()
    ie_ytd = df_ytd[df_ytd['_div'] == 'IE']['Honorarios (UF)'].sum()
    mensual = df_ytd.groupby(['_mes', '_div'])['Honorarios (UF)'].sum().unstack(fill_value=0)
    return em_ytd, ie_ytd, mensual, df_ytd

def calcular_proyeccion_em_forecast(mensual, meses_proyectados, df_pipeline):
    meses_disponibles = sorted(mensual.index) if 'EM' in mensual.columns else []
    ultimos_3 = meses_disponibles[-3:]
    valores_3m = [mensual.loc[m, 'EM'] for m in ultimos_3] if ultimos_3 else []
    prom_em_3m = (sum(valores_3m) / len(valores_3m)) if valores_3m else 0.0

    df_pip_em = df_pipeline[df_pipeline['División'].apply(clasificar_division_forecast) == 'EM']
    em_stock = pd.to_numeric(df_pip_em['Hon Probables 2026'], errors='coerce').fillna(0).sum()

    em_promedio_total = prom_em_3m * meses_proyectados
    em_proyectado = max(em_stock, em_promedio_total)

    return {
        'meses_usados': [MESES_NOMBRE[m] for m in ultimos_3],
        'prom_em_3m': prom_em_3m,
        'em_stock': em_stock,
        'em_promedio_total': em_promedio_total,
        'em_proyectado': em_proyectado,
    }

def calcular_proyeccion_ie_forecast(df_ytd, meses_reales, meses_proyectados, df_pipeline):
    ie_ytd_df = df_ytd[df_ytd['_div'] == 'IE']
    menores = ie_ytd_df[ie_ytd_df['Perdida bruta (en moneda del caso)'] < 1000]
    ie_menores_real = menores['Honorarios (UF)'].sum()
    prom_ie_menores = (ie_menores_real / meses_reales) if meses_reales else 0.0
    ie_menores_proj = prom_ie_menores * meses_proyectados

    df_pip_ie = df_pipeline[df_pipeline['División'].apply(clasificar_division_forecast) == 'IE'].copy()
    df_pip_ie['_engie'] = df_pip_ie['Nickname'].apply(es_engie_ctm_forecast)
    df_pip_ie['_perdida'] = pd.to_numeric(df_pip_ie['Perdida bruta (en moneda del caso)'], errors='coerce').fillna(0)
    df_pip_ie['_hon_prob'] = pd.to_numeric(df_pip_ie['Hon Probables 2026'], errors='coerce').fillna(0)

    pip_engie = df_pip_ie[df_pip_ie['_engie']]
    ie_engie_proj = pip_engie['_hon_prob'].sum()

    pip_mayores = df_pip_ie[(~df_pip_ie['_engie']) & (df_pip_ie['_perdida'] >= 1000)]
    ie_mayores_proj = pip_mayores['_hon_prob'].sum()

    ie_proyectado = ie_menores_proj + ie_mayores_proj + ie_engie_proj

    return {
        'ie_menores_real': ie_menores_real,
        'prom_ie_menores': prom_ie_menores,
        'ie_menores_proj': ie_menores_proj,
        'ie_mayores_proj': ie_mayores_proj,
        'ie_engie_proj': ie_engie_proj,
        'ie_proyectado': ie_proyectado,
        'n_casos_mayores': len(pip_mayores),
        'n_casos_engie': len(pip_engie),
    }

def calcular_top10_probabilidad_forecast(df_pipeline, valores_prob, n=10):
    """Top n casos del pipeline cuya 'Probabilidad cierre 2026' esté en valores_prob,
    ordenados por 'Perdida bruta (en moneda del caso)' de mayor a menor."""
    df_filtrado = df_pipeline[
        df_pipeline['Probabilidad cierre 2026'].astype(str).isin(valores_prob)
    ].copy()
    df_filtrado['_perdida_num'] = pd.to_numeric(
        df_filtrado['Perdida bruta (en moneda del caso)'], errors='coerce'
    ).fillna(0)
    df_filtrado = df_filtrado.sort_values('_perdida_num', ascending=False).head(n)
    return [
        {
            'caso': str(fila.get('Número de caso', '')),
            'nickname': str(fila.get('Nickname', '')) or '(sin nombre)',
            'division': str(fila.get('División', '')),
            'ajustador': str(fila.get('Ajustador senior', '')),
            'perdida': fila['_perdida_num'],
            'divisa': str(fila.get('Divisa', '')),
            'probabilidad': str(fila.get('Probabilidad cierre 2026', '')),
        }
        for _, fila in df_filtrado.iterrows()
    ]

st.set_page_config(page_title="JPV Pipeline y Seguimiento", layout="wide")
st.title("🚀 JPV: Pipeline de Facturación Probable")

st.sidebar.header("Carga de Documentos")

df_nube, fecha_nube = cargar_reporte_desde_nube()
if df_nube is not None:
    st.session_state["_df_nube_cache"] = df_nube
    st.session_state["_fecha_nube_cache"] = fecha_nube
elif "_df_nube_cache" in st.session_state:
    # La lectura en vivo falló en este rerun (ej. cuota de Google momentáneamente excedida).
    # Se reutiliza la última copia que sí se pudo leer en esta sesión, para no tirar abajo
    # toda la app (pestañas, filtros) por una falla transitoria en la nube.
    df_nube = st.session_state["_df_nube_cache"]
    fecha_nube = st.session_state["_fecha_nube_cache"]

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
if df_hist_nube is not None:
    st.session_state["_df_hist_nube_cache"] = df_hist_nube
    st.session_state["_fecha_hist_nube_cache"] = fecha_hist_nube
elif "_df_hist_nube_cache" in st.session_state:
    # Mismo motivo que con el reporte: evitar que una falla transitoria de lectura tire
    # abajo toda la app cuando ya tenemos una copia buena de esta sesión.
    df_hist_nube = st.session_state["_df_hist_nube_cache"]
    fecha_hist_nube = st.session_state["_fecha_hist_nube_cache"]

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

historial_manual_autorizado_nuevo = False
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
        firma_historial_manual = f"{archivo_historial.name}_{archivo_historial.size}"
        historial_manual_autorizado_nuevo = st.session_state.get("_ultima_firma_historial_manual") != firma_historial_manual
        st.session_state["_ultima_firma_historial_manual"] = firma_historial_manual
        if historial_manual_autorizado_nuevo:
            st.sidebar.caption("🔁 Este archivo reemplazará el pipeline activo y el respaldo en la nube.")

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
            # Se reemplaza el pipeline activo si es la primera carga de la sesión, o si se acaba
            # de autorizar (con contraseña) un nuevo archivo manual de Pipeline Anterior: en ese
            # caso debe reemplazar el pipeline en curso y el respaldo en la nube, no ser ignorado.
            if 'df_pipeline_activo' not in st.session_state or historial_manual_autorizado_nuevo:
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
            tab_seguimiento, tab_pipeline, tab_masiva, tab_forecast = st.tabs(
                ["🔍 Seguimiento de Caso", "📋 Pipeline General", "📝 Actualización Masiva", "📊 Línea Base Forecast"]
            )

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

                excel_bytes_pipeline, fecha_desc = generar_excel_pipeline(df_editado)

                st.sidebar.divider()
                st.sidebar.download_button(
                    label="📥 Descargar Pipeline Formateado",
                    data=excel_bytes_pipeline,
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

                if len(df_filtrado) > 0:
                    excel_bytes_resumen, fecha_resumen = generar_excel_pipeline(df_filtrado)
                    st.download_button(
                        label=f"📥 Descargar resumen de estos {len(df_filtrado)} casos",
                        data=excel_bytes_resumen,
                        file_name=f"JPV_Resumen_Casos_{fecha_resumen}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="descargar_resumen_seguimiento"
                    )

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

            # ==========================================
            # PESTAÑA ACTUALIZACIÓN MASIVA
            # ==========================================
            with tab_masiva:
                st.subheader("📝 Actualización Masiva de Casos")
                st.caption("Revisa y actualiza varios casos a la vez, en bloques de 10.")

                TAMANO_PAGINA_MASIVA = 10
                OPCIONES_PROB_MASIVA = ["0%", "25%", "50%", "75%", "100%"]

                def semaforo_emoji_masivo(prob):
                    if prob in ["75%", "100%"]:
                        return "🟢"
                    elif prob == "50%":
                        return "🟡"
                    return "🔴"

                df_base_masivo = st.session_state['df_pipeline_activo']

                if "masiva_division_aplicada" not in st.session_state:
                    st.session_state["masiva_division_aplicada"] = "Todas"
                if "masiva_ajustador_aplicado" not in st.session_state:
                    st.session_state["masiva_ajustador_aplicado"] = "Todos"
                if "masiva_pagina" not in st.session_state:
                    st.session_state["masiva_pagina"] = 0
                if "masiva_epoch" not in st.session_state:
                    st.session_state["masiva_epoch"] = 0
                epoch_masivo = st.session_state["masiva_epoch"]

                def calcular_casos_filtrados_masivo():
                    df_f = df_base_masivo.copy()
                    div_aplicada = st.session_state["masiva_division_aplicada"]
                    aj_aplicado = st.session_state["masiva_ajustador_aplicado"]
                    if div_aplicada != "Todas":
                        df_f = df_f[df_f['División'].astype(str).str.strip() == div_aplicada]
                    if aj_aplicado != "Todos":
                        df_f = df_f[df_f['Ajustador senior'].astype(str).str.strip() == aj_aplicado]
                    df_f = df_f.copy()
                    df_f['_num_caso_int'] = pd.to_numeric(
                        df_f['Número de caso'].astype(str).str.replace(r'\.0$', '', regex=True), errors='coerce'
                    )
                    return df_f.sort_values('_num_caso_int', ascending=True)

                casos_filtrados_masivo = calcular_casos_filtrados_masivo()
                total_casos_masivo = len(casos_filtrados_masivo)
                total_paginas_masivo = max(1, -(-total_casos_masivo // TAMANO_PAGINA_MASIVA))
                if st.session_state["masiva_pagina"] > total_paginas_masivo - 1:
                    st.session_state["masiva_pagina"] = total_paginas_masivo - 1
                pagina_actual_masiva = st.session_state["masiva_pagina"]
                inicio_masivo = pagina_actual_masiva * TAMANO_PAGINA_MASIVA
                fin_masivo = min(inicio_masivo + TAMANO_PAGINA_MASIVA, total_casos_masivo)
                casos_pagina_masiva = casos_filtrados_masivo.iloc[inicio_masivo:fin_masivo]

                def fecha_base_str_masivo(valor_fecha):
                    fecha_dt = pd.to_datetime(valor_fecha, errors='coerce')
                    return fecha_dt.strftime("%Y-%m-%d") if pd.notna(fecha_dt) else ""

                # --- Detectar cambios sin guardar en la página actualmente visible ---
                casos_con_cambios_masivo = []
                for _, fila_m in casos_pagina_masiva.iterrows():
                    caso_m = str(fila_m['Número de caso'])
                    prob_base_m = str(fila_m.get('Probabilidad cierre 2026', '0%'))
                    fecha_base_m = fecha_base_str_masivo(fila_m.get('Fecha probable de facturación', None))
                    prob_actual_m = st.session_state.get(f"masiva_prob_{caso_m}_{epoch_masivo}", prob_base_m)
                    obs_actual_m = st.session_state.get(f"masiva_obs_{caso_m}_{epoch_masivo}", "")
                    fecha_widget_m = st.session_state.get(f"masiva_fecha_{caso_m}_{epoch_masivo}", None)
                    fecha_actual_m = fecha_widget_m.strftime("%Y-%m-%d") if fecha_widget_m else ""
                    if prob_actual_m != prob_base_m or obs_actual_m.strip() != "" or fecha_actual_m != fecha_base_m:
                        casos_con_cambios_masivo.append(caso_m)

                hay_cambios_pendientes_masivo = len(casos_con_cambios_masivo) > 0

                # --- Filtros (con botón "Aplicar filtros") ---
                divisiones_disp_masivo = sorted(
                    df_base_masivo['División'].dropna().astype(str).str.strip()
                    .replace('', pd.NA).dropna().unique().tolist()
                )
                ajustadores_disp_masivo = sorted(
                    df_base_masivo['Ajustador senior'].dropna().astype(str).str.strip()
                    .replace('', pd.NA).dropna().unique().tolist()
                )

                col_fm1, col_fm2, col_fm3 = st.columns([2, 2, 1])
                with col_fm1:
                    division_pendiente_masiva = st.selectbox(
                        "División", options=["Todas"] + divisiones_disp_masivo,
                        index=(["Todas"] + divisiones_disp_masivo).index(st.session_state["masiva_division_aplicada"])
                        if st.session_state["masiva_division_aplicada"] in (["Todas"] + divisiones_disp_masivo) else 0,
                        key="masiva_division_pendiente"
                    )
                with col_fm2:
                    ajustador_pendiente_masiva = st.selectbox(
                        "Ajustador Senior", options=["Todos"] + ajustadores_disp_masivo,
                        index=(["Todos"] + ajustadores_disp_masivo).index(st.session_state["masiva_ajustador_aplicado"])
                        if st.session_state["masiva_ajustador_aplicado"] in (["Todos"] + ajustadores_disp_masivo) else 0,
                        key="masiva_ajustador_pendiente"
                    )
                with col_fm3:
                    st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
                    click_aplicar_masivo = st.button("Aplicar filtros", key="masiva_aplicar_filtros")

                if click_aplicar_masivo:
                    if hay_cambios_pendientes_masivo:
                        st.warning(f"⚠️ Tienes {len(casos_con_cambios_masivo)} caso(s) con cambios sin guardar en esta página. Guarda o descarta los cambios antes de aplicar otro filtro.")
                    else:
                        st.session_state["masiva_division_aplicada"] = division_pendiente_masiva
                        st.session_state["masiva_ajustador_aplicado"] = ajustador_pendiente_masiva
                        st.session_state["masiva_pagina"] = 0
                        st.rerun()

                st.markdown("---")

                # --- Paginación ---
                col_pm1, col_pm2 = st.columns([2, 1])
                with col_pm1:
                    if total_casos_masivo > 0:
                        st.markdown(f"**Mostrando casos {inicio_masivo + 1}–{fin_masivo} de {total_casos_masivo}**")
                        excel_bytes_masivo, fecha_resumen_masivo = generar_excel_pipeline(casos_filtrados_masivo)
                        st.download_button(
                            label=f"📥 Descargar resumen de los {total_casos_masivo} casos filtrados",
                            data=excel_bytes_masivo,
                            file_name=f"JPV_Resumen_Casos_{fecha_resumen_masivo}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="descargar_resumen_masiva"
                        )
                    else:
                        st.markdown("**No hay casos que coincidan con los filtros seleccionados.**")
                with col_pm2:
                    col_pant, col_psig = st.columns(2)
                    with col_pant:
                        click_anterior_masivo = st.button(
                            "⬅ Anterior", key="masiva_pagina_anterior",
                            disabled=(pagina_actual_masiva <= 0), use_container_width=True
                        )
                    with col_psig:
                        click_siguiente_masivo = st.button(
                            "Siguiente ➡", key="masiva_pagina_siguiente",
                            disabled=(pagina_actual_masiva >= total_paginas_masivo - 1), use_container_width=True
                        )

                if click_anterior_masivo or click_siguiente_masivo:
                    if hay_cambios_pendientes_masivo:
                        st.warning(f"⚠️ Tienes {len(casos_con_cambios_masivo)} caso(s) con cambios sin guardar en esta página. Guarda o descarta los cambios antes de cambiar de página.")
                    else:
                        st.session_state["masiva_pagina"] += -1 if click_anterior_masivo else 1
                        st.rerun()

                st.caption("🟢 Alta · 🟡 Media · 🔴 Baja probabilidad — 🔵 Cambios sin guardar")

                # --- Tarjetas de casos ---
                for _, fila_m in casos_pagina_masiva.iterrows():
                    caso_m = str(fila_m['Número de caso'])
                    prob_base_m = str(fila_m.get('Probabilidad cierre 2026', '0%'))
                    if prob_base_m not in OPCIONES_PROB_MASIVA:
                        prob_base_m = "0%"
                    obs_referencia_m = str(fila_m.get('Observaciones', '') or '').strip()
                    fecha_base_dt_m = pd.to_datetime(fila_m.get('Fecha probable de facturación', None), errors='coerce')
                    fecha_base_valor_m = fecha_base_dt_m.date() if pd.notna(fecha_base_dt_m) else None

                    try:
                        fecha_creacion_m = pd.to_datetime(fila_m.get('Creado en', ''), errors='coerce')
                        dias_activo_m = (ahora_cl() - fecha_creacion_m).days if pd.notna(fecha_creacion_m) else None
                    except Exception:
                        dias_activo_m = None

                    con_cambios_m = caso_m in casos_con_cambios_masivo

                    with st.container(border=True):
                        if con_cambios_m:
                            st.markdown("🔵 **Cambios sin guardar**")

                        col_c1, col_c2, col_c3, col_c4, col_c5 = st.columns([2, 0.8, 1.6, 3, 1.6])
                        with col_c1:
                            st.markdown(f"**{caso_m}**")
                            st.caption(str(fila_m.get('Asegurado', '')))
                            st.caption(str(fila_m.get('Corredora', '')))
                        with col_c2:
                            st.markdown(f"<div style='text-align:center; font-size:22px; font-weight:700;'>{dias_activo_m if dias_activo_m is not None else '—'}</div><div style='text-align:center; font-size:11px; color:#98a2b3;'>DÍAS</div>", unsafe_allow_html=True)
                        with col_c3:
                            st.selectbox(
                                f"{semaforo_emoji_masivo(st.session_state.get(f'masiva_prob_{caso_m}_{epoch_masivo}', prob_base_m))} Probabilidad 2026",
                                options=OPCIONES_PROB_MASIVA,
                                index=OPCIONES_PROB_MASIVA.index(prob_base_m),
                                key=f"masiva_prob_{caso_m}_{epoch_masivo}"
                            )
                        with col_c4:
                            if obs_referencia_m:
                                st.caption(f"🕐 Última gestión: {obs_referencia_m}")
                            else:
                                st.caption("🕐 Sin observaciones previas registradas.")
                            st.text_input(
                                "Observación (obligatoria si cambia la probabilidad)",
                                key=f"masiva_obs_{caso_m}_{epoch_masivo}",
                                placeholder="Escribe una nueva observación..."
                            )
                        with col_c5:
                            st.date_input(
                                "Fecha probable",
                                value=fecha_base_valor_m,
                                key=f"masiva_fecha_{caso_m}_{epoch_masivo}"
                            )

                if total_casos_masivo > 0:
                    st.markdown("---")
                    col_bm1, col_bm2, col_bm3 = st.columns([2.5, 1, 1.5])
                    with col_bm1:
                        if hay_cambios_pendientes_masivo:
                            st.markdown(f"🔵 **{len(casos_con_cambios_masivo)} de {len(casos_pagina_masiva)} casos con cambios sin guardar**")
                        else:
                            st.caption("Sin cambios pendientes en esta página.")
                    with col_bm2:
                        click_descartar_masivo = st.button(
                            "Descartar cambios", key="masiva_descartar", disabled=not hay_cambios_pendientes_masivo
                        )
                    with col_bm3:
                        click_guardar_masivo = st.button(
                            "💾 Guardar los 10 casos" if len(casos_pagina_masiva) == TAMANO_PAGINA_MASIVA else f"💾 Guardar los {len(casos_pagina_masiva)} casos",
                            key="masiva_guardar", type="primary", disabled=not hay_cambios_pendientes_masivo,
                            use_container_width=True
                        )

                    if click_descartar_masivo:
                        # No se puede reasignar el session_state de un widget ya instanciado en
                        # este mismo run (Streamlit lo prohíbe). En vez de eso, se incrementa la
                        # "época": los widgets de la próxima ejecución usarán claves nuevas y
                        # partirán limpios desde los valores base del pipeline.
                        st.session_state["masiva_epoch"] += 1
                        st.rerun()

                    if click_guardar_masivo:
                        errores_validacion_masivo = []
                        for caso_m in casos_con_cambios_masivo:
                            prob_actual_m = st.session_state.get(f"masiva_prob_{caso_m}_{epoch_masivo}")
                            obs_actual_m = st.session_state.get(f"masiva_obs_{caso_m}_{epoch_masivo}", "")
                            fila_original_m = casos_pagina_masiva[casos_pagina_masiva['Número de caso'].astype(str) == caso_m].iloc[0]
                            prob_base_m = str(fila_original_m.get('Probabilidad cierre 2026', '0%'))
                            if prob_actual_m != prob_base_m and not obs_actual_m.strip():
                                errores_validacion_masivo.append(caso_m)

                        if errores_validacion_masivo:
                            st.error(f"⚠️ Estos casos cambiaron de probabilidad y necesitan una observación antes de guardar: {', '.join(errores_validacion_masivo)}")
                        else:
                            try:
                                timestamp_masivo = ahora_cl().strftime("%d/%m/%Y %H:%M")
                                df_temp_masivo = st.session_state["df_pipeline_activo"].copy()
                                df_temp_masivo["Fecha probable de facturación"] = df_temp_masivo["Fecha probable de facturación"].astype(str).replace("NaT", "").replace("None", "")

                                for caso_m in casos_con_cambios_masivo:
                                    mask_m = df_temp_masivo["Número de caso"].astype(str) == caso_m
                                    if not mask_m.any():
                                        continue

                                    nueva_prob_m = st.session_state.get(f"masiva_prob_{caso_m}_{epoch_masivo}")
                                    nueva_obs_m = st.session_state.get(f"masiva_obs_{caso_m}_{epoch_masivo}", "").strip()
                                    nueva_fecha_m = st.session_state.get(f"masiva_fecha_{caso_m}_{epoch_masivo}")
                                    fecha_str_m = nueva_fecha_m.strftime("%Y-%m-%d") if nueva_fecha_m else ""

                                    hon_uf_m = float(df_temp_masivo.loc[mask_m, "Honorarios (UF)"].iloc[0] or 0)
                                    prob_decimal_m = float(nueva_prob_m.replace("%", "")) / 100
                                    hon_probables_m = hon_uf_m * prob_decimal_m

                                    df_temp_masivo.loc[mask_m, "Probabilidad cierre 2026"] = nueva_prob_m
                                    df_temp_masivo.loc[mask_m, "Indicación Probabilidad"] = PROB_MAP.get(nueva_prob_m, "")
                                    df_temp_masivo.loc[mask_m, "Fecha probable de facturación"] = fecha_str_m
                                    df_temp_masivo.loc[mask_m, "Hon Probables 2026"] = hon_probables_m
                                    if nueva_obs_m:
                                        df_temp_masivo.loc[mask_m, "Observaciones"] = f"[{timestamp_masivo}] {nueva_obs_m}"

                                st.session_state["df_pipeline_activo"] = df_temp_masivo
                                guardar_local_pipeline(df_temp_masivo)

                                errores_sync_masivo = []
                                for caso_m in casos_con_cambios_masivo:
                                    mask_m = df_temp_masivo["Número de caso"].astype(str) == caso_m
                                    if not mask_m.any():
                                        continue
                                    fila_actualizada_m = df_temp_masivo.loc[mask_m].fillna("").astype(str).iloc[0].to_dict()
                                    try:
                                        guardar_caso_en_nube(fila_actualizada_m)
                                    except Exception as error_nube_m:
                                        errores_sync_masivo.append(f"{caso_m} ({error_nube_m})")

                                # Igual que en "Descartar cambios": no se puede tocar el session_state
                                # de un widget ya instanciado en este run, así que se incrementa la
                                # época para que la próxima ejecución parta limpia desde el nuevo
                                # pipeline ya guardado (sin cambios pendientes).
                                st.session_state["masiva_epoch"] += 1

                                if errores_sync_masivo:
                                    st.session_state["_ultimo_respaldo_error"] = f"Guardado localmente, pero no se pudo sincronizar con la nube: {', '.join(errores_sync_masivo)}"

                                st.session_state["_masiva_ultimo_guardado"] = f"{len(casos_con_cambios_masivo)} casos actualizados correctamente"
                                st.session_state["_masiva_ultimo_guardado_hora"] = timestamp_masivo
                                st.rerun()
                            except Exception as error_guardado_masivo:
                                st.error(f"❌ Ocurrió un error al guardar los casos: {error_guardado_masivo}")

                if st.session_state.get("_masiva_ultimo_guardado"):
                    st.toast(f"✅ {st.session_state['_masiva_ultimo_guardado']} · {st.session_state.get('_masiva_ultimo_guardado_hora', '')}", icon="✅")
                    st.session_state["_masiva_ultimo_guardado"] = None

            # ==========================================
            # PESTAÑA LÍNEA BASE FORECAST (FASE 1: motor de cálculo + dashboard)
            # ==========================================
            with tab_forecast:
                st.subheader("📊 Línea Base Forecast — Ingeniería y Equipo Móvil")
                st.caption(
                    "Motor de cálculo, dashboard de revisión y generación de la presentación PPTX. "
                    "El historial comparativo entre forecasts llega en una fase siguiente."
                )

                archivo_produccion = st.file_uploader(
                    "Reporte de Producción del mes (Excel exportado desde Faro, hoja 'Casos')",
                    type=["xlsx"], key="forecast_reporte_produccion"
                )
                modo_forecast = st.radio(
                    "Período a emitir",
                    options=["oficial", "avance", "anterior"],
                    format_func=lambda m: {
                        "oficial": "Forecast oficial (último mes cerrado)",
                        "avance": "AVANCE del mes en curso (aún sin cerrar)",
                        "anterior": "Forecast del mes anterior",
                    }[m],
                    horizontal=True, key="forecast_modo_periodo",
                )

                corte_forecast = calcular_corte_forecast(modo_forecast)

                if corte_forecast['bloqueado']:
                    st.warning(f"⚠️ {corte_forecast['aviso']}")
                elif archivo_produccion is None:
                    st.info("Sube el Reporte de Producción del mes para calcular el forecast.")
                else:
                    try:
                        df_reporte_forecast = cargar_reporte_produccion_forecast(archivo_produccion, FORECAST_ANIO)
                        if df_reporte_forecast.empty:
                            st.warning(f"⚠️ El reporte no tiene registros facturados válidos para {FORECAST_ANIO}.")
                            df_reporte_forecast = None
                    except Exception as error_reporte_forecast:
                        st.error(f"❌ No se pudo procesar el Reporte de Producción: {error_reporte_forecast}")
                        df_reporte_forecast = None

                    if df_reporte_forecast is not None:
                        if corte_forecast['aviso']:
                            if modo_forecast == "avance":
                                st.error(corte_forecast['aviso'])
                            elif modo_forecast == "anterior":
                                st.info(corte_forecast['aviso'])
                            else:
                                st.warning(corte_forecast['aviso'])

                        # --- Panel 1: Parámetros ---
                        st.markdown("---")
                        st.markdown("#### 1️⃣ Parámetros del Forecast")
                        col_par1, col_par2 = st.columns(2)
                        with col_par1:
                            st.metric("Forecast detectado", corte_forecast['label'])
                            st.caption(f"Corte: {corte_forecast['nombre_mes_corte']} {FORECAST_ANIO}")
                        with col_par2:
                            meta_anual_forecast = st.number_input(
                                "Meta anual (UF)", min_value=0.0,
                                value=float(st.session_state.get("forecast_meta_valor", FORECAST_META_DEFAULT)),
                                step=100.0, key="forecast_meta_valor"
                            )

                        # --- Cálculos base ---
                        em_ytd, ie_ytd, mensual_forecast, df_ytd_forecast = calcular_ytd_forecast(
                            df_reporte_forecast, FORECAST_ANIO, corte_forecast['mes_corte']
                        )
                        ytd_total_forecast = em_ytd + ie_ytd
                        df_pipeline_forecast = st.session_state['df_pipeline_activo']
                        proy_em = calcular_proyeccion_em_forecast(
                            mensual_forecast, corte_forecast['meses_proyectados'], df_pipeline_forecast
                        )
                        proy_ie = calcular_proyeccion_ie_forecast(
                            df_ytd_forecast, corte_forecast['meses_reales'], corte_forecast['meses_proyectados'],
                            df_pipeline_forecast
                        )
                        if len(proy_em['meses_usados']) < 3:
                            st.caption(
                                f"ℹ️ El promedio EM se calculó con solo {len(proy_em['meses_usados'])} mes(es) "
                                f"disponible(s) en el reporte cargado — aún no hay 3 meses del año para promediar."
                            )

                        # --- Panel 2: YTD ---
                        st.markdown("---")
                        st.markdown("#### 2️⃣ Facturación YTD (Real)")
                        col_y1, col_y2, col_y3 = st.columns(3)
                        col_y1.metric("Equipo Móvil YTD", f"{em_ytd:,.2f} UF")
                        col_y2.metric("Ingeniería y Energía YTD", f"{ie_ytd:,.2f} UF")
                        col_y3.metric("Total Consolidado YTD", f"{ytd_total_forecast:,.2f} UF")
                        st.caption(
                            f"Ritmo EM ({', '.join(proy_em['meses_usados']) or '—'}): {proy_em['prom_em_3m']:,.0f} UF/mes  ·  "
                            f"Ritmo IE casos menores: {proy_ie['prom_ie_menores']:,.0f} UF/mes"
                        )

                        # --- Panel 3: Proyección editable ---
                        st.markdown("---")
                        st.markdown("#### 3️⃣ Proyección (editable antes de confirmar)")

                        col_e1, col_e2 = st.columns(2)
                        with col_e1:
                            st.markdown("**Equipo Móvil**")
                            em_stock_edit = st.number_input(
                                "EM — Stock pipeline (Hon. Probables)", value=float(proy_em['em_stock']),
                                step=10.0, key="forecast_em_stock"
                            )
                            em_promedio_edit = st.number_input(
                                f"EM — Promedio {len(proy_em['meses_usados'])}m × {corte_forecast['meses_proyectados']} meses",
                                value=float(proy_em['em_promedio_total']), step=10.0, key="forecast_em_promedio"
                            )
                            em_proyectado_edit = max(em_stock_edit, em_promedio_edit)
                            st.caption(f"EM proyectado usado (máx. de los dos anteriores): **{em_proyectado_edit:,.0f} UF**")
                        with col_e2:
                            st.markdown("**Ingeniería y Energía**")
                            ie_menores_edit = st.number_input(
                                f"IE — Casos menores ({proy_ie['prom_ie_menores']:,.0f}/mes × {corte_forecast['meses_proyectados']})",
                                value=float(proy_ie['ie_menores_proj']), step=10.0, key="forecast_ie_menores"
                            )
                            ie_mayores_edit = st.number_input(
                                f"IE — Casos mayores pipeline ({proy_ie['n_casos_mayores']} casos, Hon. Probables)",
                                value=float(proy_ie['ie_mayores_proj']), step=10.0, key="forecast_ie_mayores"
                            )
                            ie_engie_edit = st.number_input(
                                f"IE — Engie/CTM pipeline ({proy_ie['n_casos_engie']} casos, Hon. Probables)",
                                value=float(proy_ie['ie_engie_proj']), step=10.0, key="forecast_ie_engie"
                            )
                            ie_proyectado_edit = ie_menores_edit + ie_mayores_edit + ie_engie_edit
                            st.caption(f"IE proyectado total: **{ie_proyectado_edit:,.0f} UF**")

                        # --- Casos en Proceso Administrativo de Facturación ---
                        st.markdown("---")
                        if "forecast_casos_admin" not in st.session_state:
                            st.session_state["forecast_casos_admin"] = []

                        with st.expander("📋 Casos en Proceso Administrativo de Facturación", expanded=False):
                            casos_admin_ids_eliminar = []
                            for caso_admin in st.session_state["forecast_casos_admin"]:
                                col_ca1, col_ca2, col_ca3 = st.columns([3, 1.5, 0.5])
                                with col_ca1:
                                    nombre_admin_nuevo = st.text_input(
                                        "Nombre/descripción del caso", value=caso_admin["nombre"],
                                        key=f"forecast_admin_nombre_{caso_admin['id']}"
                                    )
                                with col_ca2:
                                    monto_admin_nuevo = st.number_input(
                                        "Monto UF", value=float(caso_admin["monto"]), step=10.0,
                                        key=f"forecast_admin_monto_{caso_admin['id']}"
                                    )
                                with col_ca3:
                                    st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
                                    if st.button("🗑️", key=f"forecast_admin_del_{caso_admin['id']}"):
                                        casos_admin_ids_eliminar.append(caso_admin['id'])
                                caso_admin["nombre"] = nombre_admin_nuevo
                                caso_admin["monto"] = monto_admin_nuevo

                            if casos_admin_ids_eliminar:
                                st.session_state["forecast_casos_admin"] = [
                                    c for c in st.session_state["forecast_casos_admin"]
                                    if c['id'] not in casos_admin_ids_eliminar
                                ]
                                st.rerun()

                            if st.button("➕ Agregar caso", key="forecast_admin_agregar"):
                                st.session_state["forecast_casos_admin"].append(
                                    {"id": str(uuid.uuid4()), "nombre": "", "monto": 0.0}
                                )
                                st.rerun()

                            total_admin_forecast = sum(c['monto'] for c in st.session_state["forecast_casos_admin"])
                            st.markdown(f"**Total en proceso administrativo: {total_admin_forecast:,.0f} UF**")

                        # --- Panel 4: Resumen de cierres ---
                        st.markdown("---")
                        st.markdown("#### 4️⃣ Resumen de Cierres Proyectados")

                        em_total_forecast = em_ytd + em_proyectado_edit
                        ie_total_forecast = ie_ytd + ie_proyectado_edit
                        cierre_sin_admin = em_total_forecast + ie_total_forecast
                        cierre_con_admin = cierre_sin_admin + total_admin_forecast
                        gap_sin = cierre_sin_admin - meta_anual_forecast
                        gap_con = cierre_con_admin - meta_anual_forecast
                        cum_sin = (cierre_sin_admin / meta_anual_forecast * 100) if meta_anual_forecast else 0
                        cum_con = (cierre_con_admin / meta_anual_forecast * 100) if meta_anual_forecast else 0

                        col_r1, col_r2 = st.columns(2)
                        with col_r1:
                            st.metric(
                                "Cierre sin proceso administrativo",
                                f"{cierre_sin_admin:,.0f} UF",
                                f"{gap_sin:+,.0f} UF vs. meta ({cum_sin:.1f}%)"
                            )
                        with col_r2:
                            if total_admin_forecast > 0:
                                st.metric(
                                    "Cierre con proceso administrativo",
                                    f"{cierre_con_admin:,.0f} UF",
                                    f"{gap_con:+,.0f} UF vs. meta ({cum_con:.1f}%)"
                                )
                            else:
                                st.caption("No hay casos en proceso administrativo cargados.")

                        # --- Panel 5: Gráfico de evolución (preview) ---
                        st.markdown("---")
                        st.markdown("#### 5️⃣ Evolución Real + Proyección (preview)")

                        meses_real_idx = list(range(1, corte_forecast['mes_corte'] + 1))
                        acumulado_real = []
                        acum = 0.0
                        for m in meses_real_idx:
                            valor_mes = 0.0
                            if m in mensual_forecast.index:
                                valor_mes = float(mensual_forecast.loc[m].sum())
                            acum += valor_mes
                            acumulado_real.append(acum)

                        meses_proy_idx = list(range(corte_forecast['mes_corte'] + 1, 13))
                        incremento_mensual_sin = (
                            (em_proyectado_edit + ie_proyectado_edit) / corte_forecast['meses_proyectados']
                            if corte_forecast['meses_proyectados'] else 0
                        )
                        incremento_mensual_con = (
                            (em_proyectado_edit + ie_proyectado_edit + total_admin_forecast) / corte_forecast['meses_proyectados']
                            if corte_forecast['meses_proyectados'] else 0
                        )

                        filas_grafico = []
                        for i, m in enumerate(meses_real_idx):
                            filas_grafico.append({
                                'Mes': MESES_NOMBRE[m][:3], 'Real': acumulado_real[i],
                                'Proyección (sin proc. admin.)': None, 'Proyección (con proc. admin.)': None,
                            })
                        base_sin = acumulado_real[-1] if acumulado_real else 0
                        base_con = base_sin
                        # Punto de conexión: el mes de corte también inicia las líneas de
                        # proyección (mismo valor que lo real), para que las curvas se unan
                        # visualmente sin salto en el gráfico.
                        if filas_grafico:
                            filas_grafico[-1]['Proyección (sin proc. admin.)'] = base_sin
                            if total_admin_forecast > 0:
                                filas_grafico[-1]['Proyección (con proc. admin.)'] = base_con
                        for i, m in enumerate(meses_proy_idx, start=1):
                            base_sin += incremento_mensual_sin
                            base_con += incremento_mensual_con
                            filas_grafico.append({
                                'Mes': MESES_NOMBRE[m][:3], 'Real': None,
                                'Proyección (sin proc. admin.)': base_sin,
                                'Proyección (con proc. admin.)': base_con if total_admin_forecast > 0 else None,
                            })
                        df_grafico_forecast = pd.DataFrame(filas_grafico).set_index('Mes')
                        st.line_chart(df_grafico_forecast)
                        st.caption(f"Línea de meta: {meta_anual_forecast:,.0f} UF")

                        # --- Panel 6: Generación de la presentación PPTX ---
                        st.markdown("---")
                        st.markdown("#### 🎯 Generar Presentación PPTX")

                        pipeline_bruto_total = pd.to_numeric(
                            df_pipeline_forecast['Honorarios (UF)'], errors='coerce'
                        ).fillna(0).sum()

                        valores_em_mensual = [
                            float(mensual_forecast.loc[m, 'EM']) if 'EM' in mensual_forecast.columns and m in mensual_forecast.index else 0.0
                            for m in meses_real_idx
                        ]
                        valores_ie_mensual = [
                            float(mensual_forecast.loc[m, 'IE']) if 'IE' in mensual_forecast.columns and m in mensual_forecast.index else 0.0
                            for m in meses_real_idx
                        ]

                        pct_stock_cubre = (
                            (proy_em['em_stock'] / proy_em['em_promedio_total'] * 100)
                            if proy_em['em_promedio_total'] else 0
                        )

                        top10_100_forecast = calcular_top10_probabilidad_forecast(df_pipeline_forecast, ['100%'])
                        top10_75_forecast = calcular_top10_probabilidad_forecast(df_pipeline_forecast, ['75%'])
                        top10_menor50_forecast = calcular_top10_probabilidad_forecast(df_pipeline_forecast, ['0%', '25%'])

                        datos_pptx = {
                            'anio': FORECAST_ANIO,
                            'label': corte_forecast['label'],
                            'mes_corte': corte_forecast['mes_corte'],
                            'nombre_mes_corte': corte_forecast['nombre_mes_corte'],
                            'nombre_mes_inicio': MESES_NOMBRE[1],
                            'meses_reales': corte_forecast['meses_reales'],
                            'meses_proyectados': corte_forecast['meses_proyectados'],
                            'fecha_emision': f"{ahora_cl().day} de {MESES_NOMBRE[ahora_cl().month].lower()} de {ahora_cl().year}",
                            'meta': meta_anual_forecast,
                            'em_ytd': em_ytd, 'ie_ytd': ie_ytd, 'ytd_total': ytd_total_forecast,
                            'em_proyectado': em_proyectado_edit, 'ie_proyectado': ie_proyectado_edit,
                            'em_total': em_total_forecast, 'ie_total': ie_total_forecast,
                            'cierre_sin_admin': cierre_sin_admin, 'cierre_con_admin': cierre_con_admin,
                            'gap_sin': gap_sin, 'gap_con': gap_con, 'cum_sin': cum_sin, 'cum_con': cum_con,
                            'total_admin': total_admin_forecast,
                            'casos_admin': st.session_state["forecast_casos_admin"],
                            'pipeline_bruto_total': pipeline_bruto_total,
                            'proy_em': proy_em, 'proy_ie': proy_ie,
                            'es_avance': corte_forecast['es_avance'],
                            'fuente_texto': (
                                ("AVANCE — " if corte_forecast['es_avance'] else "")
                                + f"Fuente: Reporte de Producción Faro ({MESES_NOMBRE[1][:3]}-{corte_forecast['nombre_mes_corte'][:3]} {FORECAST_ANIO}) y Pipeline JPV"
                            ),
                            'grafico_meses': [row['Mes'] for row in filas_grafico],
                            'grafico_real': [row['Real'] for row in filas_grafico],
                            'grafico_proy_sin': [row['Proyección (sin proc. admin.)'] for row in filas_grafico],
                            'grafico_proy_con': [row['Proyección (con proc. admin.)'] for row in filas_grafico],
                            'meses_mensual': [MESES_NOMBRE[m][:3] for m in meses_real_idx],
                            'valores_em_mensual': valores_em_mensual,
                            'valores_ie_mensual': valores_ie_mensual,
                            'bullets_linea_base': [
                                f"Ritmo EM últimos {len(proy_em['meses_usados'])} meses ({'-'.join(m[:3] for m in proy_em['meses_usados']) or '—'}): {proy_em['prom_em_3m']:,.0f} UF/mes.".replace(',', '.'),
                                f"Ritmo IE casos menores (<1.000 UF pérdida): {proy_ie['prom_ie_menores']:,.0f} UF/mes.".replace(',', '.'),
                                f"Total consolidado YTD: {ytd_total_forecast:,.2f} UF.".replace(',', '.'),
                            ],
                            'bullets_accion_em': [
                                f"La proyección usa el promedio de los últimos {len(proy_em['meses_usados'])} meses ({proy_em['prom_em_3m']:,.0f} UF/mes) como base.".replace(',', '.'),
                                f"El stock en pipeline ({proy_em['em_stock']:,.0f} UF) cubre el {pct_stock_cubre:.0f}% de la proyección — el resto depende de nuevas asignaciones.".replace(',', '.'),
                            ],
                            'bullets_accion_ie': [
                                f"Casos menores ({proy_ie['prom_ie_menores']:,.0f} UF/mes): mantener el ritmo histórico es clave para la proyección.".replace(',', '.'),
                                f"Casos mayores pipeline: {proy_ie['n_casos_mayores']} caso(s) ponderados por probabilidad, total {proy_ie['ie_mayores_proj']:,.0f} UF.".replace(',', '.'),
                                f"Engie/CTM: {proy_ie['n_casos_engie']} caso(s), total {proy_ie['ie_engie_proj']:,.0f} UF.".replace(',', '.'),
                            ],
                            'top10_100': top10_100_forecast,
                            'top10_75': top10_75_forecast,
                            'top10_menor50': top10_menor50_forecast,
                        }

                        texto_boton_pptx = (
                            "🎯 Generar AVANCE PPTX" if corte_forecast['es_avance']
                            else "🎯 Generar Presentación PPTX"
                        )
                        if st.button(texto_boton_pptx, type="primary", key="forecast_generar_pptx"):
                            try:
                                pptx_bytes = generar_pptx_forecast(datos_pptx)
                                fecha_archivo = ahora_cl().strftime("%d-%m-%y")
                                sufijo_avance = "_AVANCE" if corte_forecast['es_avance'] else ""
                                nombre_archivo = f"Linea_Base_{corte_forecast['label']}{sufijo_avance}_{fecha_archivo}.pptx"
                                st.session_state["_forecast_pptx_bytes"] = pptx_bytes
                                st.session_state["_forecast_pptx_nombre"] = nombre_archivo
                                if corte_forecast['es_avance']:
                                    st.success("✅ AVANCE generado correctamente. Recuerda que usa datos parciales, no es el forecast oficial.")
                                else:
                                    st.success("✅ Presentación generada correctamente.")
                            except Exception as error_pptx:
                                st.error(f"❌ No se pudo generar la presentación: {error_pptx}")

                        if st.session_state.get("_forecast_pptx_bytes"):
                            st.download_button(
                                label=f"📥 Descargar {st.session_state['_forecast_pptx_nombre']}",
                                data=st.session_state["_forecast_pptx_bytes"],
                                file_name=st.session_state["_forecast_pptx_nombre"],
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                key="forecast_descargar_pptx"
                            )

else:
    st.info("Sube los archivos para procesar el Pipeline. El sistema reportará ingresos, salidas y aplicará el formato al Excel.")
