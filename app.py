import streamlit as st
import pandas as pd
from datetime import datetime
import io
import re

# --- CONFIGURACIÓN DE ETIQUETAS ---
PROB_MAP = {
    "0%": "Nula",
    "25%": "Remota",
    "50%": "Podría Ser",
    "75%": "Altamente probable",
    "100%": "Cierta"
}

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
archivo_nuevo = st.sidebar.file_uploader("1. Nuevo Reporte de Acciones (Excel)", type=["xlsx"])
archivo_historial = st.sidebar.file_uploader("2. Pipeline Anterior (Excel Maestro)", type=["xlsx"])

if archivo_nuevo and archivo_historial:
    # 1. Cargar Reporte Nuevo (Títulos en fila 6)
    df_nuevo = pd.read_excel(archivo_nuevo, skiprows=5)
    df_nuevo.columns = [str(c).strip() for c in df_nuevo.columns]
    df_nuevo = df_nuevo.dropna(how='all', axis=0)

    # 2. Identificar Automáticamente la Última Hoja por Fecha o Posición
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
        st.error("No se pudo identificar la hoja de datos en el historial.")
    else:
        st.info(f"Última actualización detectada: **{hoja_maestra}**")
        
        # Leer historial detectando fila de encabezado
        df_hist_raw = pd.read_excel(xl_historial, sheet_name=hoja_maestra, header=None)
        fila_h = 0
        for i, row in df_hist_raw.iterrows():
            if any(str(val).strip() in ['Número de caso', 'Numero de caso', 'N° caso', 'Caso'] for val in row.values):
                fila_h = i
                break
        
        df_hist = pd.read_excel(xl_historial, sheet_name=hoja_maestra, skiprows=fila_h)
        df_hist.columns = [str(c).strip() for c in df_hist.columns]
        
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
            # Convertir decimales a strings de porcentaje (ej: 0.75 -> 75%)
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

            # --- RESUMEN DE CASOS NUEVOS ---
            st.subheader("Panel de Gestión")
            casos_viejos = set(df_hist[col_llave].unique())
            nuevos_detectados = [c for c in df_nuevo[col_llave].unique() if c not in casos_viejos]
            st.success(f"🆕 **Resumen de Actualización:** Se han identificado **{len(nuevos_detectados)} casos nuevos** que no estaban en la hoja {hoja_maestra}.")

            # --- FUNCIÓN DE ESTILO (SEMÁFORO) ---
            def color_semaforo(val):
                color = ''
                if val in ["75%", "100%"]:
                    color = 'background-color: #c6efce; color: #006100;' # Verde
                elif val == "50%":
                    color = 'background-color: #ffeb9c; color: #9c5700;' # Amarillo
                elif val in ["0%", "25%"]:
                    color = 'background-color: #ffc7ce; color: #9c0006;' # Rojo
                return color

            # Aplicar estilo al DataFrame
            df_styled = df_final.style.map(color_semaforo, subset=['Probabilidad cierre 2026'])

            # --- EDITOR DE DATOS ---
            df_editado = st.data_editor(
                df_styled,
                column_config={
                    "Probabilidad cierre 2026": st.column_config.SelectboxColumn(
                        "Probabilidad (%)", 
                        options=["0%", "25%", "50%", "75%", "100%"],
                        help="Seleccione la probabilidad de cierre"
                    ),
                    "Fecha probable de facturación": st.column_config.DateColumn("Fecha Fact."),
                    "Observaciones": st.column_config.TextColumn("Observaciones", width="large")
                },
                hide_index=True, 
                use_container_width=True
            )

            # --- RECALCULAR TRAS EDICIÓN ---
            prob_num_final = df_editado['Probabilidad cierre 2026'].str.replace('%', '').astype(float) / 100
            df_editado['Hon Probables 2026'] = df_editado['Honorarios (UF)'] * prob_num_final
            df_editado['Indicación Probabilidad'] = df_editado['Probabilidad cierre 2026'].map(PROB_MAP)
            
            st.metric("FACTURACIÓN PROBABLE TOTAL (UF)", f"{df_editado['Hon Probables 2026'].sum():,.2f}")

            # --- DESCARGA (CONVERSIÓN A DECIMAL PARA EXCEL) ---
            fecha_desc = datetime.now().strftime("%d-%m-%y")
            buffer = io.BytesIO()
            df_excel = df_editado.copy()
            df_excel['Probabilidad cierre 2026'] = df_excel['Probabilidad cierre 2026'].str.replace('%', '').astype(float) / 100

            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_excel.to_excel(writer, sheet_name=f"Casos {fecha_desc}", index=False)
            
            st.sidebar.divider()
            st.sidebar.download_button(
                label="📥 Descargar Pipeline con Formato",
                data=buffer.getvalue(),
                file_name=f"JPV_Pipeline_{fecha_desc}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("Sube los archivos para procesar el Pipeline. El sistema aplicará el semáforo de colores automáticamente.")
