import streamlit as st
import pandas as pd
from datetime import datetime
import io
import re

# --- CONFIGURACIÓN DE ETIQUETAS ---
PROB_MAP = {
    0.0: "Nula",
    0.25: "Remota",
    0.50: "Podría Ser",
    0.75: "Altamente probable",
    1.0: "Cierta"
}

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
    # 1. Cargar Reporte Nuevo (Cabecera en fila 6)
    df_nuevo = pd.read_excel(archivo_nuevo, skiprows=5)
    df_nuevo.columns = [str(c).strip() for c in df_nuevo.columns]
    df_nuevo = df_nuevo.dropna(how='all', axis=0)

    # 2. Identificar Automáticamente la "Última Actualización" en el Historial
    xl_historial = pd.ExcelFile(archivo_historial)
    hojas = xl_historial.sheet_names
    
    hoja_maestra = None
    fecha_reciente = datetime.min
    
    # Buscamos la hoja con la fecha más reciente en el nombre (formato dd-mm-yy)
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
    
    # Si no encontramos fechas, buscamos la última hoja que tenga "Número de caso"
    if not hoja_maestra:
        posibles_nombres = ['Número de caso', 'Numero de caso', 'N° caso', 'Caso']
        for h in reversed(hojas):
            df_check = pd.read_excel(xl_historial, sheet_name=h, nrows=15, header=None)
            if any(str(val).strip() in posibles_nombres for row in df_check.values for val in row):
                hoja_maestra = h
                break

    if not hoja_maestra:
        st.error("No se pudo identificar la hoja de datos en el archivo histórico.")
    else:
        st.info(f"Detectada automáticamente la última actualización: **{hoja_maestra}**")
        
        # Leer la hoja histórica detectando la fila del encabezado
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
            # Estandarizar llaves para el cruce
            df_nuevo[col_llave] = df_nuevo[col_llave].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            df_hist[col_llave] = df_hist[col_llave].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

            # Recuperar solo lo necesario del historial
            cols_persistencia = [col_llave, 'Probabilidad cierre 2026', 'Observaciones', 'Fecha probable de facturación']
            for c in cols_persistencia:
                if c not in df_hist.columns: df_hist[c] = ""

            # --- CRUCE DE DATOS ---
            df_final = pd.merge(df_nuevo, df_hist[cols_persistencia], on=col_llave, how='left')

            # Formatear para el editor
            df_final['Probabilidad cierre 2026'] = pd.to_numeric(df_final['Probabilidad cierre 2026'], errors='coerce').fillna(0.0)
            df_final['Observaciones'] = df_final['Observaciones'].astype(str).replace(['nan', 'None', '<NA>'], '')
            df_final['Fecha probable de facturación'] = pd.to_datetime(df_final['Fecha probable de facturación'], errors='coerce').dt.date

            # Limpiar columnas sobrantes y asegurar cálculos
            for col in COLUMNAS_FINALES:
                if col not in df_final.columns: df_final[col] = ""
            
            df_final['Indicación Probabilidad'] = df_final['Probabilidad cierre 2026'].map(PROB_MAP)
            if 'Honorarios (UF)' in df_final.columns:
                df_final['Honorarios (UF)'] = pd.to_numeric(df_final['Honorarios (UF)'], errors='coerce').fillna(0)
                df_final['Hon Probables 2026'] = df_final['Honorarios (UF)'] * df_final['Probabilidad cierre 2026']

            df_final = df_final[COLUMNAS_FINALES]

            st.subheader("Panel de Gestión")
            df_editado = st.data_editor(
                df_final,
                column_config={
                    "Probabilidad cierre 2026": st.column_config.SelectboxColumn("Probabilidad (%)", options=[0.0, 0.25, 0.50, 0.75, 1.0]),
                    "Fecha probable de facturación": st.column_config.DateColumn("Fecha Fact."),
                    "Observaciones": st.column_config.TextColumn("Observaciones", width="large")
                },
                hide_index=True, use_container_width=True
            )

            # Cálculo de KPI final
            df_editado['Hon Probables 2026'] = df_editado['Honorarios (UF)'] * df_editado['Probabilidad cierre 2026']
            st.metric("FACTURACIÓN PROBABLE TOTAL (UF)", f"{df_editado['Hon Probables 2026'].sum():,.2f}")

            # --- BOTÓN DE DESCARGA ---
            fecha_descarga = datetime.now().strftime("%d-%m-%y")
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_editado.to_excel(writer, sheet_name=f"Casos {fecha_descarga}", index=False)
            
            st.sidebar.divider()
            st.sidebar.download_button(
                label="📥 Descargar Pipeline Actualizado",
                data=buffer.getvalue(),
                file_name=f"JPV_Pipeline_{fecha_descarga}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("Sube los archivos; la app encontrará automáticamente la última actualización y aplicará el cruce.")
