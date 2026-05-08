import streamlit as st
import pandas as pd
from datetime import datetime
import io

# --- CONFIGURACIÓN DE ETIQUETAS ---
PROB_MAP = {
    0.0: "Nula",
    0.25: "Remota",
    0.50: "Podría Ser",
    0.75: "Altamente probable",
    1.0: "Cierta"
}

# Columnas definitivas que debe tener el archivo de salida (según tu hoja base)
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

def cargar_excel_especifico(archivo, es_reporte_acciones=False):
    if archivo is None: return None
    skip = 5 if es_reporte_acciones else 0
    df = pd.read_excel(archivo, skiprows=skip)
    df.columns = [str(c).strip() for c in df.columns]
    return df.dropna(how='all', axis=0)

st.sidebar.header("Carga de Documentos")
archivo_nuevo = st.sidebar.file_uploader("1. Nuevo Reporte de Acciones (Excel)", type=["xlsx"])
archivo_historial = st.sidebar.file_uploader("2. Pipeline Anterior (Excel Maestro)", type=["xlsx"])

if archivo_nuevo and archivo_historial:
    df_nuevo = cargar_excel_especifico(archivo_nuevo, es_reporte_acciones=True)
    
    # --- CAMBIO QUIRÚRGICO: BÚSQUEDA AUTOMÁTICA DE LA HOJA CORRECTA ---
    # Evita el KeyError buscando la hoja que realmente contiene los casos
    posibles_nombres = ['Número de caso', 'Numero de caso', 'N° caso', 'Caso']
    
    xl_hist = pd.ExcelFile(archivo_historial)
    df_hist = None
    nombre_hoja_historial = ""
    
    for sheet in xl_hist.sheet_names:
        temp_df = pd.read_excel(archivo_historial, sheet_name=sheet)
        temp_df.columns = [str(c).strip() for c in temp_df.columns]
        # Si la hoja contiene nuestra columna llave, la guardamos y detenemos la búsqueda
        if any(c in temp_df.columns for c in posibles_nombres):
            df_hist = temp_df.dropna(how='all', axis=0)
            nombre_hoja_historial = sheet
            break
            
    if df_hist is None:
        st.error("No se encontró la columna 'Número de caso' en ninguna hoja del archivo histórico.")
    else:
        st.success(f"Historial detectado y cargado exitosamente desde la hoja: **{nombre_hoja_historial}**")
        
        col_llave = next((c for c in df_nuevo.columns if c in posibles_nombres), None)
        
        if not col_llave:
            st.error(f"No se encontró la columna clave en el reporte de acciones.")
        else:
            # Aseguramos que los nombres de la columna llave coincidan exactamente en ambos archivos
            col_llave_hist = next((c for c in df_hist.columns if c in posibles_nombres), None)
            if col_llave_hist and col_llave_hist != col_llave:
                df_hist.rename(columns={col_llave_hist: col_llave}, inplace=True)

            # Estandarizar llave como texto
            df_nuevo[col_llave] = df_nuevo[col_llave].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            df_hist[col_llave] = df_hist[col_llave].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

            # 1. PREPARAR HISTORIAL: Solo nos interesan las columnas manuales para el cruce
            cols_a_recuperar = [col_llave, 'Probabilidad cierre 2026', 'Observaciones', 'Fecha probable de facturación']
            # Si alguna no existe en el historial, la creamos vacía
            for c in cols_a_recuperar:
                if c not in df_hist.columns: df_hist[c] = ""

            # 2. CRUCE (LEFT JOIN): Mantenemos todo lo nuevo, pero traemos lo manual del pasado
            df_final = pd.merge(
                df_nuevo, 
                df_hist[cols_a_recuperar], 
                on=col_llave, 
                how='left'
            )

            # 3. LIMPIEZA POST-MERGE Y TIPOS DE DATOS
            df_final['Probabilidad cierre 2026'] = pd.to_numeric(df_final['Probabilidad cierre 2026'], errors='coerce').fillna(0.0)
            df_final['Observaciones'] = df_final['Observaciones'].astype(str).replace(['nan', 'None', '<NA>'], '')
            df_final['Fecha probable de facturación'] = pd.to_datetime(df_final['Fecha probable de facturación'], errors='coerce').dt.date

            # 4. ASEGURAR COLUMNAS FALTANTES
            for col in COLUMNAS_FINALES:
                if col not in df_final.columns:
                    df_final[col] = ""

            # 5. CÁLCULOS INICIALES
            df_final['Indicación Probabilidad'] = df_final['Probabilidad cierre 2026'].map(PROB_MAP)
            if 'Honorarios (UF)' in df_final.columns:
                df_final['Honorarios (UF)'] = pd.to_numeric(df_final['Honorarios (UF)'], errors='coerce').fillna(0)
                df_final['Hon Probables 2026'] = df_final['Honorarios (UF)'] * df_final['Probabilidad cierre 2026']

            # 6. FILTRAR Y ORDENAR
            df_final = df_final[COLUMNAS_FINALES]

            st.subheader("Panel de Gestión Semanal")
            df_editado = st.data_editor(
                df_final,
                column_config={
                    "Probabilidad cierre 2026": st.column_config.SelectboxColumn("Probabilidad (%)", options=[0.0, 0.25, 0.50, 0.75, 1.0]),
                    "Fecha probable de facturación": st.column_config.DateColumn("Fecha Fact."),
                    "Observaciones": st.column_config.TextColumn("Observaciones", width="large")
                },
                hide_index=True, use_container_width=True
            )

            # Recalcular tras edición
            df_editado['Indicación Probabilidad'] = df_editado['Probabilidad cierre 2026'].map(PROB_MAP)
            df_editado['Hon Probables 2026'] = df_editado['Honorarios (UF)'] * df_editado['Probabilidad cierre 2026']
            
            st.metric("FACTURACIÓN PROBABLE TOTAL (UF)", f"{df_editado['Hon Probables 2026'].sum():,.2f}")

            # --- DESCARGA ---
            fecha_hoy = datetime.now().strftime("%d-%m-%y")
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_editado.to_excel(writer, sheet_name=f"Casos {fecha_hoy}", index=False)
            
            st.sidebar.download_button(
                label="📥 Descargar Pipeline Limpio",
                data=buffer.getvalue(),
                file_name=f"JPV_Pipeline_{fecha_hoy}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("Sube los archivos para procesar el Pipeline. El sistema detectará automáticamente las hojas correctas.")
