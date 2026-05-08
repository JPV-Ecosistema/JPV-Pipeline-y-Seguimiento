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

st.set_page_config(page_title="JPV Pipeline y Seguimiento", layout="wide")
st.title("🚀 JPV: Pipeline de Facturación Probable")

# --- FUNCIÓN DE CARGA QUIRÚRGICA ---
def cargar_excel_especifico(archivo, es_reporte_acciones=False):
    if archivo is None: return None
    
    # Si es el reporte de acciones, saltamos las 5 filas de identificación
    # (Pandas usa índice 0, así que skiprows=5 empieza a leer en la fila 6)
    skip = 5 if es_reporte_acciones else 0
    
    df = pd.read_excel(archivo, skiprows=skip)
    
    # Limpieza de nombres de columnas
    df.columns = [str(c).strip() for c in df.columns]
    
    # Eliminar filas completamente vacías que suelen quedar al final
    df = df.dropna(how='all', axis=0)
    
    return df

# --- BARRA LATERAL ---
st.sidebar.header("Carga de Documentos")
archivo_nuevo = st.sidebar.file_uploader("1. Nuevo Reporte de Acciones (Excel)", type=["xlsx"])
archivo_historial = st.sidebar.file_uploader("2. Pipeline Anterior (Excel Maestro)", type=["xlsx"])

if archivo_nuevo and archivo_historial:
    # Aplicamos el salto de filas solo al reporte nuevo
    df_nuevo = cargar_excel_especifico(archivo_nuevo, es_reporte_acciones=True)
    df_hist = cargar_excel_especifico(archivo_historial, es_reporte_acciones=False)

    # Nombres posibles de la columna clave
    posibles_nombres = ['Número de caso', 'Numero de caso', 'N° caso', 'Caso']
    col_llave = next((c for c in df_nuevo.columns if c in posibles_nombres), None)

    if not col_llave:
        st.error(f"No se encontró la columna clave. Columnas detectadas en fila 6: {list(df_nuevo.columns[:5])}...")
    else:
        # Definir columnas de persistencia
        cols_manuales = [col_llave, 'Probabilidad cierre 2026', 'Observaciones', 'Fecha probable de facturación']
        
        # Asegurar que existan en el historial para no romper el merge
        for c in cols_manuales:
            if c not in df_hist.columns:
                df_hist[c] = 0.0 if 'Probabilidad' in c else ""

        # --- CAMBIO QUIRÚRGICO: ESTANDARIZAR TIPO DE DATO PARA EL CRUCE ---
        # Evita el ValueError de pandas al mezclar texto con números
        df_nuevo[col_llave] = df_nuevo[col_llave].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        df_hist[col_llave] = df_hist[col_llave].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

        # Cruce de datos
        df_final = pd.merge(
            df_nuevo, 
            df_hist[cols_manuales], 
            on=col_llave, 
            how='left', 
            suffixes=('', '_old')
        )

        # Rellenar valores previos si existen
        df_final['Probabilidad cierre 2026'] = df_final['Probabilidad cierre 2026'].fillna(0.0)
        
        st.subheader("Panel de Gestión Semanal")
        
        # El editor de datos
        df_editado = st.data_editor(
            df_final,
            column_config={
                "Probabilidad cierre 2026": st.column_config.SelectboxColumn(
                    "Probabilidad (%)", 
                    options=[0.0, 0.25, 0.50, 0.75, 1.0]
                ),
                "Fecha probable de facturación": st.column_config.DateColumn("Fecha Fact."),
                "Observaciones": st.column_config.TextColumn("Observaciones", width="large")
            },
            hide_index=True,
            use_container_width=True
        )

        # --- LÓGICA DE CÁLCULOS ---
        df_editado['Indicación Probabilidad'] = df_editado['Probabilidad cierre 2026'].map(PROB_MAP)
        
        # Identificar columna de honorarios
        col_hon = next((c for c in df_editado.columns if 'Honorarios (UF)' in c), None)
        
        if col_hon:
            df_editado[col_hon] = pd.to_numeric(df_editado[col_hon], errors='coerce').fillna(0)
            df_editado['Hon Probables 2026'] = df_editado[col_hon] * df_editado['Probabilidad cierre 2026']
            
            # KPI
            total_uf = df_editado['Hon Probables 2026'].sum()
            st.metric("FACTURACIÓN PROBABLE TOTAL (UF)", f"{total_uf:,.2f}")

        # --- GENERACIÓN DE DESCARGA ---
        fecha_hoy = datetime.now().strftime("%d-%m-%y")
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            # La hoja nueva se guarda con la fecha de hoy
            df_editado.to_excel(writer, sheet_name=f"Casos {fecha_hoy}", index=False)
        
        st.sidebar.divider()
        st.sidebar.download_button(
            label="📥 Descargar Pipeline Actualizado",
            data=buffer.getvalue(),
            file_name=f"JPV_Pipeline_{fecha_hoy}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Favor subir los archivos Excel. El sistema procesará el reporte desde la fila 6 automáticamente.")
