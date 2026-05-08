import streamlit as st
import pandas as pd
from datetime import datetime
import io

# --- CONFIGURACIÓN Y MAPEOS ---
PROB_MAP = {
    0.0: "Nula",
    0.25: "Remota",
    0.50: "Podría Ser",
    0.75: "Altamente probable",
    1.0: "Cierta"
}

st.set_page_config(page_title="JPV Pipeline y Seguimiento", layout="wide")

st.title("🚀 JPV: Pipeline de Facturación Probable")
st.markdown("Actualización semanal de casos y seguimiento de honorarios.")

# --- 1. CARGA DE ARCHIVOS ---
st.sidebar.header("Carga de Documentos Excel")
# Configuramos para que acepte explícitamente Excel
archivo_nuevo = st.sidebar.file_uploader("1. Nuevo Reporte de Acciones", type=["xlsx", "csv"])
archivo_historial = st.sidebar.file_uploader("2. Pipeline Anterior (Archivo Maestro)", type=["xlsx"])

def cargar_datos(archivo):
    if archivo is None:
        return None
    if archivo.name.endswith('.xlsx'):
        # Leer Excel (por defecto la primera hoja)
        return pd.read_excel(archivo)
    else:
        # Leer CSV (con manejo de punto y coma o coma)
        try:
            return pd.read_csv(archivo, sep=';')
        except:
            return pd.read_csv(archivo, sep=',')

if archivo_nuevo and archivo_historial:
    df_nuevo = cargar_datos(archivo_nuevo)
    
    # Para el historial, cargamos la hoja más reciente
    xl = pd.ExcelFile(archivo_historial)
    hoja_reciente = xl.sheet_names[0]
    df_hist = pd.read_excel(archivo_historial, sheet_name=hoja_reciente)
    
    st.info(f"Cargado historial desde la hoja: {hoja_reciente}")

    # --- 2. CRUCE DE DATOS ---
    # Columnas de persistencia según tu archivo "Casos 30-04-26"
    cols_manuales = [
        'Número de caso', 
        'Probabilidad cierre 2026', 
        'Observaciones', 
        'Fecha probable de facturación'
    ]
    
    # Asegurar que las columnas existan en el histórico
    for c in cols_manuales:
        if c not in df_hist.columns:
            df_hist[c] = None if c != 'Probabilidad cierre 2026' else 0.0

    # Merge por 'Número de caso'
    df_final = pd.merge(
        df_nuevo, 
        df_hist[cols_manuales], 
        on='Número de caso', 
        how='left', 
        suffixes=('', '_old')
    )

    # Rellenar nulos para evitar errores en el editor
    df_final['Probabilidad cierre 2026'] = df_final['Probabilidad cierre 2026'].fillna(0.0)

    # --- 3. EDITOR DE DATOS ---
    st.subheader("Panel de Edición Semanal")
    
    config_columnas = {
        "Probabilidad cierre 2026": st.column_config.SelectboxColumn(
            "Probabilidad (%)",
            options=[0.0, 0.25, 0.50, 0.75, 1.0],
            width="medium"
        ),
        "Fecha probable de facturación": st.column_config.DateColumn("Fecha Facturación"),
        "Observaciones": st.column_config.TextColumn("Observaciones", width="large"),
        "Indicación Probabilidad": st.column_config.TextColumn("Estado", disabled=True),
        "Hon Probables 2026": st.column_config.NumberColumn("Hon Probables (UF)", format="%.2f", disabled=True)
    }

    df_editado = st.data_editor(
        df_final,
        column_config=config_columnas,
        hide_index=True,
        use_container_width=True
    )

    # --- 4. CÁLCULOS Y ETIQUETAS ---
    df_editado['Indicación Probabilidad'] = df_editado['Probabilidad cierre 2026'].map(PROB_MAP)
    
    # Cálculo dinámico de honorarios
    col_hon = 'Honorarios (UF)' if 'Honorarios (UF)' in df_editado.columns else None
    if col_hon:
        df_editado[col_hon] = pd.to_numeric(df_editado[col_hon], errors='coerce').fillna(0)
        df_editado['Hon Probables 2026'] = df_editado[col_hon] * df_editado['Probabilidad cierre 2026']

    # --- 5. KPI DE FACTURACIÓN ---
    total_uf = df_editado['Hon Probables 2026'].sum() if 'Hon Probables 2026' in df_editado.columns else 0
    
    st.metric(label="FACTURACIÓN PROBABLE TOTAL (UF)", value=f"{total_uf:,.2f}")

    # --- 6. EXPORTACIÓN A EXCEL ---
    fecha_hoy = datetime.now().strftime("%d-%m-%y")
    nombre_hoja = f"Casos {fecha_hoy}"

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_editado.to_excel(writer, sheet_name=nombre_hoja, index=False)
        # Mantener las hojas antiguas si es necesario
        for s in xl.sheet_names:
            if s != nombre_hoja:
                pd.read_excel(archivo_historial, sheet_name=s).to_excel(writer, sheet_name=s, index=False)

    st.sidebar.divider()
    st.sidebar.download_button(
        label="📥 Descargar Pipeline Excel",
        data=buffer.getvalue(),
        file_name=f"Pipeline_Facturacion_{fecha_hoy}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.warning("Favor subir los archivos Excel (.xlsx) en la barra lateral.")
