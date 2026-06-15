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
            # df_pipeline_activo es la fuente de verdad en memoria que comparten Tab1 y Tab2
            if 'df_pipeline_activo' not in st.session_state:
                st.session_state['df_pipeline_activo'] = df_final.copy()

            # --- BLOQUE DE FILTROS PARA PESTAÑA 2 ---
            # Los valores se leen dinámicamente desde el archivo cargado
            st.markdown("---")
            st.markdown("#### 🔎 Filtros para Seguimiento de Caso")
            col_f1, col_f2 = st.columns(2)

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

            st.markdown("---")

            # --- PESTAÑAS PRINCIPALES ---
            tab1, tab2 = st.tabs(["📋 Pipeline General", "🔍 Seguimiento de Caso"])

            # ==========================================
            # PESTAÑA 1 — PIPELINE GENERAL (SIN CAMBIOS)
            # ==========================================
            with tab1:

                # --- CAMBIO QUIRÚRGICO: REPORTE DE CASOS NUEVOS Y SALIENTES ---
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
                
                # Acordeón para revisar los casos que salieron del pipeline
                if not salientes_detectados.empty:
                    with st.expander("🔍 Ver listado de casos salientes"):
                        columnas_salientes = [col_llave, 'Nickname', 'Probabilidad cierre 2026', 'Observaciones']
                        cols_mostrar = [c for c in columnas_salientes if c in salientes_detectados.columns]
                        st.dataframe(salientes_detectados[cols_mostrar].fillna(''), hide_index=True)

                # --- FUNCIÓN DE ESTILO PARA STREAMLIT ---
                def color_semaforo(val):
                    if val in ["75%", "100%"]:
                        return 'background-color: #c6efce; color: #006100;'
                    elif val == "50%":
                        return 'background-color: #ffeb9c; color: #9c5700;'
                    elif val in ["0%", "25%"]:
                        return 'background-color: #ffc7ce; color: #9c0006;'
                    return ''

                df_styled = st.session_state['df_pipeline_activo'].style.map(color_semaforo, subset=['Probabilidad cierre 2026'])

                # --- EDITOR DE DATOS ---
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

                # --- RECALCULAR TRAS EDICIÓN ---
                prob_num_final = df_editado['Probabilidad cierre 2026'].str.replace('%', '').astype(float) / 100
                df_editado['Hon Probables 2026'] = df_editado['Honorarios (UF)'] * prob_num_final
                df_editado['Indicación Probabilidad'] = df_editado['Probabilidad cierre 2026'].map(PROB_MAP)
                
                # Sincronizar cambios del editor de vuelta al estado compartido
                st.session_state['df_pipeline_activo'] = df_editado.copy()
                
                st.metric("FACTURACIÓN PROBABLE TOTAL (UF)", f"{df_editado['Hon Probables 2026'].sum():,.2f}")

                # --- CAMBIO QUIRÚRGICO: FORMATEO CONDICIONAL NATIVO PARA EXCEL ---
                fecha_desc = datetime.now().strftime("%d-%m-%y")
                buffer = io.BytesIO()
                df_excel = df_editado.copy()
                
                # Devolver a decimal para que Excel lo calcule matemáticamente
                df_excel['Probabilidad cierre 2026'] = df_excel['Probabilidad cierre 2026'].str.replace('%', '').astype(float) / 100

                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    nombre_hoja_descarga = f"Casos {fecha_desc}"
                    df_excel.to_excel(writer, sheet_name=nombre_hoja_descarga, index=False)
                    
                    # Obtener los objetos del libro y la hoja para inyectar formatos
                    workbook = writer.book
                    worksheet = writer.sheets[nombre_hoja_descarga]
                    
                    # Crear los formatos de Excel
                    formato_pct = workbook.add_format({'num_format': '0%'})
                    formato_verde = workbook.add_format({'bg_color': '#c6efce', 'font_color': '#006100'})
                    formato_amarillo = workbook.add_format({'bg_color': '#ffeb9c', 'font_color': '#9c5700'})
                    formato_rojo = workbook.add_format({'bg_color': '#ffc7ce', 'font_color': '#9c0006'})

                    # Encontrar el índice numérico de la columna de Probabilidad (base 0)
                    idx_prob = COLUMNAS_FINALES.index('Probabilidad cierre 2026')
                    
                    # Aplicar formato de porcentaje a toda la columna
                    worksheet.set_column(idx_prob, idx_prob, 15, formato_pct)
                    
                    # Aplicar el Semáforo Condicional directo al Excel
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
            # PESTAÑA 2 — SEGUIMIENTO DE CASO (CON FILTROS)
            # ==========================================
            with tab2:
                st.subheader("🔍 Seguimiento Individual de Caso")

                # --- APLICAR FILTROS AL DATAFRAME DE ESTA PESTAÑA ---
                df_filtrado = st.session_state['df_pipeline_activo'].copy()

                if filtro_division != "Todas":
                    df_filtrado = df_filtrado[
                        df_filtrado['División'].astype(str).str.strip() == filtro_division
                    ]

                if filtro_ajustadores:
                    df_filtrado = df_filtrado[
                        df_filtrado['Ajustador senior'].astype(str).str.strip().isin(filtro_ajustadores)
                    ]

                # Indicador de cuántos casos quedan tras el filtro
                st.caption(f"Mostrando **{len(df_filtrado)}** casos según los filtros aplicados.")

                # --- SELECTOR DE CASO (sobre los casos filtrados) ---
                lista_casos = sorted(df_filtrado['Número de caso'].astype(str).unique().tolist())

                if not lista_casos:
                    st.warning("No hay casos que coincidan con los filtros seleccionados.")
                else:
                    caso_seleccionado = st.selectbox(
                        "Selecciona el Número de Caso a gestionar:",
                        options=["— Selecciona un caso —"] + lista_casos
                    )

                    if caso_seleccionado != "— Selecciona un caso —":
                        # Extraer la fila del caso desde el estado completo (no el filtrado)
                        fila_caso = st.session_state['df_pipeline_activo'][
                            st.session_state['df_pipeline_activo']['Número de caso'].astype(str) == caso_seleccionado
                        ].iloc[0]

                        st.divider()

                        # --- BLOQUE DE DATOS DE SOLO LECTURA ---
                        st.markdown("##### 📄 Datos del Caso")
                        col_a, col_b, col_c = st.columns(3)
                        with col_a:
                            st.text_input("Número de caso",         value=str(fila_caso.get('Número de caso', '')),                    disabled=True)
                            st.text_input("Número de siniestro",    value=str(fila_caso.get('Número de siniestro', '')),               disabled=True)
                            st.text_input("Nickname",               value=str(fila_caso.get('Nickname', '')),                          disabled=True)
                            st.text_input("División",               value=str(fila_caso.get('División', '')),                          disabled=True)
                            st.text_input("Compañía de seguros",    value=str(fila_caso.get('Compañía de seguros', '')),               disabled=True)
                            st.text_input("Corredora",              value=str(fila_caso.get('Corredora', '')),                         disabled=True)
                            st.text_input("Ajustador senior",       value=str(fila_caso.get('Ajustador senior', '')),                  disabled=True)
                        with col_b:
                            st.text_input("Asegurado",              value=str(fila_caso.get('Asegurado', '')),                         disabled=True)
                            st.text_input("Creado en",              value=str(fila_caso.get('Creado en', '')),                         disabled=True)
                            st.text_input("Divisa",                 value=str(fila_caso.get('Divisa', '')),                            disabled=True)
                            st.text_input("Pérdida bruta",          value=str(fila_caso.get('Perdida bruta (en moneda del caso)', '')), disabled=True)
                            st.text_input("Monto asegurado",        value=str(fila_caso.get('Monto asegurado (en moneda del caso)', '')), disabled=True)
                            st.text_input("Honorarios (UF)",        value=str(fila_caso.get('Honorarios (UF)', '')),                   disabled=True)
                        with col_c:
                            st.text_input("Último movimiento",      value=str(fila_caso.get('Último movimiento', '')),                 disabled=True)
                            st.text_area("Contenido último mov.",   value=str(fila_caso.get('Contenido último movimiento', '')),       disabled=True, height=100)
                            st.text_input("Probabilidad cierre 2026", value=str(fila_caso.get('Probabilidad cierre 2026', '')),        disabled=True)
                            st.text_input("Indicación probabilidad", value=str(fila_caso.get('Indicación Probabilidad', '')),          disabled=True)
                            st.text_input("Hon. Probables 2026 (UF)", value=str(round(float(fila_caso.get('Hon Probables 2026', 0) or 0), 2)), disabled=True)

                        st.divider()

                        # --- BLOQUE DE OBSERVACIÓN ANTERIOR ---
                        obs_anterior = str(fila_caso.get('Observaciones', '') or '')
                        if obs_anterior.strip():
                            st.markdown("##### 📌 Última Observación Registrada")
                            st.info(obs_anterior)
                        else:
                            st.markdown("##### 📌 Última Observación Registrada")
                            st.warning("Este caso no tiene observaciones previas.")

                        st.divider()

                        # --- BLOQUE EDITABLE: NUEVA OBSERVACIÓN Y FECHA ---
                        st.markdown("##### ✏️ Actualizar Seguimiento")

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

                        # --- BOTÓN DE GUARDAR ---
                        if st.button("💾 Guardar Seguimiento", type="primary"):
                            if not nueva_obs.strip():
                                st.error("⚠️ La observación es obligatoria. Por favor completa el campo antes de guardar.")
                            else:
                                # Registrar fecha y hora de actualización automáticamente
                                timestamp_ahora = datetime.now().strftime("%d/%m/%Y %H:%M")
                                obs_con_fecha = f"[{timestamp_ahora}] {nueva_obs.strip()}"
                                
                                # Actualizar la fila correspondiente en el estado compartido
                                idx = st.session_state['df_pipeline_activo'][
                                    st.session_state['df_pipeline_activo']['Número de caso'].astype(str) == caso_seleccionado
                                ].index[0]
                                
                                st.session_state['df_pipeline_activo'].at[idx, 'Observaciones'] = obs_con_fecha
                                st.session_state['df_pipeline_activo'].at[idx, 'Fecha probable de facturación'] = nueva_fecha
                                
                                st.success(f"✅ Caso **{caso_seleccionado}** actualizado correctamente el {timestamp_ahora}.")
                                st.rerun()

else:
    st.info("Sube los archivos para procesar el Pipeline. El sistema reportará ingresos, salidas y aplicará el formato al Excel.")
