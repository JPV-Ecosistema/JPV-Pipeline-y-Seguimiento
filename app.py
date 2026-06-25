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
            if 'df_pipeline_activo' not in st.session_state:
                st.session_state['df_pipeline_activo'] = df_final.copy()

            # --- BLOQUE DE FILTROS PARA PESTAÑA 2 ---
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
            # PESTAÑA 1 — PIPELINE GENERAL
            # ==========================================
            with tab1:

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
                
                st.session_state['df_pipeline_activo'] = df_editado.copy()
                
                st.metric("FACTURACIÓN PROBABLE TOTAL (UF)", f"{df_editado['Hon Probables 2026'].sum():,.2f}")

                fecha_desc = datetime.now().strftime("%d-%m-%y")
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
            # PESTAÑA 2 — SEGUIMIENTO DE CASO
            # ==========================================
            with tab2:
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

                st.caption(f"Mostrando **{len(df_filtrado)}** casos según los filtros aplicados.")

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
                            dias_activo = (datetime.now() - fecha_creacion).days if pd.notna(fecha_creacion) else None
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

                        if st.button("💾 Guardar Seguimiento", type="primary"):
                            if not nueva_obs.strip():
                                st.error("⚠️ La observación es obligatoria. Por favor completa el campo antes de guardar.")
                            else:
                                timestamp_ahora = datetime.now().strftime("%d/%m/%Y %H:%M")
                                obs_con_fecha = f"[{timestamp_ahora}] {nueva_obs.strip()}"

                                hon_uf = float(fila_caso.get('Honorarios (UF)', 0) or 0)
                                prob_decimal = float(nueva_prob.replace('%', '')) / 100
                                hon_probables_nuevo = hon_uf * prob_decimal

                                fecha_str = nueva_fecha.strftime("%Y-%m-%d") if nueva_fecha else ""

                                df_temp = st.session_state['df_pipeline_activo'].copy()
                                df_temp['Fecha probable de facturación'] = df_temp['Fecha probable de facturación'].astype(str).replace('NaT', '').replace('None', '')

                                mask = df_temp['Número de caso'].astype(str) == caso_seleccionado
                                df_temp.loc[mask, 'Observaciones'] = obs_con_fecha
                                df_temp.loc[mask, 'Fecha probable de facturación'] = fecha_str
                                df_temp.loc[mask, 'Probabilidad cierre 2026'] = nueva_prob
                                df_temp.loc[mask, 'Indicación Probabilidad'] = PROB_MAP.get(nueva_prob, '')
                                df_temp.loc[mask, 'Hon Probables 2026'] = hon_probables_nuevo

                                st.session_state['df_pipeline_activo'] = df_temp

                                st.success(f"✅ Caso **{caso_seleccionado}** actualizado correctamente el {timestamp_ahora}.")
                                st.rerun()

else:
    st.info("Sube los archivos para procesar el Pipeline. El sistema reportará ingresos, salidas y aplicará el formato al Excel.")
