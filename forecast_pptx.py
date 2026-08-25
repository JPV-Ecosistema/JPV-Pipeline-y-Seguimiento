"""
Generador de la presentación PPTX "Línea Base Forecast" (Ingeniería y Equipo Móvil).

Replica el estilo navy oscuro del deck de referencia de JPV Asociados. Las coordenadas
de layout se calibraron a partir del XML real de ese deck (10in x 5.625in), no son
un diseño inventado desde cero.

Fase 2: genera 9 slides (Portada, Agenda, Título sección, Línea Base Real, Proyección y
Pipeline, Casos en Proceso Administrativo [si aplica], Forecast Total, Análisis de
Desviación, Evolución 7+5, Facturación Mensual). Las 2 slides comparativas que dependen
del historial acumulado de forecasts (Comparativo Forecasts y Tabla Comparativa) llegan
en la Fase 3, junto con la persistencia del historial.
"""
import os
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

# --- Paleta (idéntica al deck de referencia) ---
COLOR_BG = "1A202C"
COLOR_BAR = "3182CE"
COLOR_CARD_BG = "2D3748"
COLOR_CARD_BORDER = "4A5568"
COLOR_IE = "4299E1"
COLOR_EM = "F6AD55"
COLOR_CTM = "00A896"
COLOR_VERDE = "68D391"
COLOR_ROJO = "E53E3E"
COLOR_TEXTO = "E2E8F0"
COLOR_MUTED = "A0AEC0"
COLOR_MUTED2 = "718096"
FUENTE = "Calibri"

SLIDE_W = Emu(9144000)
SLIDE_H = Emu(5143500)
MARGEN_X = Emu(457200)
CONTENIDO_W = Emu(8229600)

LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "logo_jpv_blanco.png")
LOGO_ASPECT = 1351 / 208  # ancho/alto del archivo real


def _rgb(hex_str):
    return RGBColor.from_string(hex_str)


def _sin_borde(shape):
    shape.line.fill.background()


def nueva_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fondo = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    fondo.fill.solid()
    fondo.fill.fore_color.rgb = _rgb(COLOR_BG)
    _sin_borde(fondo)
    fondo.shadow.inherit = False
    return slide


def barra(slide, y):
    b = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, y, SLIDE_W, Emu(73152))
    b.fill.solid()
    b.fill.fore_color.rgb = _rgb(COLOR_BAR)
    _sin_borde(b)
    b.shadow.inherit = False
    return b


def caja(slide, x, y, w, h, fill_hex=COLOR_CARD_BG, border_hex=COLOR_CARD_BORDER):
    c = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    c.fill.solid()
    c.fill.fore_color.rgb = _rgb(fill_hex)
    if border_hex:
        c.line.color.rgb = _rgb(border_hex)
        c.line.width = Pt(1)
    else:
        _sin_borde(c)
    c.shadow.inherit = False
    return c


def texto(slide, x, y, w, h, contenido, size, color_hex, bold=False, italic=False,
          align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE, margenes=True, wrap=True):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    if not margenes:
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

    lineas = contenido if isinstance(contenido, list) else [contenido]
    for i, linea in enumerate(lineas):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = linea
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.name = FUENTE
        run.font.color.rgb = _rgb(color_hex)
    return box


def pie_de_pagina(slide, fuente_texto):
    """Footer estándar de las slides de contenido: texto de fuente a la izquierda y logo a la derecha."""
    texto(slide, MARGEN_X, Emu(4709160), Emu(7772400), Emu(164592),
          fuente_texto, 8, COLOR_MUTED2, margenes=False)
    alto_logo = Emu(164592)
    ancho_logo = Emu(int(alto_logo * LOGO_ASPECT))
    x_logo = Emu(8503920 - ancho_logo)
    if os.path.exists(LOGO_PATH):
        slide.shapes.add_picture(LOGO_PATH, x_logo, Emu(4709160), height=alto_logo)


def encabezado_slide(slide, titulo, subtitulo, color_subtitulo=COLOR_IE):
    texto(slide, MARGEN_X, Emu(164592), CONTENIDO_W, Emu(384048), titulo, 17, COLOR_TEXTO)
    texto(slide, MARGEN_X, Emu(548640), CONTENIDO_W, Emu(228600), subtitulo, 10.5, color_subtitulo, bold=True)


def _fmt_uf(valor):
    return f"{valor:,.0f}".replace(",", ".")


def _fmt_uf2(valor):
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


# ---------------------------------------------------------
# SLIDE 1 — PORTADA
# ---------------------------------------------------------
def slide_portada(prs, datos):
    slide = nueva_slide(prs)
    barra(slide, 0)
    barra(slide, Emu(4754880))

    texto(slide, Emu(640080), Emu(1280160), Emu(7863840), Emu(640080),
          "Ingeniería y Equipo Móvil", 42, COLOR_EM, bold=True)
    texto(slide, Emu(640080), Emu(1920240), Emu(7863840), Emu(411480),
          f"Línea Base {datos['anio']}: Ingeniería y Equipos Móviles", 20, COLOR_TEXTO)
    texto(slide, Emu(640080), Emu(2423160), Emu(6000000), Emu(320040),
          f"Forecast {datos['label']} — {datos['nombre_mes_corte']} {datos['anio']}", 16, COLOR_CTM, bold=True)

    texto(slide, Emu(640080), Emu(4663440), Emu(4572000), Emu(182880),
          datos['fecha_emision'], 10, COLOR_MUTED2)

    # Logo más grande en la portada, apoyado justo sobre la barra inferior.
    alto_logo = Emu(300000)
    ancho_logo = Emu(int(alto_logo * LOGO_ASPECT))
    x_logo = Emu(8503920 - ancho_logo)
    y_logo = Emu(4754880 - 40000 - 300000)
    if os.path.exists(LOGO_PATH):
        slide.shapes.add_picture(LOGO_PATH, x_logo, y_logo, height=alto_logo)
    return slide


# ---------------------------------------------------------
# SLIDE 2 — AGENDA
# ---------------------------------------------------------
def slide_agenda(prs, datos):
    slide = nueva_slide(prs)
    texto(slide, MARGEN_X, Emu(457200), CONTENIDO_W, Emu(457200), "Agenda", 28, COLOR_TEXTO, bold=True)

    items = [
        f"Forecast {datos['label']} {datos['anio']}",
        "Análisis de Desempeño Histórico",
        "La Mirada de la Industria",
        "Plan de Acción",
    ]
    y = 1280160
    for i, item in enumerate(items, start=1):
        caja(slide, MARGEN_X, Emu(y), Emu(548640), Emu(548640), fill_hex=COLOR_CARD_BG, border_hex=None)
        texto(slide, MARGEN_X, Emu(y), Emu(548640), Emu(548640), str(i), 22, COLOR_CTM, bold=True,
              align=PP_ALIGN.CENTER, margenes=False)
        texto(slide, Emu(457200 + 640080), Emu(y), Emu(6400800), Emu(548640), item, 16, COLOR_TEXTO,
              margenes=False)
        y += 640080

    pie_de_pagina(slide, "JPV Asociados | sedgwick.")
    return slide


# ---------------------------------------------------------
# SLIDE 3 — TÍTULO DE SECCIÓN
# ---------------------------------------------------------
def slide_titulo_seccion(prs, datos):
    slide = nueva_slide(prs)
    texto(slide, Emu(640080), Emu(1920240), Emu(7863840), Emu(960120),
          f"Forecast Financiero {datos['anio']}: Análisis de\nProducción y Facturación.",
          28, COLOR_TEXTO, bold=True)
    texto(slide, Emu(640080), Emu(2880360), Emu(7863840), Emu(411480),
          "La Brecha del Forecast (La Realidad Incómoda)", 14, COLOR_MUTED)
    pie_de_pagina(slide, "JPV | sedgwick.")
    return slide


# ---------------------------------------------------------
# SLIDE 4 — LÍNEA BASE REAL
# ---------------------------------------------------------
def slide_linea_base_real(prs, datos):
    slide = nueva_slide(prs)
    encabezado_slide(
        slide,
        "1. Línea Base Real (Cierre Facturación)",
        f"Estado Actual (YTD {datos['nombre_mes_inicio']} a {datos['nombre_mes_corte']})",
        COLOR_IE,
    )

    tarjetas = [
        ("Equipo Móvil", _fmt_uf2(datos['em_ytd']) + " UF", f"{datos['nombre_mes_inicio']}–{datos['nombre_mes_corte']} {datos['anio']}"),
        ("Ingeniería y Energía", _fmt_uf2(datos['ie_ytd']) + " UF", f"{datos['nombre_mes_inicio']}–{datos['nombre_mes_corte']} {datos['anio']}"),
        ("Total Consolidado YTD", _fmt_uf2(datos['ytd_total']) + " UF", f"{datos['nombre_mes_inicio']}–{datos['nombre_mes_corte']} {datos['anio']}"),
    ]
    x = 320040
    for titulo_t, valor_t, sub_t in tarjetas:
        caja(slide, Emu(x), Emu(914400), Emu(2240280), Emu(960120))
        texto(slide, Emu(x), Emu(969264), Emu(2240280), Emu(201168), titulo_t, 9, COLOR_MUTED,
              align=PP_ALIGN.CENTER, margenes=False)
        texto(slide, Emu(x), Emu(1170432), Emu(2240280), Emu(365760), valor_t, 20, COLOR_TEXTO, bold=True,
              align=PP_ALIGN.CENTER, margenes=False)
        texto(slide, Emu(x), Emu(1536192), Emu(2240280), Emu(210312), sub_t, 8, COLOR_MUTED,
              align=PP_ALIGN.CENTER, margenes=False)
        x += 2240280 + 82296

    x = 320040
    tarjetas2 = [
        ("Honorarios en Stock", _fmt_uf(datos['pipeline_bruto_total']) + " UF", "Pipeline bruto total " + str(datos['anio'])),
        ("Meta Anual", _fmt_uf(datos['meta']) + " UF", str(datos['anio'])),
    ]
    for titulo_t, valor_t, sub_t in tarjetas2:
        caja(slide, Emu(x), Emu(2011680), Emu(2240280), Emu(960120))
        texto(slide, Emu(x), Emu(2066544), Emu(2240280), Emu(201168), titulo_t, 9, COLOR_MUTED,
              align=PP_ALIGN.CENTER, margenes=False)
        texto(slide, Emu(x), Emu(2267712), Emu(2240280), Emu(365760), valor_t, 20, COLOR_TEXTO, bold=True,
              align=PP_ALIGN.CENTER, margenes=False)
        texto(slide, Emu(x), Emu(2633472), Emu(2240280), Emu(210312), sub_t, 8, COLOR_MUTED,
              align=PP_ALIGN.CENTER, margenes=False)
        x += 2240280 + 82296

    texto(slide, MARGEN_X, Emu(3200400), Emu(4000000), Emu(228600), "Análisis de contexto", 10.5, COLOR_IE, bold=True)
    bullets = [f"• {b}" for b in datos.get('bullets_linea_base', [])]
    if bullets:
        texto(slide, MARGEN_X, Emu(3474720), Emu(8229600), Emu(457200), bullets, 9.5, COLOR_TEXTO, margenes=False)

    pie_de_pagina(slide, datos['fuente_texto'])
    return slide


# ---------------------------------------------------------
# SLIDE 5 — PROYECCIÓN Y PIPELINE
# ---------------------------------------------------------
def slide_proyeccion_pipeline(prs, datos):
    slide = nueva_slide(prs)
    encabezado_slide(
        slide,
        f"2. Proyección y Pipeline {datos['label']} (Resto del año)",
        f"Proyección (Forecast {datos['nombre_mes_corte']}—Diciembre)",
        COLOR_IE,
    )

    proy_em = datos['proy_em']
    caja(slide, Emu(320040), Emu(914400), Emu(4000000), Emu(1280160))
    texto(slide, Emu(411480), Emu(969264), Emu(3800000), Emu(228600),
          "Equipo Móvil (Pipeline Hon. Probables)", 10, COLOR_EM, bold=True, margenes=False)
    texto(slide, Emu(411480), Emu(1234440), Emu(3800000), Emu(320040),
          f"{_fmt_uf(proy_em['em_stock'])} UF (stock)  ·  {_fmt_uf(proy_em['em_promedio_total'])} UF (promedio × meses)",
          13, COLOR_TEXTO, bold=True, margenes=False)
    meses_txt = "-".join(m[:3] for m in proy_em['meses_usados']) or "—"
    texto(slide, Emu(411480), Emu(1600200), Emu(3800000), Emu(457200),
          f"Complemento EM — Promedio últimos {len(proy_em['meses_usados'])} meses\n"
          f"{_fmt_uf(proy_em['prom_em_3m'])} UF/mes ({meses_txt}) × {datos['meses_proyectados']} meses = {_fmt_uf(proy_em['em_promedio_total'])} UF",
          9, COLOR_MUTED, margenes=False)

    if datos['total_admin'] > 0:
        caja(slide, Emu(4457700), Emu(914400), Emu(4051800), Emu(1280160), border_hex=COLOR_CTM)
        texto(slide, Emu(4549140), Emu(969264), Emu(3800000), Emu(228600),
              "Proceso Administrativo de Facturación", 10, COLOR_CTM, bold=True, margenes=False)
        texto(slide, Emu(4549140), Emu(1234440), Emu(3800000), Emu(320040),
              f"{_fmt_uf(datos['total_admin'])} UF", 18, COLOR_TEXTO, bold=True, margenes=False)
        texto(slide, Emu(4549140), Emu(1600200), Emu(3800000), Emu(457200),
              f"{len(datos['casos_admin'])} caso(s) en proceso administrativo",
              9, COLOR_MUTED, margenes=False)

    proy_ie = datos['proy_ie']
    texto(slide, MARGEN_X, Emu(2377440), Emu(6000000), Emu(228600),
          f"Ingeniería — Proyección {datos['nombre_mes_corte']}—Diciembre:", 10.5, COLOR_IE, bold=True)

    ie_tarjetas = [
        ("Casos menores", "(<1.000 UF pérdida)", _fmt_uf(proy_ie['ie_menores_proj']) + " UF",
         f"{_fmt_uf(proy_ie['prom_ie_menores'])} UF/mes × {datos['meses_proyectados']} meses"),
        ("Casos mayores pipeline", "(≥1.000 UF, prob.)", _fmt_uf(proy_ie['ie_mayores_proj']) + " UF",
         f"{proy_ie['n_casos_mayores']} casos · Hon. Probables ponderados"),
        ("Engie/CTM", "(pipeline, prob.)", _fmt_uf(proy_ie['ie_engie_proj']) + " UF",
         f"{proy_ie['n_casos_engie']} caso(s) · Hon. Probables ponderados"),
    ]
    x = 320040
    for t1, t2, valor_t, sub_t in ie_tarjetas:
        caja(slide, Emu(x), Emu(2743200), Emu(2606040), Emu(1097280))
        texto(slide, Emu(x), Emu(2798064), Emu(2606040), Emu(180000), t1, 9.5, COLOR_MUTED, bold=True,
              align=PP_ALIGN.CENTER, margenes=False)
        texto(slide, Emu(x), Emu(2978064), Emu(2606040), Emu(160000), t2, 8, COLOR_MUTED,
              align=PP_ALIGN.CENTER, margenes=False)
        texto(slide, Emu(x), Emu(3150000), Emu(2606040), Emu(320040), valor_t, 17, COLOR_TEXTO, bold=True,
              align=PP_ALIGN.CENTER, margenes=False)
        texto(slide, Emu(x), Emu(3520000), Emu(2606040), Emu(280000), sub_t, 8, COLOR_MUTED,
              align=PP_ALIGN.CENTER, margenes=False)
        x += 2606040 + 55000

    pie_de_pagina(slide, datos['fuente_texto'])
    return slide


# ---------------------------------------------------------
# SLIDE — CASOS EN PROCESO ADMINISTRATIVO (condicional)
# ---------------------------------------------------------
def slide_casos_admin(prs, datos):
    slide = nueva_slide(prs)
    encabezado_slide(slide, "Casos en Proceso Administrativo de Facturación",
                      f"Total: {_fmt_uf(datos['total_admin'])} UF", COLOR_CTM)

    y = 1097280
    for caso in datos['casos_admin']:
        caja(slide, MARGEN_X, Emu(y), CONTENIDO_W, Emu(457200))
        texto(slide, Emu(457200 + 137160), Emu(y), Emu(6000000), Emu(457200),
              caso['nombre'] or "(sin nombre)", 12, COLOR_TEXTO, margenes=False)
        texto(slide, Emu(6800000), Emu(y), Emu(1372000), Emu(457200),
              f"{_fmt_uf(caso['monto'])} UF", 13, COLOR_CTM, bold=True,
              align=PP_ALIGN.RIGHT, margenes=False)
        y += 502920
        if y > 4200000:
            break

    pie_de_pagina(slide, datos['fuente_texto'])
    return slide


# ---------------------------------------------------------
# SLIDE — FORECAST TOTAL
# ---------------------------------------------------------
def slide_forecast_total(prs, datos):
    slide = nueva_slide(prs)
    encabezado_slide(slide, f"2. Forecast {datos['label']} (Total año {datos['anio']})",
                      f"Proyección (Forecast {datos['nombre_mes_corte']}—Diciembre)", COLOR_IE)

    tarjetas = [
        ("Equipo Móvil", COLOR_EM, _fmt_uf(datos['em_total']) + " UF",
         f"{_fmt_uf2(datos['em_ytd'])} YTD + {_fmt_uf(datos['em_proyectado'])} proyectado"),
        ("Ingeniería", COLOR_IE, _fmt_uf(datos['ie_total']) + " UF",
         f"{_fmt_uf2(datos['ie_ytd'])} YTD + {_fmt_uf(datos['ie_proyectado'])} proyectado"),
    ]
    x = 320040
    for titulo_t, color_t, valor_t, sub_t in tarjetas:
        caja(slide, Emu(x), Emu(1097280), Emu(2789555), Emu(1280160))
        texto(slide, Emu(x + 91440), Emu(1152144), Emu(2600000), Emu(228600), titulo_t, 10.5, color_t, bold=True, margenes=False)
        texto(slide, Emu(x + 91440), Emu(1417320), Emu(2600000), Emu(365760), valor_t, 22, COLOR_TEXTO, bold=True, margenes=False)
        texto(slide, Emu(x + 91440), Emu(1783080), Emu(2600000), Emu(320040), sub_t, 9, COLOR_MUTED, margenes=False)
        x += 2789555 + 91445

    if datos['total_admin'] > 0:
        caja(slide, Emu(x), Emu(1097280), Emu(2789555), Emu(1280160), border_hex=COLOR_CTM)
        texto(slide, Emu(x + 91440), Emu(1152144), Emu(2600000), Emu(228600),
              "Proceso Administrativo", 10.5, COLOR_CTM, bold=True, margenes=False)
        texto(slide, Emu(x + 91440), Emu(1417320), Emu(2600000), Emu(365760),
              _fmt_uf(datos['total_admin']) + " UF", 22, COLOR_TEXTO, bold=True, margenes=False)
        texto(slide, Emu(x + 91440), Emu(1783080), Emu(2600000), Emu(320040),
              "Cierre esperado próximas semanas", 9, COLOR_MUTED, margenes=False)

    texto(slide, MARGEN_X, Emu(2606040), Emu(6000000), Emu(228600),
          "Cierre de año proyectado global", 12, COLOR_TEXTO, bold=True)

    color_gap_sin = COLOR_VERDE if datos['cum_sin'] >= 100 else COLOR_EM
    caja(slide, Emu(320040), Emu(2971800), Emu(3931920), Emu(1188720))
    texto(slide, Emu(320040), Emu(3063240), Emu(3931920), Emu(365760),
          _fmt_uf(datos['cierre_sin_admin']) + " UF", 22, COLOR_TEXTO, bold=True,
          align=PP_ALIGN.CENTER, margenes=False)
    texto(slide, Emu(320040), Emu(3429000), Emu(3931920), Emu(228600),
          f"Sin proceso administrativo · {datos['cum_sin']:.1f}% de la meta", 9.5, color_gap_sin,
          align=PP_ALIGN.CENTER, margenes=False)

    if datos['total_admin'] > 0:
        color_gap_con = COLOR_VERDE if datos['cum_con'] >= 100 else COLOR_EM
        caja(slide, Emu(4343400), Emu(2971800), Emu(3931920), Emu(1188720), border_hex=COLOR_CTM)
        texto(slide, Emu(4343400), Emu(3063240), Emu(3931920), Emu(365760),
              _fmt_uf(datos['cierre_con_admin']) + " UF", 22, COLOR_TEXTO, bold=True,
              align=PP_ALIGN.CENTER, margenes=False)
        texto(slide, Emu(4343400), Emu(3429000), Emu(3931920), Emu(228600),
              f"Con proceso administrativo · {datos['cum_con']:.1f}% de la meta", 9.5, color_gap_con,
              align=PP_ALIGN.CENTER, margenes=False)

    texto(slide, MARGEN_X, Emu(4297680), CONTENIDO_W, Emu(228600),
          f"Honorarios en Stock: {_fmt_uf(datos['pipeline_bruto_total'])} UF no facturados en pipeline total (bruto, sin probabilidad)",
          9, COLOR_MUTED)

    pie_de_pagina(slide, datos['fuente_texto'])
    return slide


# ---------------------------------------------------------
# SLIDE — ANÁLISIS DE DESVIACIÓN
# ---------------------------------------------------------
def slide_desviacion(prs, datos):
    slide = nueva_slide(prs)
    encabezado_slide(slide, f"3. Análisis de Desviación (Gap) vs. Meta — Forecast {datos['label']}",
                      "Evaluación de Brecha Estratégica y Acciones Mitigadoras", COLOR_IE)

    usar_con_admin = datos['total_admin'] > 0
    cierre_ref = datos['cierre_con_admin'] if usar_con_admin else datos['cierre_sin_admin']
    gap_ref = datos['gap_con'] if usar_con_admin else datos['gap_sin']
    cum_ref = datos['cum_con'] if usar_con_admin else datos['cum_sin']
    color_gap = COLOR_VERDE if gap_ref >= 0 else COLOR_ROJO

    etiquetas = [
        ("Meta Oficial Total", _fmt_uf(datos['meta']) + " UF", str(datos['anio']), COLOR_TEXTO),
        ("Cierre Proyectado", _fmt_uf(cierre_ref) + " UF",
         "Con proc. admin." if usar_con_admin else "Sin proc. admin.", COLOR_TEXTO),
        ("Desviación (Gap)", f"{gap_ref:+,.0f}".replace(",", ".") + " UF",
         f"Cumplimiento: {cum_ref:.1f}%", color_gap),
    ]
    x = 320040
    for titulo_t, valor_t, sub_t, color_t in etiquetas:
        caja(slide, Emu(x), Emu(1005840), Emu(2606040), Emu(914400))
        texto(slide, Emu(x), Emu(1060704), Emu(2606040), Emu(201168), titulo_t, 9, COLOR_MUTED,
              align=PP_ALIGN.CENTER, margenes=False)
        texto(slide, Emu(x), Emu(1261872), Emu(2606040), Emu(320040), valor_t, 18, color_t, bold=True,
              align=PP_ALIGN.CENTER, margenes=False)
        texto(slide, Emu(x), Emu(1600200), Emu(2606040), Emu(210312), sub_t, 8.5, COLOR_MUTED,
              align=PP_ALIGN.CENTER, margenes=False)
        x += 2606040 + 55000

    texto(slide, Emu(320040), Emu(2103120), Emu(3931920), Emu(228600),
          "Acción Requerida: Equipo Móvil", 10.5, COLOR_EM, bold=True)
    bullets_em = [f"• {b}" for b in datos.get('bullets_accion_em', [])]
    texto(slide, Emu(320040), Emu(2377440), Emu(3931920), Emu(1828800), bullets_em, 9, COLOR_TEXTO, margenes=False)

    texto(slide, Emu(4457700), Emu(2103120), Emu(3931920), Emu(228600),
          "Acción Requerida: Ingeniería", 10.5, COLOR_IE, bold=True)
    bullets_ie = [f"• {b}" for b in datos.get('bullets_accion_ie', [])]
    texto(slide, Emu(4457700), Emu(2377440), Emu(3931920), Emu(1828800), bullets_ie, 9, COLOR_TEXTO, margenes=False)

    pie_de_pagina(slide, datos['fuente_texto'])
    return slide


# ---------------------------------------------------------
# Helpers de estilo para gráficos nativos (tema oscuro)
# ---------------------------------------------------------
def _estilizar_chart(chart, mostrar_leyenda=True):
    chart.has_title = False
    if mostrar_leyenda:
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(9)
        chart.legend.font.color.rgb = _rgb(COLOR_MUTED)
        chart.legend.font.name = FUENTE
    else:
        chart.has_legend = False

    cat_ax = chart.category_axis
    cat_ax.tick_labels.font.size = Pt(9)
    cat_ax.tick_labels.font.color.rgb = _rgb(COLOR_MUTED)
    cat_ax.tick_labels.font.name = FUENTE
    cat_ax.format.line.color.rgb = _rgb(COLOR_CARD_BORDER)

    val_ax = chart.value_axis
    val_ax.tick_labels.font.size = Pt(9)
    val_ax.tick_labels.font.color.rgb = _rgb(COLOR_MUTED)
    val_ax.tick_labels.font.name = FUENTE
    val_ax.format.line.color.rgb = _rgb(COLOR_CARD_BORDER)
    val_ax.has_major_gridlines = True
    val_ax.major_gridlines.format.line.color.rgb = _rgb(COLOR_CARD_BORDER)
    val_ax.major_gridlines.format.line.width = Pt(0.5)


def _colorear_series(chart, colores):
    for i, serie in enumerate(chart.series):
        color = colores[i % len(colores)]
        if chart.chart_type in (XL_CHART_TYPE.LINE, XL_CHART_TYPE.LINE_MARKERS):
            serie.format.line.color.rgb = _rgb(color)
            serie.format.line.width = Pt(2.5)
            serie.smooth = False
        else:
            serie.format.fill.solid()
            serie.format.fill.fore_color.rgb = _rgb(color)


# ---------------------------------------------------------
# SLIDE — EVOLUCIÓN N+M (gráfico de curvas acumuladas)
# ---------------------------------------------------------
def slide_evolucion(prs, datos):
    slide = nueva_slide(prs)
    encabezado_slide(slide, f"Evolución Forecast {datos['label']} año {datos['anio']}",
                      "Real acumulado + proyección hasta diciembre", COLOR_IE)

    chart_data = CategoryChartData()
    chart_data.categories = datos['grafico_meses']
    chart_data.add_series('Real', datos['grafico_real'])
    chart_data.add_series('Proyección (sin proc. admin.)', datos['grafico_proy_sin'])
    if datos['total_admin'] > 0:
        chart_data.add_series('Proyección (con proc. admin.)', datos['grafico_proy_con'])
    chart_data.add_series('Meta', [datos['meta']] * len(datos['grafico_meses']))

    x, y, cx, cy = Emu(457200), Emu(1005840), Emu(8229600), Emu(3474720)
    gf = slide.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, x, y, cx, cy, chart_data)
    chart = gf.chart
    _estilizar_chart(chart)
    colores = [COLOR_TEXTO, COLOR_EM, COLOR_CTM, COLOR_MUTED2]
    _colorear_series(chart, colores)

    pie_de_pagina(slide, datos['fuente_texto'])
    return slide


# ---------------------------------------------------------
# SLIDE — FACTURACIÓN MENSUAL POR DIVISIÓN (barras)
# ---------------------------------------------------------
def slide_facturacion_mensual(prs, datos):
    slide = nueva_slide(prs)
    encabezado_slide(slide, f"Facturación Mensual por División — {datos['nombre_mes_inicio']} a {datos['nombre_mes_corte']} {datos['anio']}",
                      "Real, por mes cerrado", COLOR_IE)

    chart_data = CategoryChartData()
    chart_data.categories = datos['meses_mensual']
    chart_data.add_series('Equipo Móvil', datos['valores_em_mensual'])
    chart_data.add_series('Ingeniería y Energía', datos['valores_ie_mensual'])

    x, y, cx, cy = Emu(457200), Emu(1005840), Emu(8229600), Emu(3474720)
    gf = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data)
    chart = gf.chart
    _estilizar_chart(chart)
    _colorear_series(chart, [COLOR_EM, COLOR_IE])

    pie_de_pagina(slide, datos['fuente_texto'])
    return slide


def generar_pptx_forecast(datos):
    """Genera la presentación completa y devuelve los bytes del archivo .pptx."""
    import io as _io
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    slide_portada(prs, datos)
    slide_agenda(prs, datos)
    slide_titulo_seccion(prs, datos)
    slide_linea_base_real(prs, datos)
    slide_proyeccion_pipeline(prs, datos)
    if datos['total_admin'] > 0:
        slide_casos_admin(prs, datos)
    slide_forecast_total(prs, datos)
    slide_desviacion(prs, datos)
    slide_evolucion(prs, datos)
    slide_facturacion_mensual(prs, datos)

    buffer = _io.BytesIO()
    prs.save(buffer)
    return buffer.getvalue()
