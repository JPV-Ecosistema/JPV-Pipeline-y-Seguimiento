"""
Generador de la presentación PPTX "Línea Base Forecast" (Ingeniería y Equipo Móvil).

Usa directamente la plantilla corporativa de JPV (assets/plantilla_jpv.pptx, un .potx
oficial con master + 10 layouts con nombre que el usuario compartió) como base de TODAS
las presentaciones que genera la app: cada slide se crea a partir del layout que
corresponde (Portada, Separador de sección, Solo encabezado, Cierre) y solo se completan
sus placeholders de texto. El header (logo, barra navy, línea de acento), el fondo y el
pie de página con el tagline de marca vienen heredados del layout — no se dibujan a mano.
Los colores y tipografías (Cambria para títulos, Calibri para texto) salen del theme real
de la plantilla ("JPV Asociados").

El contenido de cada dashboard (tarjetas, tablas, gráficos) se sigue construyendo a mano
con los helpers caja()/texto()/_tabla_casos(), dentro del área libre que deja el layout
"7 Solo encabezado" entre su header y su footer.

Fase 2: genera 15 slides (Portada, Agenda, Título sección, Línea Base Real, Proyección
y Pipeline, Casos en Proceso Administrativo [si aplica], Forecast Total, Análisis de
Desviación, Evolución 7+5, Facturación Mensual, 3 slides de Top 10 casos por Pérdida
Bruta segmentados por Probabilidad de Cierre [100%, 75% y menor a 50%], y Cierre). Las
2 slides comparativas que dependen del historial acumulado de forecasts (Comparativo
Forecasts y Tabla Comparativa) llegan en la Fase 3, junto con la persistencia del
historial.
"""
import os
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

# --- Paleta real del theme "JPV Asociados" (extraída de la plantilla corporativa) ---
COLOR_BG = "FFFFFF"           # blanco, usado para tarjetas y filas de tablas
COLOR_CARD_BORDER = "C7D0DA"  # accent6
COLOR_TEAL = "0D7377"         # accent1 — usado para Ingeniería/IE
COLOR_TEAL2 = "10AA9E"        # accent2 — usado para Proceso Administrativo/Engie
COLOR_GRIS_AZUL = "5B6B7F"    # accent4 — usado para Equipo Móvil/EM
COLOR_TEXTO = "1B2A4B"        # dk1 — texto principal sobre fondo claro
COLOR_MUTED = "5B6B7F"        # accent4 — texto secundario
COLOR_MUTED2 = "6B7C93"       # texto terciario (fuente de datos)
COLOR_VERDE = "2F855A"
COLOR_ROJO = "C53030"
FUENTE = "Calibri"
FUENTE_TITULO = "Cambria"

# Alias semánticos usados en las tarjetas/gráficos de cada división.
COLOR_IE = COLOR_TEAL
COLOR_EM = COLOR_GRIS_AZUL
COLOR_CTM = COLOR_TEAL2

# Tamaño de slide 13.333in x 7.5in (16:9 estándar) — igual al de la plantilla corporativa.
SLIDE_W = Emu(12192000)
SLIDE_H = Emu(6858000)
MARGEN_X = Emu(609600)
CONTENIDO_W = Emu(10972800)

_ASSETS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
PLANTILLA_PATH = os.path.join(_ASSETS_DIR, "plantilla_jpv.pptx")

# Índices (orden lógico, no de archivo) de los layouts de la plantilla que se usan acá.
LAYOUT_PORTADA = 1
LAYOUT_SEPARADOR = 2
LAYOUT_CONTENIDO = 7
LAYOUT_CIERRE = 9


def _rgb(hex_str):
    return RGBColor.from_string(hex_str)


def _tinte(hex_color, factor):
    """Mezcla hex_color con blanco. factor=0 -> color original, factor=1 -> blanco puro.
    Se usa para bandas de tablas y fondos suaves derivados de un color de acento."""
    r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
    r = round(r + (255 - r) * factor)
    g = round(g + (255 - g) * factor)
    b = round(b + (255 - b) * factor)
    return f"{r:02X}{g:02X}{b:02X}"


def _sin_borde(shape):
    shape.line.fill.background()


def _cargar_plantilla():
    if not os.path.exists(PLANTILLA_PATH):
        raise FileNotFoundError(
            f"No se encontró la plantilla corporativa en {PLANTILLA_PATH}. "
            "Todas las slides se generan a partir de ese archivo — sin él no se puede "
            "construir la presentación."
        )
    return Presentation(PLANTILLA_PATH)


def _eliminar_slide(prs, indice):
    """Quita una slide de la presentación (python-pptx no trae este método)."""
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    rId = slides[indice].rId
    prs.part.drop_rel(rId)
    xml_slides.remove(slides[indice])


def _texto_run(run, texto_valor, size=None, bold=None, italic=None, color_hex=None):
    """Sobrescribe solo los atributos indicados de un run (el resto queda heredado del
    layout, p.ej. fuente y alineación), para no perder el estilo oficial de un placeholder."""
    run.text = texto_valor
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if color_hex is not None:
        run.font.color.rgb = _rgb(color_hex)


def _set_placeholder(slide, idx, texto_valor, **kw):
    """Escribe texto en el placeholder idx de la slide, heredando el estilo del layout
    (fuente/tamaño/color) salvo los atributos que se pasen explícitamente en kw."""
    ph = slide.placeholders[idx]
    ph.text_frame.clear()
    run = ph.text_frame.paragraphs[0].add_run()
    _texto_run(run, texto_valor, **kw)
    return ph


def caja(slide, x, y, w, h, fill_hex=COLOR_BG, border_hex=COLOR_CARD_BORDER):
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
          align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE, margenes=True, wrap=True, fuente=None):
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
        run.font.name = fuente or FUENTE
        run.font.color.rgb = _rgb(color_hex)
    return box


def fuente_slide(slide, fuente_texto):
    """Línea pequeña con la fuente de datos, justo sobre el pie de página de la plantilla
    (que ya trae el tagline de marca y el número de slide heredados del layout)."""
    texto(slide, MARGEN_X, Emu(6185280), CONTENIDO_W, Emu(228600),
          fuente_texto, 9, COLOR_MUTED2, margenes=False)


def nueva_slide(prs, titulo, es_avance=False, etiqueta=None):
    """Crea una slide a partir del layout 'Solo encabezado' de la plantilla: el header
    (logo, barra navy, línea de acento) y el pie de página (tagline + número de slide)
    quedan heredados automáticamente del layout, acá solo se completa el título y,
    opcionalmente, la etiqueta de la esquina superior derecha."""
    layout = prs.slide_layouts[LAYOUT_CONTENIDO]
    slide = prs.slides.add_slide(layout)
    _set_placeholder(slide, 105, titulo)
    if es_avance:
        _set_placeholder(slide, 106, "AVANCE", bold=True, italic=False, color_hex=COLOR_ROJO)
    elif etiqueta:
        _set_placeholder(slide, 106, etiqueta)
    return slide


def subtitulo_slide(slide, texto_sub, color_hex=COLOR_TEAL):
    texto(slide, MARGEN_X, Emu(975360), CONTENIDO_W, Emu(228600), texto_sub, 14, color_hex, bold=True)


def _fmt_uf(valor):
    return f"{valor:,.0f}".replace(",", ".")


def _fmt_uf2(valor):
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


# ---------------------------------------------------------
# SLIDE — PORTADA (layout "1 Portada" de la plantilla)
# ---------------------------------------------------------
def slide_portada(prs, datos):
    layout = prs.slide_layouts[LAYOUT_PORTADA]
    slide = prs.slides.add_slide(layout)
    es_avance = datos.get("es_avance", False)

    antetitulo = "LÍNEA BASE FORECAST"
    if es_avance:
        antetitulo = "AVANCE — " + antetitulo
        _set_placeholder(slide, 104, antetitulo, color_hex=COLOR_ROJO)
    else:
        _set_placeholder(slide, 104, antetitulo)

    _set_placeholder(slide, 105, f"Ingeniería y Equipo Móvil {datos['anio']}")
    _set_placeholder(slide, 106, f"Forecast {datos['label']} — {datos['nombre_mes_corte']} {datos['anio']}")
    _set_placeholder(slide, 108, "JPV Asociados")
    _set_placeholder(slide, 109, datos["fecha_emision"])
    return slide


# ---------------------------------------------------------
# SLIDE — AGENDA
# ---------------------------------------------------------
def slide_agenda(prs, datos):
    slide = nueva_slide(prs, "Agenda", datos.get("es_avance", False))

    items = [
        f"Forecast {datos['label']} {datos['anio']}",
        "Análisis de Desempeño Histórico",
        "La Mirada de la Industria",
        "Plan de Acción",
    ]
    y = 1219200
    for i, item in enumerate(items, start=1):
        caja(slide, MARGEN_X, Emu(y), Emu(731520), Emu(731520), fill_hex=_tinte(COLOR_TEAL, 0.90), border_hex=None)
        texto(slide, MARGEN_X, Emu(y), Emu(731520), Emu(731520), str(i), 29.5, COLOR_TEAL, bold=True,
              align=PP_ALIGN.CENTER, margenes=False)
        texto(slide, Emu(457200 + 640080), Emu(y), Emu(8534400), Emu(731520), item, 21.5, COLOR_TEXTO,
              margenes=False)
        y += 853440

    fuente_slide(slide, "JPV Asociados · Sedgwick Alliance")
    return slide


# ---------------------------------------------------------
# SLIDE — TÍTULO DE SECCIÓN (layout "2 Separador de sección")
# ---------------------------------------------------------
def slide_titulo_seccion(prs, datos):
    layout = prs.slide_layouts[LAYOUT_SEPARADOR]
    slide = prs.slides.add_slide(layout)
    _set_placeholder(slide, 105, f"Forecast {datos['label']} {datos['anio']}")
    if datos.get("es_avance"):
        _set_placeholder(slide, 106, "AVANCE", bold=True, italic=False, color_hex=COLOR_ROJO)
    _set_placeholder(slide, 108, f"Forecast Financiero {datos['anio']}: Análisis de Producción y Facturación")
    _set_placeholder(slide, 109, "La Brecha del Forecast (La Realidad Incómoda)")
    return slide


# ---------------------------------------------------------
# SLIDE — LÍNEA BASE REAL
# ---------------------------------------------------------
def slide_linea_base_real(prs, datos):
    slide = nueva_slide(prs, "1. Línea Base Real (Cierre Facturación)", datos.get("es_avance", False))
    subtitulo_slide(slide, f"Estado Actual (YTD {datos['nombre_mes_inicio']} a {datos['nombre_mes_corte']})")

    tarjetas = [
        ("Equipo Móvil", _fmt_uf2(datos['em_ytd']) + " UF", f"{datos['nombre_mes_inicio']}–{datos['nombre_mes_corte']} {datos['anio']}"),
        ("Ingeniería y Energía", _fmt_uf2(datos['ie_ytd']) + " UF", f"{datos['nombre_mes_inicio']}–{datos['nombre_mes_corte']} {datos['anio']}"),
        ("Total Consolidado YTD", _fmt_uf2(datos['ytd_total']) + " UF", f"{datos['nombre_mes_inicio']}–{datos['nombre_mes_corte']} {datos['anio']}"),
    ]
    x = 426720
    for titulo_t, valor_t, sub_t in tarjetas:
        caja(slide, Emu(x), Emu(1219200), Emu(2987040), Emu(1280160))
        texto(slide, Emu(x), Emu(1292352), Emu(2987040), Emu(268224), titulo_t, 12.0, COLOR_MUTED,
              align=PP_ALIGN.CENTER, margenes=False)
        texto(slide, Emu(x), Emu(1560576), Emu(2987040), Emu(487680), valor_t, 26.5, COLOR_TEXTO, bold=True,
              align=PP_ALIGN.CENTER, margenes=False)
        texto(slide, Emu(x), Emu(2048256), Emu(2987040), Emu(280416), sub_t, 10.5, COLOR_MUTED,
              align=PP_ALIGN.CENTER, margenes=False)
        x += 2987040 + 109728

    x = 426720
    tarjetas2 = [
        ("Honorarios en Stock", _fmt_uf(datos['pipeline_bruto_total']) + " UF", "Pipeline bruto total " + str(datos['anio'])),
        ("Meta Anual", _fmt_uf(datos['meta']) + " UF", str(datos['anio'])),
    ]
    for titulo_t, valor_t, sub_t in tarjetas2:
        caja(slide, Emu(x), Emu(2682240), Emu(2987040), Emu(1280160))
        texto(slide, Emu(x), Emu(2755392), Emu(2987040), Emu(268224), titulo_t, 12.0, COLOR_MUTED,
              align=PP_ALIGN.CENTER, margenes=False)
        texto(slide, Emu(x), Emu(3023616), Emu(2987040), Emu(487680), valor_t, 26.5, COLOR_TEXTO, bold=True,
              align=PP_ALIGN.CENTER, margenes=False)
        texto(slide, Emu(x), Emu(3511296), Emu(2987040), Emu(280416), sub_t, 10.5, COLOR_MUTED,
              align=PP_ALIGN.CENTER, margenes=False)
        x += 2987040 + 109728

    texto(slide, MARGEN_X, Emu(4267200), Emu(5333333), Emu(304800), "Análisis de contexto", 14.0, COLOR_TEAL, bold=True)
    bullets = [f"• {b}" for b in datos.get('bullets_linea_base', [])]
    if bullets:
        texto(slide, MARGEN_X, Emu(4632960), Emu(10972800), Emu(609600), bullets, 12.5, COLOR_TEXTO, margenes=False)

    fuente_slide(slide, datos['fuente_texto'])
    return slide


# ---------------------------------------------------------
# SLIDE — PROYECCIÓN Y PIPELINE
# ---------------------------------------------------------
def slide_proyeccion_pipeline(prs, datos):
    slide = nueva_slide(prs, f"2. Proyección y Pipeline {datos['label']} (Resto del año)", datos.get("es_avance", False))
    subtitulo_slide(slide, f"Proyección (Forecast {datos['nombre_mes_corte']}—Diciembre)")

    proy_em = datos['proy_em']
    caja(slide, Emu(426720), Emu(1219200), Emu(5333333), Emu(1706880))
    texto(slide, Emu(548640), Emu(1292352), Emu(5066667), Emu(304800),
          "Equipo Móvil (Pipeline Hon. Probables)", 13.5, COLOR_EM, bold=True, margenes=False)
    texto(slide, Emu(548640), Emu(1645920), Emu(5066667), Emu(426720),
          f"{_fmt_uf(proy_em['em_stock'])} UF (stock)  ·  {_fmt_uf(proy_em['em_promedio_total'])} UF (promedio × meses)",
          17.5, COLOR_TEXTO, bold=True, margenes=False)
    meses_txt = "-".join(m[:3] for m in proy_em['meses_usados']) or "—"
    texto(slide, Emu(548640), Emu(2133600), Emu(5066667), Emu(609600),
          f"Complemento EM — Promedio últimos {len(proy_em['meses_usados'])} meses\n"
          f"{_fmt_uf(proy_em['prom_em_3m'])} UF/mes ({meses_txt}) × {datos['meses_proyectados']} meses = {_fmt_uf(proy_em['em_promedio_total'])} UF",
          12.0, COLOR_MUTED, margenes=False)

    if datos['total_admin'] > 0:
        caja(slide, Emu(5943600), Emu(1219200), Emu(5402400), Emu(1706880), border_hex=COLOR_CTM)
        texto(slide, Emu(6065520), Emu(1292352), Emu(5066667), Emu(304800),
              "Proceso Administrativo de Facturación", 13.5, COLOR_CTM, bold=True, margenes=False)
        texto(slide, Emu(6065520), Emu(1645920), Emu(5066667), Emu(426720),
              f"{_fmt_uf(datos['total_admin'])} UF", 24.0, COLOR_TEXTO, bold=True, margenes=False)
        texto(slide, Emu(6065520), Emu(2133600), Emu(5066667), Emu(609600),
              f"{len(datos['casos_admin'])} caso(s) en proceso administrativo",
              12.0, COLOR_MUTED, margenes=False)

    proy_ie = datos['proy_ie']
    texto(slide, MARGEN_X, Emu(3169920), Emu(8000000), Emu(304800),
          f"Ingeniería — Proyección {datos['nombre_mes_corte']}—Diciembre:", 14.0, COLOR_IE, bold=True)

    ie_tarjetas = [
        ("Casos menores", "(<1.000 UF pérdida)", _fmt_uf(proy_ie['ie_menores_proj']) + " UF",
         f"{_fmt_uf(proy_ie['prom_ie_menores'])} UF/mes × {datos['meses_proyectados']} meses"),
        ("Casos mayores pipeline", "(≥1.000 UF, prob.)", _fmt_uf(proy_ie['ie_mayores_proj']) + " UF",
         f"{proy_ie['n_casos_mayores']} casos · Hon. Probables ponderados"),
        ("Engie/CTM", "(pipeline, prob.)", _fmt_uf(proy_ie['ie_engie_proj']) + " UF",
         f"{proy_ie['n_casos_engie']} caso(s) · Hon. Probables ponderados"),
    ]
    x = 426720
    for t1, t2, valor_t, sub_t in ie_tarjetas:
        caja(slide, Emu(x), Emu(3657600), Emu(3474720), Emu(1463040))
        texto(slide, Emu(x), Emu(3730752), Emu(3474720), Emu(240000), t1, 12.5, COLOR_MUTED, bold=True,
              align=PP_ALIGN.CENTER, margenes=False)
        texto(slide, Emu(x), Emu(3970752), Emu(3474720), Emu(213333), t2, 10.5, COLOR_MUTED,
              align=PP_ALIGN.CENTER, margenes=False)
        texto(slide, Emu(x), Emu(4200000), Emu(3474720), Emu(426720), valor_t, 22.5, COLOR_TEXTO, bold=True,
              align=PP_ALIGN.CENTER, margenes=False)
        texto(slide, Emu(x), Emu(4693333), Emu(3474720), Emu(373333), sub_t, 10.5, COLOR_MUTED,
              align=PP_ALIGN.CENTER, margenes=False)
        x += 3474720 + 73333

    fuente_slide(slide, datos['fuente_texto'])
    return slide


# ---------------------------------------------------------
# SLIDE — CASOS EN PROCESO ADMINISTRATIVO (condicional)
# ---------------------------------------------------------
def slide_casos_admin(prs, datos):
    slide = nueva_slide(prs, "Casos en Proceso Administrativo de Facturación", datos.get("es_avance", False))
    subtitulo_slide(slide, f"Total: {_fmt_uf(datos['total_admin'])} UF", color_hex=COLOR_CTM)

    y = 1219200
    for caso in datos['casos_admin']:
        caja(slide, MARGEN_X, Emu(y), CONTENIDO_W, Emu(609600))
        texto(slide, Emu(457200 + 137160), Emu(y), Emu(8000000), Emu(609600),
              caso['nombre'] or "(sin nombre)", 16.0, COLOR_TEXTO, margenes=False)
        texto(slide, Emu(9066667), Emu(y), Emu(1829333), Emu(609600),
              f"{_fmt_uf(caso['monto'])} UF", 17.5, COLOR_CTM, bold=True,
              align=PP_ALIGN.RIGHT, margenes=False)
        y += 670560
        if y > 5600000:
            break

    fuente_slide(slide, datos['fuente_texto'])
    return slide


# ---------------------------------------------------------
# SLIDE — FORECAST TOTAL
# ---------------------------------------------------------
def slide_forecast_total(prs, datos):
    slide = nueva_slide(prs, f"2. Forecast {datos['label']} (Total año {datos['anio']})", datos.get("es_avance", False))
    subtitulo_slide(slide, f"Proyección (Forecast {datos['nombre_mes_corte']}—Diciembre)")

    tarjetas = [
        ("Equipo Móvil", COLOR_EM, _fmt_uf(datos['em_total']) + " UF",
         f"{_fmt_uf2(datos['em_ytd'])} YTD + {_fmt_uf(datos['em_proyectado'])} proyectado"),
        ("Ingeniería", COLOR_IE, _fmt_uf(datos['ie_total']) + " UF",
         f"{_fmt_uf2(datos['ie_ytd'])} YTD + {_fmt_uf(datos['ie_proyectado'])} proyectado"),
    ]
    x = 426720
    for titulo_t, color_t, valor_t, sub_t in tarjetas:
        caja(slide, Emu(x), Emu(1463040), Emu(3719407), Emu(1706880))
        texto(slide, Emu(x + 91440), Emu(1536192), Emu(3466667), Emu(304800), titulo_t, 14.0, color_t, bold=True, margenes=False)
        texto(slide, Emu(x + 91440), Emu(1889760), Emu(3466667), Emu(487680), valor_t, 29.5, COLOR_TEXTO, bold=True, margenes=False)
        texto(slide, Emu(x + 91440), Emu(2377440), Emu(3466667), Emu(426720), sub_t, 12.0, COLOR_MUTED, margenes=False)
        x += 3719407 + 121927

    if datos['total_admin'] > 0:
        caja(slide, Emu(x), Emu(1463040), Emu(3719407), Emu(1706880), border_hex=COLOR_CTM)
        texto(slide, Emu(x + 91440), Emu(1536192), Emu(3466667), Emu(304800),
              "Proceso Administrativo", 14.0, COLOR_CTM, bold=True, margenes=False)
        texto(slide, Emu(x + 91440), Emu(1889760), Emu(3466667), Emu(487680),
              _fmt_uf(datos['total_admin']) + " UF", 29.5, COLOR_TEXTO, bold=True, margenes=False)
        texto(slide, Emu(x + 91440), Emu(2377440), Emu(3466667), Emu(426720),
              "Cierre esperado próximas semanas", 12.0, COLOR_MUTED, margenes=False)

    texto(slide, MARGEN_X, Emu(3474720), Emu(8000000), Emu(304800),
          "Cierre de año proyectado global", 16.0, COLOR_TEXTO, bold=True)

    color_gap_sin = COLOR_VERDE if datos['cum_sin'] >= 100 else COLOR_EM
    caja(slide, Emu(426720), Emu(3962400), Emu(5242560), Emu(1584960))
    texto(slide, Emu(426720), Emu(4084320), Emu(5242560), Emu(487680),
          _fmt_uf(datos['cierre_sin_admin']) + " UF", 29.5, COLOR_TEXTO, bold=True,
          align=PP_ALIGN.CENTER, margenes=False)
    texto(slide, Emu(426720), Emu(4572000), Emu(5242560), Emu(304800),
          f"Sin proceso administrativo · {datos['cum_sin']:.1f}% de la meta", 12.5, color_gap_sin,
          align=PP_ALIGN.CENTER, margenes=False)

    if datos['total_admin'] > 0:
        color_gap_con = COLOR_VERDE if datos['cum_con'] >= 100 else COLOR_EM
        caja(slide, Emu(5791200), Emu(3962400), Emu(5242560), Emu(1584960), border_hex=COLOR_CTM)
        texto(slide, Emu(5791200), Emu(4084320), Emu(5242560), Emu(487680),
              _fmt_uf(datos['cierre_con_admin']) + " UF", 29.5, COLOR_TEXTO, bold=True,
              align=PP_ALIGN.CENTER, margenes=False)
        texto(slide, Emu(5791200), Emu(4572000), Emu(5242560), Emu(304800),
              f"Con proceso administrativo · {datos['cum_con']:.1f}% de la meta", 12.5, color_gap_con,
              align=PP_ALIGN.CENTER, margenes=False)

    texto(slide, MARGEN_X, Emu(5730240), CONTENIDO_W, Emu(304800),
          f"Honorarios en Stock: {_fmt_uf(datos['pipeline_bruto_total'])} UF no facturados en pipeline total (bruto, sin probabilidad)",
          12.0, COLOR_MUTED)

    fuente_slide(slide, datos['fuente_texto'])
    return slide


# ---------------------------------------------------------
# SLIDE — ANÁLISIS DE DESVIACIÓN
# ---------------------------------------------------------
def slide_desviacion(prs, datos):
    slide = nueva_slide(prs, f"3. Análisis de Desviación (Gap) vs. Meta — Forecast {datos['label']}", datos.get("es_avance", False))
    subtitulo_slide(slide, "Evaluación de Brecha Estratégica y Acciones Mitigadoras")

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
    x = 426720
    for titulo_t, valor_t, sub_t, color_t in etiquetas:
        caja(slide, Emu(x), Emu(1341120), Emu(3474720), Emu(1219200))
        texto(slide, Emu(x), Emu(1414272), Emu(3474720), Emu(268224), titulo_t, 12.0, COLOR_MUTED,
              align=PP_ALIGN.CENTER, margenes=False)
        texto(slide, Emu(x), Emu(1682496), Emu(3474720), Emu(426720), valor_t, 24.0, color_t, bold=True,
              align=PP_ALIGN.CENTER, margenes=False)
        texto(slide, Emu(x), Emu(2133600), Emu(3474720), Emu(280416), sub_t, 11.5, COLOR_MUTED,
              align=PP_ALIGN.CENTER, margenes=False)
        x += 3474720 + 73333

    texto(slide, Emu(426720), Emu(2804160), Emu(5242560), Emu(304800),
          "Acción Requerida: Equipo Móvil", 14.0, COLOR_EM, bold=True)
    bullets_em = [f"• {b}" for b in datos.get('bullets_accion_em', [])]
    texto(slide, Emu(426720), Emu(3169920), Emu(5242560), Emu(2438400), bullets_em, 12.0, COLOR_TEXTO, margenes=False)

    texto(slide, Emu(5943600), Emu(2804160), Emu(5242560), Emu(304800),
          "Acción Requerida: Ingeniería", 14.0, COLOR_IE, bold=True)
    bullets_ie = [f"• {b}" for b in datos.get('bullets_accion_ie', [])]
    texto(slide, Emu(5943600), Emu(3169920), Emu(5242560), Emu(2438400), bullets_ie, 12.0, COLOR_TEXTO, margenes=False)

    fuente_slide(slide, datos['fuente_texto'])
    return slide


# ---------------------------------------------------------
# Helpers de estilo para gráficos nativos
# ---------------------------------------------------------
def _estilizar_chart(chart, mostrar_leyenda=True):
    chart.has_title = False
    if mostrar_leyenda:
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(12)
        chart.legend.font.color.rgb = _rgb(COLOR_MUTED)
        chart.legend.font.name = FUENTE
    else:
        chart.has_legend = False

    cat_ax = chart.category_axis
    cat_ax.tick_labels.font.size = Pt(12)
    cat_ax.tick_labels.font.color.rgb = _rgb(COLOR_MUTED)
    cat_ax.tick_labels.font.name = FUENTE
    cat_ax.format.line.color.rgb = _rgb(COLOR_CARD_BORDER)

    val_ax = chart.value_axis
    val_ax.tick_labels.font.size = Pt(12)
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
    slide = nueva_slide(prs, f"Evolución Forecast {datos['label']} año {datos['anio']}", datos.get("es_avance", False))
    subtitulo_slide(slide, "Real acumulado + proyección hasta diciembre")

    chart_data = CategoryChartData()
    chart_data.categories = datos['grafico_meses']
    chart_data.add_series('Real', datos['grafico_real'])
    chart_data.add_series('Proyección (sin proc. admin.)', datos['grafico_proy_sin'])
    if datos['total_admin'] > 0:
        chart_data.add_series('Proyección (con proc. admin.)', datos['grafico_proy_con'])
    chart_data.add_series('Meta', [datos['meta']] * len(datos['grafico_meses']))

    x, y, cx, cy = Emu(609600), Emu(1341120), Emu(10972800), Emu(4632960)
    gf = slide.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, x, y, cx, cy, chart_data)
    chart = gf.chart
    _estilizar_chart(chart)
    colores = [COLOR_TEXTO, COLOR_EM, COLOR_CTM, COLOR_MUTED2]
    _colorear_series(chart, colores)

    fuente_slide(slide, datos['fuente_texto'])
    return slide


# ---------------------------------------------------------
# Helper — tabla nativa de casos (encabezado + filas alternadas)
# ---------------------------------------------------------
def _tabla_casos(slide, x, y, cx, cy_max, columnas, filas, color_header_bg, filas_max_diseno=10):
    """columnas: lista de (titulo, ancho_emu), con anchos que deben sumar cx (si no
    coinciden se reescalan proporcionalmente). filas: lista de tuplas de texto, mismo
    orden que columnas. cy_max es el alto disponible para filas_max_diseno filas de
    datos (el caso más grande, p.ej. un Top 10 completo) — el alto de fila se fija en
    base a ese máximo y la tabla se dimensiona solo a las filas que realmente tiene, en
    vez de estirarse a llenar cy_max cuando hay pocos casos. Se dibuja sobre una tarjeta
    blanca con borde, para que se lea como un componente diseñado y no como una grilla
    pegada de una planilla."""
    n_filas = len(filas) + 1
    n_cols = len(columnas)

    anchos = [ancho for _, ancho in columnas]
    suma_anchos = sum(anchos)
    if suma_anchos != int(cx):
        anchos = [round(a / suma_anchos * int(cx)) for a in anchos]
        anchos[-1] += int(cx) - sum(anchos)

    peso_header = 1.3
    alto_normal = int(int(cy_max) / (filas_max_diseno + peso_header))
    alto_header = int(alto_normal * peso_header)
    n_filas_cuerpo = n_filas - 1
    alto_tabla = alto_header + n_filas_cuerpo * alto_normal

    margen_card = Emu(27432)
    caja(slide, Emu(int(x) - margen_card), Emu(int(y) - margen_card),
         Emu(int(cx) + margen_card * 2), Emu(alto_tabla + margen_card * 2),
         fill_hex=COLOR_BG, border_hex=COLOR_CARD_BORDER)

    gf = slide.shapes.add_table(n_filas, n_cols, x, y, cx, Emu(alto_tabla))
    table = gf.table
    table.first_row = False
    table.horz_banding = False

    for c, ancho in enumerate(anchos):
        table.columns[c].width = Emu(ancho)

    table.rows[0].height = Emu(alto_header)
    for r in range(1, n_filas):
        table.rows[r].height = Emu(alto_normal)
    gf.height = Emu(alto_tabla)

    color_tinte_zebra = _tinte(color_header_bg, 0.90)
    color_tinte_destacado = _tinte(color_header_bg, 0.78)

    def _celda(r, c, texto_valor, fondo_hex, color_hex, bold, align):
        cell = table.cell(r, c)
        cell.fill.solid()
        cell.fill.fore_color.rgb = _rgb(fondo_hex)
        cell.margin_left = cell.margin_right = Emu(73152)
        cell.margin_top = cell.margin_bottom = Emu(18288)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = cell.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = align
        run = p.add_run()
        run.text = str(texto_valor)
        run.font.size = Pt(11)
        run.font.bold = bold
        run.font.name = FUENTE
        run.font.color.rgb = _rgb(color_hex)

    for c, (titulo, _) in enumerate(columnas):
        _celda(0, c, titulo.upper(), color_header_bg, "FFFFFF", True,
               PP_ALIGN.LEFT if c < 2 else PP_ALIGN.CENTER)

    ultima_col = len(columnas) - 1
    for r, fila in enumerate(filas, start=1):
        fondo_zebra = COLOR_BG if r % 2 == 1 else color_tinte_zebra
        for c, valor in enumerate(fila):
            es_destacada = c == ultima_col
            fondo = color_tinte_destacado if es_destacada else fondo_zebra
            color_texto = color_header_bg if es_destacada else COLOR_TEXTO
            _celda(r, c, valor, fondo, color_texto, es_destacada,
                   PP_ALIGN.LEFT if c < 2 else PP_ALIGN.CENTER)
    return gf


# ---------------------------------------------------------
# SLIDES — TOP 10 POR PROBABILIDAD DE CIERRE (ordenados por pérdida bruta)
# ---------------------------------------------------------
_COLUMNAS_TOP10 = [
    ("N° Caso", 1266667), ("Nickname", 3133333), ("División", 1000000),
    ("Ajustador", 1933333), ("Pérdida Bruta", 2333333), ("Prob.", 1306134),
]


def slide_top_probabilidad(prs, datos, clave, titulo, subtitulo, color_accent):
    slide = nueva_slide(prs, titulo, datos.get("es_avance", False))
    subtitulo_slide(slide, subtitulo, color_hex=color_accent)

    casos = datos.get(clave, [])
    if not casos:
        texto(slide, MARGEN_X, Emu(2438400), CONTENIDO_W, Emu(609600),
              "No hay casos que cumplan este criterio.", 16.0, COLOR_MUTED)
    else:
        filas = [
            (
                c['caso'], c['nickname'], c['division'], c['ajustador'],
                f"{c['perdida']:,.0f} {c['divisa']}".replace(",", "."),
                c['probabilidad'],
            )
            for c in casos
        ]
        _tabla_casos(slide, MARGEN_X, Emu(1341120), CONTENIDO_W, Emu(4800000),
                     _COLUMNAS_TOP10, filas, color_accent)

    fuente_slide(slide, datos['fuente_texto'])
    return slide


# ---------------------------------------------------------
# SLIDE — FACTURACIÓN MENSUAL POR DIVISIÓN (barras)
# ---------------------------------------------------------
def slide_facturacion_mensual(prs, datos):
    slide = nueva_slide(prs, f"Facturación Mensual por División — {datos['nombre_mes_inicio']} a {datos['nombre_mes_corte']} {datos['anio']}",
                         datos.get("es_avance", False))
    subtitulo_slide(slide, "Real, por mes cerrado")

    chart_data = CategoryChartData()
    chart_data.categories = datos['meses_mensual']
    chart_data.add_series('Equipo Móvil', datos['valores_em_mensual'])
    chart_data.add_series('Ingeniería y Energía', datos['valores_ie_mensual'])

    x, y, cx, cy = Emu(609600), Emu(1341120), Emu(10972800), Emu(4632960)
    gf = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data)
    chart = gf.chart
    _estilizar_chart(chart)
    _colorear_series(chart, [COLOR_EM, COLOR_IE])

    fuente_slide(slide, datos['fuente_texto'])
    return slide


# ---------------------------------------------------------
# SLIDE — CIERRE (layout "9 Cierre" de la plantilla)
# ---------------------------------------------------------
def slide_cierre(prs, datos):
    layout = prs.slide_layouts[LAYOUT_CIERRE]
    slide = prs.slides.add_slide(layout)
    _set_placeholder(slide, 104, "Gracias")
    _set_placeholder(slide, 105, "Línea Base Forecast — Ingeniería y Equipo Móvil")
    _set_placeholder(slide, 106, "JPV Asociados")
    _set_placeholder(slide, 107, datos["fecha_emision"])
    return slide


def generar_pptx_forecast(datos):
    """Genera la presentación completa y devuelve los bytes del archivo .pptx."""
    import io as _io
    prs = _cargar_plantilla()
    _eliminar_slide(prs, 0)  # quita la slide de ejemplo incluida en la plantilla

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
    slide_top_probabilidad(prs, datos, 'top10_100', "Top 10 — Casos con Probabilidad de Cierre 100%",
                            "Ordenados por Pérdida Bruta (mayor a menor)", COLOR_VERDE)
    slide_top_probabilidad(prs, datos, 'top10_75', "Top 10 — Casos con Probabilidad de Cierre 75%",
                            "Ordenados por Pérdida Bruta (mayor a menor)", COLOR_IE)
    slide_top_probabilidad(prs, datos, 'top10_menor50', "Top 10 — Casos con Probabilidad de Cierre < 50%",
                            "Ordenados por Pérdida Bruta (mayor a menor)", COLOR_ROJO)
    slide_cierre(prs, datos)

    buffer = _io.BytesIO()
    prs.save(buffer)
    return buffer.getvalue()
