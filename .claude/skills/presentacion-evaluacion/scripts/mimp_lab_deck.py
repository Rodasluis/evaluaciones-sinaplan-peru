"""Copia de la skill ppt-mimp-lab para la skill del repositorio presentacion-evaluacion, con la corrección de los effectLst duplicados (_shadow).


mimp_lab_deck — plantilla del MIMP Lab (Laboratorio de Innovación del MIMP)
para python-pptx. Hermana de `mimp_deck` (skill `ppt-mimp`), pero con una
paleta multicolor y elementos gráficos pensados para verse innovadores y
dinámicos, en vez de la paleta institucional azul/rojo estricta.

Uso mínimo:

    import sys; sys.path.insert(0, "<ruta-skill>/scripts")
    from mimp_lab_deck import MimpLabDeck, PETROL, TEAL, GOLD, RED, VIOLET, GREEN

    d = MimpLabDeck()
    d.cover("MIMP Lab", "Laboratorio de Innovación del MIMP", fecha="Agosto de 2026")
    s = d.content("Título de la diapositiva", icon="idea")
    d.text(s, d.M, d.BODY_TOP, d.CW, 0.5, [("Objetivo. ", True, GOLD), ("...", False)])
    d.closing()
    d.save("salida.pptx")

Todas las medidas están en pulgadas; la diapositiva es 13.333 × 7.5 (16:9).
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.oxml.ns import qn
from lxml import etree
import os

# ------------------------------------------------------------------ paleta
# Color oscuro de texto/títulos (equivalente al NAVY institucional, pero
# tomado del propio logo del MIMP Lab: un petróleo profundo, no un azul
# corporativo genérico).
PETROL = RGBColor(0x03, 0x47, 0x54)
PETROL2 = RGBColor(0x0A, 0x5E, 0x6E)     # variante intermedia para alternar

# Acentos rotativos — este es el cambio de fondo respecto a ppt-mimp: en vez
# de un solo color de acento (rojo) + un color secundario (naranja), hay
# acentos que se reparten entre tarjetas, pasos y filas de tablas para que
# el deck se vea variado. Elegir el acento por el significado (ver
# `references/formato.md`), no por rotación ciega dentro de un mismo bloque.
TEAL = RGBColor(0x14, 0xA7, 0xAE)        # acento primario — identidad MIMP Lab
GOLD = RGBColor(0xFE, 0xC3, 0x4D)        # acento secundario — foco / destacado
RED = RGBColor(0xEA, 0x26, 0x39)         # alerta / urgencia / hitos críticos
VIOLET = RGBColor(0x7C, 0x6F, 0xC4)      # investigación / conocimiento
GREEN = RGBColor(0x4C, 0xAF, 0x93)       # avance / resultados / cierre

# RED queda fuera de la rotación automática: es un color de alerta y usarlo
# "porque tocaba turno" en una tarjeta neutral manda una señal equivocada
# (parece un riesgo cuando no lo es). Sigue disponible para usarlo a mano
# (`accent=RED`, `icon(..., color=RED)`) cuando el contenido SÍ es una
# alerta, un riesgo o un plazo crítico.
PALETTE = [TEAL, GOLD, VIOLET, GREEN]

WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x00, 0x00, 0x00)
GRAY_TXT = RGBColor(0x5A, 0x5A, 0x5A)
GRAY_PAGE = RGBColor(0x9A, 0x9A, 0x9A)
CARD_LINE = RGBColor(0xD8, 0xD8, 0xD8)
BG = RGBColor(0xF4, 0xF4, 0xF3)          # respaldo si no hay imagen de textura
FONT = "Poppins"

# Íconos sugeridos por temática (glifos Unicode de trazo simple: se colorean
# con el color de fuente, a diferencia de los emoji a color que no se pueden
# recolorear). Ver `icon()` y la tabla completa en references/formato.md.
ICONS = {
    "idea": "\u2726",          # innovación / propuesta
    "objetivo": "\u25CE",      # objetivo / meta
    "datos": "\u25A4",         # datos / evidencia
    "investigacion": "\u2697", # investigación / evaluación
    "colaboracion": "\u2B22",  # alianzas / articulación (red / hexágono)
    "premio": "\u2606",        # reconocimiento / premiación
    "calendario": "\u25A2",    # cronograma / plazos
    "check": "\u2713",         # cumplimiento / requisito
    "engranaje": "\u2699",     # proceso / mecanismo
    "alcance": "\u25C9",       # alcance / cobertura
    "dinero": "\u20A4",        # presupuesto / financiamiento
    "alerta": "\u26A0",        # riesgo / atención
    "personas": "\u263A",      # equipo / actores / postulantes
    "documento": "\u2261",     # informe / expediente / documentación
    "tecnologia": "\u2318",    # plataforma / sistema / canal digital
    "crecimiento": "\u2197",   # avance / mejora / resultado positivo
    "ubicacion": "\u2316",     # cobertura territorial / sede
}

_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'


def tint(color, factor=0.82):
    """Mezcla `color` hacia blanco. factor=0 -> color puro, 1 -> blanco.
    Sirve para generar automáticamente la versión "pálida" de cualquiera de
    los acentos, sin tener que declarar a mano una paleta pálida por color.
    """
    r = int(color[0] + (255 - color[0]) * factor)
    g = int(color[1] + (255 - color[1]) * factor)
    b = int(color[2] + (255 - color[2]) * factor)
    return RGBColor(r, g, b)

_AVISADOS = set()


def _aviso(texto):
    """Avisa una sola vez de un recurso que falta (versión pública sin la identidad institucional)."""
    if texto not in _AVISADOS:
        _AVISADOS.add(texto)
        print("aviso:", texto)


class MimpLabDeck:
    """Constructor de presentaciones con el estilo del MIMP Lab."""

    # ---- geometría fija (idéntica a ppt-mimp: no cambiarla entre láminas)
    M = 0.92
    CW = 11.49
    TITLE_Y = 1.02
    TITLE_H = 0.86
    RULE_Y = 1.96
    BODY_TOP = 2.15
    BODY_BOT = 6.90
    W = 13.333
    H = 7.5

    def __init__(self, assets_dir=None, existing=None):
        self.assets = assets_dir or os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "..", "assets")
        if existing:
            self.prs = Presentation(existing)
        else:
            self.prs = Presentation()
            self.prs.slide_width = Inches(self.W)
            self.prs.slide_height = Inches(self.H)
        self.layout = self.prs.slide_layouts[6 if len(self.prs.slide_layouts) > 6 else 0]

    # ------------------------------------------------------------ primitivas
    def _shadow(self, shape):
        # corrección de la skill del repositorio (Fase 2): shadow.inherit = False deja un <a:effectLst/> vacío y, con
        # la sombra, el spPr queda con dos effectLst, que PowerPoint pide reparar; se quita el vacío
        spPr = shape._element.spPr
        for vacio in [e for e in spPr.findall(qn("a:effectLst")) if len(e) == 0]:
            spPr.remove(vacio)
        spPr.append(etree.fromstring(
            f'<a:effectLst {_NS}><a:outerShdw blurRad="63500" dist="12700" '
            'dir="5400000" algn="ctr" rotWithShape="0">'
            '<a:srgbClr val="B0B0B0"><a:alpha val="30000"/></a:srgbClr>'
            '</a:outerShdw></a:effectLst>'))

    def rect(self, s, x, y, w, h, fill, line=None, radius=None, shadow=False,
             dash=False):
        shp = s.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,
            Inches(x), Inches(y), Inches(w), Inches(h))
        if radius:
            shp.adjustments[0] = min(0.5, radius / min(w, h))
        if fill is None:
            shp.fill.background()
        else:
            shp.fill.solid()
            shp.fill.fore_color.rgb = fill
        if line is None:
            shp.line.fill.background()
        else:
            shp.line.color.rgb = line
            shp.line.width = Pt(0.75)
            if dash:
                shp.line.dash_style = MSO_LINE_DASH_STYLE.DASH
        shp.shadow.inherit = False
        if shadow:
            self._shadow(shp)
        return shp

    def oval(self, s, x, y, d_, fill, line=None):
        """Círculo de diámetro `d_` (pulgadas), usado como insignia de ícono."""
        shp = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y),
                                 Inches(d_), Inches(d_))
        if fill is None:
            shp.fill.background()
        else:
            shp.fill.solid()
            shp.fill.fore_color.rgb = fill
        if line is None:
            shp.line.fill.background()
        else:
            shp.line.color.rgb = line
            shp.line.width = Pt(1)
        shp.shadow.inherit = False
        return shp

    def icon(self, s, x, y, size, glyph, color=TEAL, glyph_color=WHITE):
        """Insignia circular de ícono: círculo de `size` pulgadas de diámetro,
        relleno `color`, con el glifo `glyph` centrado (clave de ICONS o un
        carácter Unicode propio). Devuelve la esquina (x, y, size) por si hay
        que alinear texto junto a ella.
        """
        self.oval(s, x, y, size, color)
        glyph = ICONS.get(glyph, glyph)
        self.text(s, x, y, size, size, glyph, size=size * 34, color=glyph_color,
                  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        return x, y, size

    def text(self, s, x, y, w, h, runs, size=11, color=PETROL, bold=False,
             italic=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
             space_after=0, bullet=False, line_spacing=None):
        """Cuadro de texto. Ver `mimp_deck.text` (misma firma y semántica)."""
        box = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = anchor

        if isinstance(runs, str):
            paras = [[(runs, bold)]]
        elif runs and isinstance(runs[0], tuple):
            paras = [runs]
        else:
            paras = runs

        expanded, cur = [], []
        for para in paras:
            cur = []
            for item in para:
                if isinstance(item, str):
                    item = (item, bold)
                for k, part in enumerate(item[0].split("\n")):
                    if k:
                        expanded.append(cur)
                        cur = []
                    if part:
                        cur.append((part,) + tuple(item[1:]))
            expanded.append(cur)
        paras = [p for p in expanded if p] or [[("", bold)]]

        for i, para in enumerate(paras):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            if space_after:
                p.space_after = Pt(space_after)
            if line_spacing:
                p.line_spacing = line_spacing
            if bullet:
                pPr = p._p.get_or_add_pPr()
                pPr.set("marL", "142875")
                pPr.set("indent", "-142875")
                pPr.append(etree.fromstring(f'<a:buFont {_NS} typeface="Arial"/>'))
                pPr.append(etree.fromstring(f'<a:buChar {_NS} char="\u2022"/>'))
            for item in para:
                r = p.add_run()
                r.text = item[0]
                f = r.font
                f.name = FONT
                f.size = Pt(size)
                f.bold = item[1] if len(item) > 1 else bold
                f.italic = italic
                f.color.rgb = item[2] if len(item) > 2 else color
        return box

    # ------------------------------------------------------- piezas del marco
    def band(self, s, bg=None):
        """Fondo + banda blanca superior con los dos logos oficiales."""
        ruta = os.path.join(self.assets, bg) if bg else None
        if ruta and os.path.exists(ruta):
            s.shapes.add_picture(ruta, 0, 0,
                                 Inches(self.W), Inches(self.H))
        else:
            if ruta:
                _aviso(f"Falta {bg} en {self.assets}: el fondo va en color sólido.")
            self.rect(s, 0, 0, self.W, self.H, {"bg_cover.jpg": BG, "bg_content.jpg": BG}.get(bg, BG))
        self.rect(s, 0, -0.5, self.W, 1.41, WHITE, radius=0.18)
        for logo, caja in (("logo_mimp.png", (0.83, 0.19, 2.32, 0.566)), ("logo_gob.png", (10.423, 0.15, 2.08, 0.654))):
            ruta = os.path.join(self.assets, logo)
            if os.path.exists(ruta):
                s.shapes.add_picture(ruta, *(Inches(v) for v in caja))
            else:
                _aviso(f"Falta {logo} en {self.assets}: la banda va sin ese logo.")

    def _footer_bar(self, s):
        """Franja de tres colores al pie — firma visual del MIMP Lab."""
        seg = self.W / 3
        for i, c in enumerate((PETROL, TEAL, GOLD)):
            self.rect(s, i * seg, self.H - 0.10, seg + 0.02, 0.10, c)

    def _page_number(self, s):
        self.text(s, 11.7, 7.02, 0.71, 0.22, str(len(self.prs.slides._sldIdLst)),
                  size=10, color=GRAY_PAGE, align=PP_ALIGN.RIGHT)

    # ------------------------------------------------------ tipos de diapositiva
    def content(self, title, rule=True, icon=None, accent=TEAL):
        """Diapositiva de contenido: textura de fondo, banda, título en
        PETROL y línea de acento (por defecto TEAL). `icon`, si se da (clave
        de ICONS), dibuja una insignia circular a la izquierda del título,
        centrada en la misma línea horizontal (mismo eje vertical: ícono y
        texto comparten centro) — úsalo para anclar visualmente el tema de
        la lámina (ver formato.md).
        """
        s = self.prs.slides.add_slide(self.layout)
        self.band(s, bg="bg_content.jpg")
        tx = self.M
        if icon:
            isize = 0.56
            iy = self.TITLE_Y + (self.TITLE_H - isize) / 2
            self.icon(s, self.M, iy, isize, icon, color=accent)
            tx = self.M + isize + 0.20
        self.text(s, tx, self.TITLE_Y, self.CW - (tx - self.M), self.TITLE_H,
                  title, size=27, bold=True, color=PETROL,
                  anchor=MSO_ANCHOR.MIDDLE if icon else MSO_ANCHOR.BOTTOM)
        if rule:
            self.rect(s, tx, self.RULE_Y, 3.4, 0.045, accent)
        self._footer_bar(s)
        self._page_number(s)
        return s

    def cover(self, title, subtitle=None, unidad=None, fecha=None):
        s = self.prs.slides.add_slide(self.layout)
        self.band(s, bg="bg_cover.jpg")
        self.text(s, 1.2, 2.35, 10.93, 1.75, title, size=44, bold=True,
                  color=PETROL, align=PP_ALIGN.CENTER)
        if subtitle:
            self.text(s, 1.5, 4.05, 10.33, 0.55, subtitle, size=20,
                      color=PETROL2, align=PP_ALIGN.CENTER)
        self.rect(s, 5.17, 4.75, 3.0, 0.045, TEAL)
        if unidad:
            self.text(s, 1.5, 5.0, 10.33, 0.4, unidad, size=13,
                      color=GRAY_TXT, align=PP_ALIGN.CENTER)
        if fecha:
            self.text(s, 1.5, 6.05, 10.33, 0.45, fecha, size=16, bold=True,
                      color=TEAL, align=PP_ALIGN.CENTER)
        return s

    def divider(self, numero, titulo, accent=TEAL):
        """Divisor de sección: fondo sólido del acento elegido (por defecto
        TEAL), alineado a la derecha como en ppt-mimp, pero con el bloque de
        color rotando según el acento que se pase (varía la sección visualmente
        sin perder consistencia)."""
        s = self.prs.slides.add_slide(self.layout)
        self.rect(s, 0, 0, self.W, self.H, accent)
        self.text(s, 4.0, 3.6, 8.41, 0.9, numero, size=46, bold=True,
                  color=tint(accent, 0.35), align=PP_ALIGN.RIGHT)
        self.text(s, 3.0, 4.5, 9.41, 1.1, titulo, size=30, bold=True,
                  color=WHITE, align=PP_ALIGN.RIGHT)
        self.rect(s, 9.41, 5.72, 3.0, 0.03, WHITE)
        self._page_number(s)
        return s

    def closing(self, titulo="\u00a1Gracias!", pie=None):
        s = self.prs.slides.add_slide(self.layout)
        self.band(s, bg="bg_cover.jpg")
        self.text(s, 1.5, 3.05, 10.33, 1.1, titulo, size=48, bold=True,
                  color=PETROL, align=PP_ALIGN.CENTER)
        self.rect(s, 5.17, 4.25, 3.0, 0.045, TEAL)
        if pie:
            self.text(s, 1.5, 4.55, 10.33, 0.45, pie, size=15, color=GRAY_TXT,
                      align=PP_ALIGN.CENTER)
        return s

    # --------------------------------------------------- componentes de cuerpo
    def card(self, s, x, y, w, h, fill=WHITE):
        return self.rect(s, x, y, w, h, fill, CARD_LINE, radius=0.05, shadow=True)

    def header_bar(self, s, x, y, w, h, texto, fill=PETROL, size=11,
                   color=WHITE, icon=None):
        """Barra de encabezado. Si se pasa `icon` (clave de ICONS o carácter
        propio), se dibuja a la izquierda del texto, dentro de la misma
        barra — úsalo en `columns()`/tarjetas para que cada una lleve el
        ícono que corresponde a lo que trata su contenido, no solo un ícono
        genérico de la lámina completa."""
        self.rect(s, x, y, w, h, fill, radius=0.05)
        tx = x + 0.12
        if icon:
            glyph = ICONS.get(icon, icon)
            self.text(s, tx, y, h, h, glyph, size=size + 4, bold=True,
                      color=color, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            tx += h * 0.85
        self.text(s, tx, y, w - (tx - x) - 0.12, h, texto, size=size, bold=True,
                  color=color, anchor=MSO_ANCHOR.MIDDLE)

    def subhead(self, s, x, y, w, texto, size=13, color=None):
        """Subtítulo de bloque. Por defecto usa TEAL (en vez del naranja de
        ppt-mimp) para que el acento de color quede consistente con la
        identidad del Lab; puede pasarse otro acento de la paleta."""
        self.text(s, x, y, w, 0.3, texto, size=size, bold=True,
                  color=color or TEAL, anchor=MSO_ANCHOR.MIDDLE)

    def bullets(self, s, x, y, w, h, items, size=11, space_after=4,
               color=PETROL):
        paras = [[(i, False)] if isinstance(i, str) else i for i in items]
        return self.text(s, x, y, w, h, paras, size=size, bullet=True,
                         space_after=space_after, color=color)

    def icon_bullets(self, s, x, y, w, h, items, size=10.5, gap=0.30,
                     accent_cycle=True):
        """Lista con marcador circular de color en vez de viñeta "•".

        `items`: lista de (glyph_key_o_texto_solo) o de tuplas
        (glyph, texto) o (glyph, color, texto). Si un item es solo texto,
        se usa el glifo "check" y el acento rota sobre PALETTE (o queda fijo
        en TEAL si accent_cycle=False).
        """
        cy = y
        for i, it in enumerate(items):
            if isinstance(it, str):
                glyph, color, txt = "check", PALETTE[i % len(PALETTE)], it
            elif len(it) == 2:
                glyph, txt = it
                color = PALETTE[i % len(PALETTE)] if accent_cycle else TEAL
            else:
                glyph, color, txt = it
            bh = 0.28
            self.oval(s, x, cy + 0.01, bh, color)
            self.text(s, x, cy - 0.005, bh, bh, ICONS.get(glyph, glyph),
                      size=bh * 26, color=WHITE, align=PP_ALIGN.CENTER,
                      anchor=MSO_ANCHOR.MIDDLE)
            self.text(s, x + bh + 0.14, cy - 0.04, w - bh - 0.14, 0.5, txt,
                      size=size, color=PETROL, anchor=MSO_ANCHOR.TOP)
            # avanza según líneas aproximadas (heurística simple, 1 línea
            # ~0.16in a este tamaño de fuente; el llamador puede pasar menos
            # items o más `h` si el texto es largo y ocupa 2-3 líneas)
            approx_lines = max(1, len(txt) // 62 + 1)
            cy += max(gap, approx_lines * 0.16 + 0.14)
        return cy

    def columns(self, s, y, h, cols, gap=0.20, header_h=0.60, header_size=10.5,
                body_size=9.5, palette=None):
        """N tarjetas iguales con cabecera y viñetas.

        cols: [{"titulo": str, "items": [str|runs], "icon": clave_opcional}]
        A diferencia de ppt-mimp (que alterna NAVY/NAVY2), cada tarjeta toma
        un acento distinto de `palette` (por defecto PALETTE) para que el
        set se vea multicolor — es la firma visual del MIMP Lab. Si además
        cada tarjeta trata un tema distinto, pasa `"icon"` por tarjeta (clave
        de ICONS) para que el ícono de cabecera corresponda a SU contenido,
        no solo al de la lámina entera. Si prefieres que todas las
        cabeceras compartan un solo color, pasa `palette=[TEAL]`.
        """
        palette = palette or PALETTE
        n = len(cols)
        cw = (self.CW - (n - 1) * gap) / n
        xs = []
        for i, c in enumerate(cols):
            x = self.M + i * (cw + gap)
            xs.append(x)
            accent = palette[i % len(palette)]
            self.card(s, x, y, cw, h)
            self.rect(s, x, y, cw, 0.06, accent)
            self.header_bar(s, x, y + 0.06, cw, header_h, c["titulo"],
                            tint(accent, 0.85), header_size, color=PETROL,
                            icon=c.get("icon"))
            if c.get("items"):
                self.bullets(s, x + 0.16, y + header_h + 0.18, cw - 0.32,
                             h - header_h - 0.3, c["items"], size=body_size,
                             space_after=5)
        return xs, cw

    def steps(self, s, y, h, pasos, gap=0.16, header_h=0.55, size=9.5,
              palette=None):
        """Fila de N pasos numerados. `pasos`: [(titulo, texto)] o
        [(titulo, texto, icono)]. Cada paso usa un acento distinto de
        `palette` (por defecto PALETTE) en vez de alternar solo dos tonos de
        azul; si se da un tercer elemento (clave de ICONS), se dibuja en la
        cabecera del paso según lo que ESE paso trata."""
        palette = palette or PALETTE
        n = len(pasos)
        sw = (self.CW - (n - 1) * gap) / n
        for i, paso in enumerate(pasos):
            t, txt = paso[0], paso[1]
            icon = paso[2] if len(paso) > 2 else None
            x = self.M + i * (sw + gap)
            accent = palette[i % len(palette)]
            self.card(s, x, y, sw, h)
            self.header_bar(s, x, y, sw, header_h, t, accent, size - 0.5,
                            icon=icon)
            self.text(s, x + 0.12, y + header_h + 0.09, sw - 0.24,
                      h - header_h - 0.18, txt, size=size, align=PP_ALIGN.CENTER)

    def gantt(self, s, meses, ejes, y=2.32, header_h=0.36, x_grid=3.75,
              label_w=2.60, nota=None, palette=None):
        """Cronograma. Firma idéntica a `mimp_deck.gantt`, pero si una fila
        no trae `color` propio, se le asigna uno de `palette` (por defecto
        PALETTE) según su posición — así cada eje del cronograma se distingue
        por color en vez de repetir un solo azul."""
        palette = palette or PALETTE
        n = len(meses)
        colw = (self.M + self.CW - x_grid) / n
        col = lambda i: x_grid + i * colw

        self.header_bar(s, self.M, y, label_w, header_h, "Eje", PETROL, 11)
        for i, m in enumerate(meses):
            self.rect(s, col(i) + 0.03, y, colw - 0.06, header_h, PETROL, WHITE)
            self.text(s, col(i), y, colw, header_h, m, size=10, bold=True,
                      color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        alto = sum(e.get("h", 1.0) + 0.08 for e in ejes)
        for i in range(n):
            self.rect(s, col(i), y + header_h + 0.08, colw, alto,
                      RGBColor(0xF7, 0xF7, 0xF6) if i % 2 == 0 else BG,
                      RGBColor(0xE2, 0xE2, 0xE2))

        cy = y + header_h + 0.14
        for idx, e in enumerate(ejes):
            eh = e.get("h", 1.0)
            accent = e.get("color") or palette[idx % len(palette)]
            self.rect(s, self.M, cy, label_w, eh, accent, radius=0.05)
            self.text(s, self.M + 0.08, cy, label_w - 0.16, eh, e["titulo"],
                      size=10.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
                      anchor=MSO_ANCHOR.MIDDLE)
            filas = e.get("filas", [])
            if filas:
                bh = min(0.30, (eh - 0.06 * (len(filas) - 1)) / len(filas))
                by = cy + (eh - (len(filas) * bh + 0.06 * (len(filas) - 1))) / 2
                for k, f in enumerate(filas):
                    for tramo in (f if isinstance(f, list) else [f]):
                        x1, x2 = tramo["ini"], tramo["fin"]
                        self.rect(s, col(x1) + 0.04, by + k * (bh + 0.06),
                                  (x2 - x1) * colw - 0.08, bh,
                                  tramo.get("fill", tint(accent)), radius=0.03)
                        if tramo.get("texto"):
                            self.text(s, col(x1) + 0.04, by + k * (bh + 0.06),
                                      (x2 - x1) * colw - 0.08, bh, tramo["texto"],
                                      size=9, align=PP_ALIGN.CENTER,
                                      anchor=MSO_ANCHOR.MIDDLE)
            cy += eh + 0.08

        if nota:
            self.text(s, self.M, min(cy + 0.12, 6.55), self.CW, 0.5, nota,
                      size=8.5, color=GRAY_TXT, space_after=3)
        return col, colw

    def table(self, s, x, y, w, h, headers, rows, col_w, header_size=10,
              body_size=9, align_cols=None, row_h=None, header_fill=PETROL,
              stripe=None):
        """Tabla nativa. `stripe`, si se da, es una lista de colores (uno por
        columna) para pintar de acento pálido cada columna de dato en vez de
        alternar filas en gris — útil para tablas donde cada columna
        representa una categoría distinta (p. ej. una política o eje) y
        conviene que el color ayude a distinguirlas."""
        n_rows = len(rows) + 1
        n_cols = len(headers)
        gshape = s.shapes.add_table(n_rows, n_cols, Inches(x), Inches(y),
                                    Inches(w), Inches(h))
        tbl = gshape.table
        for i, cw in enumerate(col_w):
            tbl.columns[i].width = Inches(cw)
        rh = row_h or (h / n_rows)
        for r in range(n_rows):
            tbl.rows[r].height = Inches(rh)
        tbl.first_row = False
        tbl.horz_banding = False
        align_cols = align_cols or [PP_ALIGN.LEFT] + [PP_ALIGN.CENTER] * (n_cols - 1)

        def _cell(r, c, text, bold, color, fill, size, align):
            cell = tbl.cell(r, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = fill
            cell.margin_left = cell.margin_right = Inches(0.06)
            cell.margin_top = cell.margin_bottom = Inches(0.02)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = align
            run = p.add_run()
            run.text = text
            run.font.size = Pt(size)
            run.font.bold = bold
            run.font.name = FONT
            run.font.color.rgb = color

        for c, htxt in enumerate(headers):
            fill = header_fill if not stripe else stripe[c % len(stripe)]
            _cell(0, c, htxt, True, WHITE, fill, header_size, PP_ALIGN.CENTER)
        for r, row in enumerate(rows, start=1):
            for c, val in enumerate(row):
                if stripe:
                    fill = tint(stripe[c % len(stripe)], 0.88 if r % 2 else 0.78)
                else:
                    fill = WHITE if r % 2 else tint(PETROL, 0.92)
                _cell(r, c, val, False, PETROL, fill, body_size, align_cols[c])
        return gshape

    # -------------------------------------------------- edición de un pptx dado
    def keep_slides(self, indices):
        lst = self.prs.slides._sldIdLst
        ids = list(lst)
        keep = [ids[i] for i in indices]
        for e in ids:
            if e not in keep:
                self.prs.part.drop_rel(e.get(qn("r:id")))
            lst.remove(e)
        for e in keep:
            lst.append(e)

    def renumber(self):
        for n, slide in enumerate(self.prs.slides, start=1):
            for shp in slide.shapes:
                if (shp.has_text_frame and shp.top and shp.top > Inches(6.9)
                        and shp.left > Inches(11)
                        and shp.text_frame.text.strip().isdigit()):
                    for p in shp.text_frame.paragraphs:
                        for r in p.runs:
                            r.text = ""
                    shp.text_frame.paragraphs[0].runs[0].text = str(n)

    def save(self, path):
        self.renumber()
        self.prs.save(path)
        self._fix_content_types(path)
        return path

    @staticmethod
    def _fix_content_types(path):
        import zipfile, shutil, re, tempfile
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            ct = z.read("[Content_Types].xml").decode("utf8")
            if 'Extension="jpg"' in ct or not any(n.endswith(".jpg") for n in names):
                return
            items = [(i, z.read(i.filename)) for i in z.infolist()]
        tmp = tempfile.mktemp(suffix=".pptx")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
            for info, data in items:
                if info.filename == "[Content_Types].xml":
                    data = re.sub(r"(<Types[^>]*>)",
                                  r'\1<Default Extension="jpg" ContentType="image/jpeg"/>',
                                  data.decode("utf8"), count=1).encode("utf8")
                out.writestr(info, data)
        shutil.move(tmp, path)
