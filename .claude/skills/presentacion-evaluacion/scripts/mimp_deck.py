"""Copia de la skill ppt-mimp para la skill del repositorio presentacion-evaluacion, con la corrección de los effectLst duplicados (_shadow).


mimp_deck — plantilla institucional MIMP para python-pptx.

Uso mínimo:

    import sys; sys.path.insert(0, "<ruta-skill>/scripts")
    from mimp_deck import MimpDeck, NAVY, ORANGE

    d = MimpDeck()
    d.cover("MIMP Lab", "Laboratorio de Innovación del MIMP", fecha="Agosto de 2026")
    s = d.content("Título de la diapositiva")
    d.text(s, d.M, d.BODY_TOP, d.CW, 0.5, [("Objetivo. ", True, ORANGE), ("...", False)])
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
NAVY = RGBColor(0x00, 0x20, 0x60)        # títulos y cuerpo de texto
NAVY2 = RGBColor(0x1F, 0x4E, 0x9C)       # variante para alternar cabeceras
RED = RGBColor(0xE3, 0x04, 0x12)         # línea de título, acentos
ORANGE = RGBColor(0xED, 0x7D, 0x31)      # subtítulos y etiquetas
BLUE = RGBColor(0x44, 0x72, 0xC4)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BG = RGBColor(0xED, 0xED, 0xED)          # fondo de diapositivas de contenido
GRAY_TXT = RGBColor(0x5A, 0x5A, 0x5A)    # notas al pie
GRAY_PAGE = RGBColor(0x9A, 0x9A, 0x9A)
CARD_LINE = RGBColor(0xD8, 0xD8, 0xD8)
NAVY_PALE = RGBColor(0xE3, 0xE8, 0xF3)
NAVY_PALE2 = RGBColor(0xC6, 0xD0, 0xE6)
RED_PALE = RGBColor(0xFB, 0xDF, 0xE1)
ORANGE_PALE = RGBColor(0xFB, 0xE6, 0xD6)
BLUE_PALE = RGBColor(0xDD, 0xE3, 0xF7)
GRAY_PALE = RGBColor(0xDF, 0xDF, 0xDF)
FONT = "Poppins"

_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'

_AVISADOS = set()


def _aviso(texto):
    """Avisa una sola vez de un recurso que falta (versión pública sin la identidad institucional)."""
    if texto not in _AVISADOS:
        _AVISADOS.add(texto)
        print("aviso:", texto)


class MimpDeck:
    """Constructor de presentaciones con el formato institucional del MIMP."""

    # ---- geometría fija (no cambiarla entre diapositivas: es lo que da unidad)
    M = 0.92          # margen izquierdo
    CW = 11.49        # ancho útil
    TITLE_Y = 1.02
    TITLE_H = 0.86
    RULE_Y = 1.96
    BODY_TOP = 2.15   # primer elemento del cuerpo
    BODY_BOT = 6.90   # último píxel utilizable del cuerpo
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
        """Rectángulo (o redondeado si se pasa `radius` en pulgadas)."""
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

    def text(self, s, x, y, w, h, runs, size=11, color=NAVY, bold=False,
             italic=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
             space_after=0, bullet=False, line_spacing=None):
        """Cuadro de texto.

        `runs` acepta:
          - "texto plano"
          - [("negrita", True), ("normal", False)]              -> un párrafo
          - [[("a", True, ORANGE), ("b", False)], [("c", False)]] -> varios
        Los saltos "\\n" se convierten en párrafos reales (si se dejan como
        salto de línea suelto, PowerPoint desalinea el texto centrado).
        """
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
            self.rect(s, 0, 0, self.W, self.H, {"bg_red.jpg": RED, "bg_blue.jpg": NAVY}.get(bg, BG))
        self.rect(s, -0.3, -0.5, self.W + 0.6, 1.41, WHITE, radius=0.18)
        for logo, caja in (("logo_mimp.png", (0.83, 0.19, 2.32, 0.566)), ("logo_gob.png", (10.423, 0.15, 2.08, 0.654))):
            ruta = os.path.join(self.assets, logo)
            if os.path.exists(ruta):
                s.shapes.add_picture(ruta, *(Inches(v) for v in caja))
            else:
                _aviso(f"Falta {logo} en {self.assets}: la banda va sin ese logo.")

    def _page_number(self, s):
        self.text(s, 11.7, 7.02, 0.71, 0.28, str(len(self.prs.slides._sldIdLst)),
                  size=10, color=GRAY_PAGE, align=PP_ALIGN.RIGHT)

    # ------------------------------------------------------ tipos de diapositiva
    def content(self, title, rule=True):
        """Diapositiva de contenido: fondo gris, banda, título y línea roja."""
        s = self.prs.slides.add_slide(self.layout)
        self.band(s)
        self.text(s, self.M, self.TITLE_Y, self.CW, self.TITLE_H, title,
                  size=28, bold=True, color=NAVY, anchor=MSO_ANCHOR.BOTTOM)
        if rule:
            self.rect(s, self.M, self.RULE_Y, 3.4, 0.035, RED)
        self._page_number(s)
        return s

    def cover(self, title, subtitle=None, unidad=None, fecha=None):
        s = self.prs.slides.add_slide(self.layout)
        self.band(s, bg="bg_red.jpg")
        self.text(s, 1.5, 2.55, 10.33, 1.1, title, size=54, bold=True,
                  color=WHITE, align=PP_ALIGN.CENTER)
        if subtitle:
            self.text(s, 1.5, 3.65, 10.33, 0.6, subtitle, size=22, color=WHITE,
                      align=PP_ALIGN.CENTER)
        self.rect(s, 5.17, 4.45, 3.0, 0.02, WHITE)
        if unidad:
            self.text(s, 1.5, 4.75, 10.33, 0.4, unidad, size=13, color=WHITE,
                      align=PP_ALIGN.CENTER)
        if fecha:
            self.text(s, 1.5, 5.9, 10.33, 0.45, fecha, size=18, color=WHITE,
                      align=PP_ALIGN.CENTER)
        return s

    def divider(self, numero, titulo):
        """Divisor de sección (fondo azul, alineado a la derecha)."""
        s = self.prs.slides.add_slide(self.layout)
        self.band(s, bg="bg_blue.jpg")
        self.text(s, 4.0, 3.6, 8.41, 0.9, numero, size=46, bold=True,
                  color=RGBColor(0x9F, 0xC3, 0xE6), align=PP_ALIGN.RIGHT)
        self.text(s, 3.0, 4.5, 9.41, 1.1, titulo, size=30, bold=True,
                  color=WHITE, align=PP_ALIGN.RIGHT)
        self.rect(s, 9.41, 5.72, 3.0, 0.02, WHITE)
        self._page_number(s)
        return s

    def closing(self, titulo="¡Gracias!", pie=None):
        s = self.prs.slides.add_slide(self.layout)
        self.band(s, bg="bg_red.jpg")
        self.text(s, 1.5, 3.1, 10.33, 1.1, titulo, size=48, bold=True,
                  color=WHITE, align=PP_ALIGN.CENTER)
        self.rect(s, 5.17, 4.3, 3.0, 0.02, WHITE)
        if pie:
            self.text(s, 1.5, 4.6, 10.33, 0.45, pie, size=15, color=WHITE,
                      align=PP_ALIGN.CENTER)
        return s

    # --------------------------------------------------- componentes de cuerpo
    def card(self, s, x, y, w, h, fill=WHITE):
        return self.rect(s, x, y, w, h, fill, CARD_LINE, radius=0.05, shadow=True)

    def header_bar(self, s, x, y, w, h, texto, fill=NAVY, size=11):
        """Barra de encabezado (para tarjetas, secciones o tablas)."""
        self.rect(s, x, y, w, h, fill, radius=0.05)
        self.text(s, x + 0.12, y, w - 0.24, h, texto, size=size, bold=True,
                  color=WHITE, anchor=MSO_ANCHOR.MIDDLE)

    def subhead(self, s, x, y, w, texto, size=13):
        """Subtítulo naranja dentro del cuerpo."""
        self.text(s, x, y, w, 0.3, texto, size=size, bold=True, color=ORANGE,
                  anchor=MSO_ANCHOR.MIDDLE)

    def bullets(self, s, x, y, w, h, items, size=11, space_after=4):
        """`items`: lista de str o de listas de runs."""
        paras = [[(i, False)] if isinstance(i, str) else i for i in items]
        return self.text(s, x, y, w, h, paras, size=size, bullet=True,
                         space_after=space_after)

    def columns(self, s, y, h, cols, gap=0.20, header_h=0.60, header_size=10.5,
                body_size=9.5):
        """N tarjetas iguales con cabecera y viñetas.

        cols: [{"titulo": str, "items": [str|runs]}]
        Devuelve la lista de coordenadas x por si hay que añadir más cosas.
        """
        n = len(cols)
        cw = (self.CW - (n - 1) * gap) / n
        xs = []
        for i, c in enumerate(cols):
            x = self.M + i * (cw + gap)
            xs.append(x)
            self.card(s, x, y, cw, h)
            self.header_bar(s, x, y, cw, header_h, c["titulo"], NAVY, header_size)
            if c.get("items"):
                self.bullets(s, x + 0.16, y + header_h + 0.12, cw - 0.32,
                             h - header_h - 0.24, c["items"], size=body_size,
                             space_after=5)
        return xs, cw

    def steps(self, s, y, h, pasos, gap=0.16, header_h=0.55, size=9.5):
        """Fila de N pasos numerados (proceso). pasos: [(titulo, texto)]"""
        n = len(pasos)
        sw = (self.CW - (n - 1) * gap) / n
        for i, (t, d) in enumerate(pasos):
            x = self.M + i * (sw + gap)
            self.card(s, x, y, sw, h)
            self.header_bar(s, x, y, sw, header_h, t, NAVY if i % 2 == 0 else NAVY2,
                            size - 0.5)
            self.text(s, x + 0.12, y + header_h + 0.09, sw - 0.24,
                      h - header_h - 0.18, d, size=size, align=PP_ALIGN.CENTER)

    def gantt(self, s, meses, ejes, y=2.32, header_h=0.36, x_grid=3.75,
              label_w=2.60, nota=None):
        """Cronograma.

        ejes: [{"titulo": "Innovación y\\nEvaluación", "color": NAVY,
                "filas": [{"ini": 0, "fin": 6, "texto": "...", "fill": NAVY_PALE}],
                "h": 1.12}]
        `ini`/`fin` son índices de columna (fin excluyente).
        """
        n = len(meses)
        colw = (self.M + self.CW - x_grid) / n
        col = lambda i: x_grid + i * colw

        self.header_bar(s, self.M, y, label_w, header_h, "Eje", NAVY, 11)
        for i, m in enumerate(meses):
            self.rect(s, col(i) + 0.03, y, colw - 0.06, header_h, NAVY, WHITE)
            self.text(s, col(i), y, colw, header_h, m, size=10, bold=True,
                      color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        alto = sum(e.get("h", 1.0) + 0.08 for e in ejes)
        for i in range(n):
            self.rect(s, col(i), y + header_h + 0.08, colw, alto,
                      RGBColor(0xF6, 0xF6, 0xF6) if i % 2 == 0 else BG,
                      RGBColor(0xE2, 0xE2, 0xE2))

        cy = y + header_h + 0.14
        for e in ejes:
            eh = e.get("h", 1.0)
            self.rect(s, self.M, cy, label_w, eh, e.get("color", NAVY), radius=0.05)
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
                                  tramo.get("fill", NAVY_PALE), radius=0.03)
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

    # -------------------------------------------------- edición de un pptx dado
    def keep_slides(self, indices):
        """Conserva y reordena las diapositivas indicadas (índices base 0)."""
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
        """Reescribe el número de página de todas las diapositivas."""
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
        """python-pptx no declara .jpg si el pptx base no lo traía."""
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
