"""Componentes de evaluación para las presentaciones, en los dos estilos de la skill (institucional con mimp_deck y
MIMP Lab con mimp_lab_deck). Salen de la presentación de resultados del piloto PNIG
(src/analisis_pnig/presentacion_resultados.py), que los usa.

Uso:
    import sys; sys.path.insert(0, ".claude/skills/presentacion-evaluacion/scripts")
    from mimp_lab_deck import MimpLabDeck, GOLD, TEAL
    import componentes as K
    d = MimpLabDeck(); s = d.content("Resultados", icon="datos", accent=GOLD)
    K.tarjeta_cifra(d, s, d.M, 2.3, 3.6, 1.75, "24 de 52", "servicios implementados al cierre", GOLD)

Componentes:
    bajada             oración de lectura bajo el título: qué muestra la lámina y qué concluye
    tarjeta_cifra      cifra grande con su etiqueta y un filete del color del tema
    caja               tarjeta con encabezado de color pálido y viñetas
    resumen_confianza  «3 (2 sólidos, 1 indicio)» a partir de las confianzas de los hallazgos de una pregunta
    tabla_valoracion   pregunta, tema, criterio, valoración y hallazgos con su confianza, con la nota de confianza
    tabla_estado       servicios por objetivo prioritario: implementados, parciales, no implementados y protocolo

Las tablas usan la tabla nativa del estilo Lab; en el estilo institucional, una tabla nativa equivalente.
"""
from collections import Counter

from mimp_deck import NAVY
from mimp_lab_deck import PETROL, MimpLabDeck, tint  # tint: versión pálida de un color, para los dos estilos
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

CONFIANZA = (("Sólido", "sólido", "sólidos"), ("Con respaldo parcial", "con respaldo parcial", "con respaldo parcial"),
             ("Indicio", "indicio", "indicios"))
# regla de confianza por origen independiente (nucleo.hallazgos.fuentes_independientes; ADR-015 del piloto)
NOTA_CONFIANZA = ("Confianza: sólido, coinciden dos o más fuentes de origen independiente; con respaldo parcial, "
                  "coinciden con discrepancias explicadas; indicio, un solo origen.")


def _lab(d):
    return isinstance(d, MimpLabDeck)


def tinta(d):
    """Color del texto del estilo: petróleo en el Lab, azul institucional en el otro."""
    return PETROL if _lab(d) else NAVY


def bajada(d, s, texto, y=2.06, h=0.4, size=12.5):
    """Oración de lectura bajo la línea del título: dice qué muestra la lámina y qué concluye, para que la presentación
    se entienda sin expositor. Una línea (unos 130 caracteres). Devuelve el tope del cuerpo debajo de ella."""
    d.text(s, d.M, y, d.CW, h, texto, size=size, color=tinta(d), italic=True)
    return round(y + h + 0.06, 2)


def tarjeta_cifra(d, s, x, y, w, h, numero, etiqueta, color):
    """Tarjeta con una cifra grande y su etiqueta."""
    d.card(s, x, y, w, h)
    d.rect(s, x, y, 0.07, h, color)
    d.text(s, x + 0.25, y + 0.12, w - 0.4, 0.75, numero, size=30, bold=True, color=color)
    d.text(s, x + 0.25, y + 0.9, w - 0.4, h - 1.0, etiqueta, size=12.5, color=tinta(d))


def caja(d, s, x, y, w, h, titulo, items, color, size=13):
    """Tarjeta con encabezado de color y viñetas (los textos ya con sus comillas “ ”)."""
    d.card(s, x, y, w, h)
    if _lab(d):
        d.header_bar(s, x, y, w, 0.55, titulo, tint(color, 0.80), 13, color=PETROL)
    else:
        d.header_bar(s, x, y, w, 0.55, titulo, color, 13)
    d.bullets(s, x + 0.2, y + 0.75, w - 0.4, h - 0.9, items, size=size, space_after=8)


def resumen_confianza(confianzas):
    """«n (k sólidos, m indicios)» para la lista de confianzas de los hallazgos de una pregunta."""
    cuenta = Counter(confianzas)
    partes = [f"{cuenta[k]} {uno if cuenta[k] == 1 else varios}" for k, uno, varios in CONFIANZA if cuenta.get(k)]
    return f"{len(confianzas)} ({', '.join(partes)})" if partes else str(len(confianzas))


def _tabla(d, s, x, y, w, h, encabezado, filas, anchos, body_size, header_size, alinear=None):
    if _lab(d):
        return d.table(s, x, y, w, h, encabezado, filas, anchos, body_size=body_size, header_size=header_size,
                       align_cols=alinear)
    forma = s.shapes.add_table(len(filas) + 1, len(encabezado), Inches(x), Inches(y), Inches(w), Inches(h))
    t = forma.table
    for j, a in enumerate(anchos):
        t.columns[j].width = Inches(a)
    for i, fila in enumerate([encabezado] + filas):
        for j, valor in enumerate(fila):
            celda = t.cell(i, j)
            celda.text = str(valor)
            celda.margin_left = celda.margin_right = Inches(0.05)
            celda.fill.solid()
            celda.fill.fore_color.rgb = NAVY if i == 0 else (RGBColor(0xFF, 0xFF, 0xFF) if i % 2 else
                                                              RGBColor(0xF2, 0xF2, 0xF2))
            for p in celda.text_frame.paragraphs:
                p.alignment = (alinear or [PP_ALIGN.LEFT] * len(encabezado))[j]
                for r in p.runs:
                    r.font.size = Pt(header_size if i == 0 else body_size)
                    r.font.bold = i == 0
                    r.font.name = "Poppins"
                    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF) if i == 0 else NAVY
    return forma


def tabla_valoracion(d, s, filas, y=2.3, h=4.45, nota=True):
    """filas: [pregunta, tema, criterio, valoración, hallazgos y confianza (resumen_confianza)]."""
    _tabla(d, s, d.M, y, d.CW, h, ["Pregunta", "Tema", "Criterio", "Valoración", "Hallazgos y confianza"], filas,
           [1.0, 3.4, 2.6, 1.5, 2.99], body_size=9.5, header_size=10,
           alinear=[PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.LEFT])
    if nota:
        d.text(s, d.M, 6.82, d.CW, 0.3, NOTA_CONFIANZA, size=8.5, color=tinta(d))


def tabla_estado(d, s, filas, x=None, y=2.3, w=6.6, h=3.6):
    """filas: [objetivo, servicios, implementados, parciales, no implementados, con protocolo], con la fila Total al
    final si corresponde."""
    _tabla(d, s, d.M if x is None else x, y, w, h,
           ["OP", "Servicios", "Implementados", "Parciales", "No implem.", "Con protocolo"], filas,
           [0.8, 1.0, 1.35, 1.1, 1.1, 1.25], body_size=12, header_size=10.5)
