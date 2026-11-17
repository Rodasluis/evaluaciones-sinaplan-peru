"""Piezas de la portada: configuración, tablas flotantes ancladas a la página, logos, ajuste del título y cierre de la
sección de portada.

Los bloques de texto y la fila de logos son tablas flotantes ancladas a la página, y todo el texto es texto de Word. Las
medidas se dan en centímetros sobre una hoja A4 (21 × 29,7 cm); los tamaños de letra, en puntos.
"""
import logging
from pathlib import Path

from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm

registro = logging.getLogger("portada")
A4_CM = (21.0, 29.7)
OBLIGATORIOS = ("titulo", "institucion", "organo_elaborador", "lugar_fecha")
# único diseño vigente (ronda 35, pedido del usuario): el centrado, elegido el 2026-11-10 entre los modelos A a K
DISENO = ("G", "centrada")
LINEAS_MAX, PT_MIN = 4, 20


class DatosFaltantes(ValueError):
    """Faltan datos obligatorios en config/portada.yaml: hay que pedirlos, no inventarlos."""


def tw(cm):
    """Centímetros a veinteavos de punto (twips)."""
    return int(round(cm * 567))


# ---------------------------------------------------------------------------------------------------- configuración
def cargar_config(ruta):
    import yaml
    ruta = Path(ruta)
    with open(ruta, encoding="utf-8") as f:
        datos = yaml.safe_load(f) or {}
    cfg = datos.get("portada") or {}
    # las rutas de los logos son relativas a la raíz del repositorio, esté el archivo en config/ o en la evaluación
    cfg["_raiz"] = Path(__file__).resolve().parents[3]
    return cfg


def validar(cfg):
    """Detiene el proceso si falta un dato obligatorio (no se inventa ninguno)."""
    faltan = [k for k in OBLIGATORIOS if not str(cfg.get(k) or "").strip()]
    if faltan:
        raise DatosFaltantes("Faltan datos obligatorios de la portada en config/portada.yaml: " + ", ".join(faltan) +
                             ". Complételos antes de generar la portada.")
    for k in ("etiqueta", "subtitulo"):
        if not str(cfg.get(k) or "").strip():
            registro.info("El campo opcional «%s» está vacío; la portada lo omite.", k)
    return cfg


def lineas_organo(cfg):
    """El órgano elaborador se escribe en una línea por unidad (separadas por « · » en la configuración)."""
    return [t.strip() for t in str(cfg["organo_elaborador"]).split("·") if t.strip()]


# ---------------------------------------------------------------------------------------------------- ajuste del título
def partir_lineas(texto, ttf, pt, ancho_cm, holgura=0.97):
    """Reparte el texto en líneas como lo haría Word (corte entre palabras) con las métricas de la letra."""
    from PIL import ImageFont
    escala = 10
    fuente = ImageFont.truetype(ttf, size=int(round(pt * escala)))
    limite = ancho_cm / 2.54 * 72 * escala * holgura
    lineas, actual = [], ""
    for palabra in texto.split():
        prueba = f"{actual} {palabra}".strip()
        if actual and fuente.getlength(prueba) > limite:
            lineas.append(actual)
            actual = palabra
        else:
            actual = prueba
    if actual:
        lineas.append(actual)
    return lineas


def ajustar_titulo(texto, ttf, pt_inicial, ancho_cm, avisos):
    """Tamaño del título para que ocupe 4 líneas o menos: baja de 1 en 1 pt hasta 20 pt; si aún no cabe, avisa."""
    pt = pt_inicial
    while True:
        lineas = partir_lineas(texto, ttf, pt, ancho_cm)
        if len(lineas) <= LINEAS_MAX:
            if pt < pt_inicial:
                registro.info("Título reducido a %s pt para caber en %s líneas.", pt, LINEAS_MAX)
            return pt, lineas
        if pt - 1 < PT_MIN:
            aviso = (f"El título no cabe en {LINEAS_MAX} líneas ni con {PT_MIN} pt ({len(lineas)} líneas): "
                     "acórtelo en config/portada.yaml.")
            registro.warning(aviso)
            avisos.append(aviso)
            return pt, lineas
        pt -= 1


# ---------------------------------------------------------------------------------------------------- texto
def _rpr(fuente, pt=None, color=None, negrita=False, espaciado=None, mayusculas=False, cursiva=False):
    x = f'<w:rPr {nsdecls("w")}>'
    if fuente:
        x += f'<w:rFonts w:ascii="{fuente}" w:hAnsi="{fuente}" w:eastAsia="{fuente}" w:cs="{fuente}"/>'
    if negrita:
        x += "<w:b/><w:bCs/>"
    if cursiva:
        x += "<w:i/><w:iCs/>"
    if mayusculas:
        x += "<w:caps/>"           # mayúsculas de formato: el texto sigue escrito como en la configuración
    if color:
        x += f'<w:color w:val="{color}"/>'
    if espaciado:
        x += f'<w:spacing w:val="{espaciado}"/>'
    if pt:
        x += f'<w:sz w:val="{int(round(pt * 2))}"/><w:szCs w:val="{int(round(pt * 2))}"/>'
    return parse_xml(x + "</w:rPr>")


def parrafo(fuente, piezas=(), jc="left", antes=0, despues=0, interlineado=1.0, exacto_pt=None, pt_marca=None,
            borde_inferior=None, sangria_der_cm=0.0, relleno=None, borde_superior=None, sangria_izq_cm=0.0):
    """Párrafo con formato directo completo (no depende del estilo Normal del documento).
    piezas: lista de (texto, {"pt", "color", "b", "i", "espaciado", "caps"}); borde_inferior y borde_superior:
    (color, grosor en octavos de pt[, separación del texto en pt])."""
    p = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    p.append(ppr)
    if borde_inferior or borde_superior:
        lados = ""
        for lado, b in (("top", borde_superior), ("bottom", borde_inferior)):   # el esquema exige top antes que bottom
            if b:
                color, sz, sep = (tuple(b) + (0,))[:3]
                lados += f'<w:{lado} w:val="single" w:sz="{sz}" w:space="{int(sep)}" w:color="{color}"/>'
        ppr.append(parse_xml(f'<w:pBdr {nsdecls("w")}>{lados}</w:pBdr>'))
    if relleno:
        ppr.append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{relleno}"/>'))
    ppr.append(parse_xml(f'<w:suppressAutoHyphens {nsdecls("w")}/>'))
    if exacto_pt:
        linea = f'w:line="{int(round(exacto_pt * 20))}" w:lineRule="exact"'
    else:
        linea = f'w:line="{int(round(interlineado * 240))}" w:lineRule="auto"'
    ppr.append(parse_xml(f'<w:spacing {nsdecls("w")} w:before="{int(round(antes * 20))}" '
                         f'w:after="{int(round(despues * 20))}" {linea}/>'))
    ppr.append(parse_xml(f'<w:ind {nsdecls("w")} w:left="{tw(sangria_izq_cm)}" w:right="{tw(sangria_der_cm)}" '
                         f'w:firstLine="0"/>'))
    ppr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="{jc}"/>'))
    tam_marca = pt_marca or (piezas[0][1].get("pt") if piezas else 2)
    ppr.append(_rpr(fuente, tam_marca))
    for texto, f in piezas:
        r = OxmlElement("w:r")
        r.append(_rpr(f.get("fuente", fuente), f.get("pt"), f.get("color"), f.get("b", False), f.get("espaciado"),
                      f.get("caps"), f.get("i", False)))
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = texto
        r.append(t)
        p.append(r)
    return p


def separador(documento, fuente):
    """Párrafo mínimo entre tablas flotantes (Word une las tablas contiguas si no hay un párrafo entre ellas)."""
    p = parrafo(fuente, exacto_pt=1, pt_marca=1)
    _cuerpo_insertar(documento, p)
    return p


def _cuerpo_insertar(documento, el):
    cuerpo = documento.element.body
    cuerpo.find(qn("w:sectPr")).addprevious(el)


# ---------------------------------------------------------------------------------------------------- logos
def imagen(documento, ruta, alto_cm, alt):
    """Run con la imagen en línea, al alto pedido y con su proporción original (sin deformar ni recolorear)."""
    from PIL import Image
    with Image.open(ruta) as im:
        ancho_px, alto_px = im.size
    ancho_cm = alto_cm * ancho_px / alto_px
    inline = documento.part.new_pic_inline(str(ruta), Cm(ancho_cm), Cm(alto_cm))
    inline.docPr.set("descr", alt)
    inline.docPr.set("title", alt)
    r = OxmlElement("w:r")
    dibujo = OxmlElement("w:drawing")
    dibujo.append(inline)
    r.append(dibujo)
    return r, ancho_cm


def logos(documento, cfg, paleta, alto_cm, avisos):
    """Celdas de la fila de logos (institución y Gobierno del Perú) con el mismo alto. Si falta un archivo, va un
    recuadro punteado con el nombre esperado y se avisa."""
    grupo = cfg.get("logos") or {}
    piezas = []
    for clave, alt in (("institucion", f"Logo de {cfg['institucion']}"), ("gobierno", "Logo del Gobierno del Perú")):
        rel = str(grupo.get(clave) or "").strip()
        ruta = (cfg["_raiz"] / rel) if rel else None
        if ruta and ruta.is_file():
            r, ancho = imagen(documento, ruta, alto_cm, alt)
            piezas.append({"run": r, "ancho": ancho, "alt": alt})
        else:
            aviso = f"Falta el logo «{clave}» ({rel or 'sin ruta'}): la portada lleva un recuadro punteado en su lugar."
            registro.warning(aviso)
            avisos.append(aviso)
            piezas.append({"run": None, "ancho": alto_cm * 3.4, "alt": alt, "falta": rel or clave})
    return piezas


def fila_logos(documento, x_cm, y_cm, piezas, alto_cm, paleta, separacion_cm=1.2, ancho_total_cm=None):
    """Tabla flotante con los logos en una fila. Con ancho_total_cm, el primero queda en x y el último termina en
    x + ancho_total_cm (logos en los extremos); sin él, van juntos a la izquierda con una separación fija."""
    if ancho_total_cm:
        separacion_cm = max(0.5, ancho_total_cm - sum(pz["ancho"] for pz in piezas))
    columnas, celdas = [], []
    for k, pz in enumerate(piezas):
        if k:
            columnas.append(separacion_cm)
            celdas.append({"parrafos": [parrafo(paleta.fuente, exacto_pt=1, pt_marca=1)]})
        columnas.append(pz["ancho"])
        if pz["run"] is not None:
            p = parrafo(paleta.fuente, pt_marca=2)
            p.append(pz["run"])
            celdas.append({"parrafos": [p], "valign": "center"})
        else:
            celdas.append({"parrafos": [parrafo(paleta.fuente, [(f"Falta el logo: {pz['falta']}",
                                                                   {"pt": 8, "color": paleta.gris_oscuro})], jc="center")],
                           "valign": "center", "borde": ("dotted", 6, paleta.gris_oscuro)})
    return tabla(documento, x_cm, y_cm, columnas, [{"alto_min": alto_cm, "celdas": celdas}])


# ---------------------------------------------------------------------------------------------------- tablas flotantes
def tabla(documento, x_cm, y_cm, columnas_cm, filas, borde_superior=None):
    """Tabla flotante anclada a la página en (x, y), insertada al final del cuerpo (ver tabla_xml)."""
    tbl = tabla_xml(columnas_cm, filas, borde_superior, (x_cm, y_cm))
    _cuerpo_insertar(documento, tbl)
    return tbl


def tabla_xml(columnas_cm, filas, borde_superior=None, posicion=None):
    """Tabla sin bordes ni márgenes de celda por defecto. Con posicion (x, y) en cm, flotante y anclada a la página; sin
    ella, en línea (para anidarla dentro de una celda). Una celda puede llevar tablas anidadas
    entre sus párrafos; la celda debe terminar en un párrafo.
    filas: [{"alto": cm (exacto) o "alto_min": cm (al menos), "celdas": [{"parrafos": [w:p…], "relleno": hex,
    "valign": top|center|bottom, "margen": (sup, der, inf, izq) en cm, "direccion": "btLr",
    "borde": (estilo, grosor, color), "span": columnas que ocupa la celda}]}]
    borde_superior: (color, grosor en octavos de pt) para una línea fina sobre la tabla."""
    ancho = sum(columnas_cm)
    sup = (f'<w:top w:val="single" w:sz="{borde_superior[1]}" w:space="0" w:color="{borde_superior[0]}"/>'
           if borde_superior else '<w:top w:val="nil"/>')
    flota = ""
    if posicion:
        x_cm, y_cm = posicion
        flota = (f'<w:tblpPr w:leftFromText="0" w:rightFromText="0" w:topFromText="0" w:bottomFromText="0" '
                 f'w:vertAnchor="page" w:horzAnchor="page" w:tblpX="{tw(x_cm)}" w:tblpY="{tw(y_cm)}"/>'
                 f'<w:tblOverlap w:val="overlap"/>')
    tbl = parse_xml(
        f'<w:tbl {nsdecls("w")}><w:tblPr>{flota}<w:tblW w:w="{tw(ancho)}" w:type="dxa"/>'
        f'<w:tblBorders>{sup}<w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/>'
        f'<w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders><w:tblLayout w:type="fixed"/>'
        f'<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>'
        f'<w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>' +
        "".join(f'<w:gridCol w:w="{tw(c)}"/>' for c in columnas_cm) + "</w:tblGrid></w:tbl>")
    for fila in filas:
        tr = OxmlElement("w:tr")
        trpr = OxmlElement("w:trPr")
        trpr.append(OxmlElement("w:cantSplit"))
        if fila.get("alto"):
            trpr.append(parse_xml(f'<w:trHeight {nsdecls("w")} w:val="{tw(fila["alto"])}" w:hRule="exact"/>'))
        elif fila.get("alto_min"):
            trpr.append(parse_xml(f'<w:trHeight {nsdecls("w")} w:val="{tw(fila["alto_min"])}" w:hRule="atLeast"/>'))
        tr.append(trpr)
        k = 0
        for celda in fila["celdas"]:
            span = celda.get("span", 1)
            c = sum(columnas_cm[k:k + span])
            k += span
            tc = OxmlElement("w:tc")
            x = f'<w:tcPr {nsdecls("w")}><w:tcW w:w="{tw(c)}" w:type="dxa"/>'
            if span > 1:
                x += f'<w:gridSpan w:val="{span}"/>'
            if celda.get("borde"):
                est, sz, col = celda["borde"]
                x += "<w:tcBorders>" + "".join(f'<w:{l} w:val="{est}" w:sz="{sz}" w:space="0" w:color="{col}"/>'
                                               for l in ("top", "left", "bottom", "right")) + "</w:tcBorders>"
            if celda.get("relleno"):
                x += f'<w:shd w:val="clear" w:color="auto" w:fill="{celda["relleno"]}"/>'
            # Word mide la posición de una tabla flotante desde el texto de la primera celda: un margen izquierdo de
            # celda correría la forma fuera de su sitio. Los márgenes laterales van como sangrías de los párrafos y
            # la celda solo lleva los márgenes superior e inferior.
            s, d, i, z = celda.get("margen") or (0, 0, 0, 0)
            if s or i:
                x += (f'<w:tcMar><w:top w:w="{tw(s)}" w:type="dxa"/><w:bottom w:w="{tw(i)}" w:type="dxa"/>'
                      f'</w:tcMar>')
            if celda.get("direccion"):
                x += f'<w:textDirection w:val="{celda["direccion"]}"/>'
            x += f'<w:vAlign w:val="{celda.get("valign", "top")}"/></w:tcPr>'
            tc.append(parse_xml(x))
            for p in celda.get("parrafos") or [parrafo(None, exacto_pt=1, pt_marca=1)]:
                ind = p.find(qn("w:pPr") + "/" + qn("w:ind"))
                if ind is not None and (d or z):
                    ind.set(qn("w:left"), str(int(ind.get(qn("w:left"), "0")) + tw(z)))
                    ind.set(qn("w:right"), str(int(ind.get(qn("w:right"), "0")) + tw(d)))
                tc.append(p)
            tr.append(tc)
        tbl.append(tr)
    return tbl


# ---------------------------------------------------------------------------------------------------- sección
def cerrar_seccion_portada(documento, paleta, relleno=None, alto_pt=1):
    """Cierra la portada como sección propia: A4, márgenes 0 (las formas llegan a los bordes), sin encabezado, sin pie y
    sin número de página. La sección siguiente conserva su encabezado, su pie y su numeración. relleno y alto_pt: relleno
    y alto del párrafo de cierre."""
    import copy
    cuerpo = documento.element.body
    sect = copy.deepcopy(cuerpo.find(qn("w:sectPr")))
    for e in (sect.findall(qn("w:headerReference")) + sect.findall(qn("w:footerReference")) +
              sect.findall(qn("w:pgNumType")) + sect.findall(qn("w:titlePg"))):
        sect.remove(e)
    tam = sect.find(qn("w:pgSz"))
    if tam is None:
        tam = OxmlElement("w:pgSz")
        sect.insert(0, tam)
    tam.set(qn("w:w"), str(tw(A4_CM[0]))); tam.set(qn("w:h"), str(tw(A4_CM[1])))
    if qn("w:orient") in tam.attrib:
        del tam.attrib[qn("w:orient")]
    mar = sect.find(qn("w:pgMar"))
    if mar is None:
        mar = OxmlElement("w:pgMar")
        tam.addnext(mar)
    for lado in ("top", "right", "bottom", "left", "header", "footer", "gutter"):
        mar.set(qn(f"w:{lado}"), "0")
    tipo = sect.find(qn("w:type"))
    if tipo is not None:
        sect.remove(tipo)
    p = parrafo(paleta.fuente, exacto_pt=alto_pt, pt_marca=1, relleno=relleno)
    ppr = p.find(qn("w:pPr"))
    ppr.append(sect)                      # sectPr es el último hijo de w:pPr
    _cuerpo_insertar(documento, p)
    return p
