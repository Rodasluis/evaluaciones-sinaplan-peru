"""Documentos Word con el formato institucional (núcleo; antes src/analisis_pnig/docx_mimp.py).

Toma como plantilla plantillas/word/informe_institucional.docx (encabezado con el logo del MIMP y el nombre de la OGMEPGD, pie centrado con número
de página, estilos de título, viñetas y tablas con cabecera azul), vacía su cuerpo y quita los comentarios, las
personas que comentaron, las notas al pie y los datos XML personalizados, para que nada de la nota pase al documento
nuevo; las notas al pie del documento nuevo se agregan con nota_pie(). Lo usan entrev_guias_entrevista.py e informe_cierre.py.

Con Documento(..., informe=True) (informe de cierre; ADR-010 y ronda 26) el formato va en estilos y no en formato
directo, para que el usuario pueda editarlo en Word: Normal, títulos 1 a 3 (el 1 empieza página), descripción (leyendas
de tablas y figuras con campo SEQ que conserva el formato al actualizarse), «Fuente MIMP», «Texto de tabla», índices
(TDC 1, TDC 2, tabla de ilustraciones y título de índice) y texto de nota al pie. Agrega además secciones (portada sin
encabezado, preliminares en romanos y cuerpo desde 1), recuadros (pregunta, respuesta y hallazgo) y una letra mínima de
7,5 pt en las tablas; la portada la agrega el módulo src/nucleo/portada (ronda 27). Sin informe=True, el formato es el de las
guías de entrevista, que no cambia.

Uso:
    from nucleo.documento import Documento
    d = Documento(pie="Guía de entrevista ENTREV · Evaluación de la PNIG · OMEP–OGMEPGD")
    d.titulo("GUÍA DE ENTREVISTA", "Evaluación de cierre de la PNIG (2019–2025)")
    d.h1("1. Propósito"); d.p("Texto con ", ("negrita", {"b": True}), ".")
    d.tabla(["Col 1", "Col 2"], [["a", "b"]], [4513, 4513], leyenda="Tabla 1. …", fuente="Fuente: …")
    d.figura("grafico.png", "Título", fuente="Fuente: …", ancho_cm=16)
    par = d.p("Texto con nota al pie."); d.nota_pie(par, "Texto de la nota.")
    d.guardar("salida.docx")

Las piezas de texto (runs) son cadenas o tuplas (texto, formato), con formato en {"b", "i", "hl", "color", "size",
"url", "caps"}; con "url", la pieza es un hipervínculo externo.
"""
import copy
import re

import docx
from docx.enum.text import WD_BREAK
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm
from lxml import etree

from .config import RAIZ_REPO
from .formato import comillas_parrafo

# plantilla institucional: estilos, encabezado con logos y pie; el cuerpo se vacía al cargarla
PLANTILLA = str(RAIZ_REPO / "plantillas/word/informe_institucional.docx")
AZUL, NEGRO, GRIS, ENLACE = "1F3864", "000000", "595959", "0563C1"
FILA_GRUPO = "D9E2F3"
REL_FUERA = ("comments", "commentsExtended", "commentsIds", "commentsExtensible", "people", "customXml")
PENDIENTE = re.compile(r"(\[(?:PENDIENTE|POR VERIFICAR|POR COMPLETAR)[^\]]*\])")
ANCHO_TEXTO_CM = 15.9           # A4 con márgenes de 2,54 cm
TAMANO_MINIMO_TABLA = 15        # 7,5 pt en el modo informe
# orden de los hijos de w:pPr (CT_PPrBase), para insertar propiedades sin romper el esquema
ORDEN_PPR = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
             "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
             "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid",
             "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
             "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"]
# recuadros del modo informe: relleno y borde izquierdo (color, grosor en octavos de punto)
RECUADROS = {"pregunta": ("DEEAF6", None), "respuesta": ("F2F2F2", (AZUL, 18)), "hallazgo": (None, (AZUL, 12))}


def _en_orden(ppr, el):
    """Inserta el elemento el en w:pPr respetando el orden del esquema (reemplaza uno igual si existe)."""
    nombre = el.tag.split("}")[1]
    previo = ppr.find(qn(f"w:{nombre}"))
    if previo is not None:
        ppr.replace(previo, el)
        return el
    k = ORDEN_PPR.index(nombre)
    for hijo in ppr:
        n = hijo.tag.split("}")[1]
        if n in ORDEN_PPR and ORDEN_PPR.index(n) > k:
            hijo.addprevious(el)
            return el
    ppr.append(el)
    return el


def _rpr(run, fmt, size_default):
    """Formato del tramo. Con size_default None (modo informe), solo se escribe lo que fmt pide: el resto viene del
    estilo del párrafo."""
    rpr = run._r.get_or_add_rPr()
    if fmt.get("b"):
        rpr.append(OxmlElement("w:b")); rpr.append(OxmlElement("w:bCs"))
    if fmt.get("i"):
        rpr.append(OxmlElement("w:i")); rpr.append(OxmlElement("w:iCs"))
    if fmt.get("caps"):
        rpr.append(OxmlElement("w:smallCaps"))
    if fmt.get("espaciado"):
        e = OxmlElement("w:spacing"); e.set(qn("w:val"), str(fmt["espaciado"])); rpr.append(e)
    if "color" in fmt or size_default is not None:
        col = OxmlElement("w:color"); col.set(qn("w:val"), fmt.get("color", NEGRO)); rpr.append(col)
    if "size" in fmt or size_default is not None:
        tam = str(fmt.get("size", size_default))
        for etq in ("w:sz", "w:szCs"):
            e = OxmlElement(etq); e.set(qn("w:val"), tam); rpr.append(e)
    if fmt.get("hl"):
        h = OxmlElement("w:highlight"); h.set(qn("w:val"), fmt["hl"]); rpr.append(h)
    # orden del esquema: b, bCs, i, iCs, smallCaps, spacing, color, sz, szCs, highlight, u
    return run


def _pieza(par, texto, fmt, size_default):
    """Agrega una pieza al párrafo; con fmt["url"], como hipervínculo externo en azul y subrayado."""
    run = _rpr(par.add_run(texto), {**fmt, "color": fmt.get("color", ENLACE)} if fmt.get("url") else fmt,
               size_default)
    if fmt.get("url"):
        u = OxmlElement("w:u"); u.set(qn("w:val"), "single"); run._r.get_or_add_rPr().append(u)
        h = OxmlElement("w:hyperlink")
        h.set(qn("r:id"), par.part.relate_to(fmt["url"], RT.HYPERLINK, is_external=True))
        h.append(run._r)  # lxml mueve el run del párrafo al hipervínculo
        par._p.append(h)
    return run


def _ppr(par, jc=None, before=None, after=None, line=None, keep_next=False, num=None, size=None, ind=None):
    ppr = par._p.get_or_add_pPr()
    if num is not None:
        np_ = OxmlElement("w:numPr")
        il = OxmlElement("w:ilvl"); il.set(qn("w:val"), "0"); np_.append(il)
        ni = OxmlElement("w:numId"); ni.set(qn("w:val"), str(num)); np_.append(ni)
        _en_orden(ppr, np_)
    if keep_next:
        _en_orden(ppr, OxmlElement("w:keepNext"))
    if any(v is not None for v in (before, after, line)):
        sp = OxmlElement("w:spacing")
        if before is not None:
            sp.set(qn("w:before"), str(before))
        if after is not None:
            sp.set(qn("w:after"), str(after))
        if line is not None:
            sp.set(qn("w:line"), str(line)); sp.set(qn("w:lineRule"), "auto")
        _en_orden(ppr, sp)
    if ind is not None:
        e = OxmlElement("w:ind"); e.set(qn("w:left"), str(ind)); _en_orden(ppr, e)
    if jc:
        e = OxmlElement("w:jc"); e.set(qn("w:val"), jc); _en_orden(ppr, e)
    return par


def _partir(piezas):
    """Normaliza las piezas y resalta en amarillo los marcadores [PENDIENTE …] y [POR VERIFICAR]."""
    salida = []
    for pz in piezas:
        texto, fmt = (pz, {}) if isinstance(pz, str) else pz
        for trozo in PENDIENTE.split(texto):
            if not trozo:
                continue
            f = dict(fmt)
            if PENDIENTE.fullmatch(trozo):
                f["hl"] = "yellow"
            salida.append((trozo, f))
    return salida


def _campo(par, instr, resultado, fmt=None, size=None):
    """Campo complejo de Word (begin, instr, separate, resultado, end); todos los tramos con el mismo formato."""
    fmt = fmt or {}
    for tipo, texto in (("begin", None), ("instr", instr), ("separate", None), ("res", resultado), ("end", None)):
        run = _rpr(par.add_run(), fmt, size)
        if tipo == "instr":
            it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = f" {texto} "
            run._r.append(it)
        elif tipo == "res":
            t = OxmlElement("w:t"); t.set(qn("xml:space"), "preserve"); t.text = texto; run._r.append(t)
        else:
            fc = OxmlElement("w:fldChar"); fc.set(qn("w:fldCharType"), tipo); run._r.append(fc)


def _estilo_xml(tipo, sid, nombre, ppr="", rpr="", basado="Normal", extra=""):
    return (f'<w:style {nsdecls("w")} w:type="{tipo}" w:styleId="{sid}"{extra}><w:name w:val="{nombre}"/>'
            f'{f"<w:basedOn w:val={chr(34)}{basado}{chr(34)}/>" if basado else ""}<w:uiPriority w:val="39"/>'
            f'<w:qFormat/><w:pPr>{ppr}</w:pPr><w:rPr>{rpr}</w:rPr></w:style>')


class Documento:
    def __init__(self, pie, titulo_doc="", autor="OMEP–OGMEPGD, MIMP", informe=False):
        self.d = docx.Document(PLANTILLA)
        self.informe = informe
        cuerpo = self.d.element.body
        for hijo in list(cuerpo):
            if hijo.tag != qn("w:sectPr"):
                cuerpo.remove(hijo)
        for rid, rel in list(self.d.part.rels.items()):
            if rel.reltype.split("/")[-1] in REL_FUERA:
                self.d.part.drop_rel(rid)
        self._notas_pie = None   # (parte, raíz) de footnotes, para nota_pie()
        self.figuras = []         # número, leyenda y archivo de cada figura (índice del anexo técnico)
        for nombre in ("footnotes", "endnotes"):
            for rel in self.d.part.rels.values():
                if rel.reltype.endswith("/" + nombre):
                    raiz = rel.target_part.element if hasattr(rel.target_part, "element") else None
                    if raiz is None:
                        raiz = etree.fromstring(rel.target_part.blob)
                        self._limpiar_notas(raiz)
                        rel.target_part._blob = etree.tostring(raiz, xml_declaration=True, encoding="UTF-8",
                                                               standalone=True)
                    else:
                        self._limpiar_notas(raiz)
                    if nombre == "footnotes":
                        self._notas_pie = (rel.target_part, raiz)
        pie_p = self.d.sections[0].footer.paragraphs[0]
        textos = [r for r in pie_p.runs if r.text and not r.text.strip().isdigit()]
        textos[0].text = f"{pie}  |  "
        cp = self.d.core_properties
        cp.title, cp.author, cp.last_modified_by = titulo_doc, autor, autor
        cp.comments = cp.subject = cp.keywords = cp.category = ""
        self.n_tabla = 0
        self.n_figura = 0
        # rastro técnico (libro y hoja) de las tablas y figuras: el texto lleva fuentes legibles (ADR-009)
        self.rastros = []
        self._seccion = ""
        self.leyendas = []   # (rótulo, número, texto) de tablas y figuras, para sus índices
        self.titulos = []    # (texto, nivel) de los títulos, para el índice general
        if informe:
            self._estilos_informe()

    # ------------------------------------------------------------------ estilos del modo informe
    def _estilos_informe(self):
        """Define los estilos del informe en styles.xml; los existentes de la plantilla se reemplazan."""
        st = self.d.styles.element
        tab = '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9016"/></w:tabs>'
        defs = [
            ("paragraph", "Normal", "Normal", '<w:spacing w:after="120" w:line="276" w:lineRule="auto"/><w:jc w:val="both"/>',
             '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/>',
             None, ' w:default="1"'),
            ("paragraph", "Ttulo1", "heading 1", '<w:keepNext/><w:keepLines/><w:pageBreakBefore/><w:spacing w:before="0" '
             'w:after="240"/><w:jc w:val="left"/><w:outlineLvl w:val="0"/>',
             f'<w:b/><w:bCs/><w:color w:val="{AZUL}"/><w:sz w:val="28"/><w:szCs w:val="28"/>', "Normal", ""),
            ("paragraph", "Ttulo2", "heading 2", '<w:keepNext/><w:keepLines/><w:spacing w:before="280" w:after="120"/>'
             '<w:jc w:val="left"/><w:outlineLvl w:val="1"/>',
             f'<w:b/><w:bCs/><w:color w:val="{AZUL}"/><w:sz w:val="24"/><w:szCs w:val="24"/>', "Normal", ""),
            ("paragraph", "Ttulo3", "heading 3", '<w:keepNext/><w:keepLines/><w:spacing w:before="200" w:after="80"/>'
             '<w:jc w:val="left"/><w:outlineLvl w:val="2"/>',
             f'<w:b/><w:bCs/><w:i/><w:iCs/><w:color w:val="{AZUL}"/><w:sz w:val="22"/><w:szCs w:val="22"/>', "Normal", ""),
            ("paragraph", "Descripcin", "caption", '<w:keepNext/><w:spacing w:before="180" w:after="60" w:line="240" '
             'w:lineRule="auto"/><w:jc w:val="left"/>',
             f'<w:b/><w:bCs/><w:color w:val="{AZUL}"/><w:sz w:val="20"/><w:szCs w:val="20"/>', "Normal", ""),
            ("paragraph", "FuenteMIMP", "Fuente MIMP", '<w:spacing w:before="40" w:after="160" w:line="240" '
             'w:lineRule="auto"/><w:jc w:val="both"/>',
             f'<w:color w:val="{GRIS}"/><w:sz w:val="16"/><w:szCs w:val="16"/>', "Normal", ""),
            ("paragraph", "TextodetablaMIMP", "Texto de tabla MIMP", '<w:spacing w:before="0" w:after="40" w:line="240" '
             'w:lineRule="auto"/><w:jc w:val="left"/>', '<w:sz w:val="16"/><w:szCs w:val="16"/>', "Normal", ""),
            ("paragraph", "TDC1", "toc 1", f'{tab}<w:spacing w:before="100" w:after="40" w:line="240" w:lineRule="auto"/>'
             '<w:ind w:right="567"/><w:jc w:val="left"/>', f'<w:b/><w:bCs/><w:color w:val="{AZUL}"/><w:sz w:val="20"/>'
             '<w:szCs w:val="20"/>', "Normal", ""),
            ("paragraph", "TDC2", "toc 2", f'{tab}<w:spacing w:before="0" w:after="30" w:line="240" w:lineRule="auto"/>'
             '<w:ind w:left="284" w:right="567"/><w:jc w:val="left"/>', '<w:sz w:val="20"/><w:szCs w:val="20"/>', "Normal",
             ""),
            ("paragraph", "Tabladeilustraciones", "table of figures", f'{tab}<w:spacing w:before="0" w:after="40" '
             'w:line="240" w:lineRule="auto"/><w:ind w:right="567"/><w:jc w:val="left"/>',
             '<w:sz w:val="20"/><w:szCs w:val="20"/>', "Normal", ""),
            ("paragraph", "TtuloTDC", "TOC Heading", '<w:keepNext/><w:spacing w:before="0" w:after="240"/><w:jc w:val="left"/>',
             f'<w:b/><w:bCs/><w:color w:val="{AZUL}"/><w:sz w:val="28"/><w:szCs w:val="28"/>', "Normal", ""),
            ("paragraph", "Textonotapie", "footnote text", '<w:spacing w:after="40" w:line="240" w:lineRule="auto"/>'
             '<w:jc w:val="both"/>', '<w:sz w:val="16"/><w:szCs w:val="16"/>', "Normal", ""),
            ("paragraph", "Prrafodelista", "List Paragraph", '<w:spacing w:after="80"/><w:jc w:val="both"/>', "",
             "Normal", ""),
            ("paragraph", "Encabezado", "header", '<w:tabs><w:tab w:val="center" w:pos="4513"/><w:tab w:val="right" '
             'w:pos="9026"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="left"/>',
             '<w:sz w:val="18"/><w:szCs w:val="18"/>', "Normal", ""),
            ("paragraph", "Piedepgina", "footer", '<w:tabs><w:tab w:val="center" w:pos="4513"/><w:tab w:val="right" '
             'w:pos="9026"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/>',
             '<w:sz w:val="16"/><w:szCs w:val="16"/>', "Normal", ""),
        ]
        for tipo, sid, nombre, ppr, rpr, basado, extra in defs:
            previo = st.find(f'{qn("w:style")}[@{qn("w:styleId")}="{sid}"]')
            if previo is None:
                previo = next((s for s in st.findall(qn("w:style"))
                               if s.find(qn("w:name")) is not None and s.find(qn("w:name")).get(qn("w:val")) == nombre),
                              None)
            nuevo = parse_xml(_estilo_xml(tipo, sid, nombre, ppr, rpr, basado, extra))
            if previo is not None:
                nuevo.set(qn("w:styleId"), previo.get(qn("w:styleId")))
                st.replace(previo, nuevo)
            else:
                st.append(nuevo)
        self.E = {n: self.d.styles[n] for n in ("Heading 1", "Heading 2", "Heading 3", "Caption", "Fuente MIMP",
                                               "Texto de tabla MIMP", "toc 1", "toc 2", "table of figures",
                                               "TOC Heading")}

    @staticmethod
    def _limpiar_notas(raiz):
        for nota in list(raiz):
            if nota.get(qn("w:type")) not in ("separator", "continuationSeparator"):
                raiz.remove(nota)

    # ------------------------------------------------------------------ bloques de texto
    def _par(self, piezas, size=None, estilo=None, **kw):
        if size is None and not self.informe:
            size = 22
        par = self.d.add_paragraph()
        if estilo is not None:
            par.style = estilo
        _ppr(par, **kw)
        for texto, fmt in _partir(piezas):
            _pieza(par, texto, fmt, size)
        return par

    def titulo(self, linea1, linea2, linea3=""):
        self._par([(linea1, {"b": True, "color": AZUL, "size": 28})], size=28, jc="center", before=240, after=100)
        self._par([(linea2, {"b": True, "color": AZUL})], jc="center", after=120 if linea3 else 240)
        if linea3:
            self._par([(linea3, {"color": GRIS, "size": 20})], size=20, jc="center", after=240)

    def h1(self, texto):
        self._seccion = texto
        self.titulos.append((texto, 1))
        if self.informe:
            return self._par([texto], estilo=self.E["Heading 1"])
        par = self._par([(texto, {"b": True, "color": AZUL})], keep_next=True, before=320, after=160)
        par.style = self.d.styles["Heading 1"]
        return par

    def h2(self, texto):
        self._seccion = texto
        self.titulos.append((texto, 2))
        if self.informe:
            return self._par([texto], estilo=self.E["Heading 2"])
        par = self._par([(texto, {"b": True, "color": AZUL})], keep_next=True, before=200, after=100)
        par.style = self.d.styles["Heading 2"]
        return par

    def h3(self, texto):
        if self.informe:
            return self._par([texto], estilo=self.E["Heading 3"])
        return self._par([(texto, {"b": True, "i": True, "color": AZUL})], keep_next=True, before=140, after=80)

    def titulo_indice(self, texto):
        """Título de un índice (estilo «TOC Heading»): no entra en el índice general."""
        return self._par([texto], estilo=self.E["TOC Heading"])

    def p(self, *piezas, jc=None, size=None, after=None, recuadro=None):
        """Párrafo de cuerpo. En el modo informe, el formato viene del estilo Normal y solo se escribe lo que se pide;
        recuadro: «pregunta», «respuesta» o «hallazgo»."""
        if self.informe:
            par = self._par(piezas, size=size, jc=jc, after=after)
            if recuadro:
                self._recuadro(par, *RECUADROS[recuadro])
            return par
        return self._par(piezas, size=size or 22, jc=jc or "both", after=120 if after is None else after, line=276)

    def _recuadro(self, par, relleno, borde):
        ppr = par._p.get_or_add_pPr()
        if borde:
            color, grosor = borde
            bdr = OxmlElement("w:pBdr")
            b = OxmlElement("w:left")
            b.set(qn("w:val"), "single"); b.set(qn("w:sz"), str(grosor)); b.set(qn("w:space"), "8")
            b.set(qn("w:color"), color)
            bdr.append(b)
            _en_orden(ppr, bdr)
        if relleno:
            sh = OxmlElement("w:shd"); sh.set(qn("w:val"), "clear"); sh.set(qn("w:color"), "auto")
            sh.set(qn("w:fill"), relleno)
            _en_orden(ppr, sh)
            if not borde:
                bdr = OxmlElement("w:pBdr")
                for lado in ("top", "left", "bottom", "right"):
                    b = OxmlElement(f"w:{lado}")
                    b.set(qn("w:val"), "single"); b.set(qn("w:sz"), "4"); b.set(qn("w:space"), "6")
                    b.set(qn("w:color"), relleno)
                    bdr.append(b)
                _en_orden(ppr, bdr)
        ind = OxmlElement("w:ind"); ind.set(qn("w:left"), "170"); ind.set(qn("w:right"), "170")
        _en_orden(ppr, ind)

    def vineta(self, *piezas, size=None):
        if self.informe:
            par = self._par(piezas, size=size, num=2, after=80)
        else:
            par = self._par(piezas, size=size or 22, num=2, after=90, line=276, jc="both")
        par.style = self.d.styles["List Paragraph"]
        return par

    def cita(self, texto, fuente):
        return self._par([(f"«{texto}» ", {"i": True}), (f"({fuente})", {})], size=20, jc="both", after=120,
                         ind=567)

    def _leyenda(self, rotulo, numero, texto):
        """«Tabla N. Texto» con el número en un campo SEQ, para que Word numere y arme los índices de tablas y de
        figuras (campo TOC \\c). En el modo informe el formato viene del estilo de descripción, de modo que el número
        conserva la negrita y el color cuando Word actualiza el campo."""
        self.leyendas.append((rotulo, numero, texto))
        if self.informe:
            par = self._par([f"{rotulo} "], estilo=self.E["Caption"])
            _campo(par, f"SEQ {rotulo} \\* ARABIC \\* MERGEFORMAT", str(numero))
            _pieza(par, f". {texto}", {}, None)
            return par
        fmt = {"b": True, "color": AZUL, "size": 20}
        par = self._par([(f"{rotulo} ", fmt)], size=20, keep_next=True, before=120, after=60)
        campo = OxmlElement("w:fldSimple")
        campo.set(qn("w:instr"), f" SEQ {rotulo} \\* ARABIC ")
        run = _rpr(par.add_run(str(numero)), fmt, 20)
        campo.append(run._r)
        par._p.append(campo)
        _pieza(par, f". {texto}", fmt, 20)
        return par

    def leyenda(self, texto):
        self.n_tabla += 1
        return self._leyenda("Tabla", self.n_tabla, texto)

    def indice(self, instr, entradas, nivel_de=None):
        """Índice como campo TOC de Word (instr, p. ej. «TOC \\o "1-2" \\h \\z \\u» o «TOC \\h \\z \\c "Tabla"»), con las
        entradas en caché: lista de (texto, página, nivel). Las entradas usan los estilos TDC 1, TDC 2 y tabla de
        ilustraciones, los mismos que Word aplica al regenerar el índice."""
        if not entradas:
            entradas = [("(Actualice el índice en Word: clic derecho, Actualizar campos.)", "", 1)]
        for k, (texto, pagina, nivel) in enumerate(entradas):
            estilo = self.E["table of figures"] if nivel_de else self.E["toc 1" if nivel == 1 else "toc 2"]
            par = self.d.add_paragraph()
            par.style = estilo
            if k == 0:
                for tipo, extra in (("begin", None), ("instr", instr), ("separate", None)):
                    r = OxmlElement("w:r")
                    if tipo == "instr":
                        it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = f" {extra} "
                        r.append(it)
                    else:
                        fc = OxmlElement("w:fldChar"); fc.set(qn("w:fldCharType"), tipo); r.append(fc)
                    par._p.append(r)
            par.add_run(texto)
            if pagina != "":
                r = par.add_run(); r._r.append(OxmlElement("w:tab"))
                par.add_run(str(pagina))
            if k == len(entradas) - 1:
                r = OxmlElement("w:r"); fc = OxmlElement("w:fldChar"); fc.set(qn("w:fldCharType"), "end")
                r.append(fc); par._p.append(r)

    # ------------------------------------------------------------------ secciones y portada
    def cerrar_seccion(self, numeracion=None):
        """Cierra la sección en curso con un salto de sección (página siguiente). numeracion: formato de página de la
        sección («lowerRoman») desde 1. La sección de portada la cierra el módulo src/nucleo/portada."""
        sect = copy.deepcopy(self.d.element.body.find(qn("w:sectPr")))
        for e in sect.findall(qn("w:pgNumType")):
            sect.remove(e)
        if numeracion:
            pg = OxmlElement("w:pgNumType"); pg.set(qn("w:fmt"), numeracion); pg.set(qn("w:start"), "1")
            sect.find(qn("w:pgMar")).addnext(pg)
        par = self.d.add_paragraph()
        _en_orden(par._p.get_or_add_pPr(), sect)
        return par

    def numerar_cuerpo_desde_1(self):
        """La última sección (cuerpo) numera en arábigos desde 1."""
        sect = self.d.element.body.find(qn("w:sectPr"))
        for e in sect.findall(qn("w:pgNumType")):
            sect.remove(e)
        pg = OxmlElement("w:pgNumType"); pg.set(qn("w:start"), "1")
        sect.find(qn("w:pgMar")).addnext(pg)

    # ------------------------------------------------------------------ figuras, fuentes y notas
    def _rastro(self, tipo, numero, leyenda, rastro):
        if rastro:
            self.rastros.append({"tipo": tipo, "numero": numero, "seccion": self._seccion, "leyenda": leyenda or "",
                                 "rastro": rastro})

    def figura(self, ruta, leyenda, fuente=None, ancho_cm=16.0, rastro=None, alt=None, nota=None):
        """Imagen centrada con leyenda numerada («Figura N.») arriba y fuente abajo, como las tablas. rastro: libro y
        hoja de donde sale la figura, que se registra aparte (self.rastros) y no va en el texto. alt: texto alternativo
        (por defecto, la leyenda). nota: «Nota: …» bajo la fuente."""
        if self.informe:
            ancho_cm = min(ancho_cm, ANCHO_TEXTO_CM)
        self.n_figura += 1
        self.figuras.append({"numero": self.n_figura, "seccion": self._seccion, "leyenda": leyenda,
                             "archivo": str(ruta)})
        self._rastro("Figura", self.n_figura, leyenda, rastro)
        self._leyenda("Figura", self.n_figura, leyenda)
        par = self.d.add_paragraph()
        _ppr(par, jc="center", after=60, keep_next=bool(fuente))
        forma = par.add_run().add_picture(ruta, width=Cm(ancho_cm))
        doc_pr = forma._inline.docPr
        doc_pr.set("descr", alt or f"Figura {self.n_figura}. {leyenda}")
        doc_pr.set("title", f"Figura {self.n_figura}")
        if fuente:
            self.fuente(fuente)
        if nota:
            self.fuente(f"Nota: {nota}")
        return par

    def fuente(self, texto):
        if self.informe:
            return self._par([texto], estilo=self.E["Fuente MIMP"])
        return self._par([(texto, {"color": GRIS, "size": 16})], size=16, after=160, jc="both")

    def nota_pie(self, par, texto):
        """Agrega al final del párrafo una llamada a una nota al pie (Word la numera)."""
        parte, raiz = self._notas_pie
        nid = max([int(n.get(qn("w:id"))) for n in raiz] + [0]) + 1
        nota = etree.SubElement(raiz, qn("w:footnote"))
        nota.set(qn("w:id"), str(nid))
        np_ = etree.SubElement(nota, qn("w:p"))
        etree.SubElement(etree.SubElement(np_, qn("w:pPr")), qn("w:pStyle")).set(
            qn("w:val"), self.d.styles["footnote text"].style_id)
        r1 = etree.SubElement(np_, qn("w:r"))
        etree.SubElement(etree.SubElement(r1, qn("w:rPr")), qn("w:rStyle")).set(
            qn("w:val"), self.d.styles["footnote reference"].style_id)
        etree.SubElement(r1, qn("w:footnoteRef"))
        r2 = etree.SubElement(np_, qn("w:r"))
        t = etree.SubElement(r2, qn("w:t"))
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = " " + texto
        if not hasattr(parte, "element"):
            parte._blob = etree.tostring(raiz, xml_declaration=True, encoding="UTF-8", standalone=True)
        run = par.add_run()
        rpr = run._r.get_or_add_rPr()
        estilo = OxmlElement("w:rStyle")
        estilo.set(qn("w:val"), self.d.styles["footnote reference"].style_id)
        rpr.append(estilo)
        ref = OxmlElement("w:footnoteReference")
        ref.set(qn("w:id"), str(nid))
        run._r.append(ref)
        return nid

    def salto(self):
        self.d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # ------------------------------------------------------------------ tablas
    def tabla(self, encabezado, filas, anchos, leyenda=None, fuente=None, size=18, grupos=(), rastro=None, nota=None):
        """filas: lista de listas de celdas (str, lista de piezas o {"piezas": …, "fill": "RRGGBB"} para una celda con
        relleno propio). grupos: índices de filas sombreadas, que en el modo informe van con la fila siguiente. rastro:
        libro y hoja de donde sale la tabla, que se registra aparte (self.rastros). nota: «Nota: …» bajo la fuente."""
        if self.informe:
            size = max(size, TAMANO_MINIMO_TABLA)
        if leyenda:
            self.leyenda(leyenda)
        self._rastro("Tabla", self.n_tabla if leyenda else "", leyenda, rastro)
        t = self.d.add_table(rows=1 + len(filas), cols=len(encabezado))
        tbl = t._tbl
        tblpr = tbl.tblPr
        for e in list(tblpr):
            if e.tag in (qn("w:tblStyle"), qn("w:tblW"), qn("w:tblLook")):
                tblpr.remove(e)
        w = OxmlElement("w:tblW"); w.set(qn("w:w"), str(sum(anchos))); w.set(qn("w:type"), "dxa"); tblpr.append(w)
        bordes = OxmlElement("w:tblBorders")
        for lado, sz in (("top", 4), ("left", 4), ("bottom", 4), ("right", 4), ("insideH", 2), ("insideV", 2)):
            b = OxmlElement(f"w:{lado}")
            b.set(qn("w:val"), "single"); b.set(qn("w:sz"), str(sz)); b.set(qn("w:space"), "0")
            b.set(qn("w:color"), "BFBFBF"); bordes.append(b)
        tblpr.append(bordes)
        lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); tblpr.append(lay)
        grid = tbl.tblGrid
        for gc, a in zip(grid.findall(qn("w:gridCol")), anchos):
            gc.set(qn("w:w"), str(a))
        estilo_celda = self.E["Texto de tabla MIMP"] if self.informe else None
        tam = size if not self.informe or size != 16 else None

        def nuevo_par(c, primero=False):
            par = c.paragraphs[0] if primero else c.add_paragraph()
            if estilo_celda is not None:
                par.style = estilo_celda
            else:
                _ppr(par, after=40)
            return par

        for i, fila in enumerate([encabezado] + list(filas)):
            tr = t.rows[i]._tr
            trpr = tr.get_or_add_trPr()
            trpr.append(OxmlElement("w:cantSplit"))
            if i == 0:
                trpr.append(OxmlElement("w:tblHeader"))
            grupo = (i - 1) in grupos
            for j, celda in enumerate(fila):
                c = t.rows[i].cells[j]
                tcpr = c._tc.get_or_add_tcPr()
                tw = tcpr.find(qn("w:tcW"))
                if tw is None:
                    tw = OxmlElement("w:tcW"); tcpr.insert(0, tw)
                tw.set(qn("w:w"), str(anchos[j])); tw.set(qn("w:type"), "dxa")
                propio = None
                if isinstance(celda, dict):
                    propio, celda = celda.get("fill"), celda["piezas"]
                relleno = AZUL if i == 0 else (FILA_GRUPO if grupo else propio)
                if relleno:
                    sh = OxmlElement("w:shd"); sh.set(qn("w:val"), "clear"); sh.set(qn("w:color"), "auto")
                    sh.set(qn("w:fill"), relleno); tcpr.append(sh)
                mar = OxmlElement("w:tcMar")
                for lado, v in (("top", 50), ("left", 80), ("bottom", 50), ("right", 80)):
                    e = OxmlElement(f"w:{lado}"); e.set(qn("w:w"), str(v)); e.set(qn("w:type"), "dxa"); mar.append(e)
                tcpr.append(mar)
                par = nuevo_par(c, primero=True)
                piezas = celda if isinstance(celda, list) else [str(celda)]
                for k, (texto, fmt) in enumerate(_partir(piezas)):
                    f = dict(fmt)
                    if i == 0:
                        f.update(b=True, color="FFFFFF")
                    elif grupo:
                        f.setdefault("b", True); f.setdefault("color", AZUL)
                    # saltos de línea dentro de la celda
                    partes = texto.split("\n")
                    for m, parte in enumerate(partes):
                        if m:
                            par = nuevo_par(c)
                        if parte:
                            _pieza(par, parte, {k: v for k, v in f.items() if k != "pie"}, tam)
                    if f.get("pie") and self._notas_pie:
                        self.nota_pie(par, f["pie"])
                if self.informe and (i == 0 or grupo):
                    for pp in c.paragraphs:
                        _en_orden(pp._p.get_or_add_pPr(), OxmlElement("w:keepNext"))
        if fuente:
            self.fuente(fuente)
        if nota:
            self.fuente(f"Nota: {nota}")
        if not fuente and not nota:
            self._par([("", {})], size=8, after=60)
        return t

    def guardar(self, ruta, comillas_inglesas=False):
        """Guarda el documento. Con comillas_inglesas (informe de cierre, ronda 35, pedido del usuario), las comillas
        « » pasan a “ ” en el cuerpo, los encabezados, los pies y las notas al pie; dentro de una cita, a ‘ ’."""
        if comillas_inglesas:
            raices = [self.d.element]
            for s in self.d.sections:
                for hf in (s.header, s.footer, s.first_page_header, s.first_page_footer):
                    raices.append(hf._element)
            if self._notas_pie:
                raices.append(self._notas_pie[1])
            for raiz in raices:
                for par in raiz.iter(qn("w:p")):
                    comillas_parrafo(par)
            if self._notas_pie and not hasattr(self._notas_pie[0], "element"):
                self._notas_pie[0]._blob = etree.tostring(self._notas_pie[1], xml_declaration=True,
                                                          encoding="UTF-8", standalone=True)
        self.d.save(ruta)
