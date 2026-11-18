"""Infografía de la evaluación de cierre de la PNIG (2019–2025): dos caras A4, estilo del MIMP Lab (pedido del usuario
del 2026-11-17, ADR-016).

Cara 1: qué se evaluó y cómo, resultados clave y valoración de las doce preguntas.
Cara 2: hallazgos principales, transición normativa y recomendaciones.

Las cifras se leen de las mismas fuentes que el informe (sintesis_hallazgos.cifras(), hallazgos, valoraciones y
recomendaciones de informe_cierre), para que coincidan con el informe y las presentaciones. El HTML es autocontenido
(logos incrustados) y Chrome sin ventana lo imprime en PDF (nucleo.figuras).

Colores (skill dataviz, validate_palette.js): estados de los servicios y de la valoración con los colores de las
figuras del informe, con el ámbar oscurecido a #e8a317 para pasar la banda de luminosidad y con etiqueta en cada
marca; acentos del MIMP Lab (teal, violeta, verde y petróleo), con el dorado claro solo como decoración.

Uso (desde la raíz del repositorio):
    uv run python src/analisis_pnig/infografia.py
Salida: evaluaciones/pnig_2026/06_informe/2026_mimp-ogmepgd_infografia-evaluacion-cierre-pnig.pdf y su HTML fuente
(06_informe/infografia/infografia_evaluacion_cierre_pnig.html).
"""
import base64
import html
import os
import re
import sys
import tempfile
from pathlib import Path

import pandas as pd
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
from nucleo.figuras import _chrome  # noqa: E402
from nucleo.formato import comillas  # noqa: E402
EVAL = _cargar_config().raiz

sys.path.insert(0, os.path.dirname(__file__))
import figuras_informe as FIGI  # noqa: E402
import informe_cierre as IC  # noqa: E402
import sintesis_hallazgos as S  # noqa: E402
from presentacion_resultados import TEMA  # noqa: E402

SALIDA = Path(f"{EVAL}/06_informe/2026_mimp-ogmepgd_infografia-evaluacion-cierre-pnig.pdf")
FUENTE_HTML = Path(f"{EVAL}/06_informe/infografia/infografia_evaluacion_cierre_pnig.html")
LOGOS = Path("assets/logos")

PETROL, TEAL, VIOLET, GREEN, GOLD = "#034754", "#14A7AE", "#7C6FC4", "#4CAF93", "#FEC34D"
INK, INK2, MUTED, LINE, PAGE = "#1f2a30", "#4a5459", "#6b7378", "#dcdcdc", "#f4f4f3"
ESTADO = {"Implementado": "#0ca30c", "Implementado parcialmente": "#e8a317", "No implementado": "#d03b3b"}
VALOR = {"Satisfactoria": "#0ca30c", "Parcial": "#e8a317", "Insuficiente": "#d03b3b"}
GRIS = "#9a9a9a"


def e(t):
    return html.escape(comillas(str(t))[0])


def logo(nombre):
    return "data:image/png;base64," + base64.b64encode((LOGOS / nombre).read_bytes()).decode("ascii")


CSS = f"""
@page {{ size: A4; margin: 0 }}
* {{ box-sizing: border-box; margin: 0; padding: 0 }}
html, body {{ -webkit-print-color-adjust: exact; print-color-adjust: exact }}
body {{ font-family: Poppins, Arial, sans-serif; color: {INK}; background: {PAGE}; font-size: 8.2pt; line-height: 1.28 }}
.pagina {{ width: 210mm; height: 297mm; position: relative; overflow: hidden; background: {PAGE};
          page-break-after: always; padding: 0 11mm }}
.pagina:last-child {{ page-break-after: auto }}
.banda {{ height: 17mm; background: #fff; margin: 0 -11mm; padding: 0 11mm; display: flex; align-items: center;
         justify-content: space-between; border-bottom: 1px solid {LINE} }}
.banda img {{ height: 10mm }}
.titulo {{ background: {PETROL}; color: #fff; margin: 0 -11mm; padding: 4.5mm 11mm 4mm }}
.titulo h1 {{ font-size: 17pt; line-height: 1.12; font-weight: 700 }}
.titulo p {{ font-size: 9.5pt; margin-top: 1.5mm; color: #d6eef0 }}
.franja {{ position: absolute; left: 0; right: 0; bottom: 0; height: 3mm; display: flex }}
.franja span {{ flex: 1 }}
.pie {{ position: absolute; left: 11mm; right: 11mm; bottom: 5.5mm; font-size: 7pt; color: {MUTED};
       display: flex; justify-content: space-between }}
h2 {{ font-size: 11pt; color: {PETROL}; font-weight: 700; margin: 3.4mm 0 1.6mm; display: flex; align-items: center;
     gap: 2.2mm }}
h2 .marca {{ width: 3mm; height: 3mm; border-radius: 50%; display: inline-block }}
.tarjeta {{ background: #fff; border: 1px solid {LINE}; border-radius: 2.5mm; padding: 2.4mm 3mm }}
.fila {{ display: grid; gap: 3mm }}
.cifra {{ border-left: 1.4mm solid; }}
.cifra .n {{ font-size: 18pt; font-weight: 700; line-height: 1.05 }}
.cifra .t {{ font-size: 8pt; color: {INK2}; margin-top: 1mm }}
.barra {{ display: flex; height: 7mm; gap: 2px; margin: 2.2mm 0 1.6mm }}
.barra span {{ display: block; height: 100% }}
.barra span:first-child {{ border-radius: 4px 0 0 4px }}
.barra span:last-child {{ border-radius: 0 4px 4px 0 }}
.ley {{ display: flex; flex-wrap: wrap; gap: 1.2mm 4mm; font-size: 7.8pt; color: {INK2} }}
.ley i {{ display: inline-block; width: 2.6mm; height: 2.6mm; border-radius: 0.6mm; margin-right: 1.2mm;
         vertical-align: -0.3mm }}
.ley b {{ color: {INK} }}
.cuadros {{ display: flex; gap: 1.2mm; margin: 2.2mm 0 1.6mm }}
.cuadros span {{ width: 6.2mm; height: 6.2mm; border-radius: 1mm; display: block }}
.puntos {{ display: flex; gap: 1.2mm; margin: 2.2mm 0 1.6mm }}
.puntos span {{ width: 5.4mm; height: 5.4mm; border-radius: 50%; display: block; border: 1.5px solid {TEAL} }}
.preg {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 1.8mm }}
.preg div {{ border-radius: 1.6mm; padding: 1.2mm 1.8mm; background: #fff; border: 1px solid {LINE};
            border-top: 1.4mm solid; min-height: 11.5mm }}
.preg b {{ font-size: 9pt; color: {PETROL} }}
.preg .tema {{ font-size: 7.4pt; color: {INK2}; display: block; line-height: 1.2; margin: 0.6mm 0 }}
.preg .val {{ font-size: 7.4pt; font-weight: 700 }}
ul {{ list-style: none }}
li {{ position: relative; padding-left: 3.6mm; margin-bottom: 1.1mm }}
li::before {{ content: ""; position: absolute; left: 0; top: 1.5mm; width: 1.6mm; height: 1.6mm; border-radius: 50%;
             background: var(--c, {TEAL}) }}
.cab {{ font-size: 9.6pt; font-weight: 700; color: #fff; border-radius: 1.6mm 1.6mm 0 0; padding: 1.6mm 3mm;
       margin: -2.4mm -3mm 2mm }}
.linea {{ position: relative; margin: 1mm 0 0 2mm; padding-left: 5mm; border-left: 2px solid {LINE} }}
.linea div {{ position: relative; margin-bottom: 1.2mm }}
.linea div::before {{ content: ""; position: absolute; left: -6.6mm; top: 0.9mm; width: 2.8mm; height: 2.8mm;
                     border-radius: 50%; background: var(--c, {TEAL}); border: 2px solid #fff }}
.linea .f {{ font-weight: 700; color: {PETROL}; font-size: 8pt }}
.rec {{ font-size: 7.5pt }}
.rec li {{ margin-bottom: 0.7mm }}
.rec .resp {{ color: {MUTED} }}
.nota {{ font-size: 7.2pt; color: {MUTED}; margin-top: 1.4mm }}
"""


def cifra(n, t, color):
    return f'<div class="tarjeta cifra" style="border-left-color:{color}"><div class="n" style="color:{color}">' \
           f'{e(n)}</div><div class="t">{e(t)}</div></div>'


def h2(texto, color):
    return f'<h2><span class="marca" style="background:{color}"></span>{e(texto)}</h2>'


def lista(items, color):
    return "<ul>" + "".join(f'<li style="--c:{color}">{e(t)}</li>' for t in items) + "</ul>"


def cabecera():
    """Banda con los logos; sin ellos (la versión pública no lleva assets/logos) queda la banda en blanco."""
    imgs = [f'<img src="{logo(f)}" alt="{alt}">' if (LOGOS / f).exists() else "<span></span>"
            for f, alt in (("logo_institucion.png", "Ministerio de la Mujer y Poblaciones Vulnerables"),
                           ("logo_gobierno.png", "Gobierno del Perú"))]
    return f'<div class="banda">{"".join(imgs)}</div>'


def cierre(n):
    return (f'<div class="pie"><span>Fuente: informe de evaluación de cierre de la PNIG (2019–2025), OGMEPGD, MIMP, '
            f'noviembre de 2026. Elaboración: OMEP.</span><span>{n} / 2</span></div>'
            f'<div class="franja"><span style="background:{PETROL}"></span><span style="background:{TEAL}"></span>'
            f'<span style="background:{GOLD}"></span></div>')


def construir():
    C = S.cifras()
    H = S.hallazgos(C)
    V = S.valoraciones(C)
    g = C["grd"]
    n = len(g)
    documentos = len(pd.read_csv("docs/fuentes/inventario_fuentes.csv"))
    conf = pd.Series([h["confianza"] for h in H]).value_counts()
    sin_prot = n - C["serv_prot_si"]
    ses = sum(len(f) for a, f, _ in FIGI.SESIONES_CONAIG if a <= 2025)
    mc = C["mic_conc"]
    disp = int(C["mic_norma"]["Disposiciones revisadas"].sum())
    m5 = C["mac5"]
    conc_div = int(m5.loc["Conceptual", "Divergencia"])
    conc_tot = int(m5.loc["Conceptual", ["Continuidad", "Precisión", "Divergencia"]].sum())
    red, sinc, amp = len(C["brechas_reduce"]), len(C["brechas_sincambio"]), len(C["brechas_amplia"])
    parc = sum(v[0] == "Parcial" for v in V.values())
    insu = sum(v[0] == "Insuficiente" for v in V.values())
    nombre_serv = dict(zip(g["Código"].map(str), g["Servicio"]))
    recs = IC.recomendaciones(C, nombre_serv)
    assert C["serv_impl"] + C["serv_parc"] + C["serv_noimpl"] == n

    # ----- cara 1 -----
    estados = [("Implementado", C["serv_impl"], "implementados"),
               ("Implementado parcialmente", C["serv_parc"], "implementados parcialmente"),
               ("No implementado", C["serv_noimpl"], "no implementados")]
    barra = '<div class="barra">' + "".join(
        f'<span style="width:{100 * k / n:.2f}%;background:{ESTADO[est]}"></span>' for est, k, _ in estados) + "</div>"
    ley_serv = '<div class="ley">' + "".join(
        f'<span><i style="background:{ESTADO[est]}"></i><b>{k}</b> {e(t)}</span>' for est, k, t in estados) + "</div>"
    cuadros = '<div class="cuadros">' + "".join(
        f'<span style="background:{c}"></span>' for c in ["#0ca30c"] * red + [GRIS] * sinc + ["#d03b3b"] * amp) + "</div>"
    ley_br = (f'<div class="ley"><span><i style="background:#0ca30c"></i><b>{red}</b> se redujeron</span>'
              f'<span><i style="background:{GRIS}"></i><b>{sinc}</b> no cambiaron</span>'
              f'<span><i style="background:#d03b3b"></i><b>{amp}</b> se amplió</span></div>')
    puntos = '<div class="puntos">' + "".join(
        f'<span style="background:{TEAL if i < C["op_mejora"] else "#fff"}"></span>' for i in range(C["op_ind"])) + "</div>"
    preguntas = '<div class="preg">' + "".join(
        f'<div style="border-top-color:{VALOR[V[p][0]]}"><b>{p}</b><span class="tema">{e(TEMA[p])}</span>'
        f'<span class="val" style="color:{INK}">{e(V[p][0])}</span></div>' for p in S.PREGUNTAS) + "</div>"
    cara1 = f"""
<section class="pagina">
{cabecera()}
<div class="titulo"><h1>Evaluación de cierre de la Política Nacional de Igualdad de Género (PNIG), 2019–2025</h1>
<p>Infografía de resultados · OGMEPGD, Ministerio de la Mujer y Poblaciones Vulnerables · Noviembre de 2026</p></div>
{h2("Qué se evaluó y cómo", TEAL)}
<div class="tarjeta" style="margin-bottom:3mm">{e(
        "La evaluación de cierre revisa el diseño y la implementación de la PNIG (D.S. N.° 008-2019-MIMP) en todo su "
        "periodo de vigencia, de 2019 a 2025, como insumo para la política que la sustituye en el marco de la Ley N.° "
        f"32535. La Política se organizó en 6 objetivos prioritarios, {C['lin_total']} lineamientos y {n} servicios.")}</div>
<div class="fila" style="grid-template-columns:repeat(4,1fr)">
{cifra("12", "preguntas de evaluación: 7 de diseño y 5 de implementación", TEAL)}
{cifra(str(documentos), "documentos revisados: normas, planes e informes", VIOLET)}
{cifra(str(C["rev_fuentes"]), "estudios sobre la eficacia de los servicios", GREEN)}
{cifra(str(len(H)), f"hallazgos: {conf.get('Sólido', 0)} sólidos, {conf.get('Con respaldo parcial', 0)} con respaldo parcial y {conf.get('Indicio', 0)} indicios", PETROL)}
</div>
<p class="nota">Seis instrumentos: revisión documental, entrevistas al nivel conductor ({e(", ".join(S.FECHA_ENTREV))}),
reportes de las entidades, coherencia y alineamiento, indicadores y recursos, y revisión de evidencia.</p>
{h2("Resultados clave", GREEN)}
<div class="fila" style="grid-template-columns:1.35fr 1fr">
<div class="tarjeta"><b style="color:{PETROL}">Implementación de los {n} servicios</b>{barra}{ley_serv}
<p class="nota">Implementado: dos de tres componentes (protocolo aprobado, recursos en el MEF e indicadores con
valor) y ejecución en 2024 o 2025. {sin_prot} de {n} servicios siguieron sin protocolo aprobado.</p></div>
<div class="tarjeta"><b style="color:{PETROL}">Las {C["brechas_n"]} brechas del diagnóstico persisten</b>{cuadros}{ley_br}
<p class="nota">La ampliación (trabajo no remunerado) surge de una comparación indicativa, porque la encuesta cambió.</p></div>
</div>
<div class="fila" style="grid-template-columns:1.35fr 1fr;margin-top:3mm">
<div class="tarjeta"><b style="color:{PETROL}">{C["op_mejora"]} de {C["op_ind"]} indicadores de objetivo prioritario mejoraron frente a su línea de base</b>{puntos}
<p class="nota">El reporte oficial no tuvo valor 2025 para {C["op_sinvalor_2025"]} de los {C["op_ind"]}; al ritmo
observado, solo la meta de embarazo adolescente se alcanzaría al 2030.</p></div>
<div class="tarjeta cifra" style="border-left-color:{VIOLET}"><div class="n" style="color:{VIOLET}">{ses}</div>
<div class="t">{e(f"sesiones ordinarias de la CONAIG entre 2021 y 2025; su acuerdo más reiterado, culminar los protocolos pendientes, no se cumplió")}</div></div>
</div>
{h2("Valoración de las doce preguntas", VIOLET)}
{preguntas}
<div class="ley" style="margin-top:2mm"><span><i style="background:{VALOR['Parcial']}"></i><b>{parc}</b> parciales</span>
<span><i style="background:{VALOR['Insuficiente']}"></i><b>{insu}</b> insuficiente (calidad de los indicadores)</span>
<span><b>0</b> satisfactorias</span></div>
{cierre(1)}
</section>"""

    # ----- cara 2 -----
    diseno = [
        f"Las {C['brechas_n']} brechas del diagnóstico persisten al cierre: el diagnóstico conserva vigencia empírica, "
        "pero su modelo del problema era incompatible con la Ley N.° 32535, según la OGAJ.",
        f"La cadena causal estuvo casi completa: {C['lin_total'] - len(C['lin_sin_serv'])} de {C['lin_total']} "
        f"lineamientos tuvieron servicio.",
        f"Las fichas de los indicadores incumplieron atributos críticos y {C['serv_sin_ind_vig']} servicios no tuvieron "
        f"indicador aprobado.",
        f"El nuevo marco diverge en lo conceptual ({conc_div} de {conc_tot} conceptos), pero los servicios se "
        f"trasladan casi completos.",
    ]
    implementacion = [
        f"{C['serv_impl']} de {n} servicios estuvieron implementados y {C['serv_parc']} lo estuvieron parcialmente; el OP2 no "
        f"tuvo servicios implementados.",
        f"Los indicadores con valor mostraron avances mixtos y {C['op_sinvalor_2025']} de {C['op_ind']} no tuvieron valor "
        f"oficial 2025.",
        "La ejecución presupuestal fue alta, pero según la DGIGND la conducción no contó con los recursos necesarios.",
        "La CONAIG funcionó con regularidad formal, pero con baja asistencia de los titulares.",
        "La rotación de personal es la brecha de capacidad más persistente.",
    ]
    hitos = [
        ("24 de diciembre de 2025", "Se publica la Ley N.° 32535, que deroga la Ley N.° 28983.", PETROL),
        ("25 de diciembre de 2025", "La PNIG pierde vigencia, según el criterio de la OGAJ.", "#d03b3b"),
        ("16 de setiembre de 2026", "La OGAJ opina que las normas siguen vigentes y pide identificar lo incompatible.",
         GRIS),
        ("9 de noviembre de 2026", f"La DGIGND, la DPIGND y la DTEG marcan como incompatibles {disp} disposiciones.",
         VIOLET),
        ("11 de noviembre de 2026", "La OGAJ revisa su criterio: la PNIG, el PEMIG y la CONAIG perdieron vigencia.",
         "#d03b3b"),
        ("25 de diciembre de 2026", "Vence la adecuación de los instrumentos de las entidades.", TEAL),
        ("31 de diciembre de 2026", "Vence la comunicación de la derogación tácita (Ley N.° 31772).", TEAL),
    ]
    linea = '<div class="linea">' + "".join(
        f'<div style="--c:{c}"><span class="f">{e(f)}</span><br>{e(t)}</div>' for f, t, c in hitos) + "</div>"
    grupos = {}
    for r in recs:
        tema, sub = r[0].split(":", 1)
        grupos.setdefault(tema.strip(), []).append((sub.strip()[0].upper() + sub.strip()[1:], r[2]))
    color_g = {"Diseño": TEAL, "Seguimiento y evaluación": VIOLET, "Implementación": GREEN}
    rec_html = '<div class="fila" style="grid-template-columns:repeat(3,1fr)">' + "".join(
        f'<div class="tarjeta rec"><div class="cab" style="background:{color_g.get(t, PETROL)}">{e(t)} ({len(v)})</div>'
        f'<ul>' + "".join(f'<li style="--c:{color_g.get(t, PETROL)}">{e(a)} <span class="resp">· {e(resp)}</span></li>'
                          for a, resp in v) + "</ul></div>" for t, v in grupos.items()) + "</div>"
    cara2 = f"""
<section class="pagina">
{cabecera()}
{h2("Hallazgos principales", TEAL)}
<div class="fila" style="grid-template-columns:1fr 1fr">
<div class="tarjeta"><div class="cab" style="background:{TEAL}">Diseño de la Política</div>{lista(diseno, TEAL)}</div>
<div class="tarjeta"><div class="cab" style="background:{GREEN}">Implementación</div>{lista(implementacion, GREEN)}</div>
</div>
{h2("Transición normativa", "#d03b3b")}
<div class="fila" style="grid-template-columns:1fr 1fr">
<div class="tarjeta">{linea}</div>
<div class="tarjeta">{lista([
        "No hay vacío normativo: las obligaciones de incorporar el enfoque, prestar los servicios y reportar su "
        "cumplimiento nacen de la Ley N.° 32535 y, en violencia, de la Ley N.° 30364.",
        f"La matriz de incompatibilidades de la DPIGND coincide con la clasificación de esta evaluación en "
        f"{int(mc['Coinciden'])} de {int(mc['Servicios'])} servicios: la mayoría puede continuar con sustento en la "
        "ley; los de educación sexual integral se rediseñan.",
        "Los protocolos aprobados conservan su vigencia; los mecanismos y los Oficiales para la Igualdad quedan sin "
        "mandato reglamentario.",
        "La nueva política debe reemplazar a la PNIG; la CONAIG requiere un mecanismo de articulación que la sustituya.",
    ], "#d03b3b")}</div>
</div>
{h2(f"Recomendaciones ({len(recs)})", VIOLET)}
{rec_html}
<p class="nota">Cada recomendación lleva su responsable, su plazo y las preguntas que la sustentan (informe, capítulo 7).</p>
{h2("Productos de la evaluación", GOLD)}
<div class="fila" style="grid-template-columns:repeat(3,1fr)">
{"".join(f'<div class="tarjeta"><b style="color:{PETROL}">{e(f)}</b><br>{e(x)}</div>' for f, x in [
        ("Informe de evaluación de cierre", "Hallazgos con su respaldo y su confianza, conclusiones por pregunta y "
         f"{len(recs)} recomendaciones."),
        ("Anexo técnico", "Libros de análisis, rastro de cada hallazgo, instrumentos y código para replicar la "
         "evaluación."),
        ("Presentaciones e infografía", "Presentación de resultados completa y resumida, y esta infografía."),
    ])}
</div>
{cierre(2)}
</section>"""
    return (f'<!doctype html><html lang="es"><head><meta charset="utf-8"><title>Infografía: evaluación de cierre de la '
            f'PNIG</title><style>{CSS}</style></head><body>{cara1}{cara2}</body></html>')


def main():
    doc = construir()
    FUENTE_HTML.parent.mkdir(parents=True, exist_ok=True)
    FUENTE_HTML.write_text(doc, encoding="utf-8")
    with tempfile.TemporaryDirectory() as perfil:
        if SALIDA.exists():
            SALIDA.unlink()
        _chrome(perfil, f"--print-to-pdf={SALIDA.resolve()}", "--no-pdf-header-footer",
                "--print-to-pdf-no-header", FUENTE_HTML.resolve().as_uri())
    assert SALIDA.exists(), "Chrome no generó el PDF"
    n = len(re.findall(rb"/Type\s*/Page(?!s)", SALIDA.read_bytes()))   # páginas del PDF, sin herramientas externas
    assert n == 2, f"la infografía tiene {n} páginas; deben ser dos caras"
    print(f"Infografía: {SALIDA} ({n} páginas A4); fuente: {FUENTE_HTML}")


if __name__ == "__main__":
    main()
