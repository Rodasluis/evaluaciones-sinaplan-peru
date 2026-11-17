"""Presentación de resultados de la evaluación de cierre de la PNIG (2019–2025), con el estilo del MIMP Lab (skill
ppt-mimp-lab, decisión del usuario del 2026-11-11), para la reunión de validación con la DGIGND y el Despacho
Viceministerial de la Mujer y para la reunión técnica con el CEPLAN.

Las cifras se leen de las mismas fuentes que el informe (sintesis_hallazgos.cifras(), hallazgos, valoraciones y libros
de 04_analisis), para que coincidan. El script sirve para la primera versión: después, la presentación se edita a mano
sobre el archivo del repositorio (CLAUDE.md), no se regenera.

Estilo: nombres de los instrumentos, no siglas (ADR-010); comillas “ ”; Poppins; un acento por lámina según su tema y
el rojo solo para riesgos.

Uso (desde la raíz del repositorio; mimp_lab_deck y los componentes vienen de la skill del repositorio
.claude/skills/presentacion-evaluacion, basada en ppt-mimp y ppt-mimp-lab):
    uv run --no-project --with pandas --with openpyxl --with python-docx --with python-pptx --with pillow
        --with pyyaml python -X utf8 src/analisis_pnig/presentacion_resultados.py
Salida: evaluaciones/pnig_2026/06_informe/2026_mimp-ogmepgd_presentacion-resultados-evaluacion-cierre-pnig.pptx
"""
import os
import sys
from pathlib import Path

import pandas as pd
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

sys.path.insert(0, os.path.dirname(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
# skill del repositorio (Fase 2): estilos institucional y MIMP Lab, y componentes de evaluación
_SKILL = Path(__file__).resolve().parents[2] / ".claude" / "skills" / "presentacion-evaluacion" / "scripts"
sys.path.insert(0, str(_SKILL))

import componentes as K  # noqa: E402
from mimp_lab_deck import GOLD, GREEN, ICONS, PETROL, RED, TEAL, VIOLET, MimpLabDeck, tint  # noqa: E402
from pptx.chart.data import CategoryChartData  # noqa: E402
from pptx.dml.color import RGBColor  # noqa: E402
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION  # noqa: E402
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN  # noqa: E402
from pptx.util import Inches, Pt  # noqa: E402

import figuras_informe as FIGI  # noqa: E402
import informe_cierre as IC  # noqa: E402
import informe_evidencia as EV  # noqa: E402
import sintesis_hallazgos as S  # noqa: E402
from docx_mimp import comillas, fmt_num, lista_es, monto  # noqa: E402

SALIDA = f"{EVAL}/06_informe/2026_mimp-ogmepgd_presentacion-resultados-evaluacion-cierre-pnig.pptx"
A = f"{EVAL}/04_analisis/"
TEMA = {"D1": "Vigencia del diagnóstico", "D2": "Situación futura deseada", "D3": "Coherencia interna y evidencia",
        "D4": "Articulación con políticas y planes", "D5": "Calidad de los indicadores", "D6": "Enfoques transversales",
        "D7": "Nuevo marco conceptual", "I1": "Implementación de los servicios", "I2": "Avance de los indicadores",
        "I3": "Recursos de la conducción", "I4": "Rectoría y CONAIG", "I5": "Sostenibilidad ante la transición"}
q = lambda t: comillas(t)[0]


stat = K.tarjeta_cifra   # tarjeta con una cifra grande y su etiqueta (componente de la skill)


def caja(d, s, x, y, w, h, titulo, items, color, size=13):
    """Tarjeta con encabezado de color y viñetas (componente de la skill), con las comillas “ ”."""
    K.caja(d, s, x, y, w, h, titulo, [q(t) for t in items], color, size)


def main():
    C = S.cifras()
    H = {h["id"]: h for h in S.hallazgos(C)}
    V = S.valoraciones(C)
    n = len(C["grd"])
    g = pd.read_excel(A + "REV_DOC_estado_implementacion_servicios.xlsx", sheet_name="Estado por servicio", dtype=str)
    nombre_serv = dict(zip(g["Código"], g["Servicio"]))
    est = "Estado asignado (preliminar)"
    tip, mt, me, ref = EV.datos()
    sol, dire = mt.Solidez.value_counts(), mt["Dirección"].value_counts()
    rc = pd.read_excel(A + "REV_ALI_coherencia_alineamiento.xlsx", sheet_name="REV_ALI.5 Conceptual", dtype=str)
    efe = rc["Efecto"].value_counts()
    conc = rc[rc.Plano == "Conceptual"]["Tipo de relación"].value_counts()
    srv_c = rc[rc.Plano == "Servicios"]["Tipo de relación"].value_counts()
    corpus = len(pd.read_csv("docs/fuentes/inventario_fuentes.csv"))
    pdg = A + "REV_IND_3_presupuesto_dgignd.xlsx"
    comp = pd.read_excel(pdg, sheet_name="Composición (%)").set_index("Genérica")
    personas = [float(comp.loc[["Personal y CAS", "Locación de servicios de personas naturales"], f"Ejecución {a}"].sum())
                for a in range(2021, 2026)]
    conc_mef = pd.read_excel(pdg, sheet_name="Conciliación MEF")
    dev, anios = C["cond_dev"], sorted(C["cond_dev"])
    cs = {a: sum(dev[a][o] for o in ("DGIGND", "DPIGND", "DPPDM")) for a in anios}
    ejec = [C["dgignd_ejec"][a] for a in anios]
    ses = [x for x in FIGI.SESIONES_CONAIG if x[0] <= 2025]
    vp, vt = C["vert_periodo"], C["vert_transicion"]
    ef = C["enfoques"].nivel.astype(str).value_counts()
    sin_prot = n - C["serv_prot_si"]

    d = MimpLabDeck()
    # 1. portada
    d.cover("Evaluación de cierre de la PNIG (2019–2025)", "Resultados para validación",
            unidad="Oficina General de Monitoreo, Evaluación de Políticas y Gestión Descentralizada · OMEP",
            fecha="Noviembre de 2026")

    # 2. objeto, alcance y método
    s = d.content("Objeto, alcance y método", icon="investigacion", accent=TEAL)
    d.columns(s, 2.3, 3.9, [
        {"titulo": "Objeto", "icon": "objetivo", "items": [
            "Evaluación de cierre del diseño y la implementación de la PNIG.",
            "Periodo completo: 2019–2025.",
            "Insumo para formular la política que la sustituye, en el marco de la Ley N.° 32535."]},
        {"titulo": "Preguntas", "icon": "check", "items": [
            "12 preguntas: 7 de diseño y 5 de implementación.",
            "Criterios: pertinencia, coherencia, calidad técnica, eficacia, eficiencia, implementación y sostenibilidad.",
            "Cada pregunta se valora como satisfactoria, parcial o insuficiente."]},
        {"titulo": "Instrumentos y fuentes", "icon": "documento", "items": [
            "Seis instrumentos: revisión documental, entrevistas, reportes de las entidades, coherencia y alineamiento, "
            "indicadores y recursos, y revisión de evidencia.",
            f"{corpus} documentos, informes LIO 2019–2025, datos del MEF, microdatos del INEI y {C['rev_fuentes']} "
            f"estudios.",
            "Cada hallazgo lleva su nivel de confianza según sus fuentes independientes."]},
    ], body_size=13, header_size=12.5, header_h=0.62)

    # 3. mensajes principales
    s = d.content("Mensajes principales", icon="objetivo", accent=GOLD)
    d.icon_bullets(s, d.M, 2.4, d.CW, 4.4, [
        ("datos", TEAL, q(f"Las {C['brechas_n']} brechas del diagnóstico persisten al cierre: {len(C['brechas_reduce'])} "
                          f"se redujeron, {len(C['brechas_sincambio'])} no cambiaron y {len(C['brechas_amplia'])} se "
                          f"amplió.")),
        ("engranaje", GOLD, q(f"{C['serv_impl']} de los {n} servicios están implementados, {C['serv_parc']} lo están "
                              f"parcialmente y {C['serv_noimpl']} no se implementó; {sin_prot} siguen sin protocolo "
                              f"aprobado.")),
        ("crecimiento", GREEN, q(f"{C['op_mejora']} de los {C['op_ind']} indicadores de objetivo prioritario mejoran "
                                 f"frente a su línea de base, pero el reporte oficial no tiene valor 2025 para "
                                 f"{C['op_sinvalor_2025']} de ellos.")),
        ("colaboracion", VIOLET, q("La CONAIG sesionó con regularidad, pero su acuerdo más reiterado, culminar los "
                                   "protocolos pendientes, no se cumplió.")),
        ("idea", TEAL, q("El nuevo marco mantiene la estructura y la mayoría de los servicios; el cambio se concentra en "
                         "el plano conceptual.")),
        ("alerta", RED, q(f"Ninguna de las 12 preguntas alcanza una valoración satisfactoria: "
                          f"{sum(v[0] == 'Parcial' for v in V.values())} son parciales y "
                          f"{sum(v[0] == 'Insuficiente' for v in V.values())} insuficiente (calidad de los "
                          f"indicadores).")),
    ], size=14.5, gap=0.72)

    # 4. valoración de las preguntas
    s = d.content("Valoración de las doce preguntas", icon="check", accent=VIOLET)
    filas = []
    for p in S.PREGUNTAS:
        hs = [h for h in H.values() if h["pregunta"] == p]
        filas.append([p, TEMA[p], S.PREGUNTAS[p][0], V[p][0], K.resumen_confianza([h["confianza"] for h in hs])])
    K.tabla_valoracion(d, s, filas)

    # 5. divisor
    d.divider("01", "Diseño de la Política", accent=TEAL)

    # 6. problema público y brechas
    s = d.content("Problema público y situación futura (D1 y D2)", icon="datos", accent=GOLD)
    dep = C["dep_resumen"].set_index("Clave")
    act = int(dep.loc["tasa_actividad", "Brecha significativa en desventaja de las mujeres"])
    tot_act = int(dep.loc["tasa_actividad", "Ámbitos con estimación (último año)"])
    sfd = pd.read_excel(A + "D1_D2_brechas_situacion_futura.xlsx", sheet_name="Situación futura deseada")
    n_enun = sfd["Enunciado (cita literal)"].nunique()
    misma = int(sfd.drop_duplicates("Enunciado (cita literal)")["Correspondencia con el indicador del OP"]
                .eq("Misma").sum())
    w3 = (d.CW - 0.4) / 3
    stat(d, s, d.M, 2.3, w3, 1.75, str(C["brechas_n"]), "brechas del diagnóstico persisten al cierre", GOLD)
    stat(d, s, d.M + w3 + 0.2, 2.3, w3, 1.75, f"{act} de {tot_act}",
         "departamentos con brecha significativa en la tasa de actividad (2025)", TEAL)
    stat(d, s, d.M + 2 * (w3 + 0.2), 2.3, w3, 1.75, f"{misma} de {n_enun}",
         "enunciados de la situación futura se miden con su mismo indicador", VIOLET)
    d.icon_bullets(s, d.M, 4.4, d.CW, 2.4, [
        ("datos", GOLD, q(f"Desde 2019, {len(C['brechas_reduce'])} brechas se redujeron, {len(C['brechas_sincambio'])} no "
                          f"cambiaron y 1 se amplió: el trabajo no remunerado, en una comparación indicativa.")),
        ("crecimiento", TEAL, q(f"Los indicadores de objetivo mejoran sin cerrar las brechas: {C['op_mejora']} de "
                                f"{C['op_ind']} frente a su línea de base.")),
        ("alcance", VIOLET, q("Al ritmo observado, solo el embarazo adolescente alcanzaría su meta al 2030.")),
        ("documento", GREEN, q("Dos fuentes del diagnóstico cambiaron: el empleo informal de la ENAHO desde 2024 y las "
                               "preguntas de tolerancia de la ENARES.")),
    ], size=13.5, gap=0.58)

    # 7. coherencia y evidencia
    s = d.content("Coherencia interna y evidencia (D3)", icon="investigacion", accent=VIOLET)
    sin_ind = n - (n - C["serv_sin_ind_vig"])
    caja(d, s, d.M, 2.3, 5.55, 3.7, "Cadena causal", [
        f"{C['lin_total'] - len(C['lin_sin_serv'])} de {C['lin_total']} lineamientos tienen servicio; el "
        f"{lista_es(C['lin_sin_serv'])} no tiene ninguno.",
        f"{n - len(C['serv_sin_causa'])} de {n} servicios se vinculan con una causa o un efecto del modelo.",
        f"{sin_ind} de {n} servicios no tienen indicador vigente, porque no tienen protocolo aprobado.",
        f"El PEMIG asignó {C['pemig_ind']} indicadores, sobre todo de cobertura, y no tuvo seguimiento propio.",
    ], VIOLET)
    caja(d, s, d.M + 5.94, 2.3, 5.55, 3.7, f"Evidencia sobre los servicios ({C['rev_fuentes']} fuentes)", [
        f"{len(mt)} tipos de intervención: {int(sol.get('Sólida', 0))} con evidencia sólida, "
        f"{int(sol.get('Moderada', 0))} moderada y {int(sol.get('Limitada', 0))} limitada.",
        f"Dirección mixta en {int(dire.get('mixta', 0))} tipos: incluso los de respaldo sólido tienen estudios con "
        f"resultados nulos.",
        "La evidencia peruana es escasa y casi toda cuasiexperimental.",
        "Los tipos con evidencia limitada son los de menor respaldo para la nueva política.",
    ], TEAL)

    # 8. articulación, indicadores y enfoques
    s = d.content("Articulación, indicadores y enfoques (D4 a D6)", icon="colaboracion", accent=GREEN)
    d.columns(s, 2.3, 3.8, [
        {"titulo": "Articulación (D4)", "icon": "colaboracion", "items": [
            f"Periodo: {vp[0]} relaciones directas, {vp[1]} indirectas y {vp[2]} sin correspondencia.",
            f"Transición: las relaciones sin correspondencia suben a {vt[2]}, sobre todo con la PGG 2026–2031.",
            f"{C['hor_dup']} duplicidades con {C['hor_pol']} políticas nacionales afines."]},
        {"titulo": "Indicadores (D5): insuficiente", "icon": "alerta", "items": [
            f"{C['ficha_lb_nc']} de {C['fichas_total']} fichas no tienen línea de base con año y valor.",
            "Ninguna ficha prevé desagregación por territorio o grupo.",
            "Las 13 fichas de objetivo prioritario quedan observadas."]},
        {"titulo": "Enfoques (D6)", "icon": "personas", "items": [
            f"{int(ef.get('Operacionalizado', 0))} de 7 enfoques se operacionalizan; solo el de género con amplitud.",
            "Territorialidad y gestión por resultados están ausentes del diseño.",
            "Los protocolos incorporan los enfoques más que el diseño."]},
    ], body_size=13, header_size=12, header_h=0.62, palette=[GREEN, RED, VIOLET])

    # 9. nuevo marco conceptual
    s = d.content("Nuevo marco conceptual (D7)", icon="idea", accent=TEAL)
    w4 = (d.CW - 0.6) / 4
    for i, (num, lab, col) in enumerate([
            (str(len(rc)), "relaciones entre la PNIG y el nuevo marco", TEAL),
            (str(efe.get("Se mantiene", 0)), "elementos se mantienen", GREEN),
            (str(efe.get("Se ajusta", 0)), "elementos se ajustan", GOLD),
            (str(efe.get("Se rediseña", 0)), "elementos se rediseñan", VIOLET)]):
        stat(d, s, d.M + i * (w4 + 0.2), 2.3, w4, 1.6, num, lab, col)
    d.icon_bullets(s, d.M, 4.3, d.CW, 2.5, [
        ("datos", TEAL, q(f"Las divergencias se concentran en el plano conceptual: {conc.get('Divergencia', 0)} de "
                          f"{int(conc.sum())} conceptos.")),
        ("idea", GOLD, q("El enfoque de género y la noción de género como construcción social se reemplazan por el "
                         "enfoque de igualdad de oportunidades y la discriminación basada en el sexo.")),
        ("engranaje", GREEN, q(f"Los servicios se trasladan casi sin divergencias: {srv_c.get('Divergencia', 0)} en "
                               f"{int(srv_c.sum())} relaciones.")),
        ("documento", VIOLET, q("El cambio afecta más la justificación y el modelo del problema que el portafolio de "
                                "servicios.")),
    ], size=13.5, gap=0.6)

    # 10. divisor
    d.divider("02", "Implementación y resultados", accent=GOLD)

    # 11. estado de los servicios
    s = d.content("Estado de los servicios por objetivo prioritario (I1)", icon="engranaje", accent=GREEN)
    filas = []
    for op, x in g.groupby("OP"):
        e_ = x[est].value_counts()
        filas.append([op.replace("OP.0", "OP"), str(len(x)), str(int(e_.get("Implementado", 0))),
                      str(int(e_.get("Implementado parcialmente", 0))), str(int(e_.get("No implementado", 0))),
                      str(int((x["Protocolo aprobado"] == "Sí").sum()))])
    e_ = g[est].value_counts()
    filas.append(["Total", str(n), str(int(e_.get("Implementado", 0))), str(int(e_.get("Implementado parcialmente", 0))),
                  str(int(e_.get("No implementado", 0))), str(C["serv_prot_si"])])
    K.tabla_estado(d, s, filas)
    parc = g[g[est] == "Implementado parcialmente"]
    d.text(s, d.M, 6.05, 6.6, 0.75, "Implementado: al menos dos de tres componentes (protocolo aprobado, recursos "
           "identificados en el MEF e indicadores con valor) y ejecución en 2024 o 2025.", size=10.5, color=PETROL)
    d.icon_bullets(s, d.M + 6.95, 2.35, d.CW - 6.95, 4.5, [
        ("alerta", RED, q("El OP2 no tiene servicios implementados: ninguno tiene protocolo ni reporta indicadores.")),
        ("check", GOLD, q(f"Solo {int((parc['Protocolo aprobado'] == 'Sí').sum())} de los {len(parc)} servicios "
                          f"parciales tiene protocolo aprobado.")),
        ("dinero", TEAL, q("Los recursos se identificaron en los datos del MEF para 34 de 52 servicios.")),
        ("personas", VIOLET, q("Dificultades más reportadas: presupuesto, contexto externo y personal.")),
    ], size=13, gap=1.05)

    # 12. recursos de la conducción
    s = d.content("Recursos de la conducción de la Política (I3)", icon="S/", accent=GOLD)
    cd = CategoryChartData()
    cd.categories = [str(a) for a in anios]
    cd.add_series("Devengado de la conducción (millones de soles)", [round(cs[a] / 1e6, 2) for a in anios])
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(d.M), Inches(2.3), Inches(6.6), Inches(4.4), cd)
    ch = gf.chart
    ch.has_legend = False
    ch.has_title = True
    ch.chart_title.text_frame.text = "Devengado de las metas de la conducción (millones de soles)"
    tp = ch.chart_title.text_frame.paragraphs[0]
    tp.runs[0].font.size, tp.runs[0].font.bold, tp.runs[0].font.name = Pt(11), True, "Poppins"
    tp.runs[0].font.color.rgb = PETROL
    pl = ch.plots[0]
    pl.gap_width = 60
    pl.has_data_labels = True
    dl = pl.data_labels
    dl.number_format, dl.number_format_is_linked = '0.00', False
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    dl.font.size, dl.font.name, dl.font.color.rgb = Pt(10), "Poppins", PETROL
    pl.series[0].format.fill.solid()
    pl.series[0].format.fill.fore_color.rgb = RGBColor(0x7F, 0x99, 0xC8)
    va = ch.value_axis
    va.visible = False
    va.has_major_gridlines = False
    ca = ch.category_axis
    ca.tick_labels.font.size, ca.tick_labels.font.name = Pt(10), "Poppins"
    ca.tick_labels.font.color.rgb = PETROL
    ca.format.line.color.rgb = RGBColor(0xC3, 0xC2, 0xB7)
    pico = max(anios, key=lambda a: cs[a])
    d.icon_bullets(s, d.M + 6.95, 2.35, d.CW - 6.95, 4.5, [
        ("crecimiento", GOLD, q(f"Máximo en {pico}: {monto(cs[pico])}; {monto(cs[2024])} en 2024 y "
                                f"{monto(cs[2025])} en 2025.")),
        ("check", GREEN, q(f"La meta de la DGIGND ejecutó entre {fmt_num(min(ejec), 0)} % y {fmt_num(max(ejec), 0)} % "
                           f"de su PIM: la restricción fue el monto asignado.")),
        ("personas", TEAL, q(f"Entre {fmt_num(min(personas), 0)} % y {fmt_num(max(personas), 0)} % del gasto de "
                             f"2021–2025 se destinó a personal, CAS y locación de servicios.")),
        ("datos", VIOLET, q(f"Las cifras del MEF coinciden al sol con las de la DGIGND en "
                            f"{int((conc_mef.Coincide == 'Sí').sum())} de {len(conc_mef)} combinaciones de órgano y "
                            f"año.")),
    ], size=13, gap=1.05)

    # 13. CONAIG
    s = d.content("CONAIG y articulación (I4)", icon="colaboracion", accent=VIOLET)
    d.steps(s, 2.3, 1.75, [(str(a), f"{len(f)} sesiones ordinarias\nAcuerdo: culminar los protocolos pendientes",
                            "calendario") for a, f, _ in ses], size=12.5, header_h=0.6)
    n_ses = sum(len(f) for _, f, _ in ses)
    d.icon_bullets(s, d.M, 4.4, d.CW, 2.4, [
        ("check", TEAL, q(f"{n_ses} sesiones ordinarias entre 2021 y 2025; desde 2022, el mínimo de tres por año de su "
                          f"reglamento.")),
        ("alerta", RED, q(f"El acuerdo de culminar los protocolos se repitió cada año, y {sin_prot} de {n} servicios "
                          f"siguen sin protocolo.")),
        ("personas", VIOLET, q("Representación mayoritaria de alternos, asistencia variable y sesiones sin quórum.")),
        ("colaboracion", GREEN, q("Baja adopción de sus recomendaciones de gobernanza: solo 2 de 16 servicios las "
                                  "incorporaron.")),
    ], size=13.5, gap=0.58)

    # 14. indicadores de OP en 2025
    s = d.content("Indicadores de objetivo prioritario en 2025 (I2)", icon="crecimiento", accent=GREEN)
    est25 = C["op_est25"].set_index("Código")
    filas = []
    for _, r in C["op_tabla"].iterrows():
        s25 = str(r["Semáforo 2025"])
        if s25.startswith(("Alto", "Medio", "Bajo")) or r["Código"] not in est25.index:
            nivel = s25
        else:
            e = est25.loc[r["Código"]]
            nivel = f"{e.Nivel}*"
        filas.append([r["Código"], q(str(r["Indicador"])), nivel, str(r["Respecto de la LB"])])
    d.table(s, d.M, 2.25, d.CW, 4.35, ["Código", "Indicador", "Nivel 2025", "Frente a la línea de base"], filas,
            [1.15, 6.2, 2.55, 1.59], body_size=9, header_size=10, row_h=0.3,
            align_cols=[PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.CENTER])
    d.text(s, d.M, 6.72, d.CW, 0.3, "* Nivel calculado en esta evaluación con microdatos del INEI o el método de la "
           "ficha, porque el reporte oficial no tiene valor 2025; no sustituye al reporte oficial.", size=8.5,
           color=PETROL)

    # 15. sostenibilidad y transición
    s = d.content("Sostenibilidad ante la transición (I5)", icon="alerta", accent=RED)
    d.columns(s, 2.3, 3.6, [
        {"titulo": "Capacidades", "icon": "personas", "items": [
            "La rotación de personal es la brecha de capacidad más persistente.",
            "La asistencia técnica nacional de la DTEG bajó de 21 a 10 entidades en 2025."]},
        {"titulo": "Riesgos de la transición", "icon": "alerta", "items": [
            "Un protocolo quedó a la espera de las acciones derivadas de la Ley N.° 32535.",
            "Normas sectoriales frenaron protocolos del Minedu.",
            "La formulación de la nueva política corre contra el plazo legal."]},
        {"titulo": "Ajustes propuestos", "icon": "idea", "items": [
            "18 segmentos de las entidades proponen ajustar un servicio y 6, un indicador.",
            "COFOPRI pidió evaluar el retiro del servicio 4.2.2.",
            "El Programa Juntos señala que su contribución al servicio 4.4.1 es indirecta."]},
    ], body_size=13, header_size=12.5, header_h=0.62, palette=[TEAL, RED, GOLD])

    # 16. recomendaciones
    s = d.content("Recomendaciones", icon="documento", accent=TEAL)
    grupos = {}
    for r in IC.recomendaciones(C, nombre_serv):
        tema, sub = r[0].split(":", 1)
        grupos.setdefault(tema.strip(), []).append(sub.strip()[0].upper() + sub.strip()[1:])
    cols = [{"titulo": f"{t} ({len(v)})", "icon": {"Diseño": "idea", "Implementación": "engranaje"}.get(t, "datos"),
             "items": v} for t, v in grupos.items()]
    d.columns(s, 2.3, 3.7, cols, body_size=13.5, header_size=12.5, header_h=0.62)
    d.text(s, d.M, 6.15, d.CW, 0.3, f"{sum(len(v) for v in grupos.values())} recomendaciones, cada una con su "
           f"responsable, su plazo y las preguntas que la sustentan (informe, capítulo 7).", size=10.5, color=PETROL)

    # 17. próximos pasos
    s = d.content("Próximos pasos", icon="calendario", accent=GOLD)
    d.steps(s, 2.4, 2.7, [
        ("Días 25 a 28", "13 a 16 de noviembre\nValidación con la DGIGND y el Despacho Viceministerial de la Mujer",
         "personas"),
        ("Días 26 a 29", "14 a 17 de noviembre\nIncorporación de las observaciones", "documento"),
        ("Día 29", "17 de noviembre\nReunión técnica con el CEPLAN", "colaboracion"),
        ("Día 30", "18 de noviembre\nEntrega del informe final y del anexo técnico a la DGIGND", "check"),
        ("Desde el día 30", "Validación por la DGIGND y la OGPPM, y remisión al CEPLAN por la OGPPM", "engranaje"),
    ], size=12.5, header_h=0.6)
    d.text(s, d.M, 5.45, d.CW, 0.9, [[("Productos: ", True), ("informe final de evaluación de cierre, anexo técnico y "
           "actas de las reuniones de validación.", False)]], size=13, color=PETROL)

    # 18. cierre
    d.closing("Gracias", pie="OGMEPGD · OMEP · Noviembre de 2026")
    d.save(SALIDA)
    print("Presentación:", SALIDA, f"({len(d.prs.slides)} láminas)")


if __name__ == "__main__":
    main()
