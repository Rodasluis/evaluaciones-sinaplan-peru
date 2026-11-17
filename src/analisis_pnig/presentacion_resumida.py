"""Presentación resumida de los resultados de la evaluación de cierre de la PNIG (2019–2025): ocho láminas con el
estilo del MIMP Lab, para exposiciones breves y para leerse sin expositor (pedidos del usuario del 2026-11-17,
ADR-016).

Cada lámina de contenido abre con una bajada (K.bajada) que dice qué muestra y qué concluye. La evaluación se presenta
como culminada: sin próximos pasos (informe validado por la DGIGND el 2026-11-17).

Las cifras se leen de las mismas fuentes que el informe (sintesis_hallazgos.cifras(), hallazgos, valoraciones y libros
de 04_analisis), para que coincidan con el informe, la presentación completa y la infografía. Como la presentación
completa, el script sirve para la primera versión: después, la presentación se edita a mano sobre el archivo del
repositorio (CLAUDE.md), no se regenera.

Uso (desde la raíz del repositorio):
    uv run python src/analisis_pnig/presentacion_resumida.py
Salida: evaluaciones/pnig_2026/06_informe/2026_mimp-ogmepgd_presentacion-resumida-resultados-evaluacion-cierre-pnig.pptx
"""
import os
import sys
from pathlib import Path

import pandas as pd
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz

sys.path.insert(0, os.path.dirname(__file__))
_SKILL = Path(__file__).resolve().parents[2] / ".claude" / "skills" / "presentacion-evaluacion" / "scripts"
sys.path.insert(0, str(_SKILL))

import componentes as K  # noqa: E402
from mimp_lab_deck import GOLD, GREEN, PETROL, RED, TEAL, VIOLET, MimpLabDeck  # noqa: E402

import informe_cierre as IC  # noqa: E402
import sintesis_hallazgos as S  # noqa: E402
from docx_mimp import comillas, fmt_num  # noqa: E402
from presentacion_resultados import TEMA  # noqa: E402

SALIDA = f"{EVAL}/06_informe/2026_mimp-ogmepgd_presentacion-resumida-resultados-evaluacion-cierre-pnig.pptx"
q = lambda t: comillas(t)[0]
# explicación breve de cada recomendación (resumen de su texto en el informe, capítulo 7)
EXPLICA = {
    "temas vigentes": "retomar las brechas que persisten",
    "enfoques": "operacionalizar los que la PNIG solo mencionó",
    "coherencia": "cada lineamiento con servicio, protocolo e indicadores",
    "evidencia": "priorizar los servicios con evidencia sólida",
    "priorización": "acotar problema y servicios a lo factible",
    "articulación horizontal": "ordenar los servicios compartidos con otras políticas",
    "alineamiento": "explicitar el vínculo con la PGG 2026–2031 y el PESEM",
    "indicadores": "fichas con línea de base y desagregación",
    "compromisos": "matriz de compromisos con las entidades",
    "gestión del conocimiento": "conservar la documentación de la PNIG",
    "componentes a preservar": "mantener los servicios implementados",
    "transición": "definir el destino de protocolos en trámite y servicios no implementados",
    "transición normativa": "usar la matriz de incompatibilidades e informar al CEPLAN y a las entidades",
    "asignación de recursos": "identificar el gasto de cada servicio",
    "coordinación": "un mecanismo que asuma las funciones de la CONAIG",
    "capacidades": "retener capacidades ante la rotación de personal",
}


def main():
    C = S.cifras()
    H = S.hallazgos(C)
    V = S.valoraciones(C)
    g = C["grd"]
    n = len(g)
    est = "Estado asignado (preliminar)"
    nombre_serv = dict(zip(g["Código"].map(str), g["Servicio"]))
    documentos = len(pd.read_csv("docs/fuentes/inventario_fuentes.csv"))
    asign = int((g["Asignación presupuestal identificable"] == "Sí").sum())
    sin_prot = n - C["serv_prot_si"]
    mc = C["mic_conc"]
    disp = int(C["mic_norma"]["Disposiciones revisadas"].sum())
    parc = sum(v[0] == "Parcial" for v in V.values())
    insu = sum(v[0] == "Insuficiente" for v in V.values())
    conf = pd.Series([h["confianza"] for h in H]).value_counts()
    w4 = (d_cw := 11.49 - 0.6) / 4

    d = MimpLabDeck()
    # 1. portada
    d.cover("Evaluación de cierre de la PNIG (2019–2025)", "Resultados de la evaluación · versión resumida",
            unidad="Oficina General de Monitoreo, Evaluación de Políticas y Gestión Descentralizada · OMEP",
            fecha="Noviembre de 2026")

    # 2. qué se evaluó y cómo
    s = d.content("Qué se evaluó y cómo", icon="investigacion", accent=TEAL)
    y = K.bajada(d, s, q("Se evaluó el diseño y la implementación de la Política Nacional de Igualdad de Género en todo "
                         "su periodo, de 2019 a 2025."))
    for i, (num, etq, col) in enumerate([
            ("12", "preguntas de evaluación: 7 sobre el diseño y 5 sobre la implementación", TEAL),
            (str(n), "servicios de los 6 objetivos prioritarios, a cargo de entidades de los tres niveles de gobierno",
             GOLD),
            (str(documentos), f"documentos revisados (normas, planes e informes) y {C['rev_fuentes']} estudios sobre "
                              f"los servicios", VIOLET),
            (str(len(H)), f"hallazgos: {conf.get('Sólido', 0)} sólidos, {conf.get('Con respaldo parcial', 0)} con "
                          f"respaldo parcial y {conf.get('Indicio', 0)} indicios", GREEN)]):
        K.tarjeta_cifra(d, s, d.M + i * (w4 + 0.2), y, w4, 1.95, num, etq, col)
    d.icon_bullets(s, d.M, y + 2.2, d.CW, 2.0, [
        ("documento", GOLD, q("Seis instrumentos: revisión documental, entrevistas a la OGPPM, la OGMEPGD y la DGIGND, "
                              "reportes de las entidades, coherencia y alineamiento, indicadores y recursos, y revisión "
                              "de evidencia.")),
        ("check", VIOLET, q("Cada hallazgo indica su confianza: sólido si lo respaldan dos o más fuentes de origen "
                            "independiente; con respaldo parcial si coinciden con discrepancias; indicio si tiene un "
                            "solo origen.")),
    ], size=12.5, gap=0.75)

    # 3. mensajes principales
    s = d.content("Mensajes principales", icon="objetivo", accent=GOLD)
    y = K.bajada(d, s, q("La PNIG respondía a brechas que persisten, pero se implementó en parte y perdió vigencia con "
                         "la Ley N.° 32535."))
    for i, (num, etq, col) in enumerate([
            (f"{C['serv_impl']} de {n}", f"servicios implementados; {C['serv_parc']} lo están en parte y "
                                         f"{C['serv_noimpl']} no se implementaron", GOLD),
            (str(C["brechas_n"]), f"brechas del diagnóstico persisten: {len(C['brechas_reduce'])} se redujeron y "
                                  f"{len(C['brechas_amplia'])} se amplió", TEAL),
            (f"{C['op_mejora']} de {C['op_ind']}", "indicadores de objetivo prioritario mejoran frente a su línea de "
                                                   "base", GREEN),
            ("0 de 12", f"preguntas alcanzan una valoración satisfactoria: {parc} parciales y {insu} insuficiente",
             RED)]):
        K.tarjeta_cifra(d, s, d.M + i * (w4 + 0.2), y, w4, 1.95, num, etq, col)
    d.icon_bullets(s, d.M, y + 2.2, d.CW, 2.1, [
        ("alerta", RED, q(f"{sin_prot} de los {n} servicios siguen sin protocolo aprobado, aunque la comisión de alto "
                          f"nivel (CONAIG) acordó cada año culminarlos.")),
        ("crecimiento", GREEN, q(f"El reporte oficial no tiene valor 2025 para {C['op_sinvalor_2025']} de los "
                                 f"{C['op_ind']} indicadores; esta evaluación estimó los que pudo con microdatos.")),
        ("idea", VIOLET, q("Según la OGAJ, la PNIG perdió vigencia el 25 de diciembre de 2025; la mayoría de sus "
                           "servicios puede continuar con sustento en la Ley N.° 32535.")),
    ], size=12.5, gap=0.62)

    # 4. valoración de las preguntas
    s = d.content("Valoración de las doce preguntas", icon="check", accent=VIOLET)
    y = K.bajada(d, s, q("Satisfactoria: el atributo se cumple; parcial: hay avances con vacíos; insuficiente: no se "
                         "cumple en lo crítico."))
    filas = []
    for p in S.PREGUNTAS:
        hs = [h for h in H if h["pregunta"] == p]
        filas.append([p, TEMA[p], S.PREGUNTAS[p][0], V[p][0], K.resumen_confianza([h["confianza"] for h in hs])])
    K.tabla_valoracion(d, s, filas, y=y, h=4.2)

    # 5. servicios y transición normativa
    s = d.content("Servicios y transición normativa", icon="engranaje", accent=GREEN)
    y = K.bajada(d, s, q(f"{C['serv_impl']} de los {n} servicios se implementaron; la mayoría puede continuar con "
                         f"sustento en la nueva ley."))
    filas = []
    for op, x in g.groupby("OP"):
        e_ = x[est].value_counts()
        filas.append([op.replace("OP.0", "OP"), str(len(x)), str(int(e_.get("Implementado", 0))),
                      str(int(e_.get("Implementado parcialmente", 0))), str(int(e_.get("No implementado", 0))),
                      str(int((x["Protocolo aprobado"] == "Sí").sum()))])
    e_ = g[est].value_counts()
    filas.append(["Total", str(n), str(int(e_.get("Implementado", 0))), str(int(e_.get("Implementado parcialmente", 0))),
                  str(int(e_.get("No implementado", 0))), str(C["serv_prot_si"])])
    K.tabla_estado(d, s, filas, y=y, h=3.3)
    d.text(s, d.M, y + 3.42, 6.6, 0.95, q(f"OP: objetivo prioritario. Implementado: al menos dos de tres componentes "
           f"(protocolo aprobado, recursos identificados en el MEF, que se hallaron para {asign} de {n} servicios, e "
           f"indicadores con valor) y ejecución en 2024 o 2025."), size=10, color=PETROL)
    K.caja(d, s, d.M + 6.95, y, d.CW - 6.95, 4.3, "Transición normativa", [q(t) for t in [
        "Según la OGAJ, la PNIG, el PEMIG y la CONAIG perdieron vigencia el 25 de diciembre de 2025.",
        f"La matriz de la DPIGND marca como incompatibles las {disp} disposiciones revisadas de esas normas.",
        f"La misma matriz coincide con esta evaluación en {int(mc['Coinciden'])} de {int(mc['Servicios'])} servicios: "
        f"la mayoría continúa con base en la Ley N.° 32535.",
        "No hay vacío normativo, pero los mecanismos y Oficiales para la igualdad quedan sin mandato.",
    ]], RED, size=11.5)

    # 6. recomendaciones
    s = d.content("Recomendaciones", icon="documento", accent=TEAL)
    recs = IC.recomendaciones(C, nombre_serv)
    y = K.bajada(d, s, q(f"{len(recs)} recomendaciones para la nueva política y la transición, cada una con su "
                         f"responsable y su plazo (informe, capítulo 7)."))
    grupos = {}
    for r in recs:
        tema, sub = r[0].split(":", 1)
        sub = sub.strip()
        grupos.setdefault(tema.strip(), []).append(f"{sub[0].upper()}{sub[1:]}: {EXPLICA[sub.lower()]}.")
    cols = [{"titulo": f"{t} ({len(v)})", "icon": {"Diseño": "idea", "Implementación": "engranaje"}.get(t, "datos"),
             "items": [q(x) for x in v]} for t, v in grupos.items()]
    d.columns(s, y, 4.25, cols, body_size=10.5, header_size=12, header_h=0.55)

    # 7. qué deja la evaluación
    s = d.content("Qué deja la evaluación para la nueva política", icon="idea", accent=VIOLET)
    y = K.bajada(d, s, q("La nueva política parte de un problema vigente y de servicios que pueden continuar, con "
                         "cambios en su marco y en su gestión."))
    w2 = (d.CW - 0.3) / 2
    for i, (titulo, texto, col) in enumerate([
            ("Un diagnóstico vigente", f"Las {C['brechas_n']} brechas que motivaron la PNIG persisten: el problema "
                                       f"sigue presente y tendrá que explicarse con el enfoque de la Ley N.° 32535.",
             TEAL),
            ("Servicios que pueden continuar", f"{C['serv_impl']} servicios están implementados y la matriz de la "
                                               f"DPIGND coincide con esta evaluación en {int(mc['Coinciden'])} de "
                                               f"{int(mc['Servicios'])}: la mayoría tiene sustento en la nueva ley.",
             GREEN),
            ("Lo que hay que fortalecer", f"Los protocolos ({sin_prot} servicios no los tienen), los indicadores con "
                                          f"línea de base y desagregación, y un mecanismo de articulación que "
                                          f"reemplace a la CONAIG.", GOLD),
            ("Una transición ordenada", "Comunicar a las entidades qué sigue vigente y qué no, y comunicar la "
                                        "derogación tácita de los decretos conforme a la Ley N.° 31772.", RED)]):
        x = d.M + (i % 2) * (w2 + 0.3)
        yy = y + (i // 2) * 2.15
        K.caja(d, s, x, yy, w2, 1.95, titulo, [q(texto)], col, size=12)

    # 8. cierre
    d.closing("Gracias", pie="OGMEPGD · OMEP · Noviembre de 2026")
    d.save(SALIDA)
    print("Presentación resumida:", SALIDA, f"({len(d.prs.slides)} láminas); acuerdo con la matriz "
          f"{fmt_num(mc['Acuerdo simple (%)'], 1)} %; ancho de tarjeta {w4:.2f} in sobre {d_cw:.2f}")


if __name__ == "__main__":
    main()
