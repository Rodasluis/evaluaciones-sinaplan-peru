"""Sección 3.7 del informe: seguimiento de las recomendaciones de evaluaciones anteriores de la PNIG.

Atiende las preguntas p3.4 (nivel de avance de las recomendaciones de la evaluación anterior), p3.5 (cómo se vienen
ejecutando) y p3.6 (qué medidas adicionales son necesarias) de la Tabla A7.1 de la Guía de seguimiento y evaluación
del CEPLAN (2024, p. 185). La nota metodológica no tiene una pregunta propia para ellas: el equipo evaluador las trata
dentro de I4 (conducción y rectoría), con REV_DOC (revisión documental) y ENTREV (contraste).

Paso 1 (extracción, reproducible): lee las recomendaciones de los cuatro informes anteriores, con su página del PDF, y
las guarda en 04_analisis/seguimiento_recomendaciones/recomendaciones_previas.csv:
  - MIMP, informe de evaluación de implementación 2020 (junio de 2021), sección 5 (pp. 116 a 119 del PDF), con el
    destinatario de cada recomendación (MIMP, entidades, MTPE, INEI, MINEDU, SERVIR, CEPLAN); se incorporó el
    2026-11-06, por decisión del usuario, después de identificarla en gob.pe;
  - MIMP, informe de evaluación de resultados de la PNIG 2021 (mayo de 2022), sección 5 (pp. 42 y 43 del PDF);
  - MIMP, informe de evaluación de la PNIG 2023 (diciembre de 2024), sección 7 (pp. 66 y 67);
  - CEPLAN, informe técnico de evaluación de la PNIG (sin fecha; la coordinación se hizo en el primer semestre de
    2024 y el PDF es del 16-01-2025): recomendaciones por servicio del OP1 en el análisis de diseño (sección 2.2) y de
    implementación (sección 3.2), y generales (sección 5, p. 70).
  La capa de texto del PDF del CEPLAN codifica la «f» y las ligaduras con caracteres de uso privado y pierde la «l»
  de la ligadura «fl»; limpiar() restituye las primeras y FL restituye la «l» en las palabras de las recomendaciones
  (verificado en la imagen de las páginas).
Paso 2 (valoración del equipo evaluador): seguimiento_recomendaciones/valoracion_recomendaciones.csv asigna a cada
recomendación su estado al corte, cómo se ejecutó (p3.5), la medida adicional o lo que se traslada a la nueva
política (p3.6), el medio de verificación con una cita que el script comprueba (en la página del PDF, en el segmento
de REP_ENT o de ENTREV, o en la hoja de un libro de análisis), y la pregunta de ENTREV que la contrasta. Las recomendaciones de protocolo valoradas «Implementada» se comprueban además con REV_DOC:
el protocolo debe estar aprobado con una norma posterior al informe que lo recomendó.
Escala (equipo evaluador): Implementada; Parcialmente implementada; No implementada; Sin evidencia (no consta medio de
verificación; se consulta en ENTREV y no se presume); Sin objeto por la transición (solo con sustento expreso).

Uso (desde la raíz del repositorio):
    uv run --no-project --with pandas --with openpyxl python -I -X utf8 src/analisis_pnig/seguimiento_recomendaciones.py
Salida: evaluaciones/pnig_2026/04_analisis/seguimiento_recomendaciones_previas.xlsx
"""
import os
import re
import subprocess
from datetime import date

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

PREVIAS = "docs/fuentes/pnig/05_evaluaciones_previas"
INFORMES = {  # evaluación: (archivo, fecha de referencia para el seguimiento, cita corta)
    "MIMP 2020": (f"{PREVIAS}/2020_mimp_informe-evaluacion-implementacion-pnig.pdf", date(2021, 6, 17), "MIMP (2021)"),
    "MIMP 2021": (f"{PREVIAS}/2021_mimp_informe-evaluacion-resultados-pnig.pdf", date(2022, 5, 31), "MIMP (2022)"),
    "MIMP 2023": (f"{PREVIAS}/2023_mimp_informe-evaluacion-pnig.pdf", date(2024, 12, 16), "MIMP (2024)"),
    "CEPLAN": (f"{PREVIAS}/2023_ceplan_informe-tecnico-evaluacion-pnig.pdf", date(2025, 1, 16), "CEPLAN (s. f.)"),
}
D = f"{EVAL}/04_analisis/seguimiento_recomendaciones"
GRD = f"{EVAL}/04_analisis/REV_DOC_estado_implementacion_servicios.xlsx"
SALIDA = f"{EVAL}/04_analisis/seguimiento_recomendaciones_previas.xlsx"
CORTE = "2026-11-07"
ESTADOS = ("Implementada", "Parcialmente implementada", "No implementada", "Sin evidencia",
           "Sin objeto por la transición")
ASPECTOS = ("Diseño", "Implementación", "Seguimiento y evaluación")  # Tabla 5 de la Guía (p. 57)
PUA = {"": "f", "": "f", "": "f", "": "f", "": "fi", "": "ff",
       "ﬁ": "fi", "ﬂ": "fl", "ﬀ": "ff"}
FL = {"infuiría": "influiría", "refeje": "refleje", "refejen": "reflejen", "refeja": "refleja",
      "refejan": "reflejan", "refejar": "reflejar", "confictos": "conflictos", "fexible": "flexible",
      "fuida": "fluida", "fuido": "fluido", "fujo": "flujo", "fujos": "flujos", "refexión": "reflexión",
      "refexionó": "reflexionó"}


def limpiar(t):
    for k, v in PUA.items():
        t = t.replace(k, v)
    return re.sub(r"\s+", " ", t).strip()


def restituir_fl(t):
    return re.sub(r"\w+", lambda m: FL.get(m.group(0), m.group(0)), t)


def pagina(ruta, p, layout=False):
    t = subprocess.run(["pdftotext", "-enc", "UTF-8"] + (["-layout"] if layout else []) +
                       ["-f", str(p), "-l", str(p), ruta, "-"], capture_output=True).stdout.decode("utf-8", "replace")
    return t


def n_paginas(ruta):
    out = subprocess.run(["pdfinfo", ruta], capture_output=True).stdout.decode("utf-8", "replace")
    return int(re.search(r"Pages:\s+(\d+)", out).group(1))


def lineas(ruta, desde, hasta, descartar, layout=False):
    """(página, línea) del rango, sin encabezados, pies ni líneas vacías."""
    for p in range(desde, hasta + 1):
        for ln in pagina(ruta, p, layout).splitlines():
            s = limpiar(ln)
            if s and not any(re.search(x, s) for x in descartar):
                yield p, s


def vinetas(items):
    """Agrupa (página, línea) en viñetas «•»; devuelve (página inicial, texto). Una viñeta que empieza con «De lo
    contrario» es la segunda rama de la anterior (CEPLAN, servicio 1.2.6) y se une a ella."""
    out = []
    for p, s in items:
        if s.startswith("•") and not s.lstrip("• ").startswith("De lo contrario"):
            out.append([p, s.lstrip("• ").strip()])
        elif out:
            out[-1][1] += " " + s.lstrip("• ")
    return [(p, re.sub(r"\s+", " ", t)) for p, t in out]


def extraer():
    filas = []

    def add(ev, n, p, texto, seccion, servicio="", responsable="", aspecto="", plazo=""):
        texto = restituir_fl(texto)
        filas.append({"id": f"{ {'MIMP 2020': 'EV2020', 'MIMP 2021': 'EV2021', 'MIMP 2023': 'EV2023', 'CEPLAN': 'CEPLAN'}[ev]}-{n:02d}",
                      "evaluacion": ev, "pagina_pdf": p, "seccion": seccion, "servicio": servicio,
                      "responsable_informe": responsable, "aspecto_informe": aspecto, "plazo_informe": plazo,
                      "texto": texto})

    # MIMP 2020: sección 5 (pp. 116 a 119), viñetas «- » bajo encabezados de destinatario y de tema (con -layout, para
    # que cada encabezado quede en su línea); el tema «Sobre la asistencia técnica …» ocupa dos líneas
    ruta = INFORMES["MIMP 2020"][0]
    destinatarios = {"Para el Ministerio de la Mujer y Poblaciones Vulnerables": "MIMP",
                     "Para las entidades intervinientes": "",
                     "A todas las entidades": "Todas las entidades intervinientes",
                     "Al Ministerio del Trabajo y Promoción del Empleo": "MTPE",
                     "Al Instituto Nacional de Estadística e Informática": "INEI",
                     "Al Ministerio de Educación": "MINEDU",
                     "A la Autoridad Nacional del Servicio Civil": "SERVIR",
                     "Al Ministerio de la Mujer y Poblaciones Vulnerables": "MIMP (como entidad interviniente)",
                     "Para el CEPLAN": "CEPLAN"}
    resp, tema, items, dentro = "", "", [], False
    for p, s in lineas(ruta, 116, 119, [r"^Política Nacional Multisectorial de Igualdad de Género$", r"^P á g i n a",
                                        r"^35 Aprobado mediante"], layout=True):
        if s == "5. Recomendaciones":
            dentro = True
        elif not dentro:
            continue
        elif s in destinatarios:
            resp, tema = destinatarios[s], ""
        elif s.startswith("Sobre "):
            tema = s
        elif s == "equivalentes" and tema.endswith("normas"):
            tema += " equivalentes"
        elif s.startswith("- "):
            items.append([p, s[2:], resp, tema])
        elif items:
            items[-1][1] += " " + s
    assert len(items) == 23, f"MIMP 2020: {len(items)} recomendaciones"
    for n, (p, t, r_, tm) in enumerate(items, 1):
        add("MIMP 2020", n, p, re.sub(r"equivalentes35\.", "equivalentes.", t),
            "5. Recomendaciones" + (f". {tm}" if tm else ""), responsable=r_)

    # MIMP 2021: sección 5, de la p. 42 hasta «Anexos» (p. 43); fuera, el encabezado ajeno, la paginación y la nota 6
    ruta = INFORMES["MIMP 2021"][0]
    items, dentro = [], False
    for p, s in lineas(ruta, 42, 43, [r"^Política Nacional Multisectorial en Discapacidad", r"^\d{1,2}$",
                                      r"^6 Estos servicios son"]):
        if s == "5. Recomendaciones":
            dentro = True
            continue
        if s.startswith("Anexos"):
            break
        if dentro:
            items.append((p, s))
    for n, (p, t) in enumerate(vinetas(items), 1):
        add("MIMP 2021", n, p, re.sub(r"(superen)6\.", r"\1.", t), "5. Recomendaciones")

    # MIMP 2023: sección 7, bloques «Unidad operativa/ Responsable», «Aspecto de mejora», «Recomendación:», «Plazo»
    # (con -layout, para que cada fila de la tabla quede en una línea)
    ruta = INFORMES["MIMP 2023"][0]
    bloque, n = None, 0
    for p, s in lineas(ruta, 66, 67, [r"^\d{2}$"], layout=True):
        if s.startswith("Unidad operativa/ Responsable"):
            bloque = {"p": p, "resp": s.split("Responsable", 1)[1].strip(), "texto": ""}
        elif bloque is None or s == "7. Recomendaciones":
            continue
        elif s.startswith("Aspecto de mejora"):
            bloque["aspecto"] = s.replace("Aspecto de mejora", "").strip()
        elif s.startswith("Recomendación:"):
            bloque["en_texto"] = True
        elif s.startswith("Plazo"):
            n += 1
            add("MIMP 2023", n, bloque["p"], bloque["texto"].strip(), "7. Recomendaciones",
                responsable=bloque["resp"], aspecto=bloque["aspecto"], plazo=s.replace("Plazo", "").strip())
            bloque = None
        elif bloque.get("en_texto"):
            bloque["texto"] += " " + s

    # CEPLAN: «Recomendaciones» de cada servicio (secciones 2.2 y 3.2) y sección 5
    ruta = INFORMES["CEPLAN"][0]
    fin = re.compile(r"^(Servicio \d\.\d\.\d|\d\.\d\. [A-ZÁÉÍÓÚ]|\d\. [A-ZÁÉÍÓÚ]|Recomendaciones$|Acciones relevantes|"
                     r"Análisis del )")
    seccion, servicio, items, n = "", "", None, 0

    def cerrar():
        nonlocal n
        for p, t in vinetas(items or []):
            n += 1
            add("CEPLAN", n, p, t, seccion, servicio=servicio)

    for p, s in lineas(ruta, 20, 70, [r"^Informe Técnico\. Evaluación de la Política", r"^Página \d+$"]):
        if items is not None and fin.match(s):
            cerrar()
            items = None
        m = re.match(r"^Servicio (\d\.\d\.\d)", s)
        if m:
            servicio = m.group(1)
        elif re.match(r"^(2\.2|3\.2)\. ", s):
            seccion = s.split(" ", 1)[0].rstrip(".") + " " + ("Diseño de los servicios" if s.startswith("2.2")
                                                             else "Implementación de los servicios priorizados")
        elif s == "5. Recomendaciones":
            seccion, servicio, items = "5 Recomendaciones generales", "", []
        elif s == "Recomendaciones":
            items = []
        elif items is not None:
            items.append((p, s))
        if s.startswith("6. Re"):
            break
    if items:
        cerrar()
    return pd.DataFrame(filas)


def verificar_textos(R):
    """Cada recomendación consta en su página (o empieza en ella y sigue en la siguiente)."""
    cache = {}
    for _, r in R.iterrows():
        ruta = INFORMES[r.evaluacion][0]
        clave = (ruta, r.pagina_pdf)
        if clave not in cache:
            cache[clave] = restituir_fl(limpiar(pagina(ruta, r.pagina_pdf) + " " + pagina(ruta, r.pagina_pdf + 1)))
        inicio = " ".join(r.texto.split()[:12])
        assert inicio in cache[clave].replace("• ", ""), f"{r.id}: el texto no consta en la p. {r.pagina_pdf}"


FIE = f"{EVAL}/04_analisis/REP_ENT_informacion_entidades.xlsx"
PSI = f"{EVAL}/04_analisis/REV_IND_base_indicadores.xlsx"
# nombre de los libros de análisis citados como medio de verificación (el texto del informe usa nombres; ADR-009)
LIBROS = {"REV_IND_2_calidad_fichas.xlsx": "Revisión de la calidad de las fichas",
          "REV_ALI_coherencia_alineamiento.xlsx": "Análisis de coherencia y alineamiento"}
NOMBRES = {  # nombre corto de los documentos citados como medio de verificación
    "2024_mimp_reporte-cumplimiento-pnig.pdf": "Reporte de cumplimiento de la PNIG 2024",
    "2025_mimp_reporte-cumplimiento-pnig.pdf": "Reporte de cumplimiento de la PNIG 2025",
    "2026_mimp-dpignd_informe-d000039-2026_ruta-actualizacion-pnig-rpcd-042-2026.pdf":
        "Informe N.° D000039-2026-MIMP-DPIGND",
    "2026_mimp-op_informe-d000112-2026_rutas-actualizacion-indicadores-rpcd-042-2026.pdf":
        "Informe N.° D000112-2026-MIMP-OP",
    "2025_congreso_ley-32535_igualdad-oportunidades-mujeres-hombres.pdf": "Ley N.° 32535",
    "2025_inpe_rp-522-2025-inpe-p_servicio-1-2-6.pdf": "Resolución Presidencial N.° 522-2025-INPE/P",
    "2026_servir_informe-mujer-servicio-civil-peruano-2026.pdf": "SERVIR, La mujer en el servicio civil peruano 2026",
}
GES = f"{EVAL}/05_entrevistas/ENTREV_codificacion.csv"
SIN_LB = {"nan", "ND", "N.D.", "N.A.", "---", "Sin valor en la ficha"}
REGLAS_NO = ("sin_protocolo", "protocolo_anterior", "sin_indicadores", "sin_valor", "n_ind", "sin_lb")
REGLAS_SI = ("protocolo_posterior", "valor_2025", "acta")


def columna(df, inicio):
    return next(c for c in df.columns if str(c).startswith(inicio))


def fecha_norma(txt):
    d, m, a = re.search(r"\((\d{2})/(\d{2})/(\d{4})\)", str(txt)).groups()
    return date(int(a), int(m), int(d))


def sin_lb(v):
    v = str(v).strip()
    return v in SIN_LB or v.startswith("N.A.")


def verificar_medio(medio, r, ctx):
    """Comprueba un medio «TIPO|referencia|cita o regla» y devuelve (regla, texto legible, rastro). El texto legible
    va en el informe; el rastro (segmento, libro y hoja o página) queda en su columna del libro."""
    tipo, ref, x = medio.split("|", 2)
    if tipo == "REP_ENT":
        s = ctx["fie"].loc[ref]
        assert limpiar(x) in limpiar(str(s.texto)), f"{r.id}: la cita no consta en el segmento {ref}"
        pag = f", p. {int(s.pagina)}" if pd.notna(s.pagina) else ""
        return None, f"{s.fuente}{pag}: «{x}»", ref
    if tipo == "ENTREV":
        s = ctx["ges"].loc[ref]
        assert limpiar(x) in limpiar(str(s.texto)), f"{r.id}: la cita no consta en el segmento {ref}"
        return None, f"Entrevista al nivel conductor ({s.informante}): «{x}»", ref
    if tipo == "HOJA":
        ruta, hoja = ref.rsplit("#", 1)
        celdas = pd.read_excel(ruta, sheet_name=hoja, dtype=str).fillna("").to_numpy().ravel()
        assert any(limpiar(x) in limpiar(c) for c in celdas), f"{r.id}: la cita no consta en {ruta}, hoja {hoja}"
        return None, f"{LIBROS[os.path.basename(ruta)]}: «{x}»", f"{os.path.basename(ruta)}, hoja «{hoja}»"
    if tipo == "PDF":
        ruta, p = ref.rsplit("#", 1)
        t = restituir_fl(limpiar(pagina(ruta, int(p))))
        assert limpiar(x) in t, f"{r.id}: la cita no consta en {ruta}, p. {p}"
        return None, f"{NOMBRES[os.path.basename(ruta)]}, p. {p}: «{x}»", f"{os.path.basename(ruta)}, p. {p}"
    if tipo == "REV_DOC":
        g = ctx["grd"].loc[ref]
        prot, norma = g[columna(ctx["grd"], "Protocolo aprobado")], g[columna(ctx["grd"], "Norma del protocolo")]
        n_ind, anios = g[columna(ctx["grd"], "Indicadores del servicio")], g[columna(ctx["grd"], "Años con valor")]
        ref_informe = INFORMES[r.evaluacion][1]
        if x == "sin_protocolo":
            assert prot == "No", f"{r.id}: {ref} tiene protocolo"
            desc = "sin protocolo aprobado"
        elif x in ("protocolo_posterior", "protocolo_anterior"):
            assert prot == "Sí", f"{r.id}: {ref} sin protocolo"
            posterior = fecha_norma(norma) > ref_informe
            assert posterior == (x == "protocolo_posterior"), f"{r.id}: fecha del protocolo de {ref} ({norma})"
            norma_sin_entidad = re.sub(r" \[.*\]$", "", str(norma))
            desc = f"protocolo aprobado por {norma_sin_entidad}"
        elif x == "sin_indicadores" or x.startswith("n_ind:"):
            k = 0 if x == "sin_indicadores" else int(x.split(":")[1])
            assert int(n_ind) == k, f"{r.id}: {ref} tiene {n_ind} indicadores"
            desc = "sin indicadores" if k == 0 else f"{k} indicador{'es' if k > 1 else ''} en el protocolo"
        elif x == "valor_2025":
            assert "2025" in str(anios), f"{r.id}: {ref} sin valor 2025"
            desc = f"valores reportados en {anios}"
        elif x == "sin_valor":
            assert pd.isna(anios) or not str(anios).strip(), f"{r.id}: {ref} tiene valores"
            desc = "sin valores reportados"
        else:
            raise ValueError(f"{r.id}: regla desconocida {x}")
        return x, f"Revisión documental, servicio {ref}: {desc}", (f"{os.path.basename(GRD)}, hoja «Estado por servicio», servicio {ref}")
    if tipo == "REV_IND":
        b = ctx["psi"]
        for cod in ref.split(";"):
            fila = b.loc[cod]
            if x == "sin_lb":
                assert sin_lb(fila["Valor LB"]), f"{r.id}: {cod} tiene línea de base"
            elif x == "acta":
                assert not sin_lb(fila["Valor LB"]) and str(fila["Estado del acta"]).startswith(
                    ("Acta firmada (ratificada)", "Aprobado mediante resolución")), f"{r.id}: {cod} sin acta"
            else:
                raise ValueError(f"{r.id}: regla desconocida {x}")
        desc = ("sin línea de base" if x == "sin_lb" else
                f"línea de base fijada por acta ratificada ({b.loc[ref.split(';')[0], 'Año LB']})")
        return x, f"Base de indicadores, {ref.replace(';', ' y ')}: {desc}", (f"{os.path.basename(PSI)}, hoja «Base REV_IND.1», {ref.replace(';', ', ')}")
    raise ValueError(f"{r.id}: tipo de medio desconocido {tipo}")


def valorar(R):
    V = pd.read_csv(f"{D}/valoracion_recomendaciones.csv", dtype=str).fillna("")
    assert list(V.id) == list(R.id), "la valoración no cubre exactamente las recomendaciones extraídas"
    ctx = {"fie": pd.read_excel(FIE, sheet_name="Codificación").set_index("id"),
           "grd": pd.read_excel(GRD, sheet_name="Estado por servicio", dtype={"Código": str}).set_index("Código"),
           "psi": pd.read_excel(PSI, sheet_name="Base REV_IND.1").set_index("Código"),
           "ges": pd.read_csv(GES).set_index("id_segmento")}
    M = R.merge(V, on="id")
    medios, rastros = [], []
    for _, r in M.iterrows():
        assert r.estado in ESTADOS, f"{r.id}: estado {r.estado}"
        assert r.aspecto in ASPECTOS, f"{r.id}: aspecto {r.aspecto}"
        if r.estado == "Sin evidencia":
            assert r.pregunta_entrev or r.nota, f"{r.id}: «Sin evidencia» sin pregunta ENTREV ni nota"
        else:
            assert r.medio1, f"{r.id}: «{r.estado}» sin medio de verificación"
        textos, rastro = [], []
        for m in (r.medio1, r.medio2):
            if not m:
                continue
            regla, txt, ras = verificar_medio(m, r, ctx)
            textos.append(txt)
            rastro.append(ras)
            if regla and r.estado == "Implementada":
                assert regla in REGLAS_SI, f"{r.id}: «Implementada» con la regla {regla}"
            if regla and r.estado == "No implementada":
                assert regla.startswith(REGLAS_NO), f"{r.id}: «No implementada» con la regla {regla}"
        for txt in (r.como, r.medida_p36, r.nota):
            assert " — " not in txt and " - " not in txt, f"{r.id}: guion explicativo"
        medios.append(" ".join(f"{i}) {t}" for i, t in enumerate(textos, 1)) if len(textos) > 1 else
                      (textos[0] if textos else "No consta medio de verificación."))
        rastros.append("; ".join(rastro))
    M["medios"], M["rastro_medios"] = medios, rastros
    return M


def guardar(M):
    orden_ev = list(INFORMES)
    res = pd.crosstab(M.evaluacion, M.estado).reindex(index=orden_ev, columns=list(ESTADOS), fill_value=0)
    res["Total"] = res.sum(axis=1)
    res.loc["Total"] = res.sum()
    matriz = M.assign(
        informe=M.evaluacion.map({k: v[2] for k, v in INFORMES.items()}),
        responsable_plazo=[" / ".join(x for x in (a, b) if x) for a, b in zip(M.responsable_informe.fillna(""),
                                                                              M.plazo_informe.fillna(""))])
    cols = {"id": "N.°", "evaluacion": "Evaluación", "informe": "Cita", "pagina_pdf": "Página (PDF)",
            "seccion": "Sección", "servicio": "Servicio", "responsable_plazo": "Responsable y plazo (informe)",
            "texto": "Recomendación (texto literal)", "aspecto": "Aspecto de mejora (Tabla 5)",
            "estado": "Estado al corte (p3.4)", "como": "Cómo se ejecutó (p3.5)", "medios": "Medio de verificación",
            "rastro_medios": "Rastro del medio de verificación",
            "medida_p36": "Medida adicional o traslado a la nueva política (p3.6)",
            "pregunta_entrev": "Pregunta ENTREV", "nota": "Nota"}
    criterios = pd.DataFrame([
        ["CEPLAN", "Preguntas", "p3.4 ¿Cuál es el nivel de avance de las recomendaciones de la evaluación anterior?; "
         "p3.5 ¿Cómo se vienen ejecutando?; p3.6 ¿Qué medidas adicionales son necesarias para su correcta "
         "implementación?", "Guía de seguimiento y evaluación (2024), Tabla A7.1, p. 185"],
        ["CEPLAN", "Aspecto de mejora", "Implementación, diseño o seguimiento y evaluación",
         "Guía de seguimiento y evaluación (2024), Tabla 5, p. 57"],
        ["Equipo evaluador", "Alcance", "Recomendaciones de las cuatro evaluaciones anteriores (MIMP 2020, MIMP 2021, "
         "MIMP 2023 y CEPLAN); una fila por viñeta del informe, salvo la segunda rama de una recomendación condicional "
         "(servicio 1.2.6), que se une a la primera. La evaluación 2020 se identificó en gob.pe después de las "
         "entrevistas, por lo que ninguna pregunta de ENTREV cubrió sus recomendaciones",
         "Decisión del usuario, 2026-11-02; evaluación 2020, 2026-11-06"],
        ["Equipo evaluador", "Escala", "Implementada: la acción recomendada consta en un documento posterior al informe; "
         "Parcialmente implementada: consta una parte de la acción; No implementada: consta que el elemento "
         "recomendado no cambió (protocolo, indicador o línea de base); Sin evidencia: no consta medio de "
         "verificación, se consulta en ENTREV y no se presume; Sin objeto por la transición: la Ley N.° 32535 deja "
         "sin objeto la recomendación", "Propuesta del equipo evaluador, aprobada por el usuario (2026-11-03)"],
        ["Equipo evaluador", "Corte", f"Documentos disponibles al {CORTE}; la fecha de referencia de cada informe es "
         "17-06-2021 (MIMP 2020, firma del informe), mayo de 2022 (MIMP 2021), 16-12-2024 (MIMP 2023) y 16-01-2025 "
         "(CEPLAN, fecha del PDF)", ""],
        ["Equipo evaluador", "Verificación", "Las citas de REP_ENT, ENTREV, los PDF y los libros de análisis constan en "
         "el segmento, la página o la hoja citada; las reglas de REV_DOC y REV_IND se comprueban con esos libros; un "
         "protocolo cuenta como implementación solo si su norma es posterior al informe",
         "seguimiento_recomendaciones.py"],
        ["Equipo evaluador", "Texto de las recomendaciones", "La capa de texto del PDF del CEPLAN pierde la «l» de "
         "la ligadura «fl»; las citas la restituyen (por ejemplo, «influiría» y «refleje», verificadas en la "
         "imagen de la p. 70)", "seguimiento_recomendaciones.py"],
    ], columns=["Plano", "Elemento", "Regla", "Fuente"])
    with pd.ExcelWriter(SALIDA, engine="openpyxl") as w:
        matriz[list(cols)].rename(columns=cols).to_excel(w, sheet_name="Matriz", index=False)
        res.reset_index().rename(columns={"evaluacion": "Evaluación"}).to_excel(
            w, sheet_name="Resumen por evaluación", index=False)
        criterios.to_excel(w, sheet_name="Criterios", index=False)
    wb = load_workbook(SALIDA)
    azul, blanco = PatternFill("solid", fgColor="1F3864"), Font(bold=True, color="FFFFFF")
    anchos = {"Matriz": [8, 11, 14, 8, 22, 8, 22, 70, 16, 18, 50, 60, 30, 45, 9, 40],
              "Resumen por evaluación": [14, 14, 16, 14, 14, 16, 8], "Criterios": [16, 22, 90, 40]}
    for ws in wb.worksheets:
        for c in ws[1]:
            c.fill, c.font = azul, blanco
            c.alignment = Alignment(wrap_text=True, vertical="center")
        for fila in ws.iter_rows(min_row=2):
            for c in fila:
                c.alignment = Alignment(wrap_text=True, vertical="top")
        for i, a in enumerate(anchos[ws.title]):
            ws.column_dimensions[chr(65 + i)].width = a
        ws.freeze_panes = "A2"
    wb.save(SALIDA)
    return res


if __name__ == "__main__":
    os.makedirs(D, exist_ok=True)
    R = extraer()
    verificar_textos(R)
    assert all(R[R.evaluacion == "MIMP 2023"][["responsable_informe", "plazo_informe"]].ne("").all()), "MIMP 2023"
    R.to_csv(f"{D}/recomendaciones_previas.csv", index=False, encoding="utf-8")
    M = valorar(R)
    res = guardar(M)
    print(res.to_string())
    for _, r in M.iterrows():
        print(r.id, r.servicio or "", "|", r.estado, "|", r.texto[:70])
