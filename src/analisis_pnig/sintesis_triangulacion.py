"""Matriz de triangulación por pregunta de evaluación (nota metodológica, numerales 7 y 9).

Reúne los hallazgos de sintesis_hallazgos.py con su evidencia, ordenada por las cuatro rutas de la
estrategia metodológica (documental, cuantitativa, información de las entidades y literatura) más el nivel conductor
(ENTREV), declara cuántas fuentes independientes los sostienen y su nivel de confianza, y propone la valoración de cada
pregunta. Las preguntas de ENTREV que contrastan cada hallazgo (las que se formularon) se leen de
05_entrevistas/ENTREV_preguntas.csv, y la evidencia del nivel conductor, de los hallazgos.

Salida: evaluaciones/pnig_2026/04_analisis/matriz_triangulacion_preguntas.xlsx, con las hojas «Valoración por
pregunta», «Matriz de triangulación», «Evidencias», «Cobertura», «Pendientes», «Cifras» y «Criterios». Las columnas en
amarillo son del equipo evaluador (valoración y confianza validadas, resultado del contraste en ENTREV): al volver a
correr el script, lo registrado en ellas se conserva. Las celdas vacías de valoración y confianza validadas se llenan
con la validación del usuario (2026-11-06); un hallazgo que cambió después queda señalado en la observación.

Planos: las preguntas, las escalas y la regla de confianza son de la nota; los enunciados, la confianza y la
valoración son del equipo evaluador, validados por el usuario el 2026-11-06.

Uso (desde la raíz del repositorio; antes, entrev_guias_entrevista.py):
    uv run --no-project --with pandas --with openpyxl --with python-docx python src/analisis_pnig/sintesis_triangulacion.py
"""
import os
import sys

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, os.path.dirname(__file__))
import sintesis_hallazgos as S  # noqa: E402
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

SALIDA = f"{S.A}/matriz_triangulacion_preguntas.xlsx"
CATALOGO = f"{EVAL}/05_entrevistas/ENTREV_preguntas.csv"
PLANOS = ["Documental", "Cuantitativa", "Entidades", "Literatura"]
EVALUADOR = {  # columnas del equipo evaluador que se conservan entre corridas: hoja -> (clave, columnas)
    "Valoración por pregunta": ("Pregunta", ["Valoración validada", "Observación del equipo evaluador"]),
    "Matriz de triangulación": ("Hallazgo", ["Resultado del contraste ENTREV", "Confianza validada",
                                             "Observación del equipo evaluador"]),
}
VALIDACION = "Validado por el usuario (2026-11-06)"


def _previo():
    """Lee lo registrado por el equipo evaluador en la versión anterior del libro."""
    previo = {}
    if not os.path.exists(SALIDA):
        return previo
    for hoja, (clave, cols) in EVALUADOR.items():
        try:
            d = pd.read_excel(SALIDA, sheet_name=hoja, dtype=str).fillna("")
        except ValueError:
            continue
        previo[hoja] = {r[clave]: {c: r.get(c, "") for c in cols} for _, r in d.iterrows()}
    return previo


def _cifra(v):
    if isinstance(v, pd.DataFrame):
        return f"Tabla de {len(v)} filas"
    if isinstance(v, (list, tuple)):
        return "; ".join(str(x) for x in v)
    if isinstance(v, dict):
        return "; ".join(f"{k}: {x if not isinstance(x, float) else round(x, 2)}" for k, x in list(v.items())[:12])
    if isinstance(v, float):
        return round(v, 2)
    return v


def construir():
    C = S.cifras()
    H = S.hallazgos(C)
    V = S.valoraciones(C)
    Q = pd.read_csv(CATALOGO, dtype=str).fillna("")
    contraste = {}
    for _, r in Q.iterrows():
        for h in filter(None, (x.strip() for x in r.contrasta.split(";"))):
            contraste.setdefault(h, []).append(r.id)
    previo = _previo()

    matriz, evid = [], []
    for h in H:
        fila = {"Pregunta": h["pregunta"], "Hallazgo": h["id"], "Aspecto": h["aspecto"],
                "Enunciado": h["enunciado"]}
        for plano in PLANOS:
            fila[plano] = "\n".join(f"{e['instrumento']}: {e['texto']} Fuente: {e['fuente']}."
                                    for e in h["evidencias"] if e["plano"] == plano)
        conductor = [e for e in h["evidencias"] if e["plano"] == "Nivel conductor"]
        fila["Nivel conductor (ENTREV)"] = ("\n".join(f"{e['instrumento']}: {e['texto']} Fuente: {e['fuente']}."
                                                      for e in conductor) or "No se contrastó en ENTREV")
        fila["Fuentes independientes"] = S.fuentes_independientes(h) if h["evidencias"] else 0
        fila["Bases"] = "; ".join(sorted({e["base"] for e in h["evidencias"]}))
        fila["Confianza propuesta"] = h["confianza"]
        fila["Estado"] = h["estado"]
        fila["Nota"] = h["nota"]
        for c in EVALUADOR["Matriz de triangulación"][1]:
            fila[c] = previo.get("Matriz de triangulación", {}).get(h["id"], {}).get(c, "")
        if h["estado"] in (S.VALIDADO, S.ACTUALIZADO) and not fila["Confianza validada"]:
            fila["Confianza validada"] = h["confianza"]
            fila["Observación del equipo evaluador"] = fila["Observación del equipo evaluador"] or (
                VALIDACION + ("; el enunciado cambió después de la validación (2026-11-07), revisar"
                              if h["estado"] == S.ACTUALIZADO else ""))
        matriz.append(fila)
        for e in h["evidencias"]:
            evid.append({"Hallazgo": h["id"], "Pregunta": h["pregunta"], "Plano": e["plano"],
                         "Instrumento": e["instrumento"], "Evidencia": e["texto"], "Fuente": e["fuente"],
                         "Base (fuente primaria)": e["base"], "Fuentes primarias que agrega": e["n_fuentes"]})
    matriz = pd.DataFrame(matriz)
    evid = pd.DataFrame(evid)

    val, cob = [], []
    for p, (crit, texto, previstos) in S.PREGUNTAS.items():
        hs = [h for h in H if h["pregunta"] == p]
        usados = sorted({e["instrumento"] for h in hs for e in h["evidencias"]})
        g = sorted({qid for h in hs for qid in contraste.get(h["id"], [])}, key=lambda x: (x[:2], int(x.split(".")[1])) if not x.startswith("PS-") else ("S", 0, x))
        valor, sustento, estado = V[p]
        fila = {"Pregunta": p, "Criterio": crit, "Texto de la pregunta": texto,
                "Hallazgos": ", ".join(h["id"] for h in hs),
                "Confianza de los hallazgos": "; ".join(f"{h['id']}: {h['confianza']}" for h in hs),
                "Valoración": valor, "Sustento": sustento, "Estado": estado,
                "Preguntas ENTREV que la contrastan": ", ".join(g)}
        for c in EVALUADOR["Valoración por pregunta"][1]:
            fila[c] = previo.get("Valoración por pregunta", {}).get(p, {}).get(c, "")
        if estado == S.VALIDADA and not fila["Valoración validada"]:
            fila["Valoración validada"] = valor
            fila["Observación del equipo evaluador"] = fila["Observación del equipo evaluador"] or VALIDACION
        val.append(fila)
        faltan = [i for i in previstos if i not in usados]
        cob.append({"Pregunta": p, "Instrumentos previstos en la nota (Tablas 4 y 5)": ", ".join(previstos),
                    "Instrumentos con evidencia en la matriz": ", ".join(usados),
                    "Previstos sin evidencia todavía": ", ".join(faltan),
                    "Dos o más instrumentos con evidencia": "Sí" if len(usados) >= 2 else "No"})
    val, cob = pd.DataFrame(val), pd.DataFrame(cob)
    assert set(val["Valoración"]) <= set(S.VALORACION) | {"Sin valorar"}

    pend = []
    for h in H:
        if h["estado"] == "Análisis pendiente":
            pend.append({"Tipo": "Análisis pendiente", "Pregunta": h["pregunta"], "Hallazgo": h["id"],
                         "Descripción": h["nota"] or h["enunciado"], "Responsable": "Equipo evaluador (OMEP)"})
    for h in H:
        if h["confianza"] == "Indicio" and h["estado"] not in (S.VALIDADO, S.ACTUALIZADO):
            pend.append({"Tipo": "Confianza baja", "Pregunta": h["pregunta"], "Hallazgo": h["id"],
                         "Descripción": "Sustentado en una sola fuente; buscar una segunda o mantenerlo como indicio.",
                         "Responsable": "Equipo evaluador (OMEP)"})
    # ronda 35: el informe de la OGAJ no llegará (usuario, 2026-11-11) y deja de figurar como fuente por llegar
    pend = pd.DataFrame(pend, columns=["Tipo", "Pregunta", "Hallazgo", "Descripción", "Responsable"])

    cif = pd.DataFrame([{"Clave": k, "Valor": _cifra(v), "Origen": C.origen[k]} for k, v in C.items()
                        if k not in ("grd", "fie_seg", "libro")])
    crit = pd.DataFrame([
        ["Nota metodológica", "Confianza del hallazgo", "Sólido: al menos dos fuentes o instrumentos independientes "
         "coinciden. Con respaldo parcial: coinciden con discrepancias explicables. Indicio: una sola fuente "
         "(numeral 7)."],
        ["Nota metodológica", "Valoración de la pregunta", "Satisfactorio, parcial o insuficiente, con la evidencia, la "
         "fuente y el nivel de confianza (numeral 9, Tabla 9)."],
        ["Equipo evaluador", "Fuentes independientes", "Dos evidencias son independientes si provienen de fuentes "
         "primarias distintas (columna «Base»). Una matriz que agrega varias fuentes primarias independientes, como "
         "REV_EVI, cuenta por cada una."],
        ["Equipo evaluador", "No determinado", "Se usa solo cuando el análisis que sustentaría el hallazgo no se ha "
         "hecho; el hallazgo queda como pendiente y no se valora."],
        ["Equipo evaluador", "Sin valorar", "La pregunta no se valora mientras falte un análisis previsto en la nota o "
         "su valoración dependa principalmente de ENTREV."],
        ["Equipo evaluador", "Rutas", "Documental (REV_DOC, REV_ALI y fichas), cuantitativa (REV_IND y panel de REP_ENT), "
         "entidades (REP_ENT: formulario, informes LIO e informes de la CONAIG), literatura (REV_EVI) y nivel conductor "
         "(ENTREV), conforme al numeral 7 de la nota."],
        ["Equipo evaluador", "Columnas en amarillo", "Las completa el equipo evaluador; se conservan al volver a correr "
         "el script."],
        ["Equipo evaluador", "Cifras", "Toda cifra de los enunciados se calcula en sintesis_hallazgos.py desde los "
         "libros de análisis (hoja «Cifras»); los datos tomados de informes se verifican contra el texto literal del "
         "segmento de REP_ENT."],
    ], columns=["Plano", "Elemento", "Regla"])
    return {"Valoración por pregunta": val, "Matriz de triangulación": matriz, "Evidencias": evid,
            "Cobertura": cob, "Pendientes": pend, "Cifras": cif, "Criterios": crit}


ANCHOS = {
    "Valoración por pregunta": [9, 16, 50, 22, 34, 14, 60, 20, 26, 16, 40],
    "Matriz de triangulación": [9, 9, 24, 70, 45, 45, 45, 40, 26, 11, 30, 16, 20, 40, 26, 16, 40],
    "Evidencias": [9, 9, 13, 11, 50, 55, 32, 12],
    "Cobertura": [9, 30, 34, 26, 16],
    "Pendientes": [18, 9, 9, 80, 24],
    "Cifras": [24, 60, 70],
    "Criterios": [18, 24, 100],
}


def escribir(hojas):
    with pd.ExcelWriter(SALIDA, engine="openpyxl") as w:
        for nombre, d in hojas.items():
            d.to_excel(w, sheet_name=nombre, index=False)
    wb = load_workbook(SALIDA)
    azul, amarillo = PatternFill("solid", fgColor="1F3864"), PatternFill("solid", fgColor="FFF2CC")
    colores = {"Sólido": "C6EFCE", "Con respaldo parcial": "FFEB9C", "Indicio": "F8CBAD", "No determinado": "D9D9D9",
               "Satisfactorio": "C6EFCE", "Parcial": "FFEB9C", "Insuficiente": "F8CBAD", "Sin valorar": "D9D9D9"}
    for nombre, ws in zip(hojas, wb.worksheets):
        for c in ws[1]:
            c.fill, c.font = azul, Font(bold=True, color="FFFFFF")
            c.alignment = Alignment(wrap_text=True, vertical="center")
        for i, a in enumerate(ANCHOS[nombre]):
            ws.column_dimensions[ws.cell(1, i + 1).column_letter].width = a
        for fila in ws.iter_rows(min_row=2):
            for c in fila:
                c.alignment = Alignment(wrap_text=True, vertical="top")
                if isinstance(c.value, str) and c.value in colores:
                    c.fill = PatternFill("solid", fgColor=colores[c.value])
        ws.freeze_panes = "C2" if nombre in EVALUADOR else "A2"
        if nombre in EVALUADOR:
            cab = [c.value for c in ws[1]]
            for col in EVALUADOR[nombre][1]:
                j = cab.index(col) + 1
                letra = ws.cell(1, j).column_letter
                for f in range(2, ws.max_row + 1):
                    ws.cell(f, j).fill = amarillo
                lista = {"Valoración validada": '"Satisfactorio,Parcial,Insuficiente"',
                         "Confianza validada": '"Sólido,Con respaldo parcial,Indicio"',
                         "Resultado del contraste ENTREV": '"Coincide,Matiza,Discrepa,Aporta información nueva,'
                                                           'No se trató"'}.get(col)
                if lista:
                    dv = DataValidation(type="list", formula1=lista, allow_blank=True)
                    ws.add_data_validation(dv)
                    dv.add(f"{letra}2:{letra}{ws.max_row}")
    wb.save(SALIDA)


if __name__ == "__main__":
    hojas = construir()
    escribir(hojas)
    v = hojas["Valoración por pregunta"]
    m = hojas["Matriz de triangulación"]
    print(v[["Pregunta", "Valoración", "Valoración validada", "Estado", "Hallazgos"]].to_string(index=False))
    print(m["Confianza propuesta"].value_counts().to_string())
    print(hojas["Cobertura"].to_string(index=False))
    print(len(hojas["Pendientes"]), "pendientes")
