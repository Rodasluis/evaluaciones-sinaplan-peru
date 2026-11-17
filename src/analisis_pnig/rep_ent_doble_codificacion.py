"""REP_ENT: doble codificación de los hallazgos críticos (nota metodológica, numeral 9: «Los hallazgos críticos se
someten a doble codificación, con resolución de discrepancias en sesión técnica documentada»).

Procedimiento aprobado por el usuario el 2026-10-31 para toda segunda codificación:
  1. un agente nuevo codifica a ciegas los segmentos candidatos (solo texto, metadatos y libro de códigos; sin la
     primera codificación) y decide si cada uno es hallazgo crítico;
  2. este script compara ambas codificaciones (porcentaje de acuerdo, índice de Jaccard y kappa de Cohen) y clasifica
     cada segmento como coincidencia, discrepancia parcial o discrepancia sustantiva;
  3. el equipo evaluador resuelve las discrepancias y revisa una muestra al azar de las coincidencias en el libro
     04_analisis/REP_ENT_doble_codificacion.xlsx, con listas desplegables (columnas en amarillo de la hoja
     «Discrepancias»; «Conforme» en la muestra) y una columna «Estado» que indica qué falta en cada fila;
  4. al volver a correr el script, las decisiones registradas en el libro se conservan y se exportan a
     04_analisis/REP_ENT_adjudicacion.csv, que es el acta en datos de la sesión técnica, y el resultado final por segmento
     a 04_analisis/REP_ENT_doble_codificacion_resultado.csv, que lee rep_ent_informacion_entidades.py.

Insumos: 04_analisis/REP_ENT_codificacion.csv (primera codificación, hallazgo_critico = «Candidato»),
04_analisis/REP_ENT_segunda_codificacion.csv (segunda codificación ciega), 03_datos/procesado/rep_ent_segmentos.csv y
04_analisis/libro_codigos_pnig.csv.

Uso (desde la raíz del repositorio):
    uv run --no-project --with openpyxl --with pandas python src/analisis_pnig/rep_ent_doble_codificacion.py
"""
import os
import re
import warnings

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

warnings.filterwarnings("ignore", category=UserWarning)
RAIZ = EVAL
SEGMENTOS = f"{RAIZ}/03_datos/procesado/rep_ent_segmentos.csv"
PRIMERA = f"{RAIZ}/04_analisis/REP_ENT_codificacion.csv"
SEGUNDA = f"{RAIZ}/04_analisis/REP_ENT_segunda_codificacion.csv"
LIBRO = f"{RAIZ}/04_analisis/libro_codigos_pnig.csv"
SALIDA = f"{RAIZ}/04_analisis/REP_ENT_doble_codificacion.xlsx"
ADJUDICACION = f"{RAIZ}/04_analisis/REP_ENT_adjudicacion.csv"
RESULTADO = f"{RAIZ}/04_analisis/REP_ENT_doble_codificacion_resultado.csv"
TAM_MUESTRA, SEMILLA = 15, 2026
SESION = {"fecha": "31/10/2026", "participantes": "Equipo evaluador (OMEP, OGMEPGD del MIMP)"}  # datos del usuario
SENTIDOS = {"Favorable", "Dificultad", "Mixto", "Propuesta", "Neutro"}

libro = pd.read_csv(LIBRO)
seg = pd.read_csv(SEGMENTOS, dtype=str).fillna("")
p1 = pd.read_csv(PRIMERA, dtype=str).fillna("")
p1 = p1[p1.hallazgo_critico == "Candidato"]
p2 = pd.read_csv(SEGUNDA, dtype=str).fillna("")

# Controles de la segunda codificación
assert set(p2.id) == set(p1.id) and p2.id.is_unique, "la segunda codificación no cubre exactamente los candidatos"
conjunto = lambda s: frozenset(c.strip() for c in s.split(";") if c.strip())
fuera = {c for s in p2.codigos for c in conjunto(s)} - set(libro.codigo)
assert not fuera, f"códigos fuera del libro: {fuera}"
assert set(p2.sentido) <= SENTIDOS and set(p2.critico) <= {"Sí", "No"}


def atribucion(nota):
    """Destino explícito de la primera codificación («Corresponde a …»), con el mismo criterio que REP_ENT."""
    if m := re.match(r"Corresponde a (\d\.\d\.\d|OP\d)", nota):
        return m.group(1)
    if nota.startswith("Corresponde a la CONAIG"):
        return "CONAIG"
    if nota.startswith(("Corresponde a la conclusión", "Corresponde a la recomendación")):
        return "Conclusiones del informe"
    return ""


def normalizar_destino(destino, r):
    """La segunda codificación anota el destino solo si difiere del servicio o indicador del segmento; en los informes de
    la CONAIG, «CONAIG» es la propia fuente y no una reatribución."""
    destino = destino.strip()
    if m := re.search(r"\d\.\d\.\d|OP\d|CONAIG", destino):
        destino = m.group(0)
    propia = destino == "CONAIG" and "CONAIG" in r.fuente
    return "" if destino in (r.servicio, r.indicador) or propia else destino


def objetivo(x):
    return x[2:] if re.fullmatch(r"OP\d", x) else x[0] if re.fullmatch(r"\d\.\d\.\d", x) else None


def mismo_destino(a, b):
    """Un servicio y el objetivo prioritario que lo contiene (2.1.4 y OP2) son la misma atribución a distinto nivel."""
    if a == b:
        return True
    return bool(a and b) and "OP" in a + b and objetivo(a) is not None and objetivo(a) == objetivo(b)


d = (seg[["id", "fuente", "anio", "pagina", "servicio", "indicador", "entidad", "texto"]]
     .merge(p1[["id", "codigos", "sentido", "nota"]], on="id")
     .merge(p2.rename(columns={"codigos": "codigos_2", "sentido": "sentido_2", "justificacion": "justificacion_2"}),
            on="id"))
d["destino_1"] = d.nota.map(atribucion)
d["destino_2"] = [normalizar_destino(x, r) for x, r in zip(d.corresponde_a, d.itertuples())]
d["misma_atribucion"] = [mismo_destino(a, b) for a, b in zip(d.destino_1, d.destino_2)]
c1, c2 = d.codigos.map(conjunto), d.codigos_2.map(conjunto)
d["en común"] = [", ".join(sorted(a & b)) for a, b in zip(c1, c2)]
d["solo 1.ª"] = [", ".join(sorted(a - b)) for a, b in zip(c1, c2)]
d["solo 2.ª"] = [", ".join(sorted(b - a)) for a, b in zip(c1, c2)]
d["Jaccard"] = [round(len(a & b) / len(a | b), 2) for a, b in zip(c1, c2)]


def tipo(r, a, b):
    if (r.sentido != r.sentido_2 or r.critico == "No" or not r.misma_atribucion or not a & b):
        return "Sustantiva"
    return "Coincidencia" if a == b else "Parcial"


d["Tipo"] = [tipo(r, a, b) for r, a, b in zip(d.itertuples(), c1, c2)]
d["Motivo"] = [", ".join(m for m, cond in (
    ("ningún código en común", not a & b), ("códigos distintos", bool(a & b) and a != b),
    ("sentido", r.sentido != r.sentido_2), ("no crítico para la 2.ª", r.critico == "No"),
    ("atribución", not r.misma_atribucion)) if cond) for r, a, b in zip(d.itertuples(), c1, c2)]


from nucleo.codificacion import kappa  # noqa: E402  (kappa de Cohen del núcleo)


por_codigo = []
for cod in sorted(set().union(*c1, *c2)):
    a, b = c1.map(lambda s: cod in s), c2.map(lambda s: cod in s)
    por_codigo.append((cod, libro.set_index("codigo").nombre.get(cod, ""), int(a.sum()), int(b.sum()),
                       int((a & b).sum()), kappa(a, b)))
por_codigo = pd.DataFrame(por_codigo, columns=["Código", "Nombre", "1.ª", "2.ª", "Ambas", "Kappa"]).sort_values(
    "1.ª", ascending=False)
pond = por_codigo.dropna(subset=["Kappa"])
resumen = pd.DataFrame([
    ("Segmentos doblemente codificados", len(d)),
    ("Coincidencias", int((d.Tipo == "Coincidencia").sum())),
    ("Discrepancias parciales", int((d.Tipo == "Parcial").sum())),
    ("Discrepancias sustantivas", int((d.Tipo == "Sustantiva").sum())),
    ("Acuerdo exacto en los códigos (%)", round(100 * (c1 == c2).mean(), 1)),
    ("Índice de Jaccard medio", round(d.Jaccard.mean(), 2)),
    ("Acuerdo en el sentido (%)", round(100 * (d.sentido == d.sentido_2).mean(), 1)),
    ("Kappa del sentido", kappa(d.sentido, d.sentido_2)),
    ("Kappa medio por código, ponderado por frecuencia", round((pond.Kappa * (pond["1.ª"] + pond["2.ª"])).sum()
                                                                 / (pond["1.ª"] + pond["2.ª"]).sum(), 2)),
    ("Hallazgos críticos confirmados por la 2.ª (%)", round(100 * (d.critico == "Sí").mean(), 1)),
    ("Acuerdo en la atribución (%)", round(100 * d.misma_atribucion.mean(), 1)),
], columns=["Indicador", "Valor"])

# Columnas que completa el evaluador. Las listas desplegables leen la hoja oculta «Listas»; donde ambas
# codificaciones coinciden, el valor viene ya puesto y puede cambiarse.
OPC_CODIGOS = ["Los de la 1.ª", "Los de la 2.ª", "Unión de ambas", "Solo los comunes", "Elegir en Código final 1 a 4"]
N_CODIGOS = 4
COLS_DEC = ["Códigos: decisión", *[f"Código final {i}" for i in range(1, N_CODIGOS + 1)], "Sentido final",
            "Crítico final", "Atribución final", "Nota de la sesión"]
SIN_REATRIBUIR = "Sin reatribuir"
servicios = pd.read_csv("src/formulario_fie03/datos/servicios.csv", dtype=str).codigo_corto.tolist()
OPC_ATRIB = [SIN_REATRIBUIR, *servicios, *[f"OP{i}" for i in range(1, 7)], "CONAIG", "Conclusiones del informe"]
OPC_LIBRO = [f"{c} ({n})" for c, n in zip(libro.codigo, libro.nombre)]


def previo(hoja, columnas):
    """Decisiones ya registradas en el libro (se conservan al regenerarlo)."""
    if not os.path.exists(SALIDA):
        return pd.DataFrame(columns=["id", *columnas])
    hojas_previas = pd.read_excel(SALIDA, sheet_name=None, dtype=str)
    if hoja not in hojas_previas:
        return pd.DataFrame(columns=["id", *columnas])
    df = hojas_previas[hoja].fillna("")
    return df[["id", *[c for c in columnas if c in df.columns]]]


vista = ["id", "Tipo", "Motivo", "fuente", "anio", "pagina", "servicio", "indicador", "entidad", "texto",
         "codigos", "sentido", "nota", "destino_1", "codigos_2", "sentido_2", "critico", "destino_2", "justificacion_2",
         "en común", "solo 1.ª", "solo 2.ª", "Jaccard"]
disc = d[d.Tipo != "Coincidencia"].sort_values(["Tipo", "Jaccard"], ascending=[False, True])[vista].copy()
sugerido = {
    "Sentido final": [a if a == b else "" for a, b in zip(disc.sentido, disc.sentido_2)],
    "Crítico final": ["Sí" if c == "Sí" else "" for c in disc.critico],
    "Atribución final": [(a if len(a) >= len(b) else b) or SIN_REATRIBUIR if mismo_destino(a, b) else ""
                         for a, b in zip(disc.destino_1, disc.destino_2)],
}
POR_REGLA = "Completado por regla"
disc = disc.merge(previo("Discrepancias", [*COLS_DEC, POR_REGLA]), on="id", how="left")
regla_previa = disc[POR_REGLA].fillna("") if POR_REGLA in disc else pd.Series("", index=disc.index)
for col in COLS_DEC:
    disc[col] = disc[col].fillna("") if col in disc else ""
    # Lo que completó la regla en una corrida anterior se recalcula, por si cambió la decisión.
    disc[col] = ["" if col in r.split("; ") else v for v, r in zip(disc[col], regla_previa)]
    if col in sugerido:
        disc[col] = [p or s for p, s in zip(disc[col], sugerido[col])]


def completar_por_regla(r):
    """Regla aprobada por el usuario (2026-10-31): si la decisión de códigos elige una codificación, los campos que el
    evaluador deja en blanco se toman de esa misma codificación. Lo que escribe el evaluador prevalece."""
    elegida = {"Los de la 1.ª": (r.sentido, "Sí", r.destino_1),
               "Los de la 2.ª": (r.sentido_2, r.critico, r.destino_2)}.get(r["Códigos: decisión"])
    llenos = []
    if elegida:
        for campo, valor in zip(("Sentido final", "Crítico final", "Atribución final"), elegida):
            if r[campo] == "":
                r[campo] = valor or SIN_REATRIBUIR
                llenos.append(campo)
    r[POR_REGLA] = "; ".join(llenos)
    return r


disc = disc.apply(completar_por_regla, axis=1)

muestra = (d[d.Tipo == "Coincidencia"].sample(n=min(TAM_MUESTRA, (d.Tipo == "Coincidencia").sum()),
                                             random_state=SEMILLA)[vista]
           .merge(previo("Muestra de coincidencias", ["Conforme", "Nota"]), on="id", how="left"))
for col in ("Conforme", "Nota"):
    muestra[col] = muestra[col].fillna("") if col in muestra else ""


def codigos_finales(r):
    """Códigos que resultan de la decisión; vacío si la fila aún no está resuelta."""
    a, b = conjunto(r.codigos), conjunto(r.codigos_2)
    elegidos = {r[f"Código final {i}"].split(" ")[0] for i in range(1, N_CODIGOS + 1)} - {""}
    final = {"Los de la 1.ª": a, "Los de la 2.ª": b, "Unión de ambas": a | b, "Solo los comunes": a & b,
             "Elegir en Código final 1 a 4": elegidos}.get(r["Códigos: decisión"], set())
    return "; ".join(sorted(final))


disc["Códigos resultantes"] = disc.apply(codigos_finales, axis=1)
completa = ((disc["Códigos resultantes"] != "") & (disc["Sentido final"] != "") & (disc["Crítico final"] != "")
            & (disc["Atribución final"] != ""))

# Exportación de las decisiones (acta en datos)
resueltas = disc[completa]
pd.concat([
    resueltas[["id", "Tipo", "Códigos: decisión", "Códigos resultantes", "Sentido final", "Crítico final",
               "Atribución final", "Nota de la sesión", POR_REGLA]].assign(Revisión="Discrepancia"),
    muestra[muestra.Conforme != ""][["id", "Conforme", "Nota"]].rename(columns={"Nota": "Nota de la sesión"})
    .assign(Tipo="Coincidencia", Revisión="Muestra"),
]).to_csv(ADJUDICACION, index=False, encoding="utf-8")

acta = pd.DataFrame([
    ("Norma", "Nota metodológica, numeral 9: los hallazgos críticos se someten a doble codificación, con resolución de "
     "discrepancias en sesión técnica documentada.", "Nota metodológica"),
    ("Primera codificación", "Claude, con nota de sustento; propone los candidatos a hallazgo crítico.",
     "Decisión del usuario, 2026-10-30"),
    ("Segunda codificación", "Agente nuevo, ciego respecto de la primera: recibió solo el texto, los metadatos y el libro "
     "de códigos, en orden aleatorio; la ceguera se aseguró por instrucción y por carpeta separada "
     "(REP_ENT_segunda_codificacion_instrucciones.md).", "Decisión del usuario, 2026-10-31"),
    ("Criterio de hallazgo crítico", "Segmento que, solo o con otros, sustenta una conclusión capaz de cambiar la valoración "
     "de una pregunta (D5, I1, I4, I5) o una recomendación sobre la continuidad, el rediseño o el retiro de un servicio o "
     "indicador. La primera codificación lo aplicó sin definición escrita; esta definición se fijó para la segunda.",
     "Equipo evaluador"),
    ("Clasificación", "Coincidencia: mismos códigos, sentido, atribución y crítico. Parcial: códigos en común pero no "
     "idénticos. Sustantiva: ningún código en común, o difieren el sentido, la atribución o la condición de crítico. "
     "En la atribución, un servicio y el objetivo que lo contiene (2.1.4 y OP2) cuentan como la misma, y «CONAIG» en "
     "un informe de la CONAIG es la propia fuente.", "Equipo evaluador"),
    ("Resolución", f"El evaluador resuelve todas las discrepancias y revisa {TAM_MUESTRA} coincidencias al azar (semilla "
     f"{SEMILLA}) para detectar errores compartidos, que dos codificaciones del mismo modelo pueden tener. Sentido, "
     "crítico y atribución vienen ya puestos donde ambas codificaciones coinciden. Si el evaluador elige los códigos de una "
     "codificación (1.ª o 2.ª), los campos que deja en blanco se toman de esa misma codificación; lo que escribe prevalece "
     "y la columna «Completado por regla» indica qué campos completó la regla.", "Decisión del usuario, 2026-10-31"),
    ("Limitación", "Ambas codificaciones las hizo el mismo modelo; el acuerdo puede estar sobrestimado. Se declara en el "
     "anexo técnico como segunda codificación asistida por IA, ciega, con adjudicación del evaluador.", "Equipo evaluador"),
    ("Sesión técnica", f"Fecha: {SESION['fecha']}. Participantes: {SESION['participantes']}.", "Equipo evaluador"),
    ("Decisiones", "Discrepancias: " + "; ".join(f"{n} con «{k}»" for k, n in
                                                resueltas["Códigos: decisión"].value_counts().items())
     + f". Campos completados por la regla: {(resueltas[POR_REGLA] != '').sum()} filas. Muestra de coincidencias: "
     f"{(muestra.Conforme == 'Sí').sum()} conformes de {len(muestra)}. Hallazgos críticos: "
     f"{(d.Tipo == 'Coincidencia').sum() + (resueltas['Crítico final'] == 'Sí').sum()} confirmados y "
     f"{(resueltas['Crítico final'] == 'No').sum()} descartados de {len(d)} candidatos.", "Equipo evaluador"),
    ("Extensión del criterio de cobertura", "Al adoptar la segunda codificación, el equipo evaluador adoptó su lectura de "
     "I1.COB: incluye el servicio que se presta sin distinguir ni priorizar a las mujeres (convocatorias o requisitos "
     "generales, titulación «sin distinción»). Para mantener un solo criterio, se extendió al resto de la codificación: "
     "se revisaron los segmentos preseleccionados por palabras clave y 5 recibieron I1.COB (columna «validacion» de "
     "REP_ENT_codificacion.csv); el libro de códigos precisa el criterio en «incluye».", "Equipo evaluador, 31/10/2026"),
    ("Avance de la resolución", f"{len(resueltas)} de {len(disc)} discrepancias resueltas; "
     f"{(muestra.Conforme != '').sum()} de {len(muestra)} coincidencias revisadas.", "Cálculo"),
], columns=["Elemento", "Detalle", "Plano"])

nombres = {"codigos": "Códigos 1.ª", "sentido": "Sentido 1.ª", "nota": "Nota 1.ª", "destino_1": "Atribución 1.ª",
           "codigos_2": "Códigos 2.ª", "sentido_2": "Sentido 2.ª", "critico": "Crítico 2.ª",
           "destino_2": "Atribución 2.ª", "justificacion_2": "Justificación 2.ª"}
# Resultado final por segmento, que lee rep_ent_informacion_entidades.py: las coincidencias quedan como hallazgo crítico
# confirmado; las discrepancias, según la adjudicación; las pendientes no se incluyen.
def nota_final(r):
    decision = r["Códigos: decisión"]
    base = r.justificacion_2 if decision == "Los de la 2.ª" else r.nota if decision == "Los de la 1.ª" else \
        f"1.ª: {r.nota} 2.ª: {r.justificacion_2}"
    return f"{base} Sesión: {r['Nota de la sesión']}" if r["Nota de la sesión"] else base


limpiar = lambda x: "Varios objetivos prioritarios" if x.startswith("Varios OP") else x
coinc = d[d.Tipo == "Coincidencia"]
pd.concat([
    pd.DataFrame({"id": coinc.id, "codigos": coinc.codigos, "sentido": coinc.sentido, "critico": "Sí",
                  "atribucion": [limpiar(x) or SIN_REATRIBUIR for x in coinc.destino_1], "nota": coinc.nota,
                  "origen": "Coincidencia de ambas codificaciones"}),
    pd.DataFrame({"id": resueltas.id, "codigos": resueltas["Códigos resultantes"], "sentido": resueltas["Sentido final"],
                  "critico": resueltas["Crítico final"], "atribucion": resueltas["Atribución final"].map(limpiar),
                  "nota": resueltas.apply(nota_final, axis=1),
                  "origen": "Discrepancia adjudicada: " + resueltas["Códigos: decisión"]}),
]).to_csv(RESULTADO, index=False, encoding="utf-8")

disc = disc.drop(columns="Códigos resultantes")
disc.insert(1, "Estado", "")
listas = pd.concat([pd.Series(OPC_CODIGOS, name="Decisión sobre los códigos"), pd.Series(OPC_LIBRO, name="Código"),
                    pd.Series(sorted(SENTIDOS), name="Sentido"), pd.Series(["Sí", "No"], name="Sí o no"),
                    pd.Series(OPC_ATRIB, name="Atribución")], axis=1)
desplegables = {"Códigos: decisión": "A", **{f"Código final {i}": "B" for i in range(1, N_CODIGOS + 1)},
                "Sentido final": "C", "Crítico final": "D", "Conforme": "D", "Atribución final": "E"}
largo = {"A": len(OPC_CODIGOS), "B": len(OPC_LIBRO), "C": len(SENTIDOS), "D": 2, "E": len(OPC_ATRIB)}

hojas = [("Resumen", resumen), ("Discrepancias", disc.rename(columns=nombres)),
         ("Muestra de coincidencias", muestra.rename(columns=nombres)), ("Concordancia por código", por_codigo),
         ("Todas", d[vista].rename(columns=nombres)), ("Acta", acta), ("Listas", listas)]
with pd.ExcelWriter(SALIDA, engine="openpyxl") as xw:
    for nombre, df in hojas:
        df.to_excel(xw, sheet_name=nombre, index=False)
        ws = xw.sheets[nombre]
        ws.freeze_panes = "B2"
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F3864")
            c.alignment = Alignment(wrap_text=True, vertical="top")
        cabecera = {c.value: c.column_letter for c in ws[1]}
        for col in ws.columns:
            titulo = str(col[0].value)
            ancho = 80 if titulo in ("texto", "Detalle") else 40 if titulo in (
                "Nota 1.ª", "Justificación 2.ª", "Indicador", "Nota de la sesión", "Nota", "Motivo") else 24 if titulo in (
                desplegables.keys() | {"Estado", "Decisión sobre los códigos", "Código", "Atribución", POR_REGLA}) else 14
            ws.column_dimensions[col[0].column_letter].width = ancho
            if titulo in COLS_DEC + ["Conforme", "Nota"]:
                for c in col[1:]:
                    c.fill = PatternFill("solid", fgColor="FFF2CC")
            if ancho >= 40:
                for c in col[1:]:
                    c.alignment = Alignment(wrap_text=True, vertical="top")
        for campo, lista in desplegables.items():
            if campo in cabecera and ws.max_row > 1:
                dv = DataValidation(type="list", formula1=f"=Listas!${lista}$2:${lista}${largo[lista] + 1}",
                                    allow_blank=True, showErrorMessage=True, errorTitle="Valor no válido",
                                    error="Elige un valor de la lista.")
                ws.add_data_validation(dv)
                dv.add(f"{cabecera[campo]}2:{cabecera[campo]}{ws.max_row}")
        if nombre == "Discrepancias":
            # Estado en vivo: indica qué falta en cada fila mientras se completa el libro.
            col = {k: cabecera[k] for k in ("Códigos: decisión", "Código final 1", "Sentido final", "Crítico final",
                                             "Atribución final")}
            for f in range(2, ws.max_row + 1):
                ws[f"{cabecera['Estado']}{f}"] = (
                    f'=IF({col["Códigos: decisión"]}{f}="","Falta: códigos",'
                    f'IF(OR({col["Códigos: decisión"]}{f}="{OPC_CODIGOS[0]}",{col["Códigos: decisión"]}{f}="{OPC_CODIGOS[1]}"),'
                    f'"Resuelta",'
                    f'IF(AND({col["Códigos: decisión"]}{f}="{OPC_CODIGOS[-1]}",{col["Código final 1"]}{f}=""),'
                    f'"Falta: elegir códigos",IF({col["Sentido final"]}{f}="","Falta: sentido",'
                    f'IF({col["Crítico final"]}{f}="","Falta: crítico",'
                    f'IF({col["Atribución final"]}{f}="","Falta: atribución","Resuelta"))))))')
        if nombre == "Listas":
            ws.sheet_state = "hidden"
print(resumen.to_string(index=False))
print(acta.Detalle.iloc[-1])
