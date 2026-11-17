"""REV_DOC, Sección B: registro del estado de implementación de los 52 servicios de la PNIG (clasificación documental).

Criterios (ADR-011, ronda 33: todo el periodo 2019–2025; ADR-012, ronda 35: dos de tres componentes):
  Implementado: al menos dos de los tres componentes (protocolo aprobado, asignación presupuestal identificable y
                reporte de indicadores, en algún año de 2019–2025), con ejecución reportada en 2024 o 2025, es decir,
                el servicio sigue operando al cierre.
  Implementado parcialmente: ejecución reportada en algún año de 2019–2025 con dos o más de esos componentes
                faltantes, o interrumpida antes de 2024.
  No implementado: sin evidencia de ejecución en ningún año de 2019–2025 (o solo acciones preparatorias).
  Sin información suficiente: solo secciones mínimas o información contradictoria.

La clasificación es preliminar: se contrasta con la lectura de las secciones del LIO y con las entrevistas.

Información detallada del LIO (regla interna): una sección del servicio de MIN_LINEAS líneas o más que no declara falta
de información; el informe la describe sin el umbral (pedido del usuario, ronda 35).

Asignación presupuestal identificable (nota v14; desde la ronda 33, todo el periodo): al menos un monto reportado para
el servicio en 2019–2025 (informes LIO o formulario de seguimiento 2023) coincide con una meta, actividad o producto de la entidad responsable en los datos
abiertos del gasto del MEF. Se lee de la hoja «Cifras verificadas» de REV_IND_3_verificacion_presupuestal.xlsx, que a su
vez toma de este libro solo la columna de responsables.

Información reportada por las entidades (REP_ENT): la nota dispone que la Sección B se completa con REP_ENT. Por servicio
se agregan los factores de dificultad y los favorables más frecuentes, las propuestas de ajuste y los hallazgos críticos
confirmados en la doble codificación, con la fuente y la página de cada cita. Se leen de la hoja «Codificación» de
REP_ENT_informacion_entidades.xlsx, que ya atribuye cada segmento a su servicio. Sirven para contrastar la clasificación;
no la modifican.

Uso (desde la raíz del repositorio, después de rev_ind_3_verificacion_presupuestal.py y de rep_ent_informacion_entidades.py,
que de este libro solo usa el servicio, la entidad y el protocolo):
    uv run --no-project --with openpyxl --with pandas python src/analisis_pnig/rev_doc_estado_servicios.py
"""
import re

import pandas as pd
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
from nucleo.config import insumo  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

RAIZ = EVAL
CONSOLIDADO = f"{RAIZ}/03_datos/raw/2026_omep_consolidado-indicadores-pnig.xlsx"
PANEL = f"{RAIZ}/03_datos/procesado/panel_formulario_seguimiento_pnig_2023_2025.csv"
LIO = f"{RAIZ}/03_datos/procesado/cobertura_lio_servicios_2019_2025.csv"
SALIDA = f"{RAIZ}/04_analisis/REV_DOC_estado_implementacion_servicios.xlsx"
VERIFICACION = f"{RAIZ}/04_analisis/REV_IND_3_verificacion_presupuestal.xlsx"
FIE03 = f"{RAIZ}/04_analisis/REP_ENT_informacion_entidades.xlsx"
LIBRO = f"{RAIZ}/04_analisis/libro_codigos_pnig.csv"
from nucleo.escalas import clasificar_implementacion  # noqa: E402
ESCALA = _cargar_config()["escalas"]["implementacion"]   # parámetros de la escala (config.yaml; ADR-011 y ADR-012)
MIN_LINEAS = ESCALA["min_lineas_lio"]   # extensión mínima de una sección del LIO para considerarla información detallada
# Coincidencias que identifican una estructura presupuestal del servicio (equipo evaluador): los totales de programa y
# los pliegos de otro sector no la identifican, y los montos redondos pueden coincidir por azar
RESULTADOS_VERIFICADOS = {"Coincide", "Cercano (≤ 0,5 % del mismo agregado)", "Coincide con el año anterior"}
NIVELES_SERVICIO = ("Meta", "Metas del mismo nombre", "Actividad", "Unidad ejecutora y actividad", "Producto")

_INS = insumo("estado_servicios.yaml")   # evaluaciones/<id>/02_insumos (Fase 2)
ETIQUETA = _INS["etiquetas_entidad"]
# Órganos del MIMP identificados a partir de los informes LIO (confirmado por la OMEP el 27-10-2026)
MIMP_LIO = _INS["organos_mimp_lio"]

# Evidencia de ejecución 2025, resumida de la sección del servicio en el informe LIO 2025
EVIDENCIA = _INS["evidencia_ejecucion_2025"]
# Ajustes por lectura cualitativa: un servicio se considera ejecutado cuando la entidad reporta su provisión a la población
# objetivo; las acciones preparatorias (diagnóstico, protocolo) o solo vinculadas no cuentan como ejecución.
AJUSTES = {c: (x["estado"], x["motivo"]) for c, x in _INS["ajustes"].items()}
# Asignaciones descartadas: la coincidencia de monto no tiene correspondencia funcional con el servicio
DESCARTADAS = dict(_INS.get("asignacion_descartada") or {})

clave = lambda c: [int(i) for i in c.split(".")]
serv = pd.read_csv("src/formulario_fie03/datos/servicios.csv")
asig = pd.read_csv("src/formulario_fie03/datos/destinatario_servicio.csv")

# Responsables
resp = {}
for _, a in asig.iterrows():
    c = a.codigo_corto
    et = MIMP_LIO[c] if a.sigla == "MIMP-PORCONFIRMAR" else ETIQUETA.get(a.sigla, a.sigla)
    resp.setdefault(c, []).append(et)

# Protocolos aprobados (hoja 3 del consolidado de la OMEP)
pr = pd.read_excel(CONSOLIDADO, sheet_name="3. Protocolos aprobados")
pr = pr[pr.iloc[:, 2].astype(str).str.match(r"^\d\.\d\.\d$")]
protocolos = {}
for _, r in pr.iterrows():
    protocolos.setdefault(r.iloc[2], []).append(f"{r.iloc[4]} [{r.iloc[1]}]")

# Valores de indicadores de servicio 2019–2025 (hoja 2)
mr = pd.read_excel(CONSOLIDADO, sheet_name="2. Metas y resultados", header=1)
mr = mr[mr["Código del indicador"].astype(str).str.startswith("ISS")]
anios_valor, n_ind = {}, {}
for _, r in mr.iterrows():
    m = re.match(r"ISS\.(\d+)\.(\d+)\.(\d+)", r["Código del indicador"])
    c = f"{int(m.group(1))}.{int(m.group(2))}.{int(m.group(3))}"
    n_ind[c] = n_ind.get(c, 0) + 1
    for a in range(2019, 2026):
        v = r[f"Valor {a}"]
        if pd.notna(v) and str(v).strip().upper() not in ("ND", "", "-", "NAN"):
            anios_valor.setdefault(c, set()).add(a)

# Formulario de seguimiento 2023–2025
pn = pd.read_csv(PANEL)
pn = pn[pn.nivel == "Servicio"]
cump25 = pn[pn.anio == 2025].groupby("servicio").cumplimiento.apply(lambda s: "; ".join(sorted(set(s)))).to_dict()
poi23 = pn[(pn.anio == 2023) & (pn.servicio_en_poi == "Sí")].servicio.unique().tolist()
poi23_detalle = pn[(pn.anio == 2023) & pn.actividades_poi.notna()].servicio.unique().tolist()

# Secciones del LIO por año
lio = pd.read_csv(LIO)


def ejec(r):
    if r.seccion != "Sí":
        return "No"
    return "Sí" if r.lineas >= MIN_LINEAS and r.sin_info == 0 else "Mínima"


lio["ejecucion"] = lio.apply(ejec, axis=1)
lio_tab = lio.pivot(index="servicio", columns="anio", values="ejecucion")
montos = lio[(lio.anio >= 2019) & (lio.monto_soles > 0)].groupby("servicio").anio.apply(lambda s: sorted(s)).to_dict()

# Montos reportados 2019–2025 verificados con los datos abiertos del MEF (REV_IND.3)
cv = pd.read_excel(VERIFICACION, sheet_name="Cifras verificadas", dtype={"Servicio": str})
cv = cv[cv.Servicio.notna() & cv["Año"].between(2019, 2025)]
cv["valida"] = (cv.Resultado.isin(RESULTADOS_VERIFICADOS) & (cv.Alcance == "Entidad responsable")
                & cv.Nivel.fillna("").str.startswith(NIVELES_SERVICIO)
                & ~cv["Observación"].fillna("").str.contains("casual"))


def estructura(g):
    """Fuente, página y estructura del MEF de las coincidencias válidas (las tres primeras)."""
    items = []
    for f, p, n, d in zip(g.Fuente, g["Página"], g.Nivel, g["Descripción MEF"]):
        fuente = f.replace("Informe ", "").replace("Formulario de seguimiento", "Formulario")
        pagina = f" p. {int(p)}" if pd.notna(p) else ""
        items.append(f"{fuente}{pagina}: {n.split(' (')[0].lower()} {d.split(' | ')[-1][:70].strip()}")
    items = list(dict.fromkeys(items))
    return "; ".join(items[:3]) + (f"; y {len(items) - 3} más" if len(items) > 3 else "")


verificadas = cv[cv.valida].groupby("Servicio").apply(estructura).to_dict()
n_valida = cv.groupby("Servicio").valida.sum().to_dict()
n_montos = cv.groupby("Servicio").size().to_dict()

filas = []
for _, s in serv.iterrows():
    c = s.codigo_corto
    tiene_prot = c in protocolos
    valores = sorted(anios_valor.get(c, []))
    reporta = any(2019 <= a <= 2025 for a in valores)
    asignacion = c in verificadas and c not in DESCARTADAS
    ej = {a: lio_tab.loc[c, a] for a in range(2019, 2026)}
    ejec_reciente = any(ej[a] == "Sí" for a in ESCALA["anios_cierre"])
    ejec_periodo = any(ej[a] == "Sí" for a in range(2019, 2026)) or bool(valores)
    faltan = [n for n, ok in zip(ESCALA["componentes"], (tiene_prot, asignacion, reporta)) if not ok]
    estado = clasificar_implementacion(faltan, ejec_reciente, ejec_periodo,
                                       all(ej[a] == "No" for a in range(2019, 2026)),
                                       ESCALA["minimo_componentes"], len(ESCALA["componentes"]))
    # sensibilidad (revisión institucional, observación 18): la regla anterior de tres de tres componentes (ADR-011)
    estado3 = clasificar_implementacion(faltan, ejec_reciente, ejec_periodo,
                                        all(ej[a] == "No" for a in range(2019, 2026)),
                                        len(ESCALA["componentes"]), len(ESCALA["componentes"]))
    automatico, ajuste = estado, None
    if c in AJUSTES:
        estado, ajuste = AJUSTES[c]
        estado3 = estado
    filas.append({
        "Código": c, "OP": s.op, "Lineamiento": s.lineamiento, "Servicio": s.servicio,
        "Entidad u órgano responsable": "; ".join(dict.fromkeys(resp.get(c, []))),
        "Protocolo aprobado": "Sí" if tiene_prot else "No",
        "Norma del protocolo": " | ".join(protocolos.get(c, [])) or None,
        "Indicadores del servicio": n_ind.get(c, 0),
        "Años con valor reportado": ", ".join(map(str, valores)) or None,
        "Nivel de cumplimiento 2025 (formulario)": cump25.get(c),
        "Asignación presupuestal identificable": "Sí" if asignacion else "No",
        "Montos 2019–2025 verificados en el MEF": f"{int(n_valida.get(c, 0))} de {n_montos.get(c, 0)}",
        "Estructura presupuestal verificada (MEF)": (f"{verificadas.get(c)} | Descartada: {DESCARTADAS[c]}"
                                                     if c in DESCARTADAS else verificadas.get(c)),
        "Declarado en el POI 2023 (formulario)": "Sí" if c in poi23 else "No",
        "LIO 2019–2025 con montos en soles": ", ".join(map(str, montos.get(c, []))) or None,
        **{f"LIO {a}": ej[a] for a in range(2019, 2026)},
        "Primer año con ejecución reportada": next((a for a in range(2019, 2026) if ej[a] == "Sí"), None),
        "Componentes faltantes": ", ".join(faltan) or None,
        "Evidencia de ejecución 2025 (LIO)": EVIDENCIA[c],
        "Clasificación por criterios": automatico,
        "Ajuste por lectura cualitativa": ajuste,
        "Estado asignado (preliminar)": estado,
        "Estado con tres de tres componentes (sensibilidad)": estado3,
    })
df = pd.DataFrame(filas).sort_values("Código", key=lambda col: col.map(clave)).reset_index(drop=True)

# Información reportada por las entidades (REP_ENT), por servicio
fie = pd.read_excel(FIE03, sheet_name="Codificación", dtype=str).fillna("")
fie = fie[fie.codigos != ""]
nombre = pd.read_csv(LIBRO, dtype=str).set_index("codigo").nombre
larga = fie.assign(codigo=fie.codigos.str.split(r";\s*")).explode("codigo")
larga = larga[larga.codigo.str.match(r"I[145]\.")]          # factores de implementación, coordinación y ajuste


def cita(r):
    fuente = r.fuente.replace("Informe ", "").replace(" de gobernanza de la CONAIG", " CONAIG")
    return f"{fuente}{' p. ' + r.pagina if r.pagina else ''} ({r.id}): {r.nota}"


def principales(g, sentidos, prefijos=("I1.", "I4."), n=3):
    """Códigos más frecuentes del servicio en los sentidos indicados, con su número de segmentos."""
    sub = g[g.sentido.isin(sentidos) & g.codigo.str.startswith(prefijos)]
    cuenta = sub.codigo.value_counts().head(n)
    return "; ".join(f"{nombre.get(k, k)} ({v})" for k, v in cuenta.items()) or None


def citas_dificultad(g):
    """Una cita por cada factor de dificultad principal: la de un hallazgo crítico si la hay, si no la primera."""
    sub = g[g.sentido.isin(["Dificultad", "Mixto"]) & g.codigo.str.startswith(("I1.", "I4."))]
    salida = []
    for k in sub.codigo.value_counts().head(3).index:
        filas_k = sub[sub.codigo == k].sort_values("hallazgo_critico", key=lambda s: s.ne("Confirmado"))
        salida.append(f"{nombre.get(k, k)}: {cita(filas_k.iloc[0])}")
    return " | ".join(salida) or None


por_servicio = {c: g for c, g in larga.groupby("atribuido a")}
criticos = fie[fie.hallazgo_critico == "Confirmado"].groupby("atribuido a")
df["Segmentos codificados (REP_ENT)"] = df["Código"].map(fie.groupby("atribuido a").size()).fillna(0).astype(int)
df["Factores de dificultad (REP_ENT)"] = df["Código"].map(
    lambda c: principales(por_servicio[c], ["Dificultad", "Mixto"]) if c in por_servicio else None)
df["Factores favorables (REP_ENT)"] = df["Código"].map(
    lambda c: principales(por_servicio[c], ["Favorable", "Mixto"]) if c in por_servicio else None)
df["Propuestas de ajuste (REP_ENT)"] = df["Código"].map(
    lambda c: principales(por_servicio[c], ["Propuesta"], prefijos=("I5.", "I1.", "I4.")) if c in por_servicio else None)
df["Citas de los factores de dificultad (REP_ENT)"] = df["Código"].map(
    lambda c: citas_dificultad(por_servicio[c]) if c in por_servicio else None)
df["Hallazgos críticos confirmados (REP_ENT)"] = df["Código"].map(
    lambda c: " | ".join(cita(r) for r in criticos.get_group(c).itertuples()) if c in criticos.groups else None)
detalle_rep_ent = (larga[larga["atribuido a"].isin(df["Código"])]
                 .groupby(["atribuido a", "codigo", "sentido"])
                 .agg(segmentos=("id", "count"), ids=("id", lambda s: ", ".join(s)))
                 .reset_index()
                 .assign(nombre=lambda d: d.codigo.map(nombre))
                 .rename(columns={"atribuido a": "Servicio", "codigo": "Código", "nombre": "Factor", "sentido": "Sentido",
                                  "segmentos": "Segmentos", "ids": "Segmentos (id)"})
                 [["Servicio", "Código", "Factor", "Sentido", "Segmentos", "Segmentos (id)"]]
                 .sort_values(["Servicio", "Segmentos"], ascending=[True, False],
                              key=lambda col: col.map(lambda c: tuple(clave(c))) if col.name == "Servicio" else col))

criterios = pd.DataFrame([
    ["Implementado", "Al menos dos de los tres componentes (protocolo aprobado, asignación presupuestal identificable y "
     "reporte de indicadores con valor en algún año de 2019–2025), con ejecución reportada en el LIO 2024 o 2025: el servicio "
     "sigue operando al cierre (ADR-011 y ADR-012)."],
    ["Implementado parcialmente", "Ejecución reportada en algún año de 2019–2025 (información detallada en el LIO o valor de "
     "indicador) con dos o más de los tres componentes faltantes, o interrumpida antes de 2024."],
    ["No implementado", "Sin información detallada del servicio en ningún LIO de 2019–2025 ni valores de indicador, o solo "
     "acciones preparatorias."],
    ["Sin información suficiente", "Solo secciones mínimas del LIO en 2019–2025, sin valores de indicador."],
    ["LIO (por año)", "Sí (información detallada): sección que describe las acciones del servicio sin declarar falta de "
     "información; Mínima: mención breve o con declaración de falta de información; No: sin sección ubicada."],
    ["Asignación presupuestal identificable", "Al menos un monto reportado para el servicio en los informes LIO 2019–2025 o "
     "en el formulario de seguimiento 2023 coincide con una meta, actividad o producto de la entidad responsable en los datos "
     "abiertos del gasto del MEF (REV_IND.3, hoja «Cifras verificadas»): coincidencia exacta, cercana a un agregado ya "
     "verificado en el mismo párrafo o con el año anterior. No cuentan las coincidencias con totales de programa, las de "
     "pliegos de otro sector ni las de montos redondos marcadas como posiblemente casuales. La declaración en el POI 2023 y "
     "los montos del LIO sin verificar se muestran como referencia, sin efecto en la clasificación."],
    ["Ajuste por lectura cualitativa", "Se considera ejecutado el servicio cuya provisión a la población objetivo reporta "
     "la entidad; las acciones preparatorias (diagnóstico, protocolo) o solo vinculadas no cuentan como ejecución. Cada "
     "ajuste se sustenta en la columna correspondiente."],
    ["Información de las entidades (REP_ENT)", "Segmentos codificados del formulario 2023–2025, de los informes LIO "
     "2019–2025 y de los informes de la CONAIG atribuidos al servicio (columna «atribuido a» de REP_ENT). Factores de "
     "dificultad: los tres códigos I1 o I4 más frecuentes en segmentos de sentido dificultad o mixto; favorables: en "
     "sentido favorable o mixto; propuestas: en sentido propuesta, incluidos los I5. Entre paréntesis, el número de "
     "segmentos. Cada factor de dificultad lleva una cita (la de un hallazgo crítico confirmado si existe). Los hallazgos "
     "críticos son los confirmados en la doble codificación. Sirven para contrastar el estado asignado; no lo modifican. "
     "El detalle por código y segmento está en la hoja «Factores REP_ENT»."],
    ["Fuentes", "Consolidado de indicadores de la OMEP (hojas 2 y 3); formulario de seguimiento 2023–2025; informes LIO "
     "2019–2025; órganos del MIMP sin asignación en la Política identificados en los informes LIO; datos abiertos del gasto "
     "del MEF (verificación de REV_IND.3)."],
], columns=["Elemento", "Criterio"])

with pd.ExcelWriter(SALIDA, engine="openpyxl") as xw:
    df.to_excel(xw, sheet_name="Estado por servicio", index=False)
    detalle_rep_ent.to_excel(xw, sheet_name="Factores REP_ENT", index=False)
    criterios.to_excel(xw, sheet_name="Criterios", index=False)
    sens = pd.DataFrame({
        "Regla": ["Dos de tres componentes y ejecución en 2024 o 2025 (vigente, ADR-012)",
                  "Tres de tres componentes y ejecución en 2024 o 2025 (anterior, ADR-011)"],
        **{e: [int((df["Estado asignado (preliminar)"] == e).sum()),
               int((df["Estado con tres de tres componentes (sensibilidad)"] == e).sum())]
           for e in ("Implementado", "Implementado parcialmente", "No implementado", "Sin información suficiente")}})
    sens.to_excel(xw, sheet_name="Sensibilidad", index=False)
    ws = xw.sheets["Estado por servicio"]
    ws.freeze_panes = "E2"
    anchos = {"Código": 8, "OP": 7, "Lineamiento": 30, "Servicio": 55, "Evidencia de ejecución 2025 (LIO)": 60,
              "Ajuste por lectura cualitativa": 45, "Estructura presupuestal verificada (MEF)": 70,
              "Factores de dificultad (REP_ENT)": 45, "Factores favorables (REP_ENT)": 45,
              "Propuestas de ajuste (REP_ENT)": 40, "Citas de los factores de dificultad (REP_ENT)": 90,
              "Hallazgos críticos confirmados (REP_ENT)": 90}
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = anchos.get(col[0].value, 16)
    xw.sheets["Criterios"].column_dimensions["A"].width = 32
    xw.sheets["Criterios"].column_dimensions["B"].width = 120
print(df["Estado asignado (preliminar)"].value_counts().to_string())
print(df[["Código", "Entidad u órgano responsable", "Protocolo aprobado", "Años con valor reportado",
          "Asignación presupuestal identificable", "LIO 2023", "LIO 2024", "LIO 2025", "Estado asignado (preliminar)"]]
      .to_string(index=False))
