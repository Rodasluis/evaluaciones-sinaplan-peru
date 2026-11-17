"""REV_IND.3: verificación de las cifras presupuestales reportadas sobre la PNIG con los datos abiertos del MEF, 2019–2025.

Cifras reportadas:
  - informes LIO 2019–2025 (docs/fuentes/pnig/06_informes_lio/): párrafos con montos en soles y palabras de presupuesto,
    con su página y el servicio de la sección donde aparecen;
  - formulario de seguimiento de la OMEP 2023: presupuesto programado y ejecutado de hasta cinco actividades operativas
    por servicio (preguntas 11 y 12; los formularios 2024 y 2025 ya no lo preguntan).
Datos del MEF: extractos del conjunto «Presupuesto y Ejecución de Gasto» (src/fuentes/filtrar_gasto_mef.py), que es la
Consulta Amigable con el detalle de meta, finalidad, fuente y genérica.

Verificación (equipo evaluador): cada monto se busca entre los totales del MEF del mismo año, en los pliegos de los
sectores de la entidad responsable del servicio (REV_DOC), en 15 niveles de agregación (pliego, unidad ejecutora,
programa, producto, actividad, meta, sus combinaciones y, para los programas presupuestales, el total de todos los
niveles de gobierno). Coincide si la diferencia es menor que la precisión con que se reportó la cifra (S/ 1 para soles
enteros); cercano si difiere a lo más 0,5 % del agregado que coincidió exacto con otra cifra del mismo párrafo. Si no hay
coincidencia en la entidad, los montos no redondos de S/ 1 millón o más se buscan en todo el Gobierno Nacional y la
coincidencia se marca para revisión. El concepto que coincide (PIA, PIM o devengado) se compara con el del informe.

Conducción: metas del pliego 039 cuyo nombre identifica a los órganos de la conducción (DGIGND, DGTEG/DTEG, DPIGND) y
las de la actividad 5001072 «Promoción de la mujer». Estructura orgánica: ROF 2012, 2021, 2023, 2024 y 2025
(docs/fuentes/planeamiento/rof_mimp/).

Uso (desde la raíz del repositorio):
    uv run --no-project --with openpyxl --with pandas python src/analisis_pnig/rev_ind_3_verificacion_presupuestal.py
"""
import re
import subprocess
from functools import lru_cache
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
from nucleo.entorno import ruta_externa  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

RAIZ = Path(EVAL)
MEF = RAIZ / "03_datos/raw/datos_abiertos_mef"
LIO = Path("docs/fuentes/pnig/06_informes_lio")
GRD = RAIZ / "04_analisis/REV_DOC_estado_implementacion_servicios.xlsx"
# libro original del formulario, en la carpeta administrativa (config/entorno.local.yaml); si no está, se usa el extracto
FORM_ORIGINAL = Path(ruta_externa("carpeta_administrativa") or "carpeta_administrativa_sin_configurar",
                     "Reportes de informacion", "Formulario de Seguimiento de Politicas MIMP 2023.xlsx")
FORM = RAIZ / "03_datos/raw/formulario_seguimiento_omep/2023_omep_formulario-seguimiento-pnig_presupuesto-ao.csv"
SALIDA = RAIZ / "04_analisis/REV_IND_3_verificacion_presupuestal.xlsx"
ANIOS = range(2019, 2026)
MONTOS = {"MONTO_PIA": "PIA", "MONTO_PIM": "PIM", "MONTO_DEVENGADO": "Devengado"}
TOL_CERCANO = 0.005
PROGRAMAS = ["0002", "0080", "0103", "0116", "0131", "1002"]  # con extracto subnacional (filtrar_gasto_mef.py)

# Sectores (código del MEF) de cada entidad responsable que nombra REV_DOC
SECTOR = {"MIMP": "39", "MINSA": "11", "MININTER": "07", "PJ": "04", "MPFN": "22", "MINJUSDH": "06", "INPE": "06",
          "SERVIR": "01", "MTPE": "12", "MINEDU": "10", "ONPE": "32", "JNE": "31", "MIDAGRI": "13", "ANA": "13",
          "COFOPRI": "37", "MVCS": "37", "MINAM": "05", "MINCETUR": "35", "PROMPERU": "35", "PRODUCE": "38",
          "FONDEPES": "38", "MIDIS": "40", "INEI": "01", "MEF": "09"}

# ---------------------------------------------------------------- cifras reportadas en los informes LIO
HEAD = [re.compile(r"^\s*servicio\s*(n[º°o]\.?\s*)?(\d\.\d\.\d)\b", re.I), re.compile(r"\(servicio\s+(\d\.\d\.\d)\)", re.I)]
CLAVE = re.compile(r"presupuest|\bPIM\b|\bPIA\b|devengad|asignaci[oó]n|ejecuci[oó]n presupuestal|ejecut", re.I)
NUM = r"(\d{1,4}(?:[ .,]\d{3})+(?:[.,]\d{1,2})?|\d+(?:[.,]\d{1,2})?)"
MONTO = re.compile(rf"S/\.?\s?{NUM}(\s?millones)?|{NUM}(\s?millones de)?\s?so-?\s?les", re.I)
CONCEPTO = [("PIA", r"\bPIA\b|apertura"),
            ("PIM", r"\bPIM\b|modificado|asignad|asignaci|program|presupuesto (total )?de|contó con"),
            ("Devengado", r"devengad|ejecut|ejecuci|gast")]


def numero(txt, millones):
    t = txt.replace(" ", "")
    if "," in t and "." in t:  # con ambos separadores, el último es el decimal (lo que sigue a 2 decimales es una nota)
        corte = max(t.rfind(","), t.rfind("."))
        v = float(re.sub(r"[.,]", "", t[:corte]) + "." + t[corte + 1:corte + 3])
        return v * 1e6 if millones else v
    if re.search(r"[.,]\d{1,2}$", t) and not re.search(r"[.,]\d{3}$", t):
        corte = t.rfind(",") if t.rfind(",") > t.rfind(".") else t.rfind(".")
        v = float(re.sub(r"[.,]", "", t[:corte]) + "." + t[corte + 1:])
    else:
        v = float(re.sub(r"[.,]", "", t))
    return v * 1e6 if millones else v


def concepto(previo):
    mejor, pos = "Sin clasificar", -1
    for c, r in CONCEPTO:
        for m in re.finditer(r, previo[-90:], re.I):
            if m.start() > pos:
                mejor, pos = c, m.start()
    return mejor


def cifras_lio():
    filas = []
    for y in ANIOS:
        txt = subprocess.run(["pdftotext", "-enc", "UTF-8", "-layout", str(LIO / f"{y}_mimp_informe-cumplimiento-lio.pdf"),
                              "-"], capture_output=True).stdout.decode("utf-8", "ignore")
        servicio = ""
        for p, pagina in enumerate(txt.split("\f"), start=1):
            bloques, actual = [], []
            for l in pagina.split("\n"):
                if "....." in l:
                    continue
                m19 = re.match(r"^\s*6\.(\d)\.(\d)\.\d+\.\s+Servicio (\d+)\.", l) if y == 2019 else None
                if m19:  # 2019: «6.X.Y.Z. Servicio N.»; el lineamiento es Y-1 salvo en el OP5
                    op, lin, n = map(int, m19.groups())
                    servicio = f"{op}.{lin - (0 if op == 5 else 1)}.{n}"
                for h in HEAD:
                    m = h.search(l)
                    if m:
                        servicio = [g for g in m.groups() if g and re.fullmatch(r"\d\.\d\.\d", g)][0]
                if l.strip():
                    actual.append((servicio, l.strip()))
                elif actual:
                    bloques.append(actual)
                    actual = []
            bloques += [actual] if actual else []
            for b in bloques:
                texto = re.sub(r"\s+", " ", " ".join(x for _, x in b))
                if not (MONTO.search(texto) and CLAVE.search(texto)) or "Recuperado de" in texto or "http" in texto:
                    continue
                for m in MONTO.finditer(texto):
                    try:
                        v = numero(m.group(1) or m.group(3), m.group(2) or m.group(4))
                    except ValueError:
                        continue
                    if v >= 1000:
                        filas.append({"Fuente": f"Informe LIO {y}", "Año": y, "Página": p, "Servicio": b[0][0],
                                      "Concepto reportado": concepto(texto[:m.start()]), "Monto reportado": v,
                                      "Precisión": precision(m.group(0), v), "Texto": texto[:700]})
    return pd.DataFrame(filas)


# ---------------------------------------------------------------- cifras del formulario 2023
def cifras_formulario():
    if FORM_ORIGINAL.exists():  # el libro original queda fuera del repositorio; se extraen solo las columnas de presupuesto
        d = pd.read_excel(FORM_ORIGINAL, "Data", dtype=str)
        d = d[d["Política"].eq("PNIG")]
        filas = []
        for _, r in d.iterrows():
            for k in range(1, 6):
                ao = str(r[f"Actividad operativa {k} (AO{k})"]).strip(" -")
                if ao in ("", "nan"):
                    continue
                filas.append({"indicador": r["Indicador"], "responsable": r["Responsable"], "servicio": r["Servicio"],
                              "ao": ao, "aei": str(r[f"AEI ({k}) asociada al servicio"]).strip(" -"),
                              "programado": r[f"Presupuesto programado AO{k}"],
                              "ejecutado": r[f"Presupuesto ejecutado AO{k}"]})
        pd.DataFrame(filas).drop_duplicates(["responsable", "ao", "programado", "ejecutado"]).to_csv(FORM, index=False)
    f = pd.read_csv(FORM, dtype=str)
    filas = []
    for _, r in f.iterrows():
        cod = re.search(r"\b(\d{2}\.\d{2}\.\d{2})\b", str(r.servicio))
        for col, conc in (("programado", "PIM"), ("ejecutado", "Devengado")):
            t = re.sub(r"[^\d.,]", "", str(r[col]))
            if not t:
                continue
            try:
                v = numero(t, None)
            except ValueError:
                continue
            filas.append({"Fuente": "Formulario de seguimiento 2023", "Año": 2023, "Página": "",
                          "Servicio": ".".join(str(int(x)) for x in cod.group(1).split(".")) if cod else "",
                          "Entidad informante": r.responsable, "Concepto reportado": conc, "Monto reportado": v,
                          "Precisión": 1.0, "Texto": f"{r.ao} | {r.aei}"})
    # la misma actividad operativa se repite en cada indicador del servicio
    return pd.DataFrame(filas).drop_duplicates(["Entidad informante", "Concepto reportado", "Monto reportado"])


# ---------------------------------------------------------------- totales del MEF por nivel de agregación
NIVELES = {
    "Pliego": ["PLIEGO"], "Unidad ejecutora": ["PLIEGO", "EJECUTORA"], "Programa": ["PLIEGO", "PROGRAMA_PPTO"],
    "Producto": ["PLIEGO", "PROGRAMA_PPTO", "PRODUCTO_PROYECTO"],
    "Actividad": ["PLIEGO", "PROGRAMA_PPTO", "PRODUCTO_PROYECTO", "ACTIVIDAD_ACCION_OBRA"],
    "Actividad (todos los productos)": ["PLIEGO", "ACTIVIDAD_ACCION_OBRA"],
    "Unidad ejecutora y actividad": ["PLIEGO", "EJECUTORA", "ACTIVIDAD_ACCION_OBRA"],
    "Unidad ejecutora y programa": ["PLIEGO", "EJECUTORA", "PROGRAMA_PPTO"],
    "Unidad ejecutora y producto": ["PLIEGO", "EJECUTORA", "PRODUCTO_PROYECTO"],
    "Meta": ["PLIEGO", "EJECUTORA", "SEC_FUNC"], "Metas del mismo nombre": ["PLIEGO", "EJECUTORA", "META_NOMBRE"],
    "Pliego y fuente": ["PLIEGO", "FUENTE_FINANCIAMIENTO"], "Programa y fuente": ["PLIEGO", "PROGRAMA_PPTO", "FUENTE_FINANCIAMIENTO"],
    "Actividad y fuente": ["PLIEGO", "PRODUCTO_PROYECTO", "ACTIVIDAD_ACCION_OBRA", "FUENTE_FINANCIAMIENTO"],
    "Sector y programa": ["SECTOR", "PROGRAMA_PPTO"]}
NOMBRE = {"PLIEGO": "PLIEGO_NOMBRE", "EJECUTORA": "EJECUTORA_NOMBRE", "PROGRAMA_PPTO": "PROGRAMA_PPTO_NOMBRE",
          "PRODUCTO_PROYECTO": "PRODUCTO_PROYECTO_NOMBRE", "ACTIVIDAD_ACCION_OBRA": "ACTIVIDAD_ACCION_OBRA_NOMBRE",
          "SEC_FUNC": "META_NOMBRE", "FUENTE_FINANCIAMIENTO": "FUENTE_FINANCIAMIENTO_NOMBRE", "SECTOR": "SECTOR_NOMBRE"}


@lru_cache(maxsize=None)
def leer(anio, nombre):
    """Extracto del MEF de un año (se lee una vez; quien lo usa filtra, no modifica)."""
    d = pd.read_csv(MEF / f"{anio}_mef_gasto_{nombre}.csv", dtype=str, keep_default_na=False)
    for c in [c for c in d.columns if c.startswith("MONTO")]:
        d[c] = d[c].astype(float)
    return d


def describir(fila, claves):
    partes = []
    for k in claves:
        n = NOMBRE.get(k)
        partes.append(f"{fila[k]} {fila[n]}".strip() if n and n in fila else str(fila[k]))
    return " | ".join(partes)


def candidatos(anio):
    nac = leer(anio, "gobierno-nacional_meta")
    out = []
    for nivel, claves in NIVELES.items():
        nombres = [NOMBRE[k] for k in claves if k in NOMBRE and NOMBRE[k] not in claves]
        g = nac.groupby(claves + ["SECTOR"] * ("SECTOR" not in claves), as_index=False).agg(
            {**{m: "sum" for m in MONTOS}, **{n: "first" for n in nombres}})
        g["Nivel"] = nivel
        g["Descripción"] = [describir(r, claves) for _, r in g.iterrows()]
        out.append(g[["SECTOR", "Nivel", "Descripción", *MONTOS]])
    # programas presupuestales: total de todos los niveles de gobierno y del nivel regional y local
    sub = leer(anio, "programas-pnig_subnacional")
    tot = leer(anio, "programas_totales")
    t = tot.groupby(["PROGRAMA_PPTO", "PROGRAMA_PPTO_NOMBRE"], as_index=False)[list(MONTOS)].sum()
    t["Nivel"], t["SECTOR"] = "Programa (todos los niveles de gobierno)", "*"
    t["Descripción"] = t.PROGRAMA_PPTO + " " + t.PROGRAMA_PPTO_NOMBRE
    out.append(t[["SECTOR", "Nivel", "Descripción", *MONTOS]])
    t = tot[tot.NIVEL_GOBIERNO.isin(["E", "R"])].groupby(["PROGRAMA_PPTO", "PROGRAMA_PPTO_NOMBRE"], as_index=False)[
        list(MONTOS)].sum()
    t["Nivel"], t["SECTOR"] = "Programa (nacional y regional)", "*"
    t["Descripción"] = t.PROGRAMA_PPTO + " " + t.PROGRAMA_PPTO_NOMBRE
    out.append(t[["SECTOR", "Nivel", "Descripción", *MONTOS]])
    # producto y actividad de los programas con extracto subnacional: suma de todos los niveles de gobierno
    for clave, nivel in (("PRODUCTO_PROYECTO", "Producto (todos los niveles de gobierno)"),
                         ("ACTIVIDAD_ACCION_OBRA", "Actividad (todos los niveles de gobierno)")):
        claves = ["PROGRAMA_PPTO", clave]
        n2 = nac[nac.PROGRAMA_PPTO.isin(sub.PROGRAMA_PPTO.unique())]
        g = pd.concat([n2.groupby(claves + [NOMBRE[clave]], as_index=False)[list(MONTOS)].sum(),
                       sub.groupby(claves + [NOMBRE[clave]], as_index=False)[list(MONTOS)].sum()])
        g = g.groupby(claves, as_index=False).agg({**{m: "sum" for m in MONTOS}, NOMBRE[clave]: "first"})
        g["Nivel"], g["SECTOR"] = nivel, "*"
        g["Descripción"] = g.PROGRAMA_PPTO + " | " + g[clave] + " " + g[NOMBRE[clave]]
        out.append(g[["SECTOR", "Nivel", "Descripción", *MONTOS]])
    for claves, nivel in ((["NIVEL_GOBIERNO", "PROGRAMA_PPTO"], "Programa (nivel regional o local)"),
                          (["NIVEL_GOBIERNO", "PROGRAMA_PPTO", "PRODUCTO_PROYECTO"], "Producto (nivel regional o local)"),
                          (["NIVEL_GOBIERNO", "PROGRAMA_PPTO", "ACTIVIDAD_ACCION_OBRA"],
                           "Actividad (nivel regional o local)")):
        nombres = [NOMBRE[k] for k in claves if k in NOMBRE]
        g = sub.groupby(claves, as_index=False).agg({**{m: "sum" for m in MONTOS}, **{n: "first" for n in nombres}})
        g["Nivel"], g["SECTOR"] = nivel, "*"
        g["Descripción"] = [describir(r, claves) for _, r in g.iterrows()]
        out.append(g[["SECTOR", "Nivel", "Descripción", *MONTOS]])
    c = pd.concat(out, ignore_index=True)
    c["PP"] = c["Descripción"].str.extract(r"(?:^|\| )(\d{4}) ", expand=False)  # programa presupuestal, si lo hay
    c = c.melt(id_vars=["SECTOR", "Nivel", "Descripción", "PP"], var_name="campo", value_name="valor")
    c = c[c.valor > 0]
    c["Concepto MEF"] = c.campo.map(MONTOS)
    return c.drop(columns="campo")


PRIORIDAD = {n: i for i, n in enumerate(["Actividad", "Producto", "Programa", "Unidad ejecutora", "Pliego",
                                          "Actividad (todos los productos)", "Unidad ejecutora y actividad",
                                          "Unidad ejecutora y producto", "Unidad ejecutora y programa",
                                          "Programa (todos los niveles de gobierno)", "Sector y programa", "Meta",
                                          "Metas del mismo nombre", "Actividad y fuente", "Programa y fuente",
                                          "Pliego y fuente"])}


def elegir(hit, tipo, alcance):
    hit = hit.assign(p=hit.Nivel.map(PRIORIDAD).fillna(99))
    h = hit.sort_values(["dif", "p"]).iloc[0]
    pliegos = hit["Descripción"].str.split(" | ", regex=False).str[0].nunique()
    return {"Resultado": tipo, "Alcance": alcance, "Concepto MEF": h["Concepto MEF"], "Nivel": h.Nivel,
            "Descripción MEF": h["Descripción"], "Monto MEF": h.valor, "Pliegos con el mismo monto": pliegos}


# Programas presupuestales cuyos totales nacionales, regionales y locales se buscan para cada sector
PROGRAMAS_SECTOR = {"11": ["0002", "0131", "1002"], "39": ["0080", "1002"], "07": ["0080", "1002"], "04": ["0080", "1002"],
                    "22": ["0080", "1002"], "06": ["0080", "1002"], "12": ["0103", "0116"]}


def totales_permitidos(c, sectores):
    pps = {p for s in sectores for p in PROGRAMAS_SECTOR.get(s, [])}
    return c.SECTOR.eq("*") & c.PP.isin(pps)


def buscar(c, monto, precision, sectores):
    """Coincide: diferencia menor que la precisión con que se reportó la cifra, en los pliegos de la entidad y en los
    totales de sus programas (o en todo el Gobierno Nacional para montos no redondos de S/ 1 millón o más, marcado para
    revisión)."""
    redondo = monto % 1000 == 0  # los montos redondos solo se buscan en los pliegos propios de la entidad
    propios = c.SECTOR.isin(sectores) | (totales_permitidos(c, sectores) & (not redondo))
    for alcance, sub, ok in (("Entidad responsable", c[propios], True),
                             ("Otra entidad (revisar)", c, monto >= 1e6 and not redondo)):
        if not ok:
            continue
        d = (sub.valor - monto).abs()
        hit = sub[d <= precision].assign(dif=d[d <= precision])
        if len(hit):
            return elegir(hit, "Coincide", alcance)
    return {"Resultado": "Sin coincidencia"}


MOTIVOS = [  # lectura automática del párrafo (equipo evaluador); la primera coincidencia gana
    (r"cr[eé]dito|subvenci|cofinanci|fondo AGROPER|planes de negocio|ingreso promedio|perciben",
     "No es gasto del pliego: créditos, subvenciones, cofinanciamientos o ingresos"),
    (r"\bLey\b|Decreto|transferencia de partidas|Ley de Presupuesto", "Monto de una norma o transferencia, no de ejecución"),
    (r"programas presupuestales vinculados|GORE|gobiernos regionales|pliegos", "Total de varias entidades o programas"),
    (r"actividad(es)? operativa|centro(s)? de costo|\bAOI?\d|\bAO\b|tarea|POI|Lima Metropolitana|protocolo",
     "Actividad operativa, centro de costo o tarea del POI: el MEF no la separa"),
]


def motivo(texto):
    for r, m in MOTIVOS:
        if re.search(r, str(texto), re.I):
            return m
    return "Sin explicación en el texto: monto agregado por la entidad o de otro corte (revisar)"


def cercano(c, monto, descripciones):
    """Cercano: a lo más 0,5 % de un agregado que coincidió exacto con otra cifra del mismo párrafo (suele ser el devengado
    o el PIM a otro corte del mismo agregado)."""
    sub = c[c["Descripción"].isin(descripciones)]
    d = (sub.valor - monto).abs()
    hit = sub[d <= TOL_CERCANO * monto].assign(dif=d[d <= TOL_CERCANO * monto])
    return elegir(hit, "Cercano (≤ 0,5 % del mismo agregado)", "Entidad responsable") if len(hit) else None


def precision(texto_monto, valor):
    """Media unidad del último dígito reportado: S/ 1 para soles enteros, S/ 5000 para «38,78 millones»."""
    if valor >= 1e6 and re.search(r"millones", texto_monto, re.I):
        dec = re.search(r"[.,](\d+)\s*millones", texto_monto)
        return 0.5 * 10 ** (6 - (len(dec.group(1)) if dec else 0))
    return 1.0


# ---------------------------------------------------------------- conducción: metas del pliego 039 por órgano
# Órgano de cada meta según su nombre (equipo evaluador); el orden importa: la primera coincidencia gana
ORGANOS = [  # (patrón, órgano, rol, grado de la asignación)
    (r"TRANSVERSALIZ", "DGTEG (2019–2023) / DTEG (desde el ROF 2023)", "Conducción: transversalización (OP5)", "Nombre"),
    (r"GESTION PUBLICA CON ENFOQUE DE GENERO", "DGTEG (2019–2023) / DTEG (desde el ROF 2023)",
     "Conducción: transversalización (OP5)", "Probable"),
    (r"DERECHOS DE LAS MUJERES", "DPPDM", "Conducción: patrones socioculturales y derechos (OP3, OP6)", "Nombre"),
    (r"PATRONES SOCIOCULTURALES|GOBERNANDO JUNTAS", "DPPDM", "Conducción: patrones socioculturales y derechos (OP3, OP6)",
     "Probable"),
    (r"AUTONOMIA ECONOMICA|CUIDADO", "DPDAEM (2021–2023) / DGPDAEM (desde el ROF 2023)", "Servicios del OP4", "Nombre"),
    (r"POLITICAS NACIONALES Y SECTORIALES EN MATERIA DE IGUALDAD|POLITICAS DE IGUALDAD DE GENERO|SEGUIMIENTO Y "
     r"EVALUACION A LA POLITICA", "DPIGND", "Conducción: diseño, seguimiento y evaluación", "Nombre"),
    (r"SEGUIMIENTO DE LOS AVANCES DE LA IGUALDAD", "DPIGND", "Conducción: diseño, seguimiento y evaluación", "Probable"),
    (r"IGUALDAD DE GENERO", "DGIGND", "Conducción (órgano conductor)", "Nombre"),
    (r"VIOLENCIA", "DGCVG", "Fuera de la conducción: rectoría en violencia (OP1)", "Nombre"),
]
# POI 2024 consistente con el PIA (formato con tareas): comprobación puntual, total anual leído en la página indicada
POI_2024 = [("AOI00108700577", "DGIGND", 52, 338046), ("AOI00108700570", "DPPDM", 54, 502802),
            ("AOI00108700657", "DPDAEM (2021–2023) / DGPDAEM (desde el ROF 2023)", 56, 984221),
            ("AOI00108700660", "DGTEG (2019–2023) / DTEG (desde el ROF 2023)", 173, 957146)]
POI_PIA = {2021: "2020_mimp_poi-2021-consistente-con-el-pia.pdf", 2022: "2021_mimp_poi-2022-consistente-con-el-pia.pdf",
           2023: "2022_mimp_poi-2023-consistente-con-el-pia.pdf"}
POI = Path("docs/fuentes/planeamiento/poi_mimp")


def organo(nombre):
    for r, o, rol, grado in ORGANOS:
        if re.search(r, nombre):
            return o, rol, grado
    return "Otro órgano", "", ""


def conduccion():
    filas = []
    for anio in ANIOS:
        n = leer(anio, "gobierno-nacional_meta")
        m = n[(n.PLIEGO == "039") & (n.EJECUTORA == "001")]
        m = m[m.ACTIVIDAD_ACCION_OBRA.eq("5001072") |
              m.META_NOMBRE.str.contains(r"IGUALDAD DE GENERO|TRANSVERSALIZ|DERECHOS DE LAS MUJERES|AUTONOMIA "
                                         r"ECONOMICA DE LA")]
        g = m.groupby(["SEC_FUNC", "PROGRAMA_PPTO", "ACTIVIDAD_ACCION_OBRA", "ACTIVIDAD_ACCION_OBRA_NOMBRE", "META_NOMBRE"],
                      as_index=False)[list(MONTOS)].sum()
        for _, r in g.iterrows():
            o, rol, grado = organo(r.META_NOMBRE)
            filas.append({"Año": anio, "Meta (sec. func.)": int(r.SEC_FUNC), "Programa": r.PROGRAMA_PPTO,
                          "Actividad": f"{r.ACTIVIDAD_ACCION_OBRA} {r.ACTIVIDAD_ACCION_OBRA_NOMBRE}",
                          "Nombre de la meta": r.META_NOMBRE, "Órgano": o, "Asignación del órgano": grado,
                          "Rol en la PNIG": rol,
                          "PIA": r.MONTO_PIA, "PIM": r.MONTO_PIM, "Devengado": r.MONTO_DEVENGADO})
    d = pd.DataFrame(filas).sort_values(["Año", "Órgano", "Meta (sec. func.)"])
    d["Ejecución (%)"] = (d.Devengado / d.PIM * 100).round(1)
    return d


def poi_pia():
    """Programación financiera anual por actividad operativa en el POI consistente con el PIA (Anexo B-5, 2021–2023)."""
    num = re.compile(r"\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+\.\d+|\d+")
    filas = []
    for anio, pdf in POI_PIA.items():
        txt = subprocess.run(["pdftotext", "-enc", "UTF-8", "-layout", str(POI / pdf), "-"],
                             capture_output=True).stdout.decode("utf-8")
        cc, ao = None, None
        for p, pag in enumerate(txt.split("\f"), start=1):
            lineas = pag.split("\n")
            for i, l in enumerate(lineas):
                m = re.search(r"Centro de Costo:\s*([\d.]+\s*-\s*)?(.+?)\s*$", l)
                if m:
                    cc = re.sub(r"\s+", " ", m.group(2)).strip()
                a = re.search(r"AOI\d{11}", l)
                interes = cc and re.search(r"IGUALDAD DE G|TRANSVERSALIZACI|DERECHOS DE LAS MUJERES|AUTONOM|"
                                           r"ARTICULACI.N (SECTORIAL|CON LOS GOB)", cc, re.I)
                if a:
                    ao = a.group(0)
                    dec = [float(x.replace(",", "")) for x in re.findall(r"\d{1,3}(?:,\d{3})*\.\d{2}\b", l)]
                    if interes and dec and "Financiero" not in l:  # sin línea «Financiero»: monto en la línea de la AO
                        filas.append({"Año": anio, "Documento": pdf, "Página": p, "Centro de costo": cc,
                                      "Actividad operativa": ao, "Programación financiera (S/)": max(dec),
                                      "Lectura": "monto anual en la línea de la actividad operativa (revisar)"})
                if not (interes and "Financiero" in l):
                    continue
                total, como = None, ""
                for j in range(i, min(i + 3, len(lineas))):
                    v = [float(x.replace(",", "")) for x in num.findall(lineas[j].split("S/.")[-1])]
                    if len(v) >= 3 and abs(sum(v[:-1]) - v[-1]) <= 2:  # meses con valor y total al final
                        total, como = v[-1], "total anual del POI (cuadra con los meses)"
                        break
                    if len(v) >= 3 and abs(sum(v[:-1]) - v[-1]) <= 0.01 * v[-1]:  # algún mes no se lee en el PDF
                        total, como = v[-1], "total anual del POI (los meses legibles suman 99 % o más)"
                        break
                    for s in range(len(v) - 12):
                        if v[s + 12] > 0 and abs(sum(v[s:s + 12]) - v[s + 12]) <= 2:
                            total, como = v[s + 12], "total anual del POI (cuadra con los 12 meses)"
                            break
                    if total is None and len(v) >= 12:
                        total, como = sum(v[-12:]), "suma de los 12 meses"
                    if total is not None:
                        break
                filas.append({"Año": anio, "Documento": pdf, "Página": p, "Centro de costo": cc,
                              "Actividad operativa": ao, "Programación financiera (S/)": total, "Lectura": como})
    d = pd.DataFrame(filas)
    # prevalece la línea «Financiero» con total; luego el monto de la línea de la AO; al final, sin monto
    d["orden"] = d.Lectura.str.startswith("monto anual en la línea").astype(int) + 2 * d[
        "Programación financiera (S/)"].isna()
    d = d.sort_values(["Año", "Actividad operativa", "orden"]).drop_duplicates(["Año", "Actividad operativa"])
    d = d.drop(columns="orden").sort_values(["Año", "Página"])
    d["Órgano"] = [organo(re.sub(r"[ÁÉÍÓÚ]", lambda x: "AEIOU"["ÁÉÍÓÚ".index(x.group(0))], c.upper()))[0]
                   if not re.search(r"ARTICULACI", c, re.I) else organo("TRANSVERSALIZACION")[0]
                   for c in d["Centro de costo"]]
    return d


def poi_vs_mef(poi, cond):
    p = poi.groupby(["Año", "Órgano"], as_index=False).agg(
        **{"Programación financiera (S/)": ("Programación financiera (S/)", "sum"),
           "AO sin monto legible": ("Programación financiera (S/)", lambda s: ", ".join(
               poi.loc[s[s.isna()].index, "Actividad operativa"]))})
    m = cond.groupby(["Año", "Órgano"], as_index=False)[["PIA", "PIM", "Devengado"]].sum()
    d = p.merge(m, on=["Año", "Órgano"], how="outer").sort_values(["Año", "Órgano"])
    d["POI − PIA del MEF"] = d["Programación financiera (S/)"] - d.PIA
    d = d[d["Año"].isin(POI_PIA)].assign(Comparación="Órgano (suma de sus actividades operativas y metas)")
    # 2024: comprobación puntual de una actividad operativa por órgano contra la meta del MEF con el mismo PIA
    m24 = cond[cond["Año"] == 2024]
    filas = []
    for ao, org, pag, monto in POI_2024:
        meta = m24[(m24["Órgano"] == org) & (m24.PIA.sub(monto).abs() <= 1)]
        filas.append({"Año": 2024, "Órgano": org, "Programación financiera (S/)": monto,
                      "AO sin monto legible": "", "PIA": meta.PIA.sum() if len(meta) else None,
                      "PIM": meta.PIM.sum() if len(meta) else None,
                      "Devengado": meta.Devengado.sum() if len(meta) else None,
                      "POI − PIA del MEF": monto - meta.PIA.sum() if len(meta) else None,
                      "Comparación": f"Actividad operativa {ao} (POI 2024 de apertura, p. {pag}) y meta "
                                     f"{', '.join(map(str, meta['Meta (sec. func.)']))} del MEF"})
    return pd.concat([d, pd.DataFrame(filas)], ignore_index=True)


ESTRUCTURA = [
    ("2019 a julio de 2021", "ROF aprobado por D.S. N.° 003-2012-MIMP, modificado por D.S. N.° 002-2015-MIMP y 004-2015-MIMP",
     "2016_mimp_rof-texto-con-modificaciones-2015.pdf", "Arts. 73 y 75 (p. 48–49); art. 78 (p. 51)",
     "DGIGND (con la Dirección de Políticas de Igualdad de Género y no Discriminación y la Dirección de Promoción y "
     "Protección de los Derechos de la Mujer) y DGTEG, ambas dependientes del Despacho Viceministerial de la Mujer. El ROF "
     "no menciona la PNIG (es anterior). POI: centros de costo 03.02 DGIGND y 03.03 DGTEG."),
    ("Julio de 2021 a octubre de 2023", "ROF aprobado por D.S. N.° 012-2021-MIMP (Sección Primera) y R.M. N.° 201-2021-MIMP "
     "(Sección Segunda); texto integrado R.M. N.° 208-2021-MIMP", "2021_mimp_rof-texto-integrado.pdf",
     "Art. 74 (p. 47); art. 76 (p. 48–49); arts. 83 y 85 (p. 52–54)",
     "La DGIGND es «el órgano de línea encargado de la conducción de la Política Nacional de Igualdad de Género» (art. 74), "
     "con DPIGND, DPPDM y la Dirección de Promoción y Desarrollo de la Autonomía Económica de las Mujeres. La DGTEG sigue "
     "como dirección general, con la Dirección de Articulación Sectorial e Interinstitucional y la de Articulación con los "
     "Gobiernos Regionales y Locales. El POI 2022 conserva los códigos antiguos (03.02 y 03.03); el de 2023 usa 07.01.0x y 07.02.0x."),
    ("Octubre de 2023 a 2025", "ROF aprobado por D.S. N.° 010-2023-MIMP (Sección Primera) y R.M. N.° 342-2023-MIMP (Sección "
     "Segunda); textos integrados R.M. N.° 362-2023-MIMP, 306-2024-MIMP y 380-2025-MIMP; modificaciones R.M. N.° 294-2024-MIMP "
     "y 367-2025-MIMP", "2023_mimp_rof-texto-integrado.pdf; 2025_mimp_rof-texto-integrado.pdf",
     "2023: art. 74 (p. 51), art. 76 (p. 53), art. 81 (p. 56), art. 83 (p. 58); 2025: art. 74 (p. 50), art. 76 (p. 52)",
     "La DGIGND conduce la PNIG y absorbe la transversalización: sus unidades son DPIGND, DPPDM y la Dirección de "
     "Transversalización del Enfoque de Género (DTEG). La autonomía económica pasa a la Dirección General de Promoción y "
     "Desarrollo de la Autonomía Económica de las Mujeres (DGPDAEM). Las modificaciones de 2024 y 2025 cambian funciones, no "
     "la estructura. POI: 07.01.01 DPIGND, 07.01.02 DPPDM, 07.01.04 DGIGND, 07.01.05 DTEG."),
]


# ---------------------------------------------------------------- conciliaciones y modificaciones presupuestarias
LIO_PDF = "Informe LIO {a}"
EVAL22 = "Informe de evaluación de resultados del PEI y del POI 2022, Tabla 8 (AEI.07.01, 5 AO)"
# (año, objeto, fuente, página, concepto, monto reportado, selección en el MEF, lectura del equipo evaluador)
CONCILIACIONES = [
    (2022, "Transversalización", EVAL22, 36, "PIM", 2492490, {"ue": "001", "organo": "DGTEG"},
     "El devengado coincide al sol; el PIM de la evaluación supera al final del MEF en S/ 10 649: corte anterior a una "
     "rebaja de fin de año, coherente con la certificación anulada en noviembre y diciembre. La meta cedió S/ 417 664 entre "
     "el PIA y el PIM. La evaluación es de 2022, no de 2023."),
    (2022, "Transversalización", EVAL22, 36, "Devengado", 2451993, {"ue": "001", "organo": "DGTEG"}, ""),
    (2023, "Transversalización", "POI 2023 consistente con el PIA (AEI.07.01)", 12, "PIA", 2763058,
     {"ue": "001", "organo": "DGTEG"}, "Suma de las metas 134, 138 y 139 (DGTEG y sus dos direcciones de articulación)."),
    (2023, "Transversalización", "Reporte de seguimiento anual del POI 2023 (AO de la AEI.07.01)", 16, "Devengado",
     2780330.32, {"ue": "001", "organo": "DGTEG"}, "El PIM del MEF es S/ 2 817 596; el reporte no informa PIM."),
    (2019, "CEM", LIO_PDF.format(a=2019), 116, "PIM", 158049395, {"ue": "009", "act": ["5006193"]}, ""),
    (2019, "CEM", LIO_PDF.format(a=2019), 116, "Devengado", 156915520.31, {"ue": "009", "act": ["5006193"]},
     "Devengado reportado a un corte anterior al cierre."),
    (2021, "CEM", LIO_PDF.format(a=2021), 106, "PIA", 190522720, {"ue": "009", "act": ["5006193"], "sin": ["35", "36", "37"]},
     "Coincide la actividad 5006193 sin las metas 35, 36 y 37 (Lima, finalidad 0258395). Diferencia identificada en las "
     "metas: el informe excluye esas tres metas y los datos públicos no indican qué servicios financian."),
    (2021, "CEM", LIO_PDF.format(a=2021), 106, "PIM", 184560207, {"ue": "009", "act": ["5006193"], "sin": ["35", "36", "37"]},
     "Con la misma agregación, el PIM reportado supera al final: corte anterior a una rebaja del PIM, que coincide con la "
     "certificación anulada en diciembre. En 2021 AURORA pasó de S/ 352,2 a 295,2 millones entre el PIA y el PIM."),
    (2021, "CEM", LIO_PDF.format(a=2021), 106, "Devengado", 167522847.35,
     {"ue": "009", "act": ["5006193"], "sin": ["35", "36", "37"]},
     "El devengado reportado está entre el acumulado a noviembre y el final: corte en diciembre."),
    (2022, "CEM", LIO_PDF.format(a=2022), 54, "PIM", 177058173, {"ue": "009", "act": ["5006193"], "sin": ["41", "42", "43"]},
     "Coincide la actividad 5006193 sin las metas 41, 42 y 43 (Lima, finalidad 0258395). Diferencia identificada en las "
     "metas: el informe excluye esas tres metas y los datos públicos no indican qué servicios financian."),
    (2022, "CEM", LIO_PDF.format(a=2022), 54, "Devengado", 175531374,
     {"ue": "009", "act": ["5006193"], "sin": ["41", "42", "43"]}, "Diferencia de corte en el devengado."),
    (2023, "CEM", LIO_PDF.format(a=2023), 42, "PIM", 188097215,
     {"ue": "009", "act": ["5006357"], "meta": r"^SERVICIOS ESPECIALIZADOS E INTERDISCIPLINARIOS"},
     "Metas de los CEM de la actividad 5006357 (sin el Programa Local de Prevención). El PIM reportado supera al final: corte "
     "anterior a una rebaja, coherente con la certificación anulada en noviembre y diciembre; no hay combinación de hasta "
     "tres metas que dé el monto."),
    (2023, "CEM", LIO_PDF.format(a=2023), 42, "Devengado", 182877558,
     {"ue": "009", "act": ["5006357"], "meta": r"^SERVICIOS ESPECIALIZADOS E INTERDISCIPLINARIOS"},
     "El devengado reportado está entre el acumulado a noviembre y el final: corte en diciembre."),
    (2024, "CEM", LIO_PDF.format(a=2024), 35, "PIM", 202964800, {"ue": "009", "act": ["5006357"]}, ""),
    (2024, "CEM", LIO_PDF.format(a=2024), 36, "Devengado", 202251880, {"ue": "009", "act": ["5006357"]}, ""),
    (2025, "CEM", LIO_PDF.format(a=2025), 39, "PIM", 220440058, {"ue": "009", "act": ["5006357"]},
     "El PIM y el devengado reportados superan a los finales del MEF; un corte anterior explicaría el PIM, pero no un "
     "devengado mayor. Diferencia identificada: la cifra incluye gasto que el MEF no registra en la actividad, y los datos "
     "públicos no permiten ubicarlo."),
    (2025, "CEM", LIO_PDF.format(a=2025), 39, "Devengado", 218041393, {"ue": "009", "act": ["5006357"]}, ""),
]
CONCEPTO_MEF = {"PIA": "MONTO_PIA", "PIM": "MONTO_PIM", "Devengado": "MONTO_DEVENGADO"}


def seleccion(n, s):
    m = n[(n.PLIEGO == "039") & (n.EJECUTORA == s["ue"])]
    if "act" in s:
        m = m[m.ACTIVIDAD_ACCION_OBRA.isin(s["act"])]
    if "sin" in s:
        m = m[~m.SEC_FUNC.isin(s["sin"])]
    if "meta" in s:
        m = m[m.META_NOMBRE.str.contains(s["meta"])]
    if "organo" in s:
        m = m[[organo(x)[0].startswith(s["organo"]) for x in m.META_NOMBRE]]
    return m


def describir_seleccion(s):
    partes = [f"UE {s['ue']}"]
    partes += [f"actividad {', '.join(s['act'])}"] if "act" in s else []
    partes += [f"sin las metas {', '.join(s['sin'])}"] if "sin" in s else []
    partes += ["metas «Servicios especializados e interdisciplinarios…»"] if "meta" in s else []
    partes += [f"metas de la {s['organo']}"] if "organo" in s else []
    return "; ".join(partes)


def conciliaciones():
    """Además del monto, dos señales mensuales del mismo agregado (detalle del pliego 039, filtrar_gasto_mef_pliego.py):
    la certificación anulada en noviembre y diciembre, que acompaña a las rebajas del PIM de fin de año, y el devengado
    acumulado a noviembre, que acota la fecha de corte de una cifra reportada."""
    filas = []
    for anio, objeto, fuente, pag, conc, monto, sel, lectura in CONCILIACIONES:
        m = seleccion(leer(anio, "gobierno-nacional_meta"), sel)
        v = m[CONCEPTO_MEF[conc]].sum()
        anulado = dev_nov = None
        if (MEF / f"{anio}_mef_gasto_pliego-039_detalle-mensual.csv").exists():
            d = seleccion(leer(anio, "pliego-039_detalle-mensual"), sel)
            mes = d.groupby(d.MES_EJE.astype(int))[["MONTO_CERTIFICADO", "MONTO_DEVENGADO"]].sum()
            anulado = mes.MONTO_CERTIFICADO.reindex([11, 12]).clip(upper=0).sum()
            dev_nov = mes.MONTO_DEVENGADO.reindex(range(1, 12)).sum()
        filas.append({"Año": anio, "Objeto": objeto, "Fuente reportada": fuente, "Página": pag, "Concepto": conc,
                      "Monto reportado": monto, "Agregado del MEF": describir_seleccion(sel), "Metas": m.SEC_FUNC.nunique(),
                      "Monto MEF": v, "Diferencia (reportado − MEF)": monto - v,
                      "Resultado": "Coincide" if abs(monto - v) <= 1 else "Difiere",
                      "Certificación anulada en nov. y dic.": anulado, "Devengado acumulado a noviembre": dev_nov,
                      "Lectura (equipo evaluador)": lectura})
    return pd.DataFrame(filas)


def modificaciones():
    """Cambio entre el PIA y el PIM por unidad ejecutora del pliego 039 y por meta de la conducción y de los CEM. El MEF
    publica el PIA y el PIM anuales (mes 0): el saldo muestra si una meta recibió o cedió recursos, no la fecha ni la
    contraparte de cada modificación."""
    ue, metas = [], []
    for anio in ANIOS:
        n = leer(anio, "gobierno-nacional_meta")
        p = n[n.PLIEGO == "039"]
        g = p.groupby(["EJECUTORA", "EJECUTORA_NOMBRE"], as_index=False)[list(MONTOS)].sum()
        ue.append(g.assign(Año=anio))
        m = p[((p.EJECUTORA == "001") & p.ACTIVIDAD_ACCION_OBRA.eq("5001072")) |
              ((p.EJECUTORA == "009") & p.ACTIVIDAD_ACCION_OBRA.isin(["5006193", "5006357"]))]
        g = m.groupby(["EJECUTORA", "SEC_FUNC", "ACTIVIDAD_ACCION_OBRA", "META_NOMBRE", "DEPARTAMENTO_META_NOMBRE"],
                      as_index=False)[list(MONTOS)].sum()
        metas.append(g.assign(Año=anio))
    out = []
    for d in (pd.concat(ue), pd.concat(metas)):
        d = d.rename(columns={**MONTOS, "EJECUTORA": "UE", "EJECUTORA_NOMBRE": "Unidad ejecutora", "SEC_FUNC": "Meta",
                              "ACTIVIDAD_ACCION_OBRA": "Actividad", "META_NOMBRE": "Nombre de la meta",
                              "DEPARTAMENTO_META_NOMBRE": "Departamento"})
        d["PIM − PIA"] = d.PIM - d.PIA
        d["Variación (%)"] = (d["PIM − PIA"] / d.PIA.where(d.PIA > 0) * 100).round(1)
        if "Meta" in d:
            d["Órgano o servicio"] = [organo(x)[0] if u == "001" else "CEM y servicios de atención (AURORA)"
                                      for x, u in zip(d["Nombre de la meta"], d.UE)]
            d["Situación"] = ["Creada en el año (PIA cero)" if a == 0 and b > 0 else "Sin recursos al cierre (PIM cero)"
                              if a > 0 and b == 0 else "Recibió recursos" if b > a else "Cedió recursos" if b < a else
                              "Sin cambio" for a, b in zip(d.PIA, d.PIM)]
        out.append(d[["Año"] + [c for c in d.columns if c != "Año"]])
    return out


def programas():
    filas = []
    for anio in ANIOS:
        t = leer(anio, "programas_totales")
        t = t[t.PROGRAMA_PPTO.isin(PROGRAMAS + ["0049"])]
        filas.append(t)
    d = pd.concat(filas)
    d["Nivel de gobierno"] = d.NIVEL_GOBIERNO.map({"E": "Nacional", "R": "Regional", "M": "Local"})
    return d.pivot_table(index=["PROGRAMA_PPTO", "PROGRAMA_PPTO_NOMBRE", "Nivel de gobierno"], columns="ANO_EJE",
                         values="MONTO_PIM", aggfunc="sum").reset_index()


def main():
    grd = pd.read_excel(GRD, "Estado por servicio", dtype=str)
    entidades = dict(zip(grd["Código"], grd["Entidad u órgano responsable"]))

    def sectores(servicio, informante=""):
        txt = informante or entidades.get(servicio, "")  # en el formulario cuenta la entidad que informa
        return sorted({s for e, s in SECTOR.items() if re.search(rf"\b{e}\b", txt.upper().replace("Ú", "U"))})

    cifras = pd.concat([cifras_lio(), cifras_formulario()], ignore_index=True)
    cifras["Entidad responsable (REV_DOC)"] = cifras.Servicio.map(entidades).fillna("")
    cifras = cifras.reset_index(drop=True)
    resultados = {}
    c = None
    for anio in ANIOS:
        c_prev, c = c, candidatos(anio)
        sel = cifras[cifras["Año"] == anio]
        for i, r in sel.iterrows():
            inf = r.get("Entidad informante")
            resultados[i] = buscar(c, r["Monto reportado"], r["Precisión"],
                                   sectores(r.Servicio, inf if isinstance(inf, str) else ""))
        for _, g in sel.groupby(["Fuente", "Página", "Texto"]):
            exactas = {resultados[i]["Descripción MEF"] for i in g.index if resultados[i]["Resultado"] == "Coincide"}
            for i in g.index:
                if resultados[i]["Resultado"] == "Sin coincidencia" and exactas:
                    resultados[i] = cercano(c, cifras.at[i, "Monto reportado"], exactas) or resultados[i]
        # totales de programa presupuestal reportados con otro corte: hasta 0,05 %, montos de S/ 10 millones o más
        prog = c[c.Nivel.isin(["Programa", "Sector y programa", "Programa (todos los niveles de gobierno)",
                               "Programa (nacional y regional)"])]
        for i, r in sel.iterrows():
            m = r["Monto reportado"]
            if resultados[i]["Resultado"] == "Sin coincidencia" and m >= 1e7:
                inf = r.get("Entidad informante")
                sec = sectores(r.Servicio, inf if isinstance(inf, str) else "")
                # solo los programas vinculados a la PNIG, en los pliegos de la entidad o en sus totales
                prog_e = prog[prog.PP.isin(PROGRAMAS) & (prog.SECTOR.isin(sec) | totales_permitidos(prog, sec))]
                d = (prog_e.valor - m).abs()
                hit = prog_e[d <= 0.0005 * m].assign(dif=d[d <= 0.0005 * m])
                if len(hit):
                    resultados[i] = elegir(hit, "Cercano (≤ 0,05 % de un total de programa)", "Programa presupuestal")
        # cifras del año anterior citadas como referencia: solo coincidencia exacta en los pliegos de la entidad
        for i, r in sel.iterrows():
            if resultados[i]["Resultado"] == "Sin coincidencia" and c_prev is not None and r["Monto reportado"] % 1000:
                inf = r.get("Entidad informante")
                sec = sectores(r.Servicio, inf if isinstance(inf, str) else "")
                sub = c_prev[c_prev.SECTOR.isin(sec)]
                d = (sub.valor - r["Monto reportado"]).abs()
                hit = sub[d <= r["Precisión"]].assign(dif=d[d <= r["Precisión"]])
                if len(hit):
                    resultados[i] = elegir(hit, "Coincide con el año anterior", "Entidad responsable")
        print(f"{anio}: {len(sel)} cifras")
    res = pd.DataFrame.from_dict(resultados, orient="index")
    cifras = cifras.join(res)
    cifras["Observación"] = [
        "; ".join(x for x in (f"el mismo monto aparece en {int(p)} pliegos" if p and p > 1 else "",
                              "monto redondo: la coincidencia puede ser casual"
                              if r == "Coincide" and m % 1000 == 0 else "") if x)
        for p, r, m in zip(cifras["Pliegos con el mismo monto"].fillna(0), cifras.Resultado, cifras["Monto reportado"])]
    cifras["Motivo probable (sin coincidencia)"] = [
        "" if r != "Sin coincidencia" else MOTIVOS[-1][1] if f.startswith("Formulario") else motivo(t)
        for t, r, f in zip(cifras.Texto, cifras.Resultado, cifras.Fuente)]  # el formulario pide montos por AO
    cifras["Diferencia"] = cifras["Monto MEF"] - cifras["Monto reportado"]
    cifras["Concepto coherente"] = [
        "" if not isinstance(m, str) or r == "Sin clasificar" else ("Sí" if (r == m or (r == "PIM" and m == "PIA")) else "No")
        for r, m in zip(cifras["Concepto reportado"], cifras["Concepto MEF"])]
    cond = conduccion()
    poi = poi_pia()
    mod_ue, mod_meta = modificaciones()
    comp = poi_vs_mef(poi, cond)
    resumen = cifras.groupby(["Fuente", "Resultado"]).size().unstack(fill_value=0).reset_index()
    motivos = cifras[cifras.Resultado == "Sin coincidencia"].groupby(
        ["Motivo probable (sin coincidencia)", "Fuente"]).size().unstack(fill_value=0).reset_index()
    por_organo = cond.pivot_table(index=["Órgano", "Rol en la PNIG"], columns="Año", values=["PIM", "Devengado"],
                                  aggfunc="sum")
    por_organo.columns = [f"{a} {c}" for c, a in por_organo.columns]
    por_organo = por_organo[sorted(por_organo.columns)].reset_index()
    criterios = pd.DataFrame([
        ("Fuente de las cifras reportadas", "Informes LIO 2019–2025 (párrafos con montos en soles y palabras de presupuesto, "
         "con página y servicio de la sección) y formulario de seguimiento 2023 (preguntas 11 y 12, por actividad operativa).",
         "Nota metodológica (REV_IND.3, REP_ENT)"),
        ("Fuente de verificación", "Portal de Datos Abiertos del MEF, conjunto «Presupuesto y Ejecución de Gasto» (la "
         "Consulta Amigable con meta, fuente y genérica): extractos de src/fuentes/filtrar_gasto_mef.py.", "Equipo evaluador"),
        ("Coincide", "El monto reportado es igual a un total del MEF del mismo año, con una diferencia menor que la precisión "
         "con que se reportó (S/ 1 para soles enteros). Se busca en los pliegos del sector de la entidad responsable (REV_DOC).",
         "Equipo evaluador"),
        ("Cercano (≤ 0,5 % del mismo agregado)", "Diferencia de hasta 0,5 % con el agregado que coincidió exacto con otra "
         "cifra del mismo párrafo (por ejemplo, el PIM coincide y el devengado reportado es preliminar). Una tolerancia "
         "general produce coincidencias casuales, por eso no se usa.", "Equipo evaluador"),
        ("Cercano (≤ 0,05 % de un total de programa)", "Montos de S/ 10 millones o más que difieren a lo más 0,05 % del "
         "total de un programa presupuestal en los pliegos de la entidad, o del total nacional y regional o de todos los "
         "niveles de los programas 0002, 0080, 0103, 0116, 0131 y 1002: mismo agregado con otro corte. Fuera de esos "
         "agregados, esta tolerancia da coincidencias casuales. No se muestra el agregado más cercano por la misma razón.",
         "Equipo evaluador"),
        ("Coincide con el año anterior", "Cifra del año previo citada como referencia en el informe: coincide exacto con "
         "un total del MEF del año anterior, en los pliegos de la entidad (montos no redondos).", "Equipo evaluador"),
        ("Otra entidad (revisar)", "Montos no redondos de S/ 1 millón o más que solo coinciden fuera del sector de la "
         "entidad.", "Equipo evaluador"),
        ("Sin coincidencia", "El monto no se ubica en ningún nivel de agregación. Corresponde a actividades operativas del "
         "POI, tareas, créditos otorgados o cifras de otro año: el clasificador presupuestal no las separa.",
         "Equipo evaluador"),
        ("Motivo probable", "Para las cifras sin coincidencia, lectura automática del párrafo (MOTIVOS): créditos o "
         "ingresos, normas, totales de varias entidades, actividades operativas o tareas del POI; si nada aplica, se marca "
         "para revisión. Es una orientación, no un dictamen.", "Equipo evaluador"),
        ("Todos los niveles de gobierno", "En los programas 0002, 0080, 0103, 0116, 0131 y 1002 se suman el nivel nacional "
         "y el regional y local por programa, producto y actividad (los reportes del MINSA suelen dar el total «por toda "
         "fuente» y todos los niveles). Estos totales solo se buscan para las entidades de los sectores que ejecutan cada "
         "programa (PROGRAMAS_SECTOR).", "Equipo evaluador"),
        ("Entidad del formulario", "En el formulario 2023 se busca solo en el sector de la entidad que informa, no en el de "
         "todas las entidades del servicio.", "Equipo evaluador"),
        ("Monto redondo", "Los montos múltiplos de S/ 1000 solo se buscan en los pliegos del sector de la entidad y su "
         "coincidencia se marca como posiblemente casual.", "Equipo evaluador"),
        ("Concepto coherente", "Sí, si el concepto que coincide en el MEF (PIA, PIM o devengado) es el que nombra el "
         "informe; «asignado» o «programado» se aceptan como PIM o PIA.", "Equipo evaluador"),
        ("Órgano de cada meta", "Por el nombre de la meta (ORGANOS). La DGCVG se muestra como referencia, fuera de la "
         "conducción.", "Equipo evaluador"),
        ("POI consistente con el PIA", "Versión de apertura del POI (Anexo B-5), comparable con el PIA del MEF. La "
         "programación se lee de la línea «Financiero» de cada actividad operativa; se usa el total anual cuando cuadra con los "
         "12 meses.", "Equipo evaluador"),
        ("Límite", "El formulario 2024 y 2025 ya no pidió presupuesto. El POI 2024 y 2025 de apertura usa un formato con "
         "tareas que no se lee de forma confiable: en 2024 se comprobó una actividad operativa por órgano (POI_2024) y 2025 "
         "no se comparó. En 2023 dos actividades de las direcciones de articulación (DGTEG) no tienen monto legible.",
         "Equipo evaluador"),
        ("Conciliaciones", "Reconstrucción de las cifras de transversalización y de los CEM que la búsqueda automática no "
         "ubica: se prueba qué agregado de metas del MEF las explica (por ejemplo, la actividad sin metas de Lima) y se "
         "registra la fuente y la página de cada cifra. La búsqueda de combinaciones se limitó a tres metas.",
         "Equipo evaluador"),
        ("Modificaciones presupuestarias", "El MEF publica el PIA y el PIM anuales de cada meta (mes 0); la certificación y "
         "el devengado son mensuales. Por eso el saldo PIM − PIA muestra si una meta, un órgano o una unidad ejecutora cedió "
         "o recibió recursos, pero no la fecha ni la contraparte de cada traslado. La evaluación trabaja solo con datos "
         "públicos (decisión del usuario, 2026-10-30): los traslados entre órganos se infieren del saldo PIM − PIA, de las "
         "metas creadas en el año y de la certificación anulada al cierre, y lo que esos datos no explican se consigna como "
         "diferencia identificada en las metas.",
         "Equipo evaluador"),
        ("Asignación del órgano", "«Nombre»: el nombre de la meta coincide con la función del órgano en el ROF. "
         "«Probable»: metas creadas durante el año (PIA cero) cuyo nombre sugiere el órgano.", "Equipo evaluador"),
    ], columns=["Elemento", "Criterio", "Plano"])
    hojas = [("Resumen", resumen), ("Sin coincidencia por motivo", motivos),
             ("Cifras verificadas", cifras.drop(columns=["Precisión"])),
             ("Conducción por órgano", por_organo), ("Conducción por meta", cond), ("POI apertura vs MEF", comp),
             ("POI apertura por AO", poi), ("Conciliaciones", conciliaciones()),
             ("Modificaciones por UE", mod_ue), ("Modificaciones por meta", mod_meta),
             ("Estructura orgánica", pd.DataFrame(ESTRUCTURA, columns=[
                 "Periodo", "Norma", "Documento", "Artículos y páginas", "Órganos de la conducción"])),
             ("Programas presupuestales", programas()), ("Criterios", criterios)]
    with pd.ExcelWriter(SALIDA, engine="openpyxl") as xw:
        for hoja, df in hojas:
            df.to_excel(xw, sheet_name=hoja, index=False)
            ws = xw.sheets[hoja]
            for c in ws[1]:
                c.font = Font(bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor="002060")
                c.alignment = Alignment(wrap_text=True, vertical="top")
            for col in ws.columns:
                cab = str(col[0].value)
                ancho = 60 if cab in ("Texto", "Descripción MEF", "Criterio", "Órganos de la conducción", "Norma") else \
                    max(10, min(45, max(len(str(c.value or "")) for c in col[:200]) + 2))
                ws.column_dimensions[col[0].column_letter].width = ancho
                for c in col[1:]:
                    if isinstance(c.value, float):
                        c.number_format = "#,##0" if abs(c.value) >= 100 else "0.0"
                    if ancho == 60:
                        c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.freeze_panes = "A2"
    print(resumen.to_string())
    print(comp.to_string())


if __name__ == "__main__":
    main()
