"""REV_ALI: matriz de coherencia y alineamiento (nota metodológica v15, numeral 5 y Anexo 2).

Arma las cinco submatrices del Anexo 2 de la nota:
  - REV_ALI.1, coherencia interna (D3): se calcula a partir del Anexo 5 de la PNIG (matriz de consistencia), de los
    indicadores vigentes de los servicios (servicios.csv, del consolidado de la OMEP: los aprobados con su protocolo; el
    PEMIG había asignado indicadores a los 52 servicios, que no se siguieron), del dictamen de fichas de REV_IND.2 y del
    estado de REV_DOC. Verifica
    las tres condiciones del numeral 5.1: lineamiento con servicio, servicio con indicador y estándar, servicio vinculado
    a una causa o efecto del modelo. La plausibilidad se toma de REV_EVI (REV_EVI_plausibilidad_por_servicio.csv, de
    rev_evi_evidencia.py).
  - REV_ALI.2, alineamiento vertical (D4): rev_ali_insumos/vertical.csv (una fila por objetivo prioritario, instrumento y
    momento), vertical_declarado.csv (alineamiento que la PNIG declaró en 2019) y vertical_indicadores_pedn.csv
    (contribución de los indicadores de los OP a los del PEDN).
  - REV_ALI.3, alineamiento horizontal (D4): rev_ali_insumos/horizontal.csv, si existe.
  - REV_ALI.4, correspondencia normativa e internacional (D4): rev_ali_insumos/normativa_internacional.csv.
  - REV_ALI.5, correspondencia conceptual (D7): rev_ali_insumos/conceptual.csv.
Los insumos los registra el equipo evaluador con referencia y cita literal. vertical.csv, vertical_declarado.csv,
normativa_internacional.csv, horizontal.csv y horizontal_universo.csv se generan con rev_ali_insumos_vertical.py,
rev_ali_insumos_normativa.py y rev_ali_insumos_horizontal.py; vertical_indicadores_pedn.csv y conceptual.csv se escribieron
directamente. El script valida las escalas y la cobertura, ubica cada cita en su documento (página del PDF o párrafo del
.docx) y se detiene si alguna no se ubica.

Uso (desde la raíz del repositorio):
    uv run --no-project --with openpyxl --with pandas --with python-docx python src/analisis_pnig/rev_ali_coherencia_alineamiento.py
"""
import re
import subprocess
import unicodedata
from functools import lru_cache
from pathlib import Path

import pandas as pd
from docx import Document
from openpyxl.styles import Alignment, Font, PatternFill
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

RAIZ = Path(EVAL)
ANALISIS = RAIZ / "04_analisis"
INSUMOS = ANALISIS / "rev_ali_insumos"
SALIDA = ANALISIS / "REV_ALI_coherencia_alineamiento.xlsx"
NORMATIVA = Path("docs/fuentes/pnig/01_normativa")
PLAN = Path("docs/fuentes/planeamiento")
INTER = Path("docs/fuentes/internacional")
PNIG = NORMATIVA / "2019_mimp_ds-008-2019-mimp_pnig.pdf"
ANEXO5 = NORMATIVA / "pnig_anexos/2019_mimp_pnig_anexo-5_matriz-consistencia.xlsx"
MARCO = Path("docs/fuentes/pnig/08_marco_conceptual_eiomh/"
             "2026_mimp-dgignd_marco-conceptual-enfoque-igualdad-oportunidades_version-final.docx")
SERVICIOS = Path("src/formulario_fie03/datos/servicios.csv")
DOCS = {
    "pnig": PNIG,
    "ley": NORMATIVA / "2025_congreso_ley-32535_igualdad-oportunidades-mujeres-hombres.pdf",
    "pe2023": PLAN / "2023_acuerdo-nacional_politicas-de-estado-vision-2050.pdf",
    "bicentenario": PLAN / "2011_ceplan_plan-bicentenario-peru-hacia-el-2021.pdf",
    "pedn2022": PLAN / "2022_pcm_ds-095-2022-pcm_anexo-pedn-2050.pdf",
    "pedn2023": PLAN / "2023_ceplan_pedn-2050_tomo-i.pdf",
    "pedn2023_ind": PLAN / "2023_ceplan_pedn-2050_tomo-ii-indicadores.pdf",
    "pgg2018": PLAN / "2018_pcm_ds-056-2018-pcm_pgg-al-2021.pdf",
    "pgg2021": PLAN / "2021_pcm_ds-164-2021-pcm_pgg-2021-2026.pdf",
    "pgg2023": PLAN / "2023_pcm_ds-042-2023-pcm_pgg-2023-2026.pdf",
    "pgg2025": PLAN / "2025_pcm_ds-141-2025-pcm_pgg-2025-2026.pdf",
    "pgg2026a": PLAN / "2026_pcm_ds-043-2026-pcm_pgg-seguridad-continuidad-transicion.pdf",
    "pgg2026b": PLAN / "2026_pcm_ds-127-2026-pcm_pgg-2026-2031.pdf",
    "pesem2022": PLAN / "2022_mimp_pesem-2018-2026_anexo.pdf",
    "pesem2024": PLAN / "2024_mimp_pesem-2024-2030_sector-mujer-y-poblaciones-vulnerables.pdf",
    "cedaw": INTER / "1979_onu_cedaw_convencion.pdf",
    "cedaw_co9": INTER / "2022_onu-cedaw_cedaw-c-per-co-9_observaciones-finales-peru.pdf",
    "belem": INTER / "1994_oea_convencion-belem-do-para.pdf",
    "beijing": INTER / "1995_onu_declaracion-y-plataforma-de-accion-de-beijing.pdf",
    "agenda2030": INTER / "2015_onu_a-res-70-1_agenda-2030.pdf",
    "montevideo": INTER / "2017_cepal_lc-crm-13-5_estrategia-montevideo.pdf",
}
PLANOS = ["Conceptual", "Problema público", "Estructura de la Política", "Servicios"]
TIPOS = ["Continuidad", "Precisión", "Divergencia"]
EFECTOS = ["Se mantiene", "Se ajusta", "Se rediseña"]
CORRESPONDENCIA = ["Directa", "Indirecta", "Sin correspondencia"]
RELACION_H = ["Complementariedad", "Duplicidad", "Vacío de atención"]
OPS = [f"OP{i}" for i in range(1, 7)]
LINEAMIENTOS = ["1.1", "1.2", "2.1", "2.2", "2.3", "3.1", "3.2", "4.1", "4.2", "4.3", "4.4", "5.1", "5.2", "5.3", "5.4",
                "5.5", "6.1", "6.2", "6.3", "6.4"]


def norm(t):
    """Normaliza para buscar citas: ligaduras, guiones de corte de línea, mayúsculas y espacios."""
    t = unicodedata.normalize("NFKC", t).replace("­", "")
    t = re.sub(r"(\w)-\s*\n\s*(\w)", r"\1\2", t)
    return re.sub(r"\s+", " ", t).strip().lower()


@lru_cache(maxsize=None)
def paginas(pdf):
    """Texto normalizado de cada página en dos lecturas (con disposición y en orden de lectura)."""
    salida = []
    for modo in (["-layout"], ["-raw"]):
        txt = subprocess.run(["pdftotext", "-enc", "UTF-8", *modo, str(pdf), "-"], capture_output=True).stdout
        salida.append([norm(p) for p in txt.decode("utf-8", "ignore").split("\f")])
    return tuple(zip(*salida))


def ubicar_pdf(cita, pdf):
    c = norm(cita)
    return [i + 1 for i, (a, b) in enumerate(paginas(pdf)) if c in a or c in b]


def ubicar_doc(df, col_doc="documento", col_cita="cita"):
    """Añade la página del PDF de cada cita; devuelve el marco y las citas no ubicadas."""
    df = df.copy()
    df["Página del PDF"] = [", ".join(map(str, ubicar_pdf(c, DOCS[d]))) if c else ""
                            for d, c in zip(df[col_doc], df[col_cita])]
    faltan = df[(df[col_cita] != "") & (df["Página del PDF"] == "")]
    return df, faltan


NO_UBICADAS = []


def registrar(nombre, faltan, col="cita"):
    for _, r in faltan.iterrows():
        NO_UBICADAS.append((nombre, r.get("op", r.get("codigo", "")), r[col]))


def leer(nombre):
    return pd.read_csv(INSUMOS / nombre, dtype=str).fillna("")


def en_escala(df, col, escala, nombre):
    fuera = set(df[col]) - set(escala)
    assert not fuera, f"{nombre}: {col} fuera de escala: {fuera}"


# ---------------------------------------------------------------- REV_ALI.1: coherencia interna
a5 = pd.read_excel(ANEXO5, "Anexo N° 5", header=None, dtype=str).iloc[4:]
a5.index = a5.index + 1  # fila de Excel
a5.columns = ["problema", "causa", "alternativa", "op", "ind_op", "lineamiento", "servicio", "persona", "proveedor",
              "cobertura", "estandar", "definicion", "ind_estandar"]
for c in ("op", "ind_op", "lineamiento"):
    a5[c] = a5[c].ffill()
a5 = a5[a5.servicio.notna()].fillna("")
a5["fila"] = a5.index
a5["cod_lin"] = a5.lineamiento.str.extract(r"(\d\.\d)", expand=False)
a5["codigo"] = a5.servicio.str.extract(r"^\s*(\d\.\d\.\d)", expand=False).fillna("")
sin_servicio = a5[a5.codigo == ""]
a5 = a5[a5.codigo != ""]
serv = pd.read_csv(SERVICIOS, dtype=str).fillna("")
assert set(a5.codigo) == set(serv.codigo_corto), "el Anexo 5 y servicios.csv no tienen los mismos servicios"
assert set(a5.cod_lin) | set(sin_servicio.lineamiento.str.extract(r"(\d\.\d)", expand=False)) == set(LINEAMIENTOS)

dictamen = pd.read_excel(ANALISIS / "REV_IND_2_calidad_fichas.xlsx", "Dictamen", dtype=str)
grd = pd.read_excel(ANALISIS / "REV_DOC_estado_implementacion_servicios.xlsx", "Estado por servicio", dtype=str).fillna("")
grd = grd.set_index("Código")
plaus = pd.read_csv(ANALISIS / "REV_EVI_plausibilidad_por_servicio.csv", dtype=str).fillna("").set_index("Código")


def servicio_de(cod):
    p = cod.split(".")
    return f"{int(p[1])}.{int(p[2])}.{int(p[3])}"


dic_s = dictamen[dictamen.Nivel == "Servicio"].assign(servicio=lambda d: d["Código"].map(servicio_de))
dic_op = dictamen[dictamen.Nivel == "OP"]


def nodos(causa):
    partes = [x.strip() for x in causa.split("\n") if x.strip()]
    tipos = {"Efecto" if x.startswith("Vulneración") else "Causa" for x in partes}
    tipo = "No consignado" if not partes else "Efecto y causa" if len(tipos) == 2 else tipos.pop()
    return "; ".join(partes), tipo


def contar(texto):
    return len(re.findall(r"(?m)^\s*\d+\s*\.", texto))


filas = []
for _, r in a5.iterrows():
    s = serv.set_index("codigo_corto").loc[r.codigo]
    nodo, tipo = nodos(r.causa)
    ind = re.findall(r"ISS\.\d\d\.\d\d\.\d\d\.\d\d", s.indicadores)
    dic = dic_s[dic_s.servicio == r.codigo]
    vacios = []
    if tipo == "No consignado":
        vacios.append("Sin causa ni alternativa consignadas en el Anexo 5")
    if not r.estandar.strip():
        vacios.append("Sin estándar de cumplimiento en el Anexo 5")
    if not ind:
        vacios.append("Sin indicador vigente (sin protocolo aprobado)")
    observ = []
    if s.anos_con_valor_reportado == "sin valores":
        observ.append("Indicador sin valores reportados en 2019–2025")
    if len(dic) and (dic["Críticas"].astype(int) > 0).any():
        observ.append(f"{(dic['Críticas'].astype(int) > 0).sum()} de {len(dic)} fichas con observación crítica (REV_IND.2)")
    if grd.loc[r.codigo, "Protocolo aprobado"] == "No":
        observ.append("Estándar sin protocolo aprobado al 31-12-2025 (REV_DOC)")
    filas.append({
        "Código": r.codigo, "OP": r.op.split(":")[0].strip(), "Lineamiento": r.lineamiento.strip(),
        "Servicio": re.sub(r"^\s*\d\.\d\.\d\.?\s*", "", r.servicio).strip(), "Proveedor (Anexo 5)": r.proveedor.strip(),
        "Causa o efecto del modelo (Anexo 5)": nodo, "Tipo de vínculo": tipo,
        "Alternativa de solución (Anexo 5)": r.alternativa.strip(),
        "Estándares (Anexo 5)": contar(r.estandar), "Indicadores del estándar (Anexo 5)": contar(r.ind_estandar),
        "Indicadores vigentes del servicio": ", ".join(ind),
        "Dictamen de las fichas (REV_IND.2)": "; ".join(f"{c}: {d}" for c, d in zip(dic["Código"], dic.Dictamen)),
        "Años con valor reportado": s.anos_con_valor_reportado,
        "Protocolo aprobado al 31-12-2025 (REV_DOC)": grd.loc[r.codigo, "Protocolo aprobado"],
        "Estado de implementación (REV_DOC)": grd.loc[r.codigo, "Estado asignado (preliminar)"],
        "Vacíos detectados": "; ".join(vacios), "Observaciones complementarias": "; ".join(observ),
        "Valoración": "Completa" if not vacios else "Con vacíos",
        "Tipo de intervención (REV_EVI)": f"{plaus.loc[r.codigo, 'Tipo']}: {plaus.loc[r.codigo, 'Tipo de intervención']}",
        "Plausibilidad (REV_EVI)": plaus.loc[r.codigo, "Plausibilidad"], "Fila del Anexo 5": r.fila})
c41 = pd.DataFrame(filas)
c41["_o"] = c41["Código"].map(lambda c: tuple(int(x) for x in c.split(".")))
c41 = c41.sort_values("_o").drop(columns="_o")
assert len(c41) == 52

lin_rows = []
for lin in LINEAMIENTOS:
    sub = c41[c41["Código"].str[:3] == lin]
    if sub.empty:
        r = sin_servicio[sin_servicio.lineamiento.str.contains(rf"^\s*{re.escape(lin)}", regex=True)].iloc[0]
        lin_rows.append({"Causa o efecto": nodos(r.causa)[0], "OP": r.op.split(":")[0].strip(),
                         "Lineamiento": r.lineamiento.strip(), "Servicios": "", "N.° de servicios": 0,
                         "Indicadores vigentes de servicio": "",
                         "Vacíos detectados": f"Lineamiento sin servicio. Anexo 5, fila {r.fila}: «{r.servicio.strip()}»",
                         "Valoración": "Rota"})
        continue
    con_vacio = sub[sub.Valoración == "Con vacíos"]
    lin_rows.append({
        "Causa o efecto": "; ".join(sorted({n for x in sub["Causa o efecto del modelo (Anexo 5)"] for n in x.split("; ") if n})),
        "OP": sub.OP.iloc[0], "Lineamiento": sub.Lineamiento.iloc[0], "Servicios": ", ".join(sub["Código"]),
        "N.° de servicios": len(sub),
        "Indicadores vigentes de servicio": f"{(sub['Indicadores vigentes del servicio'] != '').sum()} de {len(sub)} servicios",
        "Vacíos detectados": "; ".join(f"{c}: {v}" for c, v in zip(con_vacio["Código"], con_vacio["Vacíos detectados"])),
        "Valoración": "Completa" if con_vacio.empty else f"Con vacíos ({len(con_vacio)} de {len(sub)} servicios)"})
c41_lin = pd.DataFrame(lin_rows)

ind_op = a5.groupby("op").ind_op.apply(lambda s: list(dict.fromkeys(x.strip() for x in s)))
op_rows = []
for op in OPS:
    sub = c41[c41.OP == op]
    lins = c41_lin[c41_lin.OP == op]
    dop = dic_op[dic_op["Código"].str.startswith(f"IOP.0{op[-1]}")]
    nombre = next(k for k in ind_op.index if k.startswith(op))
    op_rows.append({
        "OP": op, "Objetivo prioritario": nombre.split(":", 1)[1].strip(),
        "Causa o efecto que atiende": "; ".join(sorted({n for x in sub["Causa o efecto del modelo (Anexo 5)"] for n in x.split("; ") if n})),
        "Lineamientos": len(lins), "Lineamientos sin servicio": int((lins["N.° de servicios"] == 0).sum()),
        "Servicios": len(sub), "Servicios con indicador vigente": int((sub["Indicadores vigentes del servicio"] != "").sum()),
        "Servicios con vacíos": int((sub.Valoración == "Con vacíos").sum()),
        "Indicadores del OP (Anexo 5)": "; ".join(ind_op[nombre]),
        "Dictamen de las fichas del OP (REV_IND.2)": "; ".join(f"{c}: {d}" for c, d in zip(dop["Código"], dop.Dictamen)),
    })
c41_op = pd.DataFrame(op_rows)

modelo = (c41.assign(nodo=c41["Causa o efecto del modelo (Anexo 5)"].str.split("; ")).explode("nodo")
          .query("nodo != ''").groupby("nodo")
          .agg(Servicios=("Código", "count"), OP=("OP", lambda s: ", ".join(sorted(set(s)))),
               Códigos=("Código", lambda s: ", ".join(s))).reset_index().rename(columns={"nodo": "Causa o efecto del modelo"}))
modelo.insert(1, "Tipo", modelo["Causa o efecto del modelo"].map(lambda x: "Efecto" if x.startswith("Vulneración") else "Causa"))
modelo = modelo.sort_values(["Tipo", "Servicios"], ascending=[True, False])

# ---------------------------------------------------------------- REV_ALI.2: alineamiento vertical
ver = leer("vertical.csv")
en_escala(ver, "correspondencia", CORRESPONDENCIA, "vertical")
en_escala(ver, "op", OPS, "vertical")
assert ((ver.cita == "") == (ver.correspondencia == "Sin correspondencia")).all(), "cita y correspondencia no concuerdan"
assert not ver.duplicated(["momento", "op", "instrumento"]).any(), "fila repetida en vertical.csv"
for inst, g in ver.groupby("instrumento"):
    assert set(g.op) == set(OPS), f"{inst}: faltan objetivos"
ver, f = ubicar_doc(ver)
registrar("REV_ALI.2", f)

decl = leer("vertical_declarado.csv")
decl["Página del PDF"] = decl.cita_pnig.map(lambda c: ", ".join(map(str, ubicar_pdf(c, PNIG))))
registrar("REV_ALI.2 declarado", decl[decl["Página del PDF"] == ""].rename(columns={"cita_pnig": "cita"}))

indp = leer("vertical_indicadores_pedn.csv")
indp["Página del PDF"] = indp.indicador_pedn.map(lambda c: ", ".join(map(str, ubicar_pdf(c, DOCS["pedn2023_ind"]))) if c else "")
registrar("REV_ALI.2 indicadores", indp[(indp.indicador_pedn != "") & (indp["Página del PDF"] == "")].rename(
    columns={"indicador_pedn": "cita"}))
base = pd.read_excel(ANALISIS / "REV_IND_base_indicadores.xlsx", "Base REV_IND.1", dtype=str)
indp.insert(1, "Indicador del OP", indp.codigo.map(base.set_index("Código").Indicador))
assert indp["Indicador del OP"].notna().all() and len(indp) == (base.Nivel == "OP").sum()


def celda(g):
    return "\n".join(f"{r.referencia} ({r.correspondencia.lower()})" for r in g.itertuples())


def resumen_corr(g):
    n = g.correspondencia.value_counts()
    return ", ".join(f"{n.get(k, 0)} {k.lower()}" for k in CORRESPONDENCIA)


COL_V = {"Periodo 2019–2025": ["Políticas de Estado", "Plan Bicentenario", "PEDN al 2050 (2022)",
                               "PGG al 2021 (D.S. 056-2018-PCM)", "PGG 2021–2026 (D.S. 164-2021-PCM)",
                               "PGG 2023–2026 (D.S. 042-2023-PCM)", "PGG 2025–2026 (D.S. 141-2025-PCM)", "PESEM 2018–2026"],
         "Transición": ["Políticas de Estado", "PEDN al 2050 actualizado (2023)", "PGG 2026 (D.S. 043-2026-PCM)",
                        "PGG 2026–2031 (D.S. 127-2026-PCM)", "PESEM 2024–2030"]}
ambos = ver[ver.momento == "Periodo y transición"]
ver_m = pd.concat([ver[ver.momento != "Periodo y transición"], ambos.assign(momento="Periodo 2019–2025"),
                   ambos.assign(momento="Transición")])
assert set(ver_m.instrumento) == set(COL_V["Periodo 2019–2025"]) | set(COL_V["Transición"])
v_op = []
for op in OPS:
    fila = {"OP": op}
    for m, cols in COL_V.items():
        g = ver_m[(ver_m.op == op) & (ver_m.momento == m)]
        for inst in cols:
            fila[f"{inst} [{'periodo' if m.startswith('Periodo') else 'transición'}]"] = celda(g[g.instrumento == inst])
        fila[f"Valoración ({'periodo' if m.startswith('Periodo') else 'transición'})"] = resumen_corr(g)
    v_op.append(fila)
v_op = pd.DataFrame(v_op)

# ---------------------------------------------------------------- REV_ALI.3: alineamiento horizontal (si hay insumo)
hor = None
if (INSUMOS / "horizontal.csv").exists():
    hor = leer("horizontal.csv")
    en_escala(hor, "tipo_relacion", RELACION_H, "horizontal")
    assert hor.codigo.is_unique, "código repetido en horizontal.csv"
    hor["Página del PDF"] = [", ".join(map(str, ubicar_pdf(c, Path(d)))) if c else ""
                             for d, c in zip(hor.documento, hor.cita)]
    registrar("REV_ALI.3", hor[(hor.cita != "") & (hor["Página del PDF"] == "")].rename(columns={"codigo": "op"}))
    universo = leer("horizontal_universo.csv").rename(columns={
        "politica": "Política", "norma": "Norma", "criterio_inclusion": "Criterio de inclusión",
        "antecesora_declarada_2019": "Antecesora declarada por la PNIG en 2019", "resultado": "Resultado"})
    assert set(hor.politica) <= set(universo["Política"]), "política analizada fuera del universo"
    hor_r = hor.assign(documento=hor.documento.map(lambda d: Path(d).name)).rename(columns={
        "codigo": "Código", "politica": "Política nacional afín", "norma": "Norma", "documento": "Documento",
        "elemento_politica_afin": "Objetivo, lineamiento o servicio relacionado", "entidad": "Entidad responsable",
        "op_pnig": "OP de la PNIG", "servicio_pnig": "Servicio o lineamiento de la PNIG", "cita": "Cita",
        "tipo_relacion": "Tipo de relación", "sustento": "Sustento"})
    por_servicio = (hor_r.groupby("Servicio o lineamiento de la PNIG")
                    .apply(lambda g: pd.Series({
                        **{t: int((g["Tipo de relación"] == t).sum()) for t in RELACION_H},
                        "Políticas": "; ".join(f"{p} ({t.lower()})" for p, t in zip(g["Política nacional afín"],
                                                                                    g["Tipo de relación"]))}),
                           include_groups=False)
                    .reset_index().sort_values(["Duplicidad", "Vacío de atención"], ascending=False))

# ---------------------------------------------------------------- REV_ALI.4: normativa e internacional
NOR_COLS = ["Ley N.° 32535", "CEDAW", "Observaciones finales del Comité CEDAW (2022)", "Convención de Belém do Pará",
            "Plataforma de Acción de Beijing", "ODS 5 (Agenda 2030)", "Estrategia de Montevideo"]
nor = leer("normativa_internacional.csv")
en_escala(nor, "correspondencia", CORRESPONDENCIA, "normativa")
en_escala(nor, "instrumento", NOR_COLS, "normativa")
assert len(nor.drop_duplicates(["op", "instrumento"])) == len(OPS) * len(NOR_COLS) == len(nor), "cobertura de REV_ALI.4"
nor, f = ubicar_doc(nor)
registrar("REV_ALI.4", f)
n_op = []
for op in OPS:
    g = nor[nor.op == op]
    fila = {"OP": op}
    for inst in NOR_COLS:
        fila[inst] = celda(g[g.instrumento == inst])
    fila["Valoración (internacional)"] = resumen_corr(g[g.instrumento != "Ley N.° 32535"])
    fila["Valoración (Ley N.° 32535)"] = g[g.instrumento == "Ley N.° 32535"].correspondencia.iloc[0]
    n_op.append(fila)
n_op = pd.DataFrame(n_op)

# ---------------------------------------------------------------- REV_ALI.5: correspondencia conceptual
con = leer("conceptual.csv")
assert con.codigo.is_unique, "código repetido"
for col, escala in (("plano", PLANOS), ("tipo_relacion", TIPOS), ("efecto", EFECTOS)):
    en_escala(con, col, escala, "conceptual")
assert set(con[con.plano == "Servicios"].codigo) == set(serv.codigo_corto), "faltan o sobran servicios"
assert set(con[con.plano == "Estructura de la Política"].codigo) == set(OPS) | {
    f"L{x}" for x in LINEAMIENTOS}, "faltan o sobran objetivos o lineamientos"
PARRAFOS_MARCO = [norm(p.text) for p in Document(MARCO).paragraphs if p.text.strip()]


def ubicar_nuevo(cita):
    """Busca la cita en la Ley (página del PDF) y en el marco conceptual (párrafo no vacío del .docx)."""
    lugares = [f"Ley N.° 32535, p. {p} del PDF" for p in ubicar_pdf(cita, DOCS["ley"])]
    c = norm(cita)
    lugares += [f"Marco conceptual, párr. {i + 1}" for i, p in enumerate(PARRAFOS_MARCO) if c in p]
    return "; ".join(lugares)


con["Ubicación de la cita PNIG"] = con.cita_pnig.map(
    lambda c: "; ".join(f"p. {p} del PDF" for p in ubicar_pdf(c, PNIG)) if c else "")
con["Ubicación de la cita del nuevo marco"] = con.cita_nuevo.map(lambda c: ubicar_nuevo(c) if c else "")
registrar("REV_ALI.5 PNIG", con[(con.cita_pnig != "") & (con["Ubicación de la cita PNIG"] == "")].rename(
    columns={"cita_pnig": "cita"}))
registrar("REV_ALI.5 nuevo marco", con[(con.cita_nuevo != "") & (con["Ubicación de la cita del nuevo marco"] == "")].rename(
    columns={"cita_nuevo": "cita"}))

if NO_UBICADAS:
    for n, k, c in NO_UBICADAS:
        print(f"{n} | {k} | {c}")
    raise SystemExit(f"{len(NO_UBICADAS)} citas no ubicadas")

conceptual = con.rename(columns={
    "plano": "Plano", "codigo": "Código", "elemento_pnig": "Elemento de la PNIG", "cita_pnig": "Cita de la PNIG",
    "categoria_nuevo": "Categoría del nuevo enfoque", "cita_nuevo": "Cita de la Ley o del marco",
    "tipo_relacion": "Tipo de relación", "efecto": "Efecto", "sustento": "Sustento",
    "contrastar_dgignd": "Para contrastar con la DGIGND (ENTREV)"})
conceptual = conceptual[["Plano", "Código", "Elemento de la PNIG", "Cita de la PNIG", "Ubicación de la cita PNIG",
                         "Categoría del nuevo enfoque", "Cita de la Ley o del marco",
                         "Ubicación de la cita del nuevo marco", "Tipo de relación", "Efecto", "Sustento",
                         "Para contrastar con la DGIGND (ENTREV)"]]
resumen5 = pd.concat([
    conceptual.pivot_table(index="Plano", columns="Tipo de relación", values="Código", aggfunc="count",
                           fill_value=0).reindex(PLANOS)[TIPOS],
    conceptual.pivot_table(index="Plano", columns="Efecto", values="Código", aggfunc="count",
                           fill_value=0).reindex(PLANOS)[EFECTOS]], axis=1).reset_index()
entrevista = conceptual[conceptual["Para contrastar con la DGIGND (ENTREV)"].str.startswith("¿")][
    ["Código", "Elemento de la PNIG", "Tipo de relación", "Efecto", "Para contrastar con la DGIGND (ENTREV)"]]

# ---------------------------------------------------------------- hojas de registro con nombres legibles
REN = {"momento": "Momento", "op": "OP", "instrumento": "Instrumento", "documento": "Documento",
       "referencia": "Referencia", "cita": "Cita", "correspondencia": "Correspondencia", "sustento": "Sustento"}
ver_r = ver.rename(columns=REN).assign(Documento=lambda d: d.Documento.map(lambda k: DOCS[k].name))
nor_r = nor.rename(columns=REN).assign(Documento=lambda d: d.Documento.map(lambda k: DOCS[k].name))
decl_r = decl.rename(columns={"instrumento": "Instrumento", "referencia": "Referencia", "op_declarado": "OP declarado",
                              "cita_pnig": "Cita en la PNIG (alineamiento declarado)", "valoracion": "Valoración"})
indp_r = indp.rename(columns={"codigo": "Código", "accion_pedn": "Acción estratégica del PEDN",
                              "indicador_pedn": "Indicador del PEDN (tomo II)", "relacion": "Relación",
                              "sustento": "Sustento"})

pag_modelo = ubicar_pdf("los efectos de la discriminación estructural se evidencian en las vulneraciones", PNIG)
pag_abordaje = ubicar_pdf("pasa por la atención de las causas y también de sus efectos", PNIG)
criterios = pd.DataFrame([
    ("REV_ALI.1", "Objeto", "Cadena causa o efecto, objetivo prioritario, lineamiento, servicio e indicador de la PNIG.",
     "Nota metodológica, numeral 5.1"),
    ("REV_ALI.1", "Condiciones", "1) Cada lineamiento tiene al menos un servicio. 2) Cada servicio tiene estándar de "
     "cumplimiento (Anexo 5) e indicador de servicio vigente (aprobado con su protocolo). 3) Cada servicio está vinculado a una causa o a un efecto "
     "del modelo del problema público.", "Nota metodológica, numeral 5.1"),
    ("REV_ALI.1", "Causas y efectos", "El modelo de la PNIG tiene tres causas y cuatro efectos (p. "
     f"{', '.join(map(str, pag_modelo))} del PDF) y declara que el abordaje «pasa por la atención de las causas y también "
     f"de sus efectos» (p. {', '.join(map(str, pag_abordaje))} del PDF). Por eso un servicio vinculado a un efecto cumple "
     "la condición 3; el tipo de vínculo se reporta para el análisis de plausibilidad.", "Equipo evaluador"),
    ("REV_ALI.1", "Valoración", "Servicio: completa si cumple las condiciones 2 y 3; con vacíos en otro caso. Lineamiento: "
     "rota si no tiene servicio; completa si todos sus servicios están completos; con vacíos en otro caso. Las "
     "observaciones complementarias (valores reportados, fichas con observación crítica, protocolo) no cambian la "
     "valoración: provienen de REV_IND y REV_DOC.", "Equipo evaluador"),
    ("REV_ALI.1", "Plausibilidad", "Solidez y dirección de la evidencia sobre el tipo de intervención al que pertenece el "
     "servicio, según la matriz REV_EVI (Anexo 6 de la nota): sólida, moderada, limitada o sin evidencia identificada. "
     "«Evidencia parcial» indica que la evidencia mide otro resultado o cubre solo parte de la población del servicio.",
     "Nota metodológica, numeral 5.1 y Anexo 6"),
    ("REV_ALI.2", "Objeto y momentos", "Cada OP frente a las Políticas de Estado, el plan de desarrollo nacional, la PGG y "
     "el PESEM, en dos momentos: instrumentos vigentes en el periodo 2019–2025 y vigentes para la transición (Tabla 2 de "
     "la nota). Las Políticas de Estado son las mismas en ambos momentos (edición 2023 del Acuerdo Nacional).",
     "Nota metodológica, numeral 5.2 y Tabla 2"),
    ("REV_ALI.2 y REV_ALI.4", "Escala", "Directa: el instrumento enuncia el mismo cambio sobre la misma población. "
     "Indirecta: el OP contribuye a un objetivo más amplio del instrumento. Sin correspondencia: ningún objetivo, eje o "
     "lineamiento guarda relación.", "Nota metodológica, numerales 5.2 y 5.4"),
    ("REV_ALI.2", "Alineamiento declarado", "Se parte del alineamiento que la PNIG declaró al aprobarse (numeral 1.5.1) y "
     "se valora si cada acción se asignó al OP que corresponde.", "Nota metodológica, numeral 5.2"),
    ("REV_ALI.2", "Indicadores del PEDN", "Mismo indicador, indicador equivalente (mismo concepto con otra definición "
     "operativa), indicador relacionado o sin indicador en el tomo II del PEDN actualizado.", "Equipo evaluador"),
    ("REV_ALI.2", "PESEM", "Solo se revisa el PESEM del sector Mujer y Poblaciones Vulnerables, como indica la Tabla 2; el "
     "OP2 corresponde al sector Salud.", "Nota metodológica, Tabla 2"),
    ("REV_ALI.3", "Universo", "Sucesoras vigentes de los planes y políticas que la PNIG declaró en su alineamiento horizontal "
     "(numeral 1.5.2) y políticas aprobadas del registro del CEPLAN (base de datos al 15-10-2026; lista del anexo 1 de la "
     "RPCD 042-2026) que comparten población objetivo o servicios con la PNIG. Los documentos están en "
     "docs/fuentes/politicas_afines/.", "Nota metodológica, numeral 5.3"),
    ("REV_ALI.3", "Escala", "Complementariedad: la política afín atiende otra población, otro componente o el mismo problema "
     "con otro abordaje. Duplicidad: el mismo servicio, con el mismo proveedor, figura como servicio en ambas políticas. "
     "Vacío de atención: una población o necesidad queda sin servicio en ninguna de las dos políticas, o quedará sin él en la "
     "transición.", "Nota metodológica, numeral 5.3; criterio de duplicidad del equipo evaluador"),
    ("REV_ALI.4", "Instrumentos", "Ley N.° 32535; CEDAW y observaciones finales al Perú (CEDAW/C/PER/CO/9, 2022); Convención "
     "de Belém do Pará; Plataforma de Acción de Beijing; ODS 5; Estrategia de Montevideo (CEPAL, LC/CRM.13/5). La "
     "plantilla del Anexo 2 incorpora la columna de la Estrategia de Montevideo desde la nota v15.",
     "Nota metodológica, numeral 5.4 y Anexo 2"),
    ("REV_ALI.4", "Valoración", "Se informa por separado la correspondencia con la Ley N.° 32535 y el conteo de "
     "correspondencias con los seis instrumentos internacionales.", "Equipo evaluador"),
    ("REV_ALI.5", "Objeto", "Correspondencia entre la PNIG y el enfoque de igualdad de oportunidades entre mujeres y hombres "
     "en cuatro planos: conceptual, problema público, estructura (6 objetivos y 20 lineamientos) y servicios (52).",
     "Nota metodológica, numeral 5.5 y Tabla 3"),
    ("REV_ALI.5", "Referencia", "Marco conceptual de la DGIGND, versión final del 16-09-2026, y Ley N.° 32535 (El Peruano, "
     "24-12-2025).", "Nota metodológica, numeral 5.5"),
    ("REV_ALI.5", "Tipo de relación", "Continuidad: la categoría se mantiene sin variación sustantiva. Precisión: el nuevo "
     "marco delimita o especifica su contenido. Divergencia: ambos marcos no resultan equivalentes, incluido el caso en que "
     "el nuevo marco no tiene categoría equivalente o introduce una nueva.", "Nota metodológica, numeral 5.5"),
    ("REV_ALI.5", "Efecto", "Se mantiene, se ajusta o se rediseña. En el plano del problema equivalen a «se mantiene», «se "
     "precisa» y «se reformula».", "Nota metodológica, Tabla 9"),
    ("REV_ALI.5", "Regla para los servicios", "Se mantiene: la definición del servicio cabe en un mandato de la Ley sin "
     "cambio de contenido; la mención al enfoque de género se lee como enfoque de igualdad de oportunidades (Tercera DCF). "
     "Se ajusta: la definición usa una categoría que cambia (violencia basada en género, derechos sexuales y reproductivos, "
     "transversalización del enfoque de género, igualdad de género en la educación) o la Ley añade condiciones. Se "
     "rediseña: la Ley dispone algo contrario a la definición del servicio.", "Equipo evaluador"),
    ("REV_ALI.5", "Contraste", "Las filas con pregunta se contrastan con la DGIGND en la entrevista al nivel conductor "
     "(ENTREV, bloque 1).", "Nota metodológica, numeral 5.5"),
    ("Todas", "Citas", "Cada cita es literal y se ubica automáticamente en el texto del documento (página del PDF; párrafo "
     "no vacío en el marco conceptual, que es un .docx). Una cita no ubicada detiene la ejecución. La página es la del "
     "PDF, que puede diferir de la numeración impresa.", "Equipo evaluador"),
    ("Todas", "Carácter", "Valoración del equipo evaluador, validada por el usuario (2026-10-31 y 2026-11-06).",
     "Equipo evaluador"),
], columns=["Submatriz", "Elemento", "Criterio", "Plano"])

hojas = [("REV_ALI.1 Por servicio", c41), ("REV_ALI.1 Por lineamiento", c41_lin), ("REV_ALI.1 Por objetivo", c41_op),
         ("REV_ALI.1 Modelo causal", modelo), ("REV_ALI.2 Por objetivo", v_op), ("REV_ALI.2 Registro", ver_r),
         ("REV_ALI.2 Declarado 2019", decl_r), ("REV_ALI.2 Indicadores PEDN", indp_r)]
if hor is not None:
    hojas += [("REV_ALI.3 Horizontal", hor_r), ("REV_ALI.3 Por servicio PNIG", por_servicio),
              ("REV_ALI.3 Universo", universo)]
hojas += [("REV_ALI.4 Por objetivo", n_op), ("REV_ALI.4 Registro", nor_r), ("REV_ALI.5 Conceptual", conceptual),
          ("Resumen REV_ALI.5", resumen5), ("Para la entrevista DGIGND", entrevista), ("Criterios", criterios)]
ANCHO = {"Sustento": 60, "Criterio": 80, "Para contrastar con la DGIGND (ENTREV)": 60, "Cita": 50, "Vacíos detectados": 50,
         "Observaciones complementarias": 45, "Servicio": 45, "Lineamiento": 40, "Elemento de la PNIG": 45,
         "Cita de la PNIG": 45, "Categoría del nuevo enfoque": 45, "Cita de la Ley o del marco": 45}
with pd.ExcelWriter(SALIDA, engine="openpyxl") as xw:
    for nombre, df in hojas:
        df.to_excel(xw, sheet_name=nombre, index=False)
        ws = xw.sheets[nombre]
        ws.freeze_panes = "C2"
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F3864")
            c.alignment = Alignment(wrap_text=True, vertical="top")
        for col in ws.columns:
            titulo = str(col[0].value)
            ws.column_dimensions[col[0].column_letter].width = ANCHO.get(
                titulo, 30 if "[" in titulo or titulo in NOR_COLS or titulo.startswith("Causa") else 16)
            for c in col[1:]:
                c.alignment = Alignment(wrap_text=True, vertical="top")

print("REV_ALI.1:", c41.Valoración.value_counts().to_dict(), "| lineamientos:", c41_lin.Valoración.str.split(" ").str[0].value_counts().to_dict())
print("REV_ALI.2:", ver.groupby("momento").correspondencia.value_counts().to_dict())
print("REV_ALI.4:", nor.correspondencia.value_counts().to_dict())
print("REV_ALI.5:", len(conceptual), "filas |", len(entrevista), "preguntas")
