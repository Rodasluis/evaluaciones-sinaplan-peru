"""REV_DOC: contraste de la clasificación de los servicios de D7 con la matriz técnica de incompatibilidades de la DPIGND.

La matriz (anexa al Informe N.° D000041-2026-MIMP-DGIGND, 09-11-2026) revisa, disposición por disposición, la
compatibilidad de la PNIG, el PEMIG y la CONAIG con la Ley N.° 32535 y, para cada grupo de servicios de la PNIG, dice
con qué base legal su contenido puede seguir ejecutándose o por qué es incompatible. La evaluación clasificó antes, por
su cuenta, cada servicio según su correspondencia con el nuevo marco (REV_ALI.5: se mantiene, se ajusta o se
rediseña). Este script compara las dos clasificaciones (decisión del usuario del 2026-11-17, ADR-016):
  1. expande los grupos de servicios de la matriz («Servicios 1.1.1 a 1.1.6 y 1.1.8») a sus códigos;
  2. clasifica cada grupo con la regla fija de 02_insumos/matriz_incompatibilidades.yaml (texto de su observación);
  3. cruza esa clasificación con la de REV_ALI.5: tabla cruzada, concordancia y kappa de Cohen.
La matriz no modifica la clasificación de la evaluación; sirve para contrastarla.

Uso (desde la raíz del repositorio, después de rev_ali_coherencia_alineamiento.py):
    uv run python src/analisis_pnig/rev_doc_matriz_incompatibilidades.py
Salida: evaluaciones/<id>/04_analisis/REV_DOC_matriz_incompatibilidades.xlsx
"""
import re

import pandas as pd
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.codificacion import acuerdo, kappa  # noqa: E402
from nucleo.config import cargar as _cargar_config  # noqa: E402
from nucleo.config import insumo  # noqa: E402

EVAL = _cargar_config().raiz
CFG = insumo("matriz_incompatibilidades.yaml")
SERVICIOS = insumo("servicios.csv")
MAC = f"{EVAL}/04_analisis/REV_ALI_coherencia_alineamiento.xlsx"
SALIDA = f"{EVAL}/04_analisis/REV_DOC_matriz_incompatibilidades.xlsx"
EFECTOS = ["Se mantiene", "Se ajusta", "Se rediseña"]


def clasificar(observacion):
    for r in CFG["regla"]:
        if re.search(r["patron"], observacion, flags=re.IGNORECASE):
            return r["efecto"]
    raise ValueError(observacion)


def expandir(cabecera, codigos):
    """Códigos de servicio que nombra la cabecera de un grupo de la matriz (rangos «a» y códigos sueltos)."""
    orden = {c: i for i, c in enumerate(codigos)}
    salida = set(re.findall(r"\d\.\d\.\d", cabecera))
    for a, b in re.findall(r"(\d\.\d\.\d)\s+a\s+(\d\.\d\.\d)", cabecera):
        salida |= {c for c in codigos if orden[a] <= orden[c] <= orden[b]}
    return salida


codigos = SERVICIOS["codigo_corto"].tolist()
m = pd.read_excel(CFG["matriz"], sheet_name=CFG["hoja_pnig"], header=2)
disp = m["Disposición"].map(str)
i0 = m.index[disp.str.startswith(CFG["desde"])][0]
i1 = m.index[disp.str.startswith(CFG["hasta"])][0]
filas = []
for _, r in m.loc[i0 + 1:i1 - 1].iterrows():
    texto = str(r["Tema y texto literal de la norma (extracto)"])
    cabecera = texto.split("“")[0].strip()
    obs = str(r["Observación"]) if pd.notna(r["Observación"]) else ""
    cods = expandir(cabecera, codigos)
    if not cods:
        continue   # introducción del numeral 3 (lineamiento 3.2 sin servicio)
    for c in cods:
        filas.append({"Código": c, "Grupo de la matriz": cabecera, "¿Compatible? (matriz)": r["¿Compatible?"],
                      "Observación de la matriz": obs, "Clasificación según la matriz": clasificar(obs)})
mat = pd.DataFrame(filas)
assert not mat["Código"].duplicated().any(), mat[mat["Código"].duplicated()]
assert set(mat["Código"]) == set(codigos), set(codigos) ^ set(mat["Código"])

ali = pd.read_excel(MAC, sheet_name="REV_ALI.5 Conceptual")
ali = ali[ali.Plano == "Servicios"].assign(Código=lambda x: x["Código"].map(str))
ali = ali[["Código", "Elemento de la PNIG", "Categoría del nuevo enfoque", "Efecto"]].rename(
    columns={"Elemento de la PNIG": "Servicio", "Categoría del nuevo enfoque": "Base en el nuevo marco (evaluación)",
             "Efecto": "Clasificación de la evaluación (REV_ALI.5)"})
df = ali.merge(mat, on="Código", how="left")
df["Coinciden"] = (df["Clasificación de la evaluación (REV_ALI.5)"] == df["Clasificación según la matriz"]).map(
    {True: "Sí", False: "No"})
orden = {c: i for i, c in enumerate(codigos)}
df = df.sort_values("Código", key=lambda s: s.map(orden)).reset_index(drop=True)

ev, mx = df["Clasificación de la evaluación (REV_ALI.5)"], df["Clasificación según la matriz"]
cruce = pd.crosstab(ev, mx, margins=True, margins_name="Total").reindex(
    index=EFECTOS + ["Total"], columns=EFECTOS + ["Total"], fill_value=0)
cruce.index.name, cruce.columns.name = "Evaluación (filas) / matriz (columnas)", None
k, pa = kappa(ev, mx), acuerdo(ev, mx)
resumen_contraste = pd.DataFrame({
    "Medida": ["Servicios", "Coinciden", "Acuerdo simple (%)", "Kappa de Cohen",
               "Coinciden en «Se rediseña»", "Discrepancias entre «Se mantiene» y «Se ajusta»"],
    "Valor": [len(df), int((df.Coinciden == "Sí").sum()), round(pa, 1), k,
              int(((ev == "Se rediseña") & (mx == "Se rediseña")).sum()),
              int(((ev != mx) & ev.isin(EFECTOS[:2]) & mx.isin(EFECTOS[:2])).sum())]})

res = pd.read_excel(CFG["matriz"], sheet_name=CFG["hoja_resumen"], header=2)
res.columns = [str(c).strip() for c in res.columns]
res = res[pd.to_numeric(res["N°"], errors="coerce").notna()][
    ["Norma", "Revisadas", "% de disposiciones marcadas “NO”", "Situación jurídica (Informe Técnico)"]].rename(
    columns={"Revisadas": "Disposiciones revisadas", "% de disposiciones marcadas “NO”": "Incompatibles (%)"})
res["Incompatibles (%)"] = (res["Incompatibles (%)"] * 100).round(1)

criterios = pd.DataFrame(
    [("Fuente", "Matriz técnica de incompatibilidades de la DPIGND, anexa al Informe N.° D000041-2026-MIMP-DGIGND "
                "(09-11-2026), hoja de la PNIG, numeral 3 de su Anexo (provisión de servicios)."),
     ("Expansión", "Cada grupo de la matriz nombra sus servicios por código o por rango («1.1.1 a 1.1.6 y 1.1.8»); "
                   "se expande con el orden de los 52 servicios. Cada servicio queda en un solo grupo.")]
    + [(f"Regla: {r['efecto']}", f"{r['criterio']} Patrón en la observación: {r['patron']}") for r in CFG["regla"]]
    + [("Contraste", "La clasificación de la evaluación es la de REV_ALI.5 (correspondencia de cada servicio con la "
                     "Ley N.° 32535 y el marco conceptual de la DGIGND), hecha antes de conocer la matriz. Se mide el "
                     "acuerdo simple y el kappa de Cohen (nucleo.codificacion). La matriz no cambia la clasificación."),
       ("Decisión", "Contraste aprobado por el usuario el 2026-11-17 (ADR-016).")],
    columns=["Criterio", "Descripción"])

with pd.ExcelWriter(SALIDA, engine="openpyxl") as xw:
    res.to_excel(xw, sheet_name="Resumen por norma", index=False)
    df.to_excel(xw, sheet_name="Servicios", index=False)
    cruce.to_excel(xw, sheet_name="Tabla cruzada")
    resumen_contraste.to_excel(xw, sheet_name="Concordancia", index=False)
    df[df.Coinciden == "No"].to_excel(xw, sheet_name="Discrepancias", index=False)
    criterios.to_excel(xw, sheet_name="Criterios", index=False)
    anchos = {"Código": 8, "Servicio": 60, "Base en el nuevo marco (evaluación)": 22, "Grupo de la matriz": 50,
              "Observación de la matriz": 90, "Norma": 60, "Situación jurídica (Informe Técnico)": 90,
              "Descripción": 120, "Criterio": 24, "Medida": 45}
    for hoja in xw.sheets.values():
        for col in hoja.columns:
            hoja.column_dimensions[col[0].column_letter].width = anchos.get(col[0].value, 22)
        hoja.freeze_panes = "A2"
print(cruce.to_string())
print(f"Coinciden {int((df.Coinciden == 'Sí').sum())} de {len(df)}; acuerdo {pa:.1f} %; kappa {k}")
