"""D1: brechas entre productoras y productores agropecuarios con la Encuesta Nacional Agropecuaria (ENA), 2019–2025.

El diagnóstico de la PNIG (p. 19) señala que las mujeres «tienen poco acceso a recursos productivos y créditos» y que
«pocas poseen tierras», con datos del CENAGRO 2012 (1,8 ha agrícolas en promedio para las productoras frente a 3 ha
para los productores). La ENA permite seguir esas brechas entre productores agropecuarios que son personas naturales
(fila del productor, P1102 = 1, en el capítulo 1100):
  - productora_mujer: porcentaje de mujeres entre quienes conducen una unidad agropecuaria (2019, 2021 a 2025);
  - credito_solicito: productores que solicitaron algún crédito en los últimos 12 meses (P901; capítulo 900; 2019 y
    2022 a 2025; la ENA 2021 no publica ese capítulo);
  - propietario: productores que son propietarios de al menos una parcela (P111_1 en el capítulo 100A desde 2023;
    P110_1 en el capítulo 100 de 2019 a 2022, misma pregunta);
  - titulo: propietarios con título de propiedad, inscrito o no, de al menos una parcela (P112 = 1 o 2 desde 2023; P111
    de 2019 a 2022, mismas opciones; el INEI calcula el 26,0 % de 2024 sobre los propietarios);
  - superficie_ha y superficie_mediana_ha: superficie total de las parcelas que trabaja o conduce (P105_SUP_ha), en
    promedio y en mediana ponderada; el promedio es sensible a parcelas muy grandes.
Los capítulos de parcelas de 2019 a 2022 no se descargaron: su estructura difiere y el INEI publica la serie desde 2023.

Diseño: estratos departamento × ESTRATO y conglomerado (CONGLOMERADO hasta 2022, NSEGM desde 2023); factor FACTOR o
FACTOR_PRODUCTOR. Los errores estándar son una aproximación: el INEI no publica la variable de estrato geográfico.
Validación con el INEI, «Perú: Brechas de Género, 2025», pp. 117 a 119 (ENA 2024).

Uso (desde la raíz del repositorio):
    uv run --no-project --with pandas --with pyreadstat python -I -X utf8 src/analisis_pnig/d1_brechas_ena.py
Salidas: evaluaciones/pnig_2026/04_analisis/d1_brechas/brechas_ena.csv, brechas_ena_departamental.csv y validacion_ena.csv
"""
import os
import re
import sys

import numpy as np
import pandas as pd
import pyreadstat

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import svy  # noqa: E402
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

RAIZ = f"{EVAL}/03_datos/raw/microdatos_inei/ena"
SALIDA = f"{EVAL}/04_analisis/d1_brechas"
ANIOS = (2019, 2021, 2022, 2023, 2024, 2025)
ETIQUETAS = {
    "productora_mujer": "Mujeres entre quienes conducen una unidad agropecuaria (personas naturales)",
    "credito_solicito": "Productores que solicitaron algún crédito en los últimos 12 meses",
    "propietario": "Productores propietarios de al menos una parcela",
    "titulo": "Propietarios con título de propiedad de al menos una parcela",
    "superficie_ha": "Superficie promedio de las parcelas que trabaja o conduce (ha)",
    "superficie_mediana_ha": "Superficie mediana de las parcelas que trabaja o conduce (ha)",
}
# INEI, Perú: Brechas de Género 2025, ENA 2024: (indicador, sexo o total): (valor, página)
OFICIAL = {("productora_mujer", "Total"): (31.5, 117), ("credito_solicito", "Mujer"): (6.1, 119),
           ("credito_solicito", "Hombre"): (10.5, 119), ("credito_solicito", "Total"): (9.1, 119),
           ("propietario", "Total"): (70.8, 118), ("propietario", "Mujeres entre propietarios"): (33.7, 118),
           ("titulo", "Total"): (26.0, 118), ("titulo", "Mujeres entre titulares"): (33.0, 118)}


def archivo(anio, patron):
    for raiz, _, nombres in os.walk(f"{RAIZ}/{anio}"):
        for n in nombres:
            if re.fullmatch(patron, n, re.I):
                return os.path.join(raiz, n)
    return None


def leer(p, cols):
    _, meta = pyreadstat.read_dta(p, metadataonly=True, encoding="latin1")
    usar = [c for c in cols if c in meta.column_names]
    df, _ = pyreadstat.read_dta(p, usecols=usar, apply_value_formats=False, encoding="latin1")
    for c in df.columns:
        if c != "CCDD":
            df[c] = pd.to_numeric(df[c], errors="coerce")
    df["CCDD"] = df.CCDD.astype(str).str.zfill(2)
    return df


def productores(anio):
    p = archivo(anio, r"\d+_cap1100\.dta")
    psu, ident, peso = ("CONGLOMERADO", "NSELUA", "FACTOR") if anio <= 2022 else ("NSEGM", "ID_PROD",
                                                                                    "FACTOR_PRODUCTOR")
    clave = ["CCDD", psu, ident, "UA"]
    d = leer(p, clave + ["ESTRATO", peso, "RESFIN", "P1102", "P1103"])
    d = d[(d.P1102 == 1) & d.P1103.isin([1, 2])].drop_duplicates(clave)
    d["sexo"] = np.where(d.P1103 == 2, "Mujer", "Hombre")
    # la ENA 2021 y 2022 no publican ESTRATO: el estrato es solo el departamento
    d["estrato"] = d.CCDD + "-" + (d.ESTRATO.map(str) if "ESTRATO" in d else "")
    d["psu"] = d.CCDD + "-" + d[psu].map(str)
    d["peso"] = d[peso]
    d["productora_mujer"] = (d.P1103 == 2).astype(float)
    p9 = archivo(anio, r"\d+_cap900\.dta")
    if p9:
        c = leer(p9, clave + ["P901"]).drop_duplicates(clave)
        d = d.merge(c[clave + ["P901"]], on=clave, how="left")
        d["credito_solicito"] = np.where(d.P901.isin([1, 2]), (d.P901 == 1).astype(float), np.nan)
    # mismas preguntas con otros nombres: P110_1 y P111 en el capítulo 100 (2019 a 2022), P111_1 y P112 en el 100A
    p1 = archivo(anio, r"\d+_cap100a_03\.dta") or archivo(anio, r"\d+_cap100_2\.dta")
    prop, tit = ("P111_1", "P112") if anio >= 2023 else ("P110_1", "P111")
    if p1:
        par = leer(p1, clave + [prop, tit, "P105_SUP_ha"])
        par["prop"] = (par[prop] == 1).astype(float)
        par["tit"] = par[tit].isin([1, 2]).astype(float)
        g = par.groupby(clave).agg(propietario=("prop", "max"), titulo=("tit", "max"),
                                   superficie_ha=("P105_SUP_ha", "sum")).reset_index()
        d = d.merge(g, on=clave, how="left")
        d["propietario"] = d.propietario.fillna(0.0)
        d["titulo"] = d.titulo.where(d.propietario == 1)
        d["superficie_ha"] = d.superficie_ha.where(d.superficie_ha > 0)
    return d.reset_index(drop=True)


def mediana(x, w):
    """Mediana ponderada."""
    x, w = np.asarray(x, float), np.asarray(w, float)
    o = np.argsort(x)
    c = np.cumsum(w[o]) / w.sum()
    return float(x[o][np.searchsorted(c, 0.5)])


def por_departamento(d, ind, anio):
    """La ENA de productores tiene inferencia departamental (ficha técnica 2024, num. 8, p. 7)."""
    if ind == "productora_mujer":
        filas = []
        for r in svy.por_dominio(d, "CCDD", "peso", "estrato", "psu", d[ind], escala=100):
            filas.append({"ambito": r["ambito"], "valor_a": r["estimacion"], "ee_a": r["ee"], "cv_a": r["cv"],
                          "n_a": r["n"], "valor_b": 100 - r["estimacion"], "ee_b": r["ee"], "n_b": r["n"],
                          "diferencia": 2 * r["estimacion"] - 100, "ee_diferencia": 2 * r["ee"]})
    else:
        filas = svy.por_dominio(d, "CCDD", "peso", "estrato", "psu", d[ind], "sexo", escala=100)
    return [{"indicador": ind, "anio": anio, **f} for f in filas]


def calcular():
    filas, val, deps = [], [], []
    for anio in ANIOS:
        d = productores(anio)
        D = svy.Diseno(d, "peso", "estrato", "psu")
        print(f"ENA {anio}: {len(d)} productores (personas naturales)")
        for ind in ETIQUETAS:
            if ind not in d or d[ind].notna().sum() == 0:
                continue
            escala = 1 if ind == "superficie_ha" else 100
            if ind != "superficie_ha":
                deps += por_departamento(d, ind, anio)
            if ind == "productora_mujer":
                e = svy.estimar(D, d[ind], escala=100)
                filas.append({"indicador": ind, "anio": anio, "valor_a": e["estimacion"], "ee_a": e["ee"],
                              "n_a": e["n"], "valor_b": 100 - e["estimacion"], "ee_b": e["ee"], "n_b": e["n"],
                              "diferencia": 2 * e["estimacion"] - 100, "ee_diferencia": 2 * e["ee"],
                              "nota": "valor_a: % de mujeres; valor_b: % de hombres entre quienes conducen"})
                if anio == 2024:
                    val.append(("productora_mujer", "Total", e["estimacion"], e["ee"]))
                continue
            b = svy.brecha(D, d[ind], d.sexo, "Mujer", "Hombre", escala=escala)
            filas.append({"indicador": ind, "anio": anio, **b})
            if ind == "superficie_ha":
                m = {s: mediana(g.superficie_ha, g.peso) for s, g in d[d.superficie_ha.notna()].groupby("sexo")}
                filas.append({"indicador": "superficie_mediana_ha", "anio": anio, "valor_a": m["Mujer"],
                              "valor_b": m["Hombre"], "diferencia": m["Mujer"] - m["Hombre"],
                              "razon": m["Mujer"] / m["Hombre"],
                              "nota": "Mediana ponderada, sin error estándar; el promedio es sensible a parcelas "
                                      "muy grandes y su razón oscila entre años"})
            if anio == 2024 and ind != "superficie_ha":
                t = svy.estimar(D, d[ind], escala=100)
                val += [(ind, "Mujer", b["valor_a"], b["ee_a"]), (ind, "Hombre", b["valor_b"], b["ee_b"]),
                        (ind, "Total", t["estimacion"], t["ee"])]
                if ind in ("propietario", "titulo"):
                    s = svy.estimar(D, d.productora_mujer, universo=d[ind] == 1, escala=100)
                    clave = "Mujeres entre propietarios" if ind == "propietario" else "Mujeres entre titulares"
                    val.append((ind, clave, s["estimacion"], s["ee"]))
    t = pd.DataFrame(filas)
    t.insert(1, "etiqueta", t.indicador.map(ETIQUETAS))
    t.insert(2, "unidad", np.where(t.indicador.str.startswith("superficie"), "ha", "%"))
    t.insert(3, "encuesta", "ENA")
    v = []
    for ind, sexo, calc, ee in val:
        if (ind, sexo) in OFICIAL:
            ref, pag = OFICIAL[(ind, sexo)]
            v.append({"indicador": ind, "anio": 2024, "sexo": sexo, "referencia": ref,
                      "origen": f"INEI, Brechas de Género 2025, p. {pag}", "calculado": round(calc, 2),
                      "diferencia": round(calc - ref, 2), "dentro_ic95": abs(calc - ref) <= svy.Z95 * ee + 0.05})
    dep = pd.DataFrame(deps)
    dep.insert(0, "encuesta", "ENA")
    dep.insert(3, "departamento", dep.ambito.map(svy.AMBITOS))
    dep.insert(4, "unidad", "%")
    return t, pd.DataFrame(v), dep


if __name__ == "__main__":
    t, v, dep = calcular()
    t.to_csv(f"{SALIDA}/brechas_ena.csv", index=False, encoding="utf-8")
    dep.to_csv(f"{SALIDA}/brechas_ena_departamental.csv", index=False, encoding="utf-8")
    v.to_csv(f"{SALIDA}/validacion_ena.csv", index=False, encoding="utf-8")
    pd.set_option("display.width", 200)
    print(v.to_string(index=False))
    print(t[["indicador", "anio", "valor_a", "valor_b", "diferencia", "ee_diferencia", "razon"]].round(3)
          .to_string(index=False))
