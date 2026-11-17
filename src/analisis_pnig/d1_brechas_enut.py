"""D1: tiempo semanal de trabajo no remunerado por sexo, ENUT 2010 y 2024 (12 años a más).

ENUT 2010 (lista de actividades con horas de lunes a viernes y de sábado a domingo): trabajo doméstico no remunerado =
secciones C a J (cocina, aseo, ropa, reparaciones, cuidado de niños y de enfermos, compras, gestión del hogar), N
(apoyo a otros hogares), O (trabajo voluntario y comunal) y P (cuidado de miembros con dificultades). Con esa
agrupación se reproduce la cifra del diagnóstico de la PNIG (39 h 28 min y 15 h 54 min semanales; Freyre y López,
2011, p. 13; p. 17 del documento de la PNIG).

ENUT 2024 (diario por franjas de 10 minutos de un día de semana, un sábado y un domingo): traducción a Python de la
sintaxis «02_calculo_indicadores_ENUT.do» de la carpeta administrativa (grupo de la CEPAL = primer dígito del código
de la actividad principal). Trabajo no remunerado = grupos 3 (doméstico para el propio hogar), 4 (cuidado de
miembros del hogar) y 5 (otros hogares, comunidad y voluntariado). Horas semanales = (5 × día de semana + sábado +
domingo) / 60.

Los instrumentos difieren (lista de actividades frente a diario), por lo que los niveles de 2010 y 2024 no son
estrictamente comparables; sí lo son la brecha relativa y la razón entre mujeres y hombres dentro de cada año.
Universo: residentes de 12 años a más con entrevista completa. Diseño: conglomerado y, como estrato, región o
departamento por área (aproximación; el INEI no publica el estrato).

Uso (desde la raíz del repositorio):
    uv run --no-project --with pandas --with pyreadstat python -I -X utf8 src/analisis_pnig/d1_brechas_enut.py
Salida: evaluaciones/pnig_2026/04_analisis/d1_brechas/uso_tiempo_enut.csv
"""
import glob
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

RAIZ = f"{EVAL}/03_datos/raw/microdatos_inei/enut"
SALIDA = f"{EVAL}/04_analisis/d1_brechas"
SECCIONES_2010 = set("CDEFGHIJNOP")
DIAGNOSTICO = {"Mujer": 39 + 28 / 60, "Hombre": 15 + 54 / 60}


def leer(patron, **kw):
    p = glob.glob(f"{RAIZ}/**/{patron}", recursive=True)[0]
    lector = pyreadstat.read_sav if p.endswith(".sav") else pyreadstat.read_dta
    df, _ = lector(p, apply_value_formats=False, encoding="latin1", **kw)
    df.columns = [c.upper() for c in df.columns]
    return df


def enut_2010():
    c5 = leer("05_CAPITULO_500.sav")
    c2 = leer("02_CAPITULO_200.sav", usecols=None)
    k = ["CONGLOMERADO", "NSELV", "HOGAR", "TOT_HOGAR", "P201"]
    c5["P201"] = c5["P500_A"]
    for df in (c5, c2):
        for c in k:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    d = c5.merge(c2[k + ["P204", "P205"]], on=k, how="left")
    d = d[(d.RESULT == 1) & (pd.to_numeric(d.P205, errors="coerce") >= 12) & d.P204.isin([1, 2])].copy()
    horas = pd.Series(0.0, index=d.index)
    for c in d.columns:
        m = re.match(r"P501([A-Z])_(\d+)$", c)
        if m and m.group(1) in SECCIONES_2010:
            s, i = m.groups()
            for x in ("2", "3"):
                for u, f in (("H", 1), ("M", 1 / 60)):
                    col = f"P50{x}{s}_{i}_{u}"
                    if col in d:
                        horas += pd.to_numeric(d[col], errors="coerce").fillna(0) * f
    d["horas_no_remunerado"] = horas
    d["sexo"] = np.where(d.P204 == 2, "Mujer", "Hombre")
    d["peso"] = pd.to_numeric(d.FAC500, errors="coerce")
    d["estrato"] = d.REGION.astype(str) + "-" + d.AREA.astype(str)
    return d.reset_index(drop=True), "CONGLOMERADO"


def enut_2024():
    act = [f"P600_{x}_ACT1" for x in ("DSEM", "SAB", "DOM")]
    k = ["CONG", "NSELV", "HOGAR", "P201"]
    t = leer("V_ENUT2024_600.dta", usecols=k + act + ["FACTORFINAL", "CCDD", "AREA"])
    for c in k:
        t[c] = pd.to_numeric(t[c], errors="coerce")
    grupo = {x: pd.to_numeric(t[f"P600_{x}_ACT1"], errors="coerce").floordiv(
        10 ** (np.floor(np.log10(pd.to_numeric(t[f"P600_{x}_ACT1"], errors="coerce").clip(lower=1))))) for x in
        ("DSEM", "SAB", "DOM")}
    for x in ("DSEM", "SAB", "DOM"):
        t[f"nr_{x}"] = grupo[x].isin([3, 4, 5]).astype(int) * 10
    por_persona = t.groupby(k).agg(nr_DSEM=("nr_DSEM", "sum"), nr_SAB=("nr_SAB", "sum"), nr_DOM=("nr_DOM", "sum"),
                                   peso=("FACTORFINAL", "first"), CCDD=("CCDD", "first"),
                                   AREA=("AREA", "first")).reset_index()
    por_persona["horas_no_remunerado"] = (5 * por_persona.nr_DSEM + por_persona.nr_SAB + por_persona.nr_DOM) / 60
    m = leer("V_ENUT2024_200.dta")
    for c in k:
        m[c] = pd.to_numeric(m[c], errors="coerce")
    cols = [c for c in ("P204", "P205_A", "P206", "P207", "P208", "PER_RESFIN") if c in m.columns]
    d = por_persona.merge(m[k + cols], on=k, how="left")
    resi = ((d.P206 == 1) & (d.P207 == 2)) | ((d.P206 == 2) & (d.P208 == 1))
    ok = resi & (pd.to_numeric(d.P205_A, errors="coerce") >= 12) & d.P204.isin([1, 2])
    if "PER_RESFIN" in d.columns:
        ok &= d.PER_RESFIN == 1
    d = d[ok].copy()
    d["sexo"] = np.where(d.P204 == 2, "Mujer", "Hombre")
    d["peso"] = pd.to_numeric(d.peso, errors="coerce")
    d["estrato"] = d.CCDD.astype(str) + "-" + d.AREA.astype(str)
    return d.reset_index(drop=True), "CONG"


def calcular():
    filas = []
    for anio, f in ((2010, enut_2010), (2024, enut_2024)):
        d, psu = f()
        D = svy.Diseno(d, "peso", "estrato", psu)
        r = svy.brecha(D, d.horas_no_remunerado, d.sexo, "Mujer", "Hombre")
        filas.append({"indicador": "horas_no_remunerado", "anio": anio, **r, "n": len(d)})
        print(f"ENUT {anio}: {len(d)} personas de 12 años a más")
    t = pd.DataFrame(filas)
    t["etiqueta"] = "Horas semanales de trabajo no remunerado (doméstico, cuidado, otros hogares y voluntariado)"
    t["encuesta"] = "ENUT"
    t["diagnostico_mujer"] = np.where(t.anio == 2010, DIAGNOSTICO["Mujer"], np.nan)
    t["diagnostico_hombre"] = np.where(t.anio == 2010, DIAGNOSTICO["Hombre"], np.nan)
    t["nota"] = np.where(t.anio == 2010, "Lista de actividades; reproduce el diagnóstico de la PNIG",
                         "Diario de actividades (10 minutos, actividad principal); nivel no comparable con 2010")
    return t


if __name__ == "__main__":
    os.makedirs(SALIDA, exist_ok=True)
    t = calcular()
    t.to_csv(f"{SALIDA}/uso_tiempo_enut.csv", index=False, encoding="utf-8")
    print(t[["anio", "valor_a", "ee_a", "valor_b", "ee_b", "diferencia", "ee_diferencia", "razon",
             "diagnostico_mujer", "diagnostico_hombre"]].round(2).to_string(index=False))
