"""D1 y D2: tolerancia social a la violencia contra las mujeres por sexo, ENARES 2019 y 2024 (18 años a más).

Traducción a Python de la sintaxis Stata «tolerancia_violencia.do» de la carpeta administrativa: una persona tolera
la violencia contra las mujeres si está de acuerdo o muy de acuerdo con al menos una afirmación que la justifica, o en
desacuerdo con una que la rechaza (ítems 3, 4 y 16 de la pregunta 401). Se calculan dos versiones:
  - índice completo de 2024, con los 22 ítems de la pregunta 401 (la de la sintaxis), contrastado con el valor
    oficial de IOP.06.01 en 2024;
  - índice comparable 2019–2024, con los 10 ítems que se repiten con el mismo texto en ambos cuestionarios
    (ítems 1 a 10; los ítems 3 y 4 en sentido inverso).
La ENARES 2015 usa otro instrumento (afirmaciones distintas), por lo que su valor oficial (54,8 %) se cita como
referencia y no se recalcula.

Ninguna de las dos versiones reproduce el índice oficial de IOP.06.01 (58,9 % en 2019 y 50,0 % en 2024): el INEI no
publica cómo agrega las afirmaciones, y ningún umbral de respuestas tolerantes reproduce ambos valores. Por eso la
serie oficial se mantiene como referencia y el índice comparable se usa solo para la brecha entre mujeres y hombres y
para la dirección del cambio entre 2019 y 2024, como cálculo del equipo evaluador.

Diseño: factor de la persona seleccionada (factor_hym) y conglomerado. Como estrato se usa departamento por área en
ambos años: en 2019 el estrato del INEI deja un conglomerado por estrato y en 2024 no se publica.

Uso (desde la raíz del repositorio):
    uv run --no-project --with pandas --with pyreadstat --with openpyxl python -I -X utf8 src/analisis_pnig/d1_brechas_enares.py
Salidas: evaluaciones/pnig_2026/04_analisis/d1_brechas/tolerancia_enares.csv y tolerancia_enares_departamental.csv
(solo 2024: las ediciones 2013, 2015 y 2019 tienen inferencia nacional y la de 2024, nacional y departamental; INEI,
Informe de los Principales Resultados ENARES 2024, p. 21)
"""
import glob
import os
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

RAIZ = f"{EVAL}/03_datos/raw/microdatos_inei/enares"
SALIDA = f"{EVAL}/04_analisis/d1_brechas"
LLAVE = ["ID", "HOGAR_ID", "PERSONA_ID"]
INVERSOS = {3, 4, 16}          # tolerar = estar en desacuerdo
COMUNES = range(1, 11)          # mismo texto en 2019 y 2024
OFICIAL = {2015: 54.8, 2019: 58.9, 2024: 50.0}  # IOP.06.01, reportes de seguimiento (REV_IND, hoja «Serie anual»)
ARCHIVOS = {2019: ("12_v_c2cap200.dta", "14_v_c2cap400.dta"), 2024: ("12_CRS02_CAP200.dta", "14_CRS02_CAP400.dta")}


def leer(anio, nombre):
    p = glob.glob(f"{RAIZ}/{anio}/**/{nombre}", recursive=True)[0]
    df, _ = pyreadstat.read_dta(p, apply_value_formats=False, encoding="latin1")
    df.columns = [c.upper() for c in df.columns]
    return df


def tolera(df, items):
    cols = [f"C2P401_{i}" for i in items]
    resp = df[cols].apply(pd.to_numeric, errors="coerce")
    t = pd.Series(False, index=df.index)
    for i, c in zip(items, cols):
        t |= resp[c].isin([1, 2]) if i in INVERSOS else resp[c].isin([3, 4])
    valido = resp.notna().any(axis=1)
    return np.where(valido, t.astype(float), np.nan)


def base(anio):
    f200, f400 = ARCHIVOS[anio]
    roster = leer(anio, f200)[LLAVE + ["C1P207", "C1P208_A"]]
    d = leer(anio, f400).merge(roster, on=LLAVE, how="left")
    d["sexo"] = np.where(d.C1P207 == 2, "Mujer", np.where(d.C1P207 == 1, "Hombre", None))
    d = d[(pd.to_numeric(d.C1P208_A, errors="coerce") >= 18) & d.sexo.notna()].reset_index(drop=True)
    d["ESTRATO"] = d.CCDD.astype(str) + "-" + d.AREA.astype(str)
    if "CONGLOMERADO" not in d.columns:
        d["CONGLOMERADO"] = d.CONGLOME
    d["tolerancia_10"] = tolera(d, COMUNES)
    if anio == 2024:
        d["tolerancia_22"] = tolera(d, range(1, 23))
    return d


def calcular():
    filas, deps = [], []
    for anio in (2019, 2024):
        d = base(anio)
        D = svy.Diseno(d, "FACTOR_HYM", "ESTRATO", "CONGLOMERADO")
        for ind in ("tolerancia_10", "tolerancia_22"):
            if ind not in d.columns:
                continue
            total = svy.estimar(D, d[ind], escala=100)
            r = svy.brecha(D, d[ind], d.sexo, "Mujer", "Hombre", escala=100)
            filas.append({"indicador": ind, "anio": anio, "total": total["estimacion"], "ee_total": total["ee"],
                          "n": total["n"], **r})
        if anio == 2024:
            cc = pd.to_numeric(d.CCDD).astype(int).astype(str).str.zfill(2)
            pr = pd.to_numeric(d.CCPP).astype(int)
            d["ambito"] = np.where(cc == "15", np.where(pr == 1, "LM", "RL"), cc)
            deps += [{"indicador": "tolerancia_10", "anio": anio, **r} for r in
                     svy.por_dominio(d, "ambito", "FACTOR_HYM", "ESTRATO", "CONGLOMERADO", d["tolerancia_10"], "sexo",
                                     escala=100)]
        print(f"ENARES {anio}: {len(d)} personas de 18 años a más")
    t = pd.DataFrame(filas)
    t["etiqueta"] = t.indicador.map({"tolerancia_10": "Tolerancia social a la violencia contra las mujeres "
                                                      "(10 afirmaciones comunes a 2019 y 2024)",
                                     "tolerancia_22": "Tolerancia social a la violencia contra las mujeres "
                                                      "(22 afirmaciones del cuestionario 2024)"})
    t["oficial_IOP0601"] = t.anio.map(OFICIAL)
    t["encuesta"] = "ENARES"
    dep = pd.DataFrame(deps)
    dep.insert(0, "encuesta", "ENARES")
    dep.insert(3, "departamento", dep.ambito.map(svy.AMBITOS))
    dep.insert(4, "unidad", "%")
    return t, dep


if __name__ == "__main__":
    os.makedirs(SALIDA, exist_ok=True)
    t, dep = calcular()
    t.to_csv(f"{SALIDA}/tolerancia_enares.csv", index=False, encoding="utf-8")
    dep.to_csv(f"{SALIDA}/tolerancia_enares_departamental.csv", index=False, encoding="utf-8")
    pd.set_option("display.width", 220)
    print(t[["indicador", "anio", "total", "ee_total", "valor_a", "valor_b", "diferencia", "ee_diferencia",
             "oficial_IOP0601", "n"]].round(2).to_string(index=False))
