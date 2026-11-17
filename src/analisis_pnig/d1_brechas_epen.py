"""D1: brechas laborales entre mujeres y hombres con la Encuesta Permanente de Empleo Nacional (EPEN), 2023–2025.

Desde 2024 la ENAHO no publica la variable de empleo informal; el INEI la mide con la EPEN, cuya base anual
departamental (módulo 76) se usa aquí. Traducción a Python de la sintaxis «Indicadores OP4.do» de la carpeta
administrativa (DGPDAEM): informalidad de la ocupación principal (Informal_P), afiliación a pensiones (C364_1 a C364_4),
ocupados de 14 años a más (C208 > 13 y OCUP300 = 1), factor FAC300_ANUAL. Se agregan la tasa de actividad y el ingreso
laboral mensual (ingtrabw) para comparar con las series de la ENAHO. Diseño: ESTRATO y CONGLOMERADO (svy.py).

La base anual de 2022 está dañada en el portal del INEI (error de CRC; registro en descargas_microdatos-inei_*.tsv),
por lo que el valor de 2022 se toma del INEI, «Perú: Brechas de Género, 2025», cuadro 6.14. Las cifras de la EPEN no
se empalman con las de la ENAHO: son encuestas distintas y sus niveles difieren.

Uso (desde la raíz del repositorio):
    uv run --no-project --with pandas --with pyreadstat python -I -X utf8 src/analisis_pnig/d1_brechas_epen.py
Salidas: evaluaciones/pnig_2026/04_analisis/d1_brechas/brechas_epen.csv y brechas_epen_departamental.csv
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

RAIZ = f"{EVAL}/03_datos/raw/microdatos_inei/epen"
SALIDA = f"{EVAL}/04_analisis/d1_brechas"
COLS = ["CONGLOMERADO", "ESTRATO", "CCDD", "DIVISION_LIMA", "C207", "C208", "OCUP300", "RESIDENT", "Informal_P", "ingtrabw", "FAC300_ANUAL",
        "C364_1", "C364_2", "C364_3", "C364_4"]
ETIQUETAS = {
    "informal_epen": "Empleo informal (ocupados de 14 años a más, EPEN)",
    "tasa_actividad_epen": "Tasa de actividad (14 años a más, EPEN)",
    "ingreso_laboral_epen": "Ingreso laboral mensual promedio (ocupados con ingreso, EPEN)",
    "afiliado_pension_epen": "Ocupados afiliados a un sistema de pensiones (EPEN)",
}
# INEI, Perú: Brechas de Género 2025: cuadro 6.14 (informalidad) y gráfico 6.1 (tasa de actividad)
OFICIAL = {("informal_epen", 2022): (76.8, 71.7), ("informal_epen", 2023): (73.4, 69.3),
           ("informal_epen", 2024): (73.3, 69.1), ("tasa_actividad_epen", 2023): (62.0, 78.4),
           ("tasa_actividad_epen", 2024): (61.6, 78.2)}


def leer(anio):
    p = glob.glob(f"{RAIZ}/{anio}/**/*.dta", recursive=True)
    if not p:
        return None
    _, meta = pyreadstat.read_dta(p[0], metadataonly=True, encoding="latin1")
    usar = [c for c in COLS if c in meta.column_names]
    df, _ = pyreadstat.read_dta(p[0], usecols=usar, apply_value_formats=False, encoding="latin1")
    for c in df.columns:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


def calcular():
    filas, deps = [], []
    for anio in (2023, 2024, 2025):
        d = leer(anio)
        if d is None:
            continue
        d = d[(d.C208 >= 14) & d.C207.isin([1, 2])].reset_index(drop=True)
        d["sexo"] = np.where(d.C207 == 2, "Mujer", "Hombre")
        # la EPEN no separa el Callao: DIVISION_LIMA 1 es la «Provincia de Lima, incluye la Provincia Constitucional del
        # Callao» y 2 la «Región de Lima»; el INEI publica así sus cifras (Brechas de Género 2025, p. 83)
        cc = d.CCDD.astype(int).astype(str).str.zfill(2)
        d["ambito"] = np.where(cc == "15", np.where(d.DIVISION_LIMA == 1, "LMC", "RL"), cc)
        ocupado = d.OCUP300 == 1
        d["informal_epen"] = np.where(ocupado & d.Informal_P.isin([1, 2]), (d.Informal_P == 1).astype(float), np.nan)
        d["tasa_actividad_epen"] = d.OCUP300.isin([1, 2]).astype(float)
        d["ingreso_laboral_epen"] = np.where(ocupado & (d.ingtrabw > 0), d.ingtrabw, np.nan)
        pens = d[["C364_1", "C364_2", "C364_3", "C364_4"]]
        d["afiliado_pension_epen"] = np.where(ocupado, (pens == 1).any(axis=1).astype(float), np.nan)
        D = svy.Diseno(d, "FAC300_ANUAL", "ESTRATO", "CONGLOMERADO")
        for ind in ETIQUETAS:
            escala = 1 if ind.startswith("ingreso") else 100
            filas.append({"indicador": ind, "anio": anio, **svy.brecha(D, d[ind], d.sexo, "Mujer", "Hombre",
                                                                        escala=escala)})
            # la EPEN Departamentos tiene inferencia anual por departamento (INEI, Indicadores del Mercado Laboral a
            # nivel departamental 2022-2024, num. 6.5, p. 104)
            deps += [{"indicador": ind, "anio": anio, **r} for r in
                     svy.por_dominio(d, "ambito", "FAC300_ANUAL", "ESTRATO", "CONGLOMERADO", d[ind], "sexo",
                                     escala=escala)]
        print(f"EPEN {anio}: {len(d)} personas de 14 años a más")
    a, b = OFICIAL[("informal_epen", 2022)]
    filas.append({"indicador": "informal_epen", "anio": 2022, "valor_a": a, "valor_b": b, "diferencia": a - b,
                  "razon": a / b, "nota": "Valor publicado por el INEI (Brechas de Género 2025, cuadro 6.14); base "
                                          "anual 2022 dañada en el portal"})
    t = pd.DataFrame(filas).sort_values(["indicador", "anio"])
    t.insert(1, "etiqueta", t.indicador.map(ETIQUETAS))
    t.insert(2, "unidad", np.where(t.indicador.str.startswith("ingreso"), "S/", "%"))
    t.insert(3, "encuesta", "EPEN")
    dep = pd.DataFrame(deps)
    dep.insert(0, "encuesta", "EPEN")
    dep.insert(3, "departamento", dep.ambito.map(svy.AMBITOS))
    dep.insert(4, "unidad", np.where(dep.indicador.str.startswith("ingreso"), "S/", "%"))
    return t, dep


def validar(t):
    filas = []
    for (ind, anio), (m, h) in OFICIAL.items():
        r = t[(t.indicador == ind) & (t.anio == anio)]
        if r.empty or anio == 2022:
            continue
        r = r.iloc[0]
        for sexo, ref, v, ee in (("Mujer", m, r.valor_a, r.ee_a), ("Hombre", h, r.valor_b, r.ee_b)):
            filas.append({"indicador": ind, "anio": anio, "sexo": sexo, "referencia": ref,
                          "origen": "INEI, Brechas de Género 2025", "calculado": round(v, 2),
                          "diferencia": round(v - ref, 2), "dentro_ic95": abs(v - ref) <= svy.Z95 * ee + 0.05})
    return pd.DataFrame(filas)


if __name__ == "__main__":
    t, dep = calcular()
    t.to_csv(f"{SALIDA}/brechas_epen.csv", index=False, encoding="utf-8")
    dep.to_csv(f"{SALIDA}/brechas_epen_departamental.csv", index=False, encoding="utf-8")
    v = validar(t)
    v.to_csv(f"{SALIDA}/validacion_epen.csv", index=False, encoding="utf-8")
    pd.set_option("display.width", 200)
    print(v.to_string(index=False))
    print(t[["indicador", "anio", "valor_a", "valor_b", "diferencia", "ee_diferencia", "razon"]].round(2)
          .to_string(index=False))
