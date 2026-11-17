"""D1 y D2: manifestaciones del problema medidas con la ENDES (mujeres de 15 a 49 años), 2019–2025.

La ENDES entrevista solo a mujeres, de modo que mide las manifestaciones del problema que la PNIG describe en el
diagnóstico y en sus indicadores de OP, no brechas entre mujeres y hombres. Se calculan con códigos numéricos estándar
(no con etiquetas, que cambian entre años):
  - uso de métodos anticonceptivos de las mujeres actualmente unidas (V313 > 0; IOP.02.01) y de métodos modernos
    (V313 = 3; el diagnóstico cita 54,5 % en 2017);
  - demanda insatisfecha de planificación familiar de las unidas (V626 en 1 y 2; IOP.02.02);
  - adolescentes de 15 a 19 años alguna vez embarazadas (V201 > 0 o V213 = 1; IOP.02.03);
  - el esposo o compañero decide solo sobre la salud de la mujer (V743A = 4; el diagnóstico cita 11,5 % en 2017).
La violencia física o sexual del esposo o compañero en los últimos 12 meses (IOP.01.01) se toma del repositorio
«Indicadores ENDES» del usuario (validado contra el INEI) para 2021–2025; para 2019 y 2020 se aplican las mismas
funciones del repositorio a los módulos del INEI.

Cada serie se contrasta con los valores oficiales de los reportes de seguimiento (REV_IND, hoja «Serie anual»).
Diseño: V005/1 000 000, estrato V022, conglomerado V001 (estimador de svy.py). Por departamento (V024; la ficha
técnica 2025 declara inferencia anual para «cada uno de los 24 departamentos del país y la Provincia Constitucional
del Callao», p. 8) se repite el cálculo con el diseño del departamento (svy.por_dominio).

Uso (desde la raíz del repositorio):
    uv run --no-project --with pandas --with pyreadstat --with pyarrow --with openpyxl --with shapely --with requests \
        python -I -X utf8 src/analisis_pnig/d1_brechas_endes.py
Salidas: evaluaciones/pnig_2026/04_analisis/d1_brechas/manifestaciones_endes.csv y
manifestaciones_endes_departamental.csv
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
from nucleo.entorno import ruta_externa  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

# repositorio «Indicadores ENDES» del equipo (variable REPO_ENDES o config/entorno.local.yaml)
REPO_ENDES = ruta_externa("repo_endes") or ""
RAICES = [f"{EVAL}/03_datos/raw/microdatos_inei/endes"] + ([REPO_ENDES + "/data/raw"] if REPO_ENDES else [])
SALIDA = f"{EVAL}/04_analisis/d1_brechas"
PSI = f"{EVAL}/04_analisis/REV_IND_base_indicadores.xlsx"
ANIOS = range(2019, 2026)
ARCHIVOS = {"REC0111": ["CASEID", "V001", "V005", "V012", "V022", "V024", "V025", "UBIGEO"],
            "RE223132": ["CASEID", "V201", "V213", "V313"],
            "RE516171": ["CASEID", "V502", "V626", "V743A"]}
INDICADORES = {
    "uso_anticonceptivo": ("Mujeres actualmente unidas que usan algún método anticonceptivo", "IOP.02.01", ""),
    "uso_moderno": ("Mujeres actualmente unidas que usan un método moderno", "",
                    "54,5 % en 2017 (INEI, 2018b); p. 22 del documento de la PNIG"),
    "demanda_insatisfecha": ("Mujeres actualmente unidas con demanda insatisfecha de planificación familiar",
                             "IOP.02.02", ""),
    "embarazo_adolescente": ("Adolescentes de 15 a 19 años alguna vez embarazadas", "IOP.02.03",
                             "13,4 % en 2017 (INEI, 2018); p. 22"),
    "esposo_decide_salud": ("Mujeres unidas cuya salud decide solo el esposo o compañero", "",
                            "11,5 % en 2017 (INEI); p. 22"),
    "violencia_fisica_o_sexual_12m": ("Violencia física o sexual del esposo o compañero en los últimos 12 meses "
                                      "(mujeres alguna vez unidas)", "IOP.01.01",
                                      "10,6 % en 2017; meta de la situación futura deseada: 2,4 % al 2030 (p. 20)"),
}


def ruta(anio, nombre):
    for raiz in RAICES:
        for ext in ("dta", "sav"):
            hallado = [p for p in glob.glob(f"{raiz}/{anio}/**/{nombre}*.{ext}", recursive=True)
                       if os.path.basename(p).upper().startswith(nombre)]
            if hallado:
                return hallado[0]
    return None


def leer(anio, nombre):
    p = ruta(anio, nombre)
    if p is None:
        raise FileNotFoundError(f"ENDES {anio}: falta {nombre}")
    lector = pyreadstat.read_sav if p.endswith(".sav") else pyreadstat.read_dta
    _, meta = lector(p, metadataonly=True)
    mapa = {c.upper(): c for c in meta.column_names}
    usar = [mapa[c] for c in ARCHIVOS[nombre] if c in mapa]
    df, _ = lector(p, usecols=usar, apply_value_formats=False)
    df.columns = [c.upper() for c in df.columns]
    df["CASEID"] = df.CASEID.astype(str).str.strip()
    return df


def base(anio):
    df = leer(anio, "REC0111")
    for nombre in ("RE223132", "RE516171"):
        df = df.merge(leer(anio, nombre), on="CASEID", how="left")
    for c in df.columns:
        if c != "CASEID":
            df[c] = pd.to_numeric(df[c], errors="coerce")
    df["peso"] = df.V005 / 1_000_000
    df["ambito"] = df.UBIGEO.map(svy.ambito_ubigeo)  # Lima Metropolitana, Región Lima y Callao por separado
    mujer_15_49 = df.V012.between(15, 49)
    unida = mujer_15_49 & (df.V502 == 1)
    df["uso_anticonceptivo"] = np.where(unida & df.V313.notna(), (df.V313 > 0).astype(float), np.nan)
    df["uso_moderno"] = np.where(unida & df.V313.notna(), (df.V313 == 3).astype(float), np.nan)
    df["demanda_insatisfecha"] = np.where(unida & df.V626.notna(), df.V626.isin([1, 2]).astype(float), np.nan)
    adol = df.V012.between(15, 19)
    df["embarazo_adolescente"] = np.where(adol, ((df.V201 > 0) | (df.V213 == 1)).astype(float), np.nan)
    df["esposo_decide_salud"] = np.where(unida & df.V743A.notna(), (df.V743A == 4).astype(float), np.nan)
    return df


def nivel(r):
    """Indicador de mujeres: valor_a es el nivel; no hay valor_b."""
    return {"ambito": r["ambito"], "valor_a": r["estimacion"], "ee_a": r["ee"], "cv_a": r["cv"], "n_a": r["n"]}


def violencia_microdatos(anio, config, indicadores):
    """Violencia 2019 y 2020: los módulos del INEI (REC0111, REC91, RE516171 y REC84DV, en SPSS) con las mismas funciones
    del repositorio, que trabajan con las etiquetas de valor como texto."""
    base = None
    for nombre in ("REC0111", "REC91", "RE516171", "REC84DV"):
        p = ruta(anio, nombre)
        lector = pyreadstat.read_sav if p.lower().endswith(".sav") else pyreadstat.read_dta
        df, _ = lector(p, apply_value_formats=True)
        df.columns = [c.upper() for c in df.columns]
        df["CASEID"] = df.CASEID.astype(str).str.strip()
        if base is None:
            base = df
        else:
            dup = [c for c in df.columns if c in base.columns and c != "CASEID"]
            base = base.merge(df.drop(columns=dup), on="CASEID", how="left", validate="one_to_one")
    base[config.COL_ANIO] = anio
    base = indicadores.marcar_universo(base)
    return indicadores.construir_violencia(base, copiar=False)


def violencia_repositorio():
    """Violencia física o sexual 12 meses, 2019–2025: 2021–2025 con la base del repositorio «Indicadores ENDES»; 2019 y
    2020 con los módulos del INEI y las funciones del mismo repositorio."""
    if not REPO_ENDES:
        raise SystemExit("Falta la ruta del repositorio «Indicadores ENDES» (REPO_ENDES o config/entorno.local.yaml).")
    sys.path.insert(0, REPO_ENDES + "/src")
    from endes_violencia import carga, config, indicadores
    df = indicadores.construir_todo(carga.cargar_base())
    grupos = [(a, violencia_microdatos(a, config, indicadores)) for a in (2019, 2020)]
    grupos += [(int(a), g.reset_index(drop=True)) for a, g in df.groupby(config.COL_ANIO)]
    filas, deps = [], []
    for anio, g in grupos:
        g = g.reset_index(drop=True)
        d = svy.Diseno(g, config.COL_PESO, config.VAR_ESTRATO, config.VAR_CONGLOMERADO)
        nota = ("Repositorio «Indicadores ENDES» (Zenodo 10.5281/zenodo.22823100)" if anio >= 2021 else
                "Módulos del INEI con las funciones del repositorio «Indicadores ENDES»")
        filas.append({"indicador": "violencia_fisica_o_sexual_12m", "anio": int(anio),
                      **svy.estimar(d, g["violencia_fisica_o_sexual_12m"], escala=100), "nota": nota})
        # el repositorio no guarda el UBIGEO: se toma de REC0111 del mismo año por CASEID
        if "UBIGEO" not in g:
            ub = leer(int(anio), "REC0111")[["CASEID", "UBIGEO"]]
            g = g.assign(CASEID=g.CASEID.astype(str).str.strip()).merge(ub, on="CASEID", how="left",
                                                                         validate="one_to_one")
        assert g.UBIGEO.notna().all(), f"ENDES {anio}: mujeres sin UBIGEO"
        g["ambito"] = g.UBIGEO.map(svy.ambito_ubigeo)
        deps += [{"indicador": "violencia_fisica_o_sexual_12m", "anio": int(anio), **nivel(r)} for r in
                 svy.por_dominio(g, "ambito", config.COL_PESO, config.VAR_ESTRATO, config.VAR_CONGLOMERADO,
                                 g["violencia_fisica_o_sexual_12m"], escala=100)]
    return filas, deps


def oficiales():
    s = pd.read_excel(PSI, sheet_name="Serie anual")
    s = s[s.Código.isin([v[1] for v in INDICADORES.values() if v[1]])]
    return {(r.Código, int(r.Año)): r.Valor for _, r in s.iterrows() if pd.notna(r.Valor)}


def calcular():
    filas, deps = [], []
    for anio in ANIOS:
        df = base(anio)
        d = svy.Diseno(df, "peso", "V022", "V001")
        for ind in ("uso_anticonceptivo", "uso_moderno", "demanda_insatisfecha", "embarazo_adolescente",
                    "esposo_decide_salud"):
            filas.append({"indicador": ind, "anio": anio, **svy.estimar(d, df[ind], escala=100)})
            deps += [{"indicador": ind, "anio": anio, **nivel(r)} for r in
                     svy.por_dominio(df, "ambito", "peso", "V022", "V001", df[ind], escala=100)]
        print(f"ENDES {anio}: {len(df)} mujeres entrevistadas")
    v, vd = violencia_repositorio()
    filas += v
    deps += vd
    t = pd.DataFrame(filas)
    of = oficiales()
    t["codigo_op"] = t.indicador.map(lambda k: INDICADORES[k][1])
    t["oficial"] = [of.get((c, a), np.nan) if c else np.nan for c, a in zip(t.codigo_op, t.anio)]
    t["diferencia_oficial"] = t.estimacion - t.oficial
    t["dentro_ic95"] = np.where(t.oficial.notna(), (t.diferencia_oficial.abs() <= svy.Z95 * t.ee), None)
    t.insert(1, "etiqueta", t.indicador.map(lambda k: INDICADORES[k][0]))
    t.insert(2, "encuesta", "ENDES")
    t["referencia_diagnostico"] = t.indicador.map(lambda k: INDICADORES[k][2])
    dep = pd.DataFrame(deps)
    dep.insert(0, "encuesta", "ENDES")
    dep.insert(3, "departamento", dep.ambito.map(svy.AMBITOS))
    dep.insert(4, "unidad", "%")
    return t, dep


if __name__ == "__main__":
    os.makedirs(SALIDA, exist_ok=True)
    t, dep = calcular()
    t.to_csv(f"{SALIDA}/manifestaciones_endes.csv", index=False, encoding="utf-8")
    dep.to_csv(f"{SALIDA}/manifestaciones_endes_departamental.csv", index=False, encoding="utf-8")
    pd.set_option("display.width", 200)
    print(t[["indicador", "anio", "estimacion", "ee", "oficial", "diferencia_oficial", "dentro_ic95"]]
          .round(2).to_string(index=False))
