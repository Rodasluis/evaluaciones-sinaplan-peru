"""D1: brechas entre mujeres y hombres con la ENAHO, 2016–2025.

Recalcula, con la misma definición en todos los años, las brechas que el diagnóstico de la PNIG cita con la ENAHO
(tasa de actividad, población sin ingresos propios, informalidad, trabajo familiar no remunerado, trabajo del hogar,
jóvenes que no estudian ni trabajan, uso de internet, ingresos laborales) y otras del mismo ámbito (desempleo,
analfabetismo, afiliación a pensiones, sector público). 2016 y 2017 sirven para validar las definiciones contra las
cifras del diagnóstico; 2019–2025 es el periodo de la Política.

Traducción a Python de las sintaxis Stata de la carpeta administrativa («Sintaxis de indicadores de PN.do», sintaxis
ENAHO 2025 de la OMEP e «Indicadores OP4.do»): residentes habituales ((p204 = 1 y p205 = 2) o (p204 = 2 y p206 = 1)),
factor fac500a en el módulo 500 y factora07 en el 300, sin ingresos propios según las preguntas p501 a p50411,
ingreso laboral mensual = (i524a1 + d529t + i530a + d536 + i538a1 + d540t + i541a + d543 + d544t) / 12.
Empleo informal: desde 2024 la ENAHO no publica ocupinf (el INEI pasó la medición a la EPEN); se reconstruye con la
regla del INEI para la ENAHO 2012–2020 (Cuenta Satélite de la Economía Informal 2022–2024, anexo A2.2, pp. 87–90):
sector de los hogares por p507, p510, p510a1, p506r4 y p512b; sector informal por registro en la SUNAT o rama
extractiva; asalariados (p507 = 3, 4 o 6) informales si el centro de trabajo no aporta su seguro de salud (p419a1 a
p419a8 del módulo 400). Se usa ocupinf hasta 2023 y la regla en 2024 y 2025; validacion_informal_enaho.csv compara
ambas en 2019–2023.
Diseño muestral: estrato y conglomerado (conglome), estimador de svy.py. Por departamento (los dos primeros dígitos
del ubigeo; la ficha técnica 2025 declara «24 departamentos, cada uno como dominio de estudio», p. 10) se repite el
cálculo con el diseño del departamento (svy.por_dominio).

Fuentes de microdatos: ENAHO 2018–2023 en la carpeta local del usuario (Microdatos INEI); 2016, 2017, 2024 y 2025
descargadas con src/fuentes/descargar_microdatos_inei.py (hash en docs/fuentes/descargas_microdatos-inei_*.tsv).

Uso (desde la raíz del repositorio):
    uv run --no-project --with pandas --with pyreadstat python -I src/analisis_pnig/d1_brechas_enaho.py
Salidas: evaluaciones/pnig_2026/04_analisis/d1_brechas/brechas_enaho.csv y brechas_enaho_departamental.csv
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

# microdatos: copia en 03_datos o carpeta del equipo (variable MICRODATOS_INEI o config/entorno.local.yaml)
RAICES = [f"{EVAL}/03_datos/raw/microdatos_inei/enaho"] + [r for r in [ruta_externa("microdatos_inei")] if r]
SALIDA = f"{EVAL}/04_analisis/d1_brechas"
ANIOS = range(2016, 2026)
LLAVE = ["conglome", "vivienda", "hogar", "codperso"]
ING = ["i524a1", "d529t", "i530a", "d536", "i538a1", "d540t", "i541a", "d543", "d544t"]
P511_MON = [f"p511{i}" for i in range(1, 10)] + ["p51111"]          # pago monetario en la ocupación principal
P511_NOMON = ["p51110", "p51112"]                                    # en especie o no recibe pago
P537 = [f"p537{i}" for i in range(1, 10)] + ["p53711"]               # pago en la ocupación secundaria
P556 = [f"p556{i}a" for i in range(1, 15) if i != 6]                 # transferencias (sin Juntos, p5566a)
COLS500 = LLAVE + ["ubigeo", "estrato", "p204", "p205", "p206", "p207", "p208a", "ocu500", "fac500a", "ocupinf", "p507",
                   "p510", "p510a1", "p506r4", "p512b", "p517", "p558a5"] + ING + P511_MON + P511_NOMON + P537 + P556
P419A = [f"p419a{i}" for i in range(1, 9)]                           # quién aporta el seguro de salud (1 = centro de trabajo)
COLS400 = LLAVE + ["ubigeo"] + P419A
COLS300 = LLAVE + ["ubigeo", "estrato", "p204", "p205", "p206", "p207", "p208a", "p301a", "p302", "p306", "p307", "p314a",
                   "factora07"]

# Indicador: (etiqueta, base, unidad, cita del diagnóstico de la PNIG)
INDICADORES = {
    "tasa_actividad": ("Tasa de actividad (PEA/PET, 14 años a más)", "500", "%",
                       "M 64,0 y H 81,0 en 2017 (INEI, 2018e); p. 26 del documento de la PNIG"),
    "tasa_desempleo": ("Tasa de desempleo (desocupados/PEA, 14 años a más)", "500", "%",
                       "Mujeres urbanas 5,4 en 2017 (INEI, 2018); p. 26"),
    "sin_ingresos": ("Población de 14 años a más sin ingresos propios que no estudia (definición del INEI)",
                     "500+300", "%", "M 31,5 y H 12,0 en 2016 (INEI, 2017b); p. 26"),
    "informal": ("Empleo informal (ocupados)", "500", "%", "Mujeres 75,1 en 2016 (INEI, 2017b); p. 26"),
    "tfnr": ("Trabajadores familiares no remunerados (ocupados)", "500", "%",
             "Mujeres 17,0 (INEI, 2017b); p. 26"),
    "trabajo_hogar": ("Trabajadores del hogar (ocupados)", "500", "%", "Mujeres 5,1 (INEI, 2017b); p. 26"),
    "afiliado_pension": ("Ocupados afiliados a un sistema de pensiones", "500", "%", ""),
    "ingreso_laboral": ("Ingreso laboral mensual promedio (ocupados con ingreso, soles corrientes)", "500", "S/",
                        "Mujeres 29,3 % menos que hombres en 2017 (INEI, 2018e); paridad 80 % en 2017 (situación "
                        "futura deseada, p. 20)"),
    "ingreso_publico": ("Ingreso laboral mensual promedio en el sector público (asalariados, soles corrientes)",
                        "500", "S/", "Brecha en el servicio civil 16 % (SERVIR, 2018); p. 26"),
    "nini_15_29": ("Jóvenes de 15 a 29 años que no estudian ni están ocupados", "300+500", "%",
                   "M 24,0 y H 11,8 en 2017 (ENAHO 2017); p. 24"),
    "analfabetismo": ("Analfabetismo, 15 años a más", "300", "%", ""),
    "internet": ("Uso de internet en el mes anterior, 6 años a más", "300", "%",
                 "M 42,6 y H 48,3 en 2016 (ENAHO 2016); p. 19"),
}
# Cifras del diagnóstico para validar las definiciones: (indicador, año, sexo o «M urbana», valor)
DIAGNOSTICO = [("tasa_actividad", 2017, "Mujer", 64.0), ("tasa_actividad", 2017, "Hombre", 81.0),
               ("sin_ingresos", 2016, "Mujer", 31.5), ("sin_ingresos", 2016, "Hombre", 12.0),
               ("informal", 2016, "Mujer", 75.1), ("tfnr", 2016, "Mujer", 17.0), ("trabajo_hogar", 2016, "Mujer", 5.1),
               ("nini_15_29", 2017, "Mujer", 24.0), ("nini_15_29", 2017, "Hombre", 11.8),
               ("internet", 2016, "Mujer", 42.6), ("internet", 2016, "Hombre", 48.3)]


def ruta(anio, patron):
    for raiz in RAICES:
        hallado = glob.glob(f"{raiz}/{anio}/**/{patron}", recursive=True)
        if hallado:
            return hallado[0]
    return None


def leer(anio, modulo, cols):
    patron = f"enaho01a-{anio}-{modulo}.dta"
    p = ruta(anio, patron)
    if p is None:
        return None
    _, meta = pyreadstat.read_dta(p, metadataonly=True)
    mapa = {c.lower(): c for c in meta.column_names}
    usar = [mapa[c] for c in cols if c in mapa]
    df, _ = pyreadstat.read_dta(p, usecols=usar, apply_value_formats=False)
    df.columns = [c.lower() for c in df.columns]
    for c in cols:
        if c not in df.columns:
            df[c] = np.nan
    for c in LLAVE:
        df[c] = df[c].astype(str).str.strip()
    u = df.pop("ubigeo").map(lambda x: str(x).strip().split(".")[0].zfill(6))
    df["ccdd"] = u.str[:2]
    df["ambito"] = u.map(svy.ambito_ubigeo)  # Lima Metropolitana, Región Lima y Callao por separado
    return df


def num(s):
    return pd.to_numeric(s, errors="coerce")


def preparar_500(df):
    df = df.copy()
    for c in df.columns:
        if c not in LLAVE + ["ccdd", "ambito"]:
            df[c] = num(df[c])
    df["resi"] = ((df.p204 == 1) & (df.p205 == 2)) | ((df.p204 == 2) & (df.p206 == 1))
    df["sexo"] = np.where(df.p207 == 2, "Mujer", np.where(df.p207 == 1, "Hombre", None))
    ocu = df.ocu500
    df["pea"] = np.where(df.p208a >= 14, ocu.isin([1, 2]).astype(float), np.nan)
    df["desocupado"] = np.where(ocu.isin([1, 2]), (ocu == 2).astype(float), np.nan)
    ocupado = ocu == 1
    # Sin ingresos propios (ficha técnica del INEI, indicador 2.3): ocupados cuya ocupación principal se paga solo en
    # especie o no se paga y cuya ocupación secundaria, si la tienen, tampoco se paga; desocupados e inactivos sin
    # transferencias monetarias (p5561a a p55614a, salvo Juntos). Luego se excluye a quien estudia (en preparar_union).
    mon_p = (df[P511_MON] == 1).any(axis=1)
    sin_p = (df[P511_NOMON] == 1).any(axis=1) & ~mon_p
    tiene_sec = df.p517.isin(range(1, 8))
    sin_s = ~(df[P537] == 1).any(axis=1)
    ocup_sin = ocupado & df.p507.isin(range(1, 8)) & sin_p & (~tiene_sec | sin_s)
    p556 = [c for c in P556 if df[c].notna().any()]  # p55610a a p55614a no existen antes de 2019
    no_ocup_sin = ocu.isin([2, 3, 4]) & (df[p556] == 2).all(axis=1)
    df["sin_ingresos_base"] = np.where(df.p208a >= 14, (ocup_sin | no_ocup_sin).astype(float), np.nan)
    con_400 = df[P419A].notna().any().any()
    df["informal_regla"] = np.where(ocupado & df.p507.notna() & con_400, informal_inei(df), np.nan)
    oficial = np.where(ocupado & df.ocupinf.isin([1, 2]), (df.ocupinf == 1).astype(float), np.nan)
    df["informal"] = oficial if df.ocupinf.notna().any() else df.informal_regla
    df["tfnr"] = np.where(ocupado & df.p507.notna(), (df.p507 == 5).astype(float), np.nan)
    df["trabajo_hogar"] = np.where(ocupado & df.p507.notna(), (df.p507 == 6).astype(float), np.nan)
    df["afiliado_pension"] = np.where(ocupado & df.p558a5.notna(), (df.p558a5 != 5).astype(float), np.nan)
    ing = df[ING].sum(axis=1, min_count=1) / 12
    df["ingreso_laboral"] = np.where(ocupado & (ing > 0), ing, np.nan)
    publico = ocupado & df.p510.isin([1, 2, 3]) & df.p507.isin([3, 4])
    df["ingreso_publico"] = np.where(publico & (ing > 0), ing, np.nan)
    return df


def informal_inei(df):
    """Empleo informal en la ocupación principal según el INEI (Cuenta Satélite 2022–2024, figuras A2.2.2b a A2.2.5).

    Paso 1: trabajan en sociedades los asalariados del sector público, cooperativas y SERVICE (p510 = 1 a 5) y, en
    unidades privadas, quienes declaran persona jurídica (p510a1 = 1) o, sin serlo y fuera del agro, más de 30
    trabajadores (cuasisociedad). Paso 2: dentro de los hogares productores de mercado, es sector informal la unidad
    no registrada en la SUNAT (p510a1 = 3) o de rama extractiva (CIIU rev. 4, divisiones 01 a 09). Pasos 3 y 4: es
    informal todo el sector informal, el trabajador familiar no remunerado y «otro», y el asalariado (incluido el
    trabajador del hogar) cuyo seguro de salud no aporta su centro de trabajo.
    """
    division = df.p506r4 // 100
    agro = division.between(1, 3)
    extractiva = division.between(1, 9)
    publico = df.p507.isin([3, 4]) & df.p510.isin([1, 2, 3, 4, 5])
    sociedad = publico | (df.p507.isin([1, 2, 3, 4]) &
                          ((df.p510a1 == 1) | (df.p510a1.isin([2, 3]) & ~agro & (df.p512b > 30))))
    mercado = df.p507.isin([5, 7]) | (df.p507.isin([1, 2, 3, 4]) & ~sociedad)
    sector_informal = mercado & (extractiva | (df.p510a1 == 3))
    seguro_empleador = (df[P419A] == 1).any(axis=1)
    return (sector_informal | df.p507.isin([5, 7]) | (df.p507.isin([3, 4, 6]) & ~seguro_empleador)).astype(float)


def validar_informal(df, anio):
    """Compara la regla con ocupinf (residentes ocupados, ponderado): tasas por sexo y concordancia individual."""
    d = df[df.informal_regla.notna() & df.ocupinf.isin([1, 2])]
    if d.empty:
        return []
    o = (d.ocupinf == 1).astype(float)
    filas = []
    for sexo in ("Mujer", "Hombre"):
        m = d.sexo == sexo
        filas.append({"anio": anio, "sexo": sexo, "ocupinf": np.average(o[m], weights=d.fac500a[m]) * 100,
                      "regla_inei": np.average(d.informal_regla[m], weights=d.fac500a[m]) * 100})
    conc = np.average(d.informal_regla == o, weights=d.fac500a) * 100
    for f in filas:
        f.update(diferencia=f["regla_inei"] - f["ocupinf"], concordancia=conc,
                 discrepancias=int((d.informal_regla != o).sum()), casos=len(d))
    return filas


def preparar_300(df):
    df = df.copy()
    for c in df.columns:
        if c not in LLAVE + ["ccdd", "ambito"]:
            df[c] = num(df[c])
    df["resi"] = ((df.p204 == 1) & (df.p205 == 2)) | ((df.p204 == 2) & (df.p206 == 1))
    df["sexo"] = np.where(df.p207 == 2, "Mujer", np.where(df.p207 == 1, "Hombre", None))
    # p302 solo se pregunta a quien no aprobó más que primaria: el resto sabe leer y escribir
    df["analfabetismo"] = np.where((df.p208a >= 15) & (df.p302.notna() | df.p301a.notna()),
                                   (df.p302 == 2).astype(float), np.nan)
    # p307 (asiste) solo se pregunta a quien está matriculado (p306 = 1)
    df["estudia"] = np.where(df.p306.isin([1, 2]), ((df.p306 == 1) & (df.p307 == 1)).astype(float), np.nan)
    df["internet"] = np.where((df.p208a >= 6) & df.p314a.isin([1, 2]), (df.p314a == 1).astype(float), np.nan)
    return df


def por_departamento(df, ind, clave, anio, peso, escala):
    return [{"indicador": clave, "anio": anio, **r}
            for r in svy.por_dominio(df, "ambito", peso, "estrato", "conglome", df[ind], "sexo", escala=escala)]


def calcular():
    filas, deps = [], []
    for anio in ANIOS:
        m5 = leer(anio, "500", COLS500)
        m3 = leer(anio, "300", COLS300)
        if m5 is None or m3 is None:
            print(f"ENAHO {anio}: falta el módulo 500 o 300")
            continue
        m4 = leer(anio, "400", COLS400)  # 2016 y 2017 sin módulo 400: usan ocupinf
        if m4 is not None:
            m5 = m5.merge(m4[LLAVE + P419A], on=LLAVE, how="left")
        else:
            m5 = m5.assign(**{c: np.nan for c in P419A})
        d3 = preparar_300(m3)
        d3 = d3[d3.resi & d3.sexo.notna()].reset_index(drop=True)
        d5 = preparar_500(m5)
        d5 = d5[d5.resi & d5.sexo.notna()].reset_index(drop=True)
        VALIDACION_INFORMAL.extend(validar_informal(d5, anio))
        d5 = d5.merge(d3[LLAVE + ["estudia"]], on=LLAVE, how="left")
        d5["sin_ingresos"] = np.where(d5.sin_ingresos_base.notna(),
                                      ((d5.sin_ingresos_base == 1) & (d5.estudia != 1)).astype(float), np.nan)
        D5 = svy.Diseno(d5, "fac500a", "estrato", "conglome")
        for ind in ("pea", "desocupado", "sin_ingresos", "informal", "tfnr", "trabajo_hogar", "afiliado_pension",
                    "ingreso_laboral", "ingreso_publico"):
            clave = {"pea": "tasa_actividad", "desocupado": "tasa_desempleo"}.get(ind, ind)
            if d5[ind].notna().sum() == 0:
                filas.append({"indicador": clave, "anio": anio, "nota": "Variable no publicada en la base del año"})
                continue
            escala = 1 if ind.startswith("ingreso") else 100
            r = svy.brecha(D5, d5[ind], d5.sexo, "Mujer", "Hombre", escala=escala)
            filas.append({"indicador": clave, "anio": anio, **r})
            deps += por_departamento(d5, ind, clave, anio, "fac500a", escala)
        D3 = svy.Diseno(d3, "factora07", "estrato", "conglome")
        for ind in ("analfabetismo", "internet"):
            filas.append({"indicador": ind, "anio": anio, **svy.brecha(D3, d3[ind], d3.sexo, "Mujer", "Hombre",
                                                                         escala=100)})
            deps += por_departamento(d3, ind, ind, anio, "factora07", 100)
        # Jóvenes que no estudian ni trabajan: asistencia (300) y ocupación (500) de la misma persona
        j = d5[(d5.p208a >= 15) & (d5.p208a <= 29)].reset_index(drop=True)
        j["nini"] = np.where(j.estudia.notna() & j.ocu500.notna(),
                             ((j.estudia == 0) & (j.ocu500 != 1)).astype(float), np.nan)
        DJ = svy.Diseno(j, "fac500a", "estrato", "conglome")
        filas.append({"indicador": "nini_15_29", "anio": anio, **svy.brecha(DJ, j.nini, j.sexo, "Mujer", "Hombre",
                                                                          escala=100)})
        deps += por_departamento(j, "nini", "nini_15_29", anio, "fac500a", 100)
        print(f"ENAHO {anio}: {len(d5)} residentes en el módulo 500, {len(d3)} en el 300")
    t = pd.DataFrame(filas)
    t.insert(1, "etiqueta", t.indicador.map(lambda k: INDICADORES[k][0]))
    t.insert(2, "unidad", t.indicador.map(lambda k: INDICADORES[k][2]))
    t.insert(3, "encuesta", "ENAHO")
    t["nota"] = t.get("nota", pd.Series(index=t.index, dtype=object))
    t.loc[t.anio == 2020, "nota"] = t.loc[t.anio == 2020, "nota"].fillna(
        "2020: entrevista mayormente telefónica por la emergencia sanitaria; comparar con cautela")
    sin_ocupinf = (t.indicador == "informal") & (t.anio >= 2024)
    t.loc[sin_ocupinf, "nota"] = t.loc[sin_ocupinf, "nota"].fillna(
        "Reconstruido con la regla del INEI (Cuenta Satélite 2022–2024, pp. 87–90): la base no publica ocupinf")
    t.loc[t.indicador == "nini_15_29", "nota"] = t.loc[t.indicador == "nini_15_29", "nota"].fillna(
        "Definición estándar (no matriculado o sin asistir, y no ocupado); no reproduce la cifra del diagnóstico")
    dep = pd.DataFrame(deps)
    dep.insert(0, "encuesta", "ENAHO")
    dep.insert(3, "departamento", dep.ambito.map(svy.AMBITOS))
    dep.insert(4, "unidad", dep.indicador.map(lambda k: INDICADORES[k][2]))
    return t, dep


REFERENCIAS = f"{SALIDA}/referencias_oficiales.csv"
VALIDACION_INFORMAL = []


def validar(t):
    filas = []
    if os.path.exists(REFERENCIAS):
        ref = pd.read_csv(REFERENCIAS)
        for _, r in ref[ref.encuesta == "ENAHO"].iterrows():
            DIAGNOSTICO.append((r.indicador, int(r.anio), r.sexo, float(r.valor), "Serie oficial del INEI"))
    for item in DIAGNOSTICO:
        ind, anio, sexo, ref = item[:4]
        origen = item[4] if len(item) > 4 else "Diagnóstico de la PNIG"
        r = t[(t.indicador == ind) & (t.anio == anio)]
        base = {"indicador": ind, "anio": anio, "sexo": sexo, "referencia": ref, "origen": origen}
        if r.empty or pd.isna(r.iloc[0].get("valor_a")):
            filas.append({**base, "calculado": np.nan})
            continue
        v = r.iloc[0]["valor_a" if sexo == "Mujer" else "valor_b"]
        ee = r.iloc[0]["ee_a" if sexo == "Mujer" else "ee_b"]
        filas.append({**base, "calculado": round(v, 2), "diferencia": round(v - ref, 2),
                      "dentro_ic95": abs(v - ref) <= svy.Z95 * ee})
    return pd.DataFrame(filas)


if __name__ == "__main__":
    os.makedirs(SALIDA, exist_ok=True)
    t, dep = calcular()
    t.to_csv(f"{SALIDA}/brechas_enaho.csv", index=False, encoding="utf-8")
    dep.to_csv(f"{SALIDA}/brechas_enaho_departamental.csv", index=False, encoding="utf-8")
    vi = pd.DataFrame(VALIDACION_INFORMAL)
    vi.round(3).to_csv(f"{SALIDA}/validacion_informal_enaho.csv", index=False, encoding="utf-8")
    assert (vi.concordancia >= 98).all() and (vi.diferencia.abs() <= 0.3).all(), vi
    print(vi.round(2).to_string(index=False))
    v = validar(t)
    v.to_csv(f"{SALIDA}/validacion_enaho_diagnostico.csv", index=False, encoding="utf-8")
    pd.set_option("display.width", 220)
    print(v.to_string(index=False))
    print(t[t.anio.isin([2019, 2025])][["indicador", "anio", "valor_a", "valor_b", "diferencia", "ee_diferencia",
                                        "razon"]].round(2).to_string(index=False))
