"""D1 y D2: brechas por ámbito (departamentos, Lima Metropolitana, Región Lima y Callao), series 2019–2025 y mapas.

Lee las estimaciones por ámbito de los scripts por encuesta (brechas_enaho_departamental.csv,
brechas_epen_departamental.csv, brechas_ena_departamental.csv, manifestaciones_endes_departamental.csv y
tolerancia_enares_departamental.csv), la participación política por ámbito (d2_participacion_departamental.py) y las
series nacionales, y produce:
  - evaluaciones/pnig_2026/04_analisis/D1_D2_brechas_departamentales.xlsx, con las hojas «Resumen por indicador»,
    «Brechas por ámbito», «Cambio por ámbito», «Participación política», «Validación con el INEI», «Diseño muestral» y
    «Criterios»;
  - evaluaciones/pnig_2026/04_analisis/d1_brechas/figuras/*.png: series nacionales 2019–2025 por sexo con su intervalo
    de confianza, un mapa por indicador (2019, 2025 y cambio) y mapas de participación política, con un índice
    (indice_figuras.csv) que lee informe_cierre.py; la nota de cada figura dice cuántos ámbitos tienen una estimación
    confiable y en qué año cambió el instrumento.

Reglas (decisiones del usuario del 2026-11-02; el umbral del coeficiente de variación es el del INEI):
  - se mapean todos los indicadores que se pueden estimar; una estimación es confiable si el coeficiente de variación
    de mujeres y de hombres es de 15 % o menos (INEI, Brechas de Género 2025, p. 178); si es mayor, es referencial: el
    mapa la marca con trama y la nota lista los ámbitos y su coeficiente; con menos de 25 casos sin ponderar en un sexo
    no se estima;
  - la brecha de un año es significativa si difiere de la paridad al 95 %;
  - el cambio compara 2019 con 2025 (o el año más cercano con datos) con la misma regla que la clasificación nacional
    (svy.cambio); con 26 pruebas al 5 % se esperan alrededor de 1,3 cambios significativos por azar;
  - el empleo informal se compara con la ENAHO 2019 y 2025: en 2024 y 2025 se reconstruye con la regla del INEI, que
    reproduce la variable ocupinf de 2019 a 2023 (d1_brechas_enaho.py); la EPEN queda como contraste nacional;
  - las brechas que solo mide la EPEN desde 2023 (actividad, ingreso, pensiones) se mapean con la ENAHO 2019–2025.
Ámbitos: 23 departamentos, Callao, Lima Metropolitana (provincia de Lima) y Región Lima, como los publica el INEI. La
EPEN solo distingue «Lima Metropolitana y Callao» y Región Lima; la ENA publica Lima entera.

Cartografía: Peru maps (github.com/Rodasluis/Peru-maps), capa provincial simplificada copiada a docs/fuentes/cartografia/
con su hash; la provincia de Lima forma Lima Metropolitana y las otras nueve provincias, Región Lima. Uso estadístico.

Uso (desde la raíz del repositorio, después de los scripts por encuesta, d1_d2_brechas.py y
d2_participacion_departamental.py):
    uv run --no-project --with pandas --with openpyxl --with geopandas --with matplotlib python -I -X utf8 \\
        src/analisis_pnig/d1_d2_departamental.py
"""
import os
import re
import subprocess
import sys
import unicodedata

import geopandas as gpd
import matplotlib
import numpy as np
import pandas as pd

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib import ticker  # noqa: E402
from matplotlib.lines import Line2D  # noqa: E402
from matplotlib.patches import Patch  # noqa: E402
from openpyxl import load_workbook  # noqa: E402
from openpyxl.styles import Alignment, Font, PatternFill  # noqa: E402

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import svy  # noqa: E402
from d2_autoridades_2026 import NOTA_2026  # noqa: E402
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

D = f"{EVAL}/04_analisis/d1_brechas"
FIG = f"{D}/figuras"
SALIDA = f"{EVAL}/04_analisis/D1_D2_brechas_departamentales.xlsx"
NACIONAL = f"{EVAL}/04_analisis/D1_D2_brechas_situacion_futura.xlsx"
GEO = "docs/fuentes/cartografia/2026_peru-maps_provincia-simplificado.geojson"
INEI_BRECHAS = "docs/fuentes/pnig/04_seguimiento/fuentes_indicadores/2025_inei_brechas-de-genero-2025.pdf"
MET = "docs/fuentes/inei_metodologia"
CV_MAX, N_MIN, INICIO, FIN = 15.0, 25, 2019, 2025

# Paleta: divergente azul y rojo con gris neutro, secuencial azul (skill dataviz; validada con validate_palette.js)
AZUL = ["#86b6ef", "#3987e5", "#1c5cab"]
ROJO = ["#ee8c84", "#e34948", "#a3262a"]
SECUENCIAL = ["#cde2fb", "#86b6ef", "#3987e5", "#1c5cab", "#0d366b"]
NEUTRO, SIN_DATO = "#dcdbd5", "#ffffff"
TINTA, TINTA2, BORDE = "#0b0b0b", "#52514e", "#898781"
MUJERES, HOMBRES = "#eb6834", "#2a78d6"   # ranuras categóricas 2 y 1
plt.rcParams.update({"font.family": "Arial", "font.size": 7, "axes.edgecolor": BORDE, "axes.labelcolor": TINTA2,
                     "xtick.color": TINTA2, "ytick.color": TINTA2, "hatch.linewidth": 0.5, "hatch.color": TINTA2,
                     "axes.spines.top": False, "axes.spines.right": False})

# (encuesta, indicador): (nombre corto, tipo, signo o sentido)
#   tipo «brecha»: mujeres menos hombres en puntos porcentuales; «razon»: mujeres/hombres; «nivel»: solo mujeres.
#   signo: +1 si un valor positivo de la brecha es desventaja de las mujeres, −1 si lo es uno negativo, 0 si la
#   brecha no expresa desventaja. sentido (nivel): «desc» si bajar es mejorar, «asc» si subir es mejorar.
META = {
    ("ENAHO", "tasa_actividad"): ("Tasa de actividad", "brecha", -1),
    ("ENAHO", "tasa_desempleo"): ("Tasa de desempleo", "brecha", 1),
    ("ENAHO", "sin_ingresos"): ("Población sin ingresos propios", "brecha", 1),
    ("ENAHO", "informal"): ("Empleo informal", "brecha", 1),
    ("ENAHO", "tfnr"): ("Trabajo familiar no remunerado", "brecha", 1),
    ("ENAHO", "trabajo_hogar"): ("Trabajo del hogar remunerado", "brecha", 1),
    ("ENAHO", "afiliado_pension"): ("Afiliación a un sistema de pensiones", "brecha", -1),
    ("ENAHO", "ingreso_laboral"): ("Ingreso laboral mensual", "razon", -1),
    ("ENAHO", "ingreso_publico"): ("Ingreso laboral en el sector público", "razon", -1),
    ("ENAHO", "nini_15_29"): ("Jóvenes que no estudian ni están ocupados", "brecha", 1),
    ("ENAHO", "analfabetismo"): ("Analfabetismo", "brecha", 1),
    ("ENAHO", "internet"): ("Uso de internet", "brecha", -1),
    ("EPEN", "informal_epen"): ("Empleo informal (EPEN)", "brecha", 1),
    ("EPEN", "tasa_actividad_epen"): ("Tasa de actividad (EPEN)", "brecha", -1),
    ("EPEN", "ingreso_laboral_epen"): ("Ingreso laboral mensual (EPEN)", "razon", -1),
    ("EPEN", "afiliado_pension_epen"): ("Afiliación a pensiones (EPEN)", "brecha", -1),
    ("ENA", "productora_mujer"): ("Mujeres entre quienes conducen una unidad agropecuaria", "brecha", -1),
    ("ENA", "credito_solicito"): ("Productores que solicitaron crédito", "brecha", -1),
    ("ENA", "propietario"): ("Productores propietarios de alguna parcela", "brecha", -1),
    ("ENA", "titulo"): ("Propietarios con título de alguna parcela", "brecha", -1),
    ("ENARES", "tolerancia_10"): ("Tolerancia social a la violencia contra las mujeres", "brecha", 0),
    ("ENDES", "uso_anticonceptivo"): ("Uso de algún método anticonceptivo (unidas)", "nivel", "asc"),
    ("ENDES", "uso_moderno"): ("Uso de un método moderno (unidas)", "nivel", "asc"),
    ("ENDES", "demanda_insatisfecha"): ("Demanda insatisfecha de planificación familiar", "nivel", "desc"),
    ("ENDES", "embarazo_adolescente"): ("Adolescentes de 15 a 19 años alguna vez embarazadas", "nivel", "desc"),
    ("ENDES", "esposo_decide_salud"): ("El esposo decide solo sobre la salud de la mujer", "nivel", "desc"),
    ("ENDES", "violencia_fisica_o_sexual_12m"): ("Violencia física o sexual de pareja, últimos 12 meses", "nivel",
                                                  "desc"),
}
# no se mapean: los indicadores de la EPEN, que la ENAHO mide de 2019 a 2025, se mapean con la ENAHO (decisión del
# usuario, 2026-11-02; el empleo informal de la ENAHO 2024 y 2025 se reconstruye con la regla del INEI)
NO_MAPA = {("EPEN", "informal_epen"), ("EPEN", "tasa_actividad_epen"),
           ("EPEN", "ingreso_laboral_epen"), ("EPEN", "afiliado_pension_epen")}
# cambios de instrumento y años sin dato, para la nota de cada mapa y de cada serie
NOTA_INSTRUMENTO = {
    "informal": "La ENAHO no publica la variable de empleo informal desde 2024; 2025 se reconstruye con la regla del "
                "INEI (Cuenta Satélite de la Economía Informal 2022–2024, pp. 87–90), que reproduce la variable de "
                "2019 a 2023 con una concordancia de 99,97 % o más.",
    "propietario": "En 2023 la ENA cambió el capítulo de parcelas (de las trabajadas en los últimos 12 meses a las que se "
                   "trabajan al momento de la entrevista), por lo que la comparación 2019–2025 es indicativa. No hubo "
                   "ENA en 2020.",
    "titulo": "En 2023 la ENA cambió el capítulo de parcelas (de las trabajadas en los últimos 12 meses a las que se "
              "trabajan al momento de la entrevista): el nivel baja de 35 a 38 % en 2019–2022 a 26 a 28 % desde 2023, "
              "por lo que la comparación 2019–2025 es indicativa. No hubo ENA en 2020.",
    "credito_solicito": "La ENA 2021 no publicó el capítulo de crédito y no hubo ENA en 2020.",
    "productora_mujer": "No hubo ENA en 2020.",
    "violencia_fisica_o_sexual_12m": "2019 y 2020 con los módulos del INEI y las funciones del repositorio «Indicadores "
                                     "ENDES»; el valor de 2020 no reproduce el oficial (7,9 frente a 8,8 %).",
    "tolerancia_10": "La ENARES 2019 solo tiene inferencia nacional; 2024 es el único año con inferencia por "
                     "departamento.",
}
ARCHIVOS = ["brechas_enaho_departamental", "brechas_epen_departamental", "brechas_ena_departamental",
            "manifestaciones_endes_departamental", "tolerancia_enares_departamental"]
# Nivel de inferencia de cada encuesta: (encuesta, inferencia por ámbito, cita literal, documento, página del PDF)
DISENO = [
    ("ENAHO", "Sí", "24 departamentos, cada uno como dominio de estudio", "2025_inei_enaho_ficha-tecnica.pdf", 10),
    ("ENDES", "Sí", "Cada uno de los 24 departamentos del país y la Provincia Constitucional del Callao",
     "2025_inei_endes_ficha-tecnica.pdf", 8),
    ("EPEN", "Sí (Lima Metropolitana con el Callao y Región Lima)", "La muestra está diseñada para dar resultados "
     "confiables a nivel", "2025_inei_epen_mercado-laboral-departamental-2022-2024.pdf", 104),
    ("ENA", "Sí (productores, marco de áreas; Lima entera)", "tendrá los siguientes niveles de inferencia",
     "2024_inei_ena_ficha-tecnica.pdf", 7),
    ("ENARES", "Solo en 2024 (2013, 2015 y 2019: nacional)", "Nacional y departamental, donde el departamento de Lima "
     "se divide en tres ámbitos", "2025_inei_enares-2024_principales-resultados.pdf", 12),
    ("ENUT", "No (nacional, urbano y rural, resto costa, sierra, selva y Lima Metropolitana)",
     "La muestra está diseñada para dar resultados confiables a nivel de", "2025_inei_enut-2024_principales-"
     "resultados.pdf", 19),
]
# Gráficos por ámbito de 2024 del INEI para validar: (página, encuesta, indicador, bloques de valores en orden).
# Los valores aparecen en el orden de las barras; el bloque ordenado es el de las mujeres (o el de 2024) y las
# etiquetas siguen ese orden (comprobado con el texto de cada página y con las coordenadas de las etiquetas).
VALIDA = [(77, "ENAHO", "analfabetismo", ("Hombre", "Mujer")),
          (83, "EPEN", "tasa_actividad_epen", ("Hombre", "Mujer")),
          (105, "ENAHO", "sin_ingresos", ("Hombre", "Mujer")),
          (114, "ENA", "productora_mujer", ("2014", "Mujer")),
          (120, "ENA", "credito_solicito", ("Hombre", "Mujer")),
          (140, "ENAHO", "internet", ("Mujer", "Hombre"))]
ETIQUETAS_INEI = {"lima metropolitana": "LM", "region lima": "RL", "lima metrop./prov. const. callao": "LMC",
                  "prov. const. del callao": "07", "lima": "15", "nacional": "nacional"}


def fmt(x, d=1):
    return f"{x:,.{d}f}".replace(",", " ").replace(".", ",")


def texto_pagina(pdf, pag, raw=False):
    args = ["pdftotext", "-enc", "UTF-8"] + (["-raw"] if raw else []) + ["-f", str(pag), "-l", str(pag), pdf, "-"]
    return subprocess.run(args, capture_output=True, check=True).stdout.decode("utf-8")


def norm(x):
    return unicodedata.normalize("NFKD", str(x)).encode("ascii", "ignore").decode().strip().lower()


# ---------------------------------------------------------------------------------------------- datos
def leer():
    t = pd.concat([pd.read_csv(f"{D}/{a}.csv", dtype={"ambito": str}) for a in ARCHIVOS], ignore_index=True)
    t = t[[k in META for k in zip(t.encuesta, t.indicador)]].copy()
    t["nombre"] = [META[k][0] for k in zip(t.encuesta, t.indicador)]
    t["tipo"] = [META[k][1] for k in zip(t.encuesta, t.indicador)]
    t["departamento"] = t.ambito.map(svy.AMBITOS)
    b, r, n = t.tipo == "brecha", t.tipo == "razon", t.tipo == "nivel"
    t["valor"] = np.select([b, r, n], [t.diferencia, t.razon, t.valor_a])
    t["ee"] = np.select([b, r, n], [t.ee_diferencia, t.ee_razon, t.ee_a])
    t["paridad"] = np.select([b, r], [0.0, 1.0], np.nan)
    dos = t.tipo != "nivel"
    n_b = t.n_b.where(dos, N_MIN)
    t["estimable"] = (t.n_a >= N_MIN) & (n_b >= N_MIN) & t.valor.notna()
    t["cv"] = np.where(dos & t.cv_b.notna(), np.fmax(t.cv_a, t.cv_b), t.cv_a)
    t["confiable"] = t.estimable & (t.cv <= CV_MAX)
    t["condicion"] = np.where(~t.estimable, f"Sin estimación (menos de {N_MIN} casos)",
                              np.where(t.confiable, "Confiable", f"Referencial (CV > {fmt(CV_MAX, 0)} %)"))
    t["significativa"] = np.where(dos & t.estimable, (t.valor - t.paridad).abs() > svy.Z95 * t.ee, np.nan)
    return t


def periodos(g):
    """2019 y 2025, o los años más cercanos con datos del indicador."""
    anios = sorted(g.anio.unique())
    inicio = INICIO if INICIO in anios else min(a for a in anios if a >= INICIO) if max(anios) >= INICIO else anios[0]
    return inicio, max(anios)


def cambios(t):
    filas = []
    for (enc, ind), g in t.groupby(["encuesta", "indicador"], sort=False):
        inicio, fin = periodos(g)
        if inicio == fin:
            continue
        _, tipo, signo = META[(enc, ind)]
        a, b = g[g.anio == inicio].set_index("ambito"), g[g.anio == fin].set_index("ambito")
        for cc in sorted(set(a.index) & set(b.index)):
            base = {"encuesta": enc, "indicador": ind, "nombre": META[(enc, ind)][0], "ambito": cc,
                    "departamento": svy.AMBITOS.get(cc), "inicio": inicio, "fin": fin}
            if not (a.loc[cc, "estimable"] and b.loc[cc, "estimable"]):
                filas.append({**base, "clase": "Sin estimación"})
                continue
            v0, e0, v1, e1 = a.loc[cc, "valor"], a.loc[cc, "ee"], b.loc[cc, "valor"], b.loc[cc, "ee"]
            dif, ee, sig = svy.cambio(v0, e0, v1, e1)
            if tipo == "nivel":
                mejora = dif < 0 if signo == "desc" else dif > 0
                clase = ("Mejora" if mejora else "Empeora") if sig else "Sin cambio significativo"
            else:
                p = a.loc[cc, "paridad"]
                clase = ("Se reduce" if abs(v1 - p) < abs(v0 - p) else "Se amplía") if sig else \
                    "Sin cambio significativo"
            filas.append({**base, "valor_inicio": v0, "valor_fin": v1, "cambio": dif, "ee_cambio": ee,
                          "significativo": sig, "clase": clase,
                          "confiable_ambos": bool(a.loc[cc, "confiable"] and b.loc[cc, "confiable"])})
    return pd.DataFrame(filas)


def nacional_clases():
    try:
        e = pd.read_excel(NACIONAL, sheet_name=0)
        return dict(zip(e.Clave, e["Clasificación"]))
    except Exception:  # noqa: BLE001 el libro nacional se genera antes; si falta, la columna queda vacía
        return {}


def ambitos_mapa(u):
    """Ámbitos que se pintan: si hay Lima Metropolitana y Callao por separado, no se usa el ámbito conjunto."""
    if {"LM", "07"} <= set(u.index):
        u = u.drop(index="LMC", errors="ignore")
    return u


def nota_confiabilidad(g, anios):
    """Cuántos ámbitos tienen estimación confiable en cada año y cuáles no (con su coeficiente de variación)."""
    partes = []
    for a in anios:
        u = ambitos_mapa(g[g.anio == a].set_index("ambito"))
        est, conf = u[u.estimable], u[u.confiable]
        ref = est[~est.confiable].sort_values("cv")
        txt = (f"{a}: confiable (coeficiente de variación de 15 % o menos) en {len(conf)} de los {len(est)} "
               f"departamentos con estimación")
        if len(ref) and len(ref) <= 8:
            txt += "; referenciales: " + ", ".join(f"{svy.AMBITOS[k]} ({fmt(r.cv, 1)} %)" for k, r in ref.iterrows())
        elif len(ref):
            txt += (f"; los otros {len(ref)} son referenciales (coeficiente de variación de {fmt(ref.cv.min(), 1)} a "
                    f"{fmt(ref.cv.max(), 1)} %)")
            if len(conf):
                txt += "; confiables: " + ", ".join(f"{svy.AMBITOS[k]} ({fmt(r.cv, 1)} %)"
                                                    for k, r in conf.sort_values("cv").iterrows())
        sin = u[~u.estimable]
        if len(sin):
            txt += "; sin estimación: " + ", ".join(svy.AMBITOS[k] for k in sin.index)
        partes.append(txt)
    return ". ".join(partes) + "."


def resumen(t, c):
    nac = nacional_clases()
    filas = []
    for (enc, ind), g in t.groupby(["encuesta", "indicador"], sort=False):
        nombre, tipo, signo = META[(enc, ind)]
        inicio, fin = periodos(g)
        u = ambitos_mapa(g[g.anio == fin].set_index("ambito")).reset_index()
        k = c[(c.encuesta == enc) & (c.indicador == ind)] if not c.empty else pd.DataFrame()
        f = {"Encuesta": enc, "Indicador": nombre, "Clave": ind, "Tipo": {"brecha": "Brecha M − H (pp)",
             "razon": "Razón M/H", "nivel": "Nivel en mujeres (%)"}[tipo],
             "Años comparados": f"{inicio}–{fin}" if inicio != fin else str(fin),
             "Ámbitos con estimación (último año)": int(u.estimable.sum()),
             "Confiables (último año)": int(u.confiable.sum()),
             "CV mínimo y máximo (último año, %)": f"{fmt(u.cv.min(), 0)} a {fmt(u.cv.max(), 0)}"}
        if tipo != "nivel":
            s = u[u.estimable & (u.significativa == 1)]
            lado = np.sign(s.valor - s.paridad) * (signo if signo else 1)
            f["Brecha significativa en desventaja de las mujeres" if signo else "Brecha significativa, mayor en "
              "mujeres"] = int((lado > 0).sum())
            f["Brecha significativa en desventaja de los hombres" if signo else "Brecha significativa, mayor en "
              "hombres"] = int((lado < 0).sum())
            f["Brecha no significativa"] = int((u.estimable & (u.significativa == 0)).sum())
        if not k.empty:
            for cl in (["Se reduce", "Sin cambio significativo", "Se amplía"] if tipo != "nivel" else
                       ["Mejora", "Sin cambio significativo", "Empeora"]):
                f[f"Cambio: {cl.lower()}"] = int((k.clase == cl).sum())
            f["Cambios significativos"] = int(k.clase.isin(["Se reduce", "Se amplía", "Mejora", "Empeora"]).sum())
            f["Esperados por azar (5 %)"] = round(0.05 * int((k.clase != "Sin estimación").sum()), 2)
        f["Clasificación nacional"] = nac.get(ind, "")
        f["Mapa"] = "No (se mapea con la ENAHO)" if (enc, ind) in NO_MAPA else "Sí"
        filas.append(f)
    return pd.DataFrame(filas)


def validar(t):
    filas = []
    for pag, enc, ind, bloques in VALIDA:
        lineas = [l.strip() for l in texto_pagina(INEI_BRECHAS, pag, raw=True).splitlines()]
        vals, labs = [], []
        for l in lineas:
            if re.fullmatch(r"\d+,\d", l):
                vals.append(float(l.replace(",", ".")))
                continue
            k = norm(re.sub(r"\s+\d/$", "", l))
            if k in ETIQUETAS_INEI:
                labs.append(ETIQUETAS_INEI[k])
            else:
                try:
                    labs.append(svy.codigo_departamento(k))
                except KeyError:
                    pass
        labs = [x for x in labs if x != "nacional"]
        assert len(vals) == len(bloques) * len(labs), (pag, len(vals), len(labs))
        n = len(labs)
        g = t[(t.encuesta == enc) & (t.indicador == ind) & (t.anio == 2024)].set_index("ambito")
        for i, bloque in enumerate(bloques):
            if bloque not in ("Mujer", "Hombre"):
                continue
            for cc, ref in zip(labs, vals[i * n:(i + 1) * n]):
                if cc not in g.index:
                    continue
                r = g.loc[cc]
                v, e = (r.valor_a, r.ee_a) if bloque == "Mujer" else (r.valor_b, r.ee_b)
                filas.append({"Encuesta": enc, "Indicador": META[(enc, ind)][0], "Año": 2024,
                              "Ámbito": svy.AMBITOS[cc], "Sexo": bloque, "INEI": ref,
                              "Calculado": round(v, 2), "Diferencia": round(v - ref, 2), "EE": round(e, 2),
                              "Dentro del IC 95 %": "Sí" if abs(v - ref) <= svy.Z95 * e + 0.05 else "No",
                              "Fuente": f"INEI, Brechas de Género 2025, p. {pag}"})
    v = pd.DataFrame(filas)
    # emparejamiento comprobado con las coordenadas de cada etiqueta en la página (pdfplumber): el desajuste no viene
    # de la lectura del gráfico
    nota = ("El valor nacional 2024 coincide (INEI 78,1 y 80,8; cálculo 78,0 y 80,7) y las cifras de mujeres "
            "coinciden; con la misma definición para ambos sexos, la diferencia en estos hombres no se explica")
    v["Nota"] = np.where((v.Indicador == "Uso de internet") & (v.Sexo == "Hombre") & (v["Dentro del IC 95 %"] == "No"),
                         nota, "")
    return v


def diseno():
    filas = []
    for enc, inf, cita, doc, pag in DISENO:
        ok = re.sub(r"\s+", " ", cita) in re.sub(r"\s+", " ", texto_pagina(f"{MET}/{doc}", pag))
        assert ok, f"{enc}: la cita no consta en la p. {pag} de {doc}"
        filas.append({"Encuesta": enc, "Inferencia por ámbito": inf, "Cita": f"«{cita}»",
                      "Documento": f"docs/fuentes/inei_metodologia/{doc}", "Página del PDF": pag})
    cita = "variación mayor al 15%"
    assert cita in re.sub(r"\s+", " ", texto_pagina(INEI_BRECHAS, 178))
    filas.append({"Encuesta": "Criterio de confiabilidad", "Inferencia por ámbito": "Coeficiente de variación de 15 % o menos",
                  "Cita": "El INEI no estima por departamento el embarazo adolescente de 2024 porque las estimaciones "
                          "«presentan un coeficiente de variación mayor al 15%»",
                  "Documento": INEI_BRECHAS, "Página del PDF": 178})
    return pd.DataFrame(filas)


def criterios():
    return pd.DataFrame([
        ["INEI", "Confiabilidad", f"Coeficiente de variación de {fmt(CV_MAX, 0)} % o menos en mujeres y en hombres "
         "(INEI, Brechas de Género 2025, p. 178). Con más de 15 %, la estimación es referencial: el mapa la marca con "
         "trama y la nota de la figura lista los ámbitos referenciales con su coeficiente."],
        ["Equipo evaluador", "Casos mínimos", f"Con menos de {N_MIN} casos sin ponderar en un sexo no se estima (regla "
         "de presentación de la ENDES para estimaciones basadas en menos de 25 casos)."],
        ["Equipo evaluador", "Estimación", "Mismas definiciones que la estimación nacional; el diseño se arma con los "
         "casos de cada ámbito (svy.por_dominio), de modo que los estratos quedan anidados en él."],
        ["Equipo evaluador", "Brecha significativa", "Diferencia con la paridad (0 puntos o razón 1) mayor que 1,96 "
         "errores estándar."],
        ["Usuario (2026-11-02)", "Comparación", "2019 frente a 2025, o los años más cercanos con datos, con nota del "
         "cambio de instrumento. Significativo al 95 % y hacia la paridad: se reduce; en contra: se amplía. Para los "
         "indicadores de mujeres de la ENDES: mejora o empeora según el sentido del indicador. Con 26 pruebas al 5 % se "
         "esperan alrededor de 1,3 cambios significativos por azar."],
        ["Usuario (2026-11-02)", "Indicadores que se mapean", "Todos los que se pueden estimar, con la nota de "
         "confiabilidad. Los de la EPEN que la ENAHO mide desde 2019 se mapean con la ENAHO."],
        ["Equipo evaluador", "Empleo informal", "ENAHO 2019 frente a ENAHO 2025. La ENAHO no publica la variable ocupinf "
         "desde 2024; se reconstruye con la regla del INEI para la ENAHO (Cuenta Satélite de la Economía Informal "
         "2022–2024, anexo A2.2, pp. 87–90), que reproduce ocupinf en 2019–2023 con una concordancia individual de "
         "99,97 % o más (hoja «Validación informalidad»). La EPEN queda como contraste nacional."],
        ["Usuario (2026-11-02)", "Ámbitos", "23 departamentos, Callao, Lima Metropolitana (provincia de Lima) y Región "
         "Lima, como los publica el INEI. La EPEN solo distingue Lima Metropolitana con el Callao y Región Lima (variable "
         "DIVISION_LIMA); la ENA publica Lima entera y no tiene casos suficientes en el Callao."],
        ["Equipo evaluador", "Encuestas sin inferencia por ámbito", "ENUT 2024 (nacional, urbano y rural, regiones "
         "naturales y Lima Metropolitana) y ENARES 2019; la ENARES 2024 sí la tiene."],
        ["Equipo evaluador", "Participación política", "Registros completos: JNE y ONPE compilados por el INEI (anexo, "
         "cuadros 4.2 a 4.4) hasta 2022; gobernadoras 2019–2026 de los informes de cumplimiento de la LIO; Senado "
         "(distrito múltiple) 2026–2031 de la Resolución N.° 1412-2026-JNE; gobiernos regionales 2027–2030 de la ONPE "
         "con la regla del 30 % y la segunda elección; Cámara de Diputados 2026–2031 de la Resolución N.° "
         "1411-2026-JNE; alcaldías 2027–2030 del conteo de la ONPE de las ERM 2026. El sexo de las autoridades 2026 se "
         "identifica por el nombre (validado con la Tabla 3 de la ONPE para diputados y senadores). " + NOTA_2026 +
         ". Sin error muestral; con pocos cargos por ámbito, un caso cambia mucho el porcentaje."],
        ["Equipo evaluador", "Cartografía", "Peru maps, capa provincial simplificada (límites del INEI); uso solo "
         "estadístico."],
    ], columns=["Plano", "Elemento", "Regla"])


# ---------------------------------------------------------------------------------------------- figuras
def guardar(fig, nombre):
    os.makedirs(FIG, exist_ok=True)
    ruta = f"{FIG}/{nombre}.png"
    fig.savefig(ruta, dpi=300, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return ruta


def coma(ax):
    """Marcas en pasos redondos (1, 2 o 5 por potencia de 10) y decimales solo si la marca no es entera."""
    ax.yaxis.set_major_locator(ticker.MaxNLocator(nbins=5, steps=[1, 2, 5, 10]))
    ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, _: fmt(x, 0 if float(x).is_integer() else 1)))


def serie_panel(ax, g, titulo, sexo=True, estilo="-", desde=INICIO):
    g = g[g.anio >= desde].sort_values("anio")
    if g.empty:
        return
    anios = list(range(int(g.anio.min()), int(g.anio.max()) + 1))
    g = g.set_index("anio").reindex(anios)
    pares = [("valor_a", "ee_a", MUJERES)] + ([("valor_b", "ee_b", HOMBRES)] if sexo else [])
    for v, e, col in pares:
        ax.plot(g.index, g[v], color=col, lw=1.4, ls=estilo, marker="o", ms=2.5)
        if e in g and g[e].notna().any():
            ax.fill_between(g.index, g[v] - svy.Z95 * g[e], g[v] + svy.Z95 * g[e], color=col, alpha=0.18, lw=0)
            # un año sin vecinos no forma banda: su intervalo va como barra de error
            hay = g[v].notna()
            solo = hay & ~hay.shift(1, fill_value=False) & ~hay.shift(-1, fill_value=False)
            if solo.any():
                ax.errorbar(g.index[solo], g[v][solo], yerr=svy.Z95 * g[e][solo], fmt="none", ecolor=col,
                            elinewidth=0.9, capsize=2, alpha=0.6)
    ax.set_title(titulo, fontsize=7, color=TINTA, loc="left")
    ax.grid(axis="y", color="#e8e7e2", lw=0.5)
    ax.tick_params(labelsize=6, length=2)
    ax.set_xlim(desde - 0.4, FIN + 0.4)
    ax.xaxis.set_major_locator(ticker.MultipleLocator(2 if desde < 2015 else 1))
    coma(ax)


def leyenda_sexo(fig, sexo=True, extra=()):
    h = [Line2D([], [], color=MUJERES, lw=1.4, marker="o", ms=2.5, label="Mujeres")]
    if sexo:
        h.append(Line2D([], [], color=HOMBRES, lw=1.4, marker="o", ms=2.5, label="Hombres"))
    h.append(Patch(color=TINTA2, alpha=0.18, label="Intervalo de confianza al 95 %"))
    h += list(extra)
    fig.legend(handles=h, loc="lower center", ncol=len(h), frameon=False, fontsize=7, bbox_to_anchor=(0.5, -0.02))


def series_enaho():
    e = pd.read_csv(f"{D}/brechas_enaho.csv")
    claves = ["tasa_actividad", "tasa_desempleo", "sin_ingresos", "informal", "tfnr", "trabajo_hogar",
              "afiliado_pension", "ingreso_laboral", "ingreso_publico", "nini_15_29", "analfabetismo", "internet"]
    fig, axs = plt.subplots(4, 3, figsize=(6.3, 7.4))
    for ax, k in zip(axs.flat, claves):
        serie_panel(ax, e[e.indicador == k], META[("ENAHO", k)][0])
    fig.tight_layout(h_pad=1.2, rect=(0, 0.03, 1, 1))
    leyenda_sexo(fig)
    return guardar(fig, "series_enaho")


def series_otras():
    p = pd.read_csv(f"{D}/brechas_epen.csv")
    a = pd.read_csv(f"{D}/brechas_ena.csv")
    r = pd.read_csv(f"{D}/tolerancia_enares.csv")
    u = pd.read_csv(f"{D}/uso_tiempo_enut.csv")
    paneles = [(a[a.indicador == "credito_solicito"], "Productores que solicitaron crédito (ENA)", INICIO),
               (a[a.indicador == "propietario"], "Productores propietarios (ENA)", INICIO),
               (a[a.indicador == "titulo"], "Propietarios con título (ENA)", INICIO),
               (p[p.indicador == "tasa_actividad_epen"], "Tasa de actividad (EPEN)", INICIO),
               (r[r.indicador == "tolerancia_10"], "Tolerancia social a la violencia (ENARES)", INICIO),
               (u[u.indicador == "horas_no_remunerado"], "Horas semanales de trabajo no remunerado (ENUT)", 2010)]
    fig, axs = plt.subplots(2, 3, figsize=(6.3, 3.9))
    for ax, (g, tit, desde) in zip(axs.flat, paneles):
        serie_panel(ax, g, tit, desde=desde)
    fig.tight_layout(h_pad=1.2, rect=(0, 0.05, 1, 1))
    leyenda_sexo(fig)
    return guardar(fig, "series_otras")


def series_endes():
    n = pd.read_csv(f"{D}/manifestaciones_endes.csv").rename(columns={"estimacion": "valor_a", "ee": "ee_a"})
    claves = [k for (enc, k) in META if enc == "ENDES"]
    fig, axs = plt.subplots(2, 3, figsize=(6.3, 3.9))
    for ax, k in zip(axs.flat, claves):
        serie_panel(ax, n[n.indicador == k], META[("ENDES", k)][0], sexo=False)
    fig.tight_layout(h_pad=1.2, rect=(0, 0.05, 1, 1))
    leyenda_sexo(fig, sexo=False)
    return guardar(fig, "series_endes")


def serie_participacion():
    pp = pd.read_csv(f"{D}/participacion_politica_jne.csv")
    pp["anio"] = pp.periodo.str.extract(r"(\d{4})").astype(float)
    series = [("Congreso de la República", "Congreso, composición en funciones (INEI)", "#6250d6", "-"),
              ("Congreso de la República, electas (ONPE)", "Congreso, electas (ONPE)", "#6250d6", ":"),
              ("Alcaldías (provinciales y distritales)", "Alcaldías", "#1baf7a", "-"),
              ("Regidurías", "Regidurías", "#eda100", "-"),
              ("Gobiernos regionales: gobernadoras", "Gobernadoras regionales", "#e34948", "-")]
    fig, ax = plt.subplots(figsize=(6.3, 3.0))
    for cargo, etiqueta, col, est in series:
        g = pp[pp.cargo.str.startswith(cargo) & (pp.cargo.str.contains("ONPE") == ("ONPE" in cargo))]
        g = g[g.porcentaje_mujeres.notna()].sort_values("anio")
        ax.plot(g.anio, g.porcentaje_mujeres, color=col, ls=est, lw=1.4, marker="o", ms=2.5, label=etiqueta)
    ax.axhline(40, color="#6250d6", lw=0.8, ls="--")
    ax.axhline(20, color="#1baf7a", lw=0.8, ls="--")
    ax.text(1951, 41, "Meta al 2030: 40 % de congresistas", fontsize=6.5, color=TINTA2)
    ax.text(1951, 21, "Meta al 2030: 20 % de alcaldesas", fontsize=6.5, color=TINTA2)
    ax.set_ylim(0, 56)
    ax.set_ylabel("% de mujeres", fontsize=7)
    ax.set_xlabel("Año de inicio del periodo (electas de la ONPE: año de la elección)", fontsize=6.5)
    ax.grid(axis="y", color="#e8e7e2", lw=0.5)
    ax.tick_params(labelsize=6.5, length=2)
    coma(ax)
    ax.legend(frameon=False, fontsize=6.5, loc="upper left", bbox_to_anchor=(0, 1.0), ncol=2)
    return guardar(fig, "series_participacion_politica")


# ---------------------------------------------------------------------------------------------- mapas
def mapa_base():
    """26 ámbitos: departamentos, Callao, Lima Metropolitana (provincia 1501) y Región Lima (resto del 15)."""
    g = gpd.read_file(GEO)[["ccdd", "ccpp", "geometry"]]
    g["ambito"] = np.where(g.ccdd == "15", np.where(g.ccpp == "01", "LM", "RL"), g.ccdd)
    g = g.dissolve(by="ambito", as_index=False)[["ambito", "geometry"]]
    assert len(g) == 26
    return g


def expandir(valores):
    """Pinta los ámbitos conjuntos sobre sus partes: «LMC» sobre Lima Metropolitana y el Callao; «15» sobre Lima."""
    out = dict(valores)
    for conj, partes in (("LMC", ("LM", "07")), ("15", ("LM", "RL"))):
        if conj in out:
            for k in partes:
                out.setdefault(k, out[conj])
    return {k: v for k, v in out.items() if k not in ("LMC", "15")}


def cortes(valores, n=3, d=1):
    v = np.asarray(sorted(valores))
    if len(v) < n:
        return []
    return sorted({round(float(np.quantile(v, q)), d) for q in np.linspace(0, 1, n + 1)[1:-1]})


def clase_magnitud(x, cs):
    return int(np.searchsorted(cs, x, side="right"))


def etiquetas_rango(cs, d, unidad):
    lim = [0.0] + list(cs)
    out = []
    for i in range(len(cs) + 1):
        if i == 0:
            out.append(f"hasta {fmt(lim[1], d)}{unidad}" if cs else "cualquier magnitud")
        elif i < len(cs):
            out.append(f"{fmt(lim[i], d)} a {fmt(lim[i + 1], d)}{unidad}")
        else:
            out.append(f"más de {fmt(lim[i], d)}{unidad}")
    return out


def pintar(ax, geo, colores, trama):
    colores = expandir(colores)
    trama = set(expandir({k: 1 for k in trama}))
    geo.plot(ax=ax, color=geo.ambito.map(colores).fillna(SIN_DATO), edgecolor=BORDE, linewidth=0.3)
    t = geo[geo.ambito.isin(trama)]
    if not t.empty:
        t.plot(ax=ax, facecolor="none", edgecolor=TINTA2, linewidth=0, hatch="//////")
        t.plot(ax=ax, facecolor="none", edgecolor=BORDE, linewidth=0.3)
    ax.set_axis_off()


def leyendas(fig, h, hc, titulo="Brecha en el año"):
    """Dos bloques: clases del año bajo los dos primeros mapas y clases del cambio bajo el tercero."""
    estilo = dict(loc="upper center", frameon=False, fontsize=6, title_fontsize=6.5, alignment="left")
    fig.legend(handles=h, ncol=2, title=titulo, bbox_to_anchor=(0.36, 0.07), **estilo)
    if hc:
        fig.legend(handles=hc, ncol=1, title="Cambio", bbox_to_anchor=(0.84, 0.07), **estilo)


def titulo_anio(g, anio):
    return str(anio)


def mapa_brecha(t, c, enc, ind, geo):
    nombre, tipo, signo = META[(enc, ind)]
    g = t[(t.encuesta == enc) & (t.indicador == ind)]
    inicio, fin = periodos(g)
    d = 2 if tipo == "razon" else 1
    unidad = "" if tipo == "razon" else " pp"
    usar = g[g.anio.isin([inicio, fin]) & g.estimable & (g.significativa == 1)]
    cs = cortes((usar.valor - usar.paridad).abs(), d=d)
    unico = inicio == fin  # la ENARES solo tiene un año con inferencia por ámbito
    fig, axs = plt.subplots(1, 1 if unico else 3, figsize=(3.4, 3.6) if unico else (6.3, 3.3))
    axs = np.atleast_1d(axs)
    usados, hay_trama, hay_sin = set(), False, False
    for ax, anio in zip(axs[:2], (fin,) if unico else (inicio, fin)):
        u = ambitos_mapa(g[g.anio == anio].set_index("ambito"))
        col, trama = {}, []
        for cc, r in u.iterrows():
            if not r.estimable:
                continue
            if r.significativa != 1:
                col[cc] = NEUTRO
            else:
                lado = np.sign(r.valor - r.paridad) * (signo if signo else 1)
                col[cc] = (ROJO if lado > 0 else AZUL)[clase_magnitud(abs(r.valor - r.paridad), cs)]
            if not r.confiable:
                trama.append(cc)
        usados |= set(col.values())
        hay_trama |= bool(trama)
        hay_sin |= len(expandir(col)) < 26
        pintar(ax, geo, col, trama)
        ax.set_title(titulo_anio(g, anio), fontsize=8, color=TINTA)
    cambios_usados = set()
    if not unico:
        k = c[(c.encuesta == enc) & (c.indicador == ind)].set_index("ambito")
        colc = {cc: {"Se reduce": AZUL[1], "Se amplía": ROJO[1], "Sin cambio significativo": NEUTRO}.get(cl)
                for cc, cl in k.clase.items() if cl != "Sin estimación"}
        tramac = [cc for cc, r in k.iterrows() if r.clase != "Sin estimación" and not r.confiable_ambos]
        cambios_usados = set(k.clase)
        pintar(axs[2], geo, colc, tramac)
        axs[2].set_title(f"Cambio {inicio}–{fin}", fontsize=8, color=TINTA)
    rangos = etiquetas_rango(cs, d, unidad)
    malo, bueno = (("Desventaja de las mujeres", "Desventaja de los hombres") if signo else
                   ("Mayor en mujeres", "Mayor en hombres"))
    h = [Patch(facecolor=ROJO[i], label=f"{malo}: {rangos[i]}") for i in range(len(rangos)) if ROJO[i] in usados]
    h += [Patch(facecolor=AZUL[i], label=f"{bueno}: {rangos[i]}") for i in range(len(rangos)) if AZUL[i] in usados]
    if NEUTRO in usados:
        h.append(Patch(facecolor=NEUTRO, edgecolor=BORDE, lw=0.3, label="Brecha no significativa"))
    if hay_trama:
        h.append(Patch(facecolor="white", edgecolor=TINTA2, hatch="//////", lw=0.3,
                       label="Referencial (coeficiente de variación > 15 %)"))
    if hay_sin:
        h.append(Patch(facecolor=SIN_DATO, edgecolor=BORDE, lw=0.3, label="Sin estimación"))
    hc = [Patch(facecolor=color, edgecolor=BORDE, lw=0.3, label=cl) for cl, color in
          (("Se reduce", AZUL[1]), ("Se amplía", ROJO[1]), ("Sin cambio significativo", NEUTRO)) if cl in cambios_usados]
    if unico:
        fig.legend(handles=h, ncol=1, title="Brecha en el año", bbox_to_anchor=(0.5, 0.07), loc="upper center",
                   frameon=False, fontsize=6, title_fontsize=6.5, alignment="left")
    else:
        leyendas(fig, h, hc)
    medida = "razón mujeres/hombres; la paridad es 1" if tipo == "razon" else "mujeres menos hombres, en puntos"
    titulo = nombre if f"({enc})" in nombre else f"{nombre} ({enc})"
    fig.suptitle(f"{titulo}:\nbrecha por departamento ({medida})" if unico else
                 f"{titulo}: brecha por departamento ({medida})", fontsize=8, color=TINTA, y=0.99)
    fig.subplots_adjust(wspace=0.02, top=0.83 if unico else 0.9, bottom=0.06)
    return guardar(fig, f"mapa_{ind}"), inicio, fin


def mapa_nivel(t, c, enc, ind, geo):
    nombre, _, sentido = META[(enc, ind)]
    g = t[(t.encuesta == enc) & (t.indicador == ind)]
    inicio, fin = periodos(g)
    usar = g[g.anio.isin([inicio, fin]) & g.estimable]
    cs = cortes(usar.valor, n=5)
    fig, axs = plt.subplots(1, 3, figsize=(6.3, 3.3))
    for ax, anio in zip(axs[:2], (inicio, fin)):
        u = g[(g.anio == anio) & g.estimable].set_index("ambito")
        col = {cc: SECUENCIAL[clase_magnitud(r.valor, cs)] for cc, r in u.iterrows()}
        pintar(ax, geo, col, [cc for cc, r in u.iterrows() if not r.confiable])
        ax.set_title(str(anio), fontsize=8, color=TINTA)
    k = c[(c.encuesta == enc) & (c.indicador == ind)].set_index("ambito")
    hay_trama = bool((~usar.confiable).any())
    colc = {cc: {"Mejora": AZUL[1], "Empeora": ROJO[1], "Sin cambio significativo": NEUTRO}.get(cl)
            for cc, cl in k.clase.items() if cl != "Sin estimación"}
    pintar(axs[2], geo, colc, [cc for cc, r in k.iterrows() if r.clase != "Sin estimación" and not r.confiable_ambos])
    axs[2].set_title(f"Cambio {inicio}–{fin}", fontsize=8, color=TINTA)
    lim = [usar.valor.min()] + cs + [usar.valor.max()]
    h = [Patch(facecolor=SECUENCIAL[i], label=f"{fmt(lim[i])} a {fmt(lim[i + 1])} %") for i in range(len(lim) - 1)]
    if hay_trama:
        h.append(Patch(facecolor="white", edgecolor=TINTA2, hatch="//////", lw=0.3,
                       label="Referencial (coeficiente de variación > 15 %)"))
    hc = [Patch(facecolor=color, edgecolor=BORDE, lw=0.3, label=cl) for cl, color in
          (("Mejora", AZUL[1]), ("Empeora", ROJO[1]), ("Sin cambio significativo", NEUTRO)) if cl in set(k.clase)]
    leyendas(fig, h, hc, titulo="Nivel en el año")
    fig.suptitle(f"{nombre} (ENDES, mujeres): nivel por departamento", fontsize=8, color=TINTA, y=0.99)
    fig.subplots_adjust(wspace=0.02, top=0.9, bottom=0.06)
    return guardar(fig, f"mapa_{ind}"), inicio, fin


# periodos en funciones en 2019 y en 2025, y la composición que regirá en 2030 (elecciones de 2026)
ALC, REG, CONG = "Alcaldías (provinciales y distritales)", "Regidurías", "Congreso de la República"
# (clave, paneles (cargo, periodo, rótulo), cortes de las clases)
PP_MAPAS = [("alcaldias", [(ALC, "2019/2022", "2019/2022"), (ALC, "2023/2026", "2023/2026"),
                           (ALC, "2027/2030", "2027/2030 (ONPE)")], [2.5, 5, 10]),
            ("regidurias", [(REG, "2019/2022", "2019/2022"), (REG, "2023/2026", "2023/2026")], [30, 40, 50]),
            ("congreso", [(CONG, "2016/2020", "2016/2020"), (CONG, "2021/2026", "2021/2026"),
                          (CONG, "2026/2031", "2026/2031, diputados"),
                          ("Senado, distrito múltiple", "2026/2031", "2026/2031, Senado\n(distrito múltiple)")],
             [20, 40, 50])]


def mapas_participacion(pd_dep, geo):
    rutas = []
    for clave, paneles, cs in PP_MAPAS:
        fig, axs = plt.subplots(1, len(paneles), figsize=(2.15 * len(paneles), 3.4))
        for ax, (cargo, per, rotulo) in zip(axs, paneles):
            u = pd_dep[(pd_dep.cargo == cargo) & (pd_dep.periodo == per) & pd_dep.total.notna()].set_index("ambito")
            col = {cc: SECUENCIAL[1 + clase_magnitud(r.porcentaje_mujeres, cs)] if r.mujeres > 0 else SECUENCIAL[0]
                   for cc, r in u.iterrows()}
            pintar(ax, geo, col, [])
            ax.set_title(f"{rotulo}\n{fmt(100 * u.mujeres.sum() / u.total.sum())} % en los departamentos", fontsize=7,
                         color=TINTA)
        lim = [0] + cs
        h = [Patch(facecolor=SECUENCIAL[0], label="Ninguna mujer")]
        h += [Patch(facecolor=SECUENCIAL[1 + i], label=(f"Más de 0 a {fmt(lim[1], 1)} %" if i == 0 else
                                                        f"{fmt(lim[i], 1)} a {fmt(lim[i + 1], 1)} %"))
              for i in range(len(cs))]
        h.append(Patch(facecolor=SECUENCIAL[4], label=f"Más de {fmt(cs[-1], 1)} %"))
        fig.legend(handles=h, loc="upper center", ncol=3, frameon=False, fontsize=6, bbox_to_anchor=(0.5, 0.06))
        nombre = {"alcaldias": "Alcaldesas", "regidurias": "Regidoras", "congreso": "Congresistas"}[clave]
        fig.suptitle(f"{nombre}: % de mujeres entre las autoridades electas por departamento", fontsize=8,
                     color=TINTA, y=0.99)
        fig.subplots_adjust(wspace=0.02, top=0.85, bottom=0.06)
        rutas.append((guardar(fig, f"mapa_participacion_{clave}"), clave, [x[2].split("\n")[0] for x in paneles]))
    return rutas


def mapa_gobiernos_regionales(pd_dep, geo):
    """Gobernadora o gobernador por circunscripción regional; 2027–2030 con el estado de la segunda elección."""
    g = pd_dep[pd_dep.cargo == "Gobiernos regionales"]
    pers = ("2019/2022", "2023/2026", "2027/2030")
    GOBERNADORA, GOBERNADOR = AZUL[2], SECUENCIAL[0]
    fig, axs = plt.subplots(1, 3, figsize=(6.45, 3.4))
    for ax, per in zip(axs, pers):
        u = g[g.periodo == per].set_index("ambito")
        assert len(u) == 25, (per, len(u))
        col = {c: GOBERNADORA if r.mujeres == 1 else GOBERNADOR for c, r in u.iterrows()}
        col["LM"] = NEUTRO
        segunda = [c for c, r in u.iterrows() if str(r.nota).startswith("Segunda elección")]
        pintar(ax, geo, col, segunda)
        rot = f"{per} (ONPE)" if per == "2027/2030" else per
        ax.set_title(f"{rot}\n{int(u.mujeres.sum())} gobernadora{'s' if u.mujeres.sum() != 1 else ''} de 25",
                     fontsize=7, color=TINTA)
    h = [Patch(facecolor=GOBERNADORA, edgecolor=BORDE, label="Gobernadora"),
         Patch(facecolor=GOBERNADOR, edgecolor=BORDE, label="Gobernador"),
         Patch(facecolor=GOBERNADOR, edgecolor=TINTA2, hatch="//////",
               label="Segunda elección entre dos fórmulas encabezadas por hombres (6-12-2026)"),
         Patch(facecolor=NEUTRO, edgecolor=BORDE, label="No elige gobernador (Lima Metropolitana)")]
    fig.legend(handles=h, loc="upper center", ncol=2, frameon=False, fontsize=6, bbox_to_anchor=(0.5, 0.08))
    fig.suptitle("Gobiernos regionales: sexo de quien encabeza la fórmula ganadora por circunscripción", fontsize=8,
                 color=TINTA, y=0.99)
    fig.subplots_adjust(wspace=0.02, top=0.85, bottom=0.08)
    return guardar(fig, "mapa_participacion_gobiernos_regionales")


# ---------------------------------------------------------------------------------------------- libro
def escribir(hojas):
    with pd.ExcelWriter(SALIDA, engine="openpyxl") as w:
        for nombre, df in hojas.items():
            df.to_excel(w, sheet_name=nombre, index=False)
    wb = load_workbook(SALIDA)
    for ws in wb.worksheets:
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F3864")
            c.alignment = Alignment(wrap_text=True, vertical="top")
        for col in ws.columns:
            largo = max(len(str(c.value)) if c.value is not None else 0 for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(max(10, largo * 0.9), 60)
        ws.freeze_panes = "A2"
    wb.save(SALIDA)


def main():
    t = leer()
    c = cambios(t)
    r = resumen(t, c)
    v = validar(t)
    geo = mapa_base()
    pd_dep = pd.read_csv(f"{D}/participacion_politica_departamental.csv", dtype={"ambito": str})
    figuras = [("series_enaho", series_enaho(), "Brechas laborales, económicas y educativas por sexo, 2019–2025",
                "ENAHO; 2020 se levantó en buena parte por teléfono. Empleo informal: la ENAHO no publica la variable "
                "desde 2024; 2024 y 2025 se reconstruyen con la regla del INEI (Cuenta Satélite de la Economía Informal "
                "2022–2024, pp. 87–90), que la reproduce de 2019 a 2023. Ingresos en soles corrientes."),
               ("series_otras", series_otras(), "Brechas por sexo con la ENA, la EPEN, la ENARES y la ENUT",
                "ENA: no hubo encuesta en 2020, el capítulo de crédito no se publicó en 2021 y el de parcelas cambió en "
                "2023 (de las trabajadas en los últimos 12 meses a las que se trabajan al momento de la entrevista). "
                "EPEN: base anual de 2023 a 2025. ENARES y ENUT: dos ediciones con instrumentos distintos (la ENUT, "
                "2010 y 2024); la comparación es indicativa."),
               ("series_endes", series_endes(), "Indicadores de mujeres de la ENDES, 2019–2025",
                "2020 se levantó en buena parte por teléfono; la violencia de pareja de 2020 no reproduce el valor "
                "oficial (7,9 frente a 8,8 %)."),
               ("series_participacion_politica", serie_participacion(),
                "Mujeres en el Congreso, las alcaldías, las regidurías y los gobiernos regionales, 1956–2030",
                "El INEI informa la composición del Congreso en funciones (actualizada a agosto de 2025) y la ONPE las "
                "electas; por eso difieren en 2016 y 2021. Gobernadoras 2015–2026: informes de cumplimiento de la LIO "
                "(MIMP; fuente primaria: JNE). Alcaldías y gobernadoras 2027–2030: conteo de la ONPE con sexo por "
                f"nombre. {NOTA_2026}. Metas de la situación futura deseada (PNIG, p. 20).")]
    for _, f in r[r.Mapa == "Sí"].iterrows():
        enc, ind = f.Encuesta, f.Clave
        fn = mapa_nivel if META[(enc, ind)][1] == "nivel" else mapa_brecha
        ruta, inicio, fin = fn(t, c, enc, ind, geo)
        g = t[(t.encuesta == enc) & (t.indicador == ind)]
        anios = (inicio, fin) if inicio != fin else (fin,)
        nota = " ".join(x for x in (NOTA_INSTRUMENTO.get(ind, ""), "Trama: estimación referencial. Confiabilidad:",
                                    nota_confiabilidad(g, anios)) if x)
        figuras.append((f"mapa_{ind}", ruta, f"{META[(enc, ind)][0]} por departamento, "
                        + (f"{inicio} y {fin}" if inicio != fin else str(fin)), f"{enc}. {nota}"))
    for ruta, clave_pp, pers in mapas_participacion(pd_dep, geo):
        clave = os.path.basename(ruta)[:-4]
        cargo_txt = {"alcaldias": "alcaldías provinciales y distritales", "regidurias": "regidurías",
                     "congreso": "Congreso de la República"}[clave_pp]
        extra = {"alcaldias": f" Alcaldías 2027–2030: conteo de la ONPE (ERM 2026; sin segunda vuelta). {NOTA_2026}.",
                 "congreso": " Cámara de Diputados y Senado 2026–2031: Resoluciones N.° 1411-2026-JNE y 1412-2026-JNE; "
                             "el Senado solo muestra los 30 escaños del distrito múltiple (Lima Metropolitana 4, uno "
                             "en cada otro departamento); los 30 del distrito único nacional no tienen departamento (11 mujeres)."
                 }.get(clave_pp, "")
        figuras.append((clave, ruta, f"Mujeres electas por departamento: {cargo_txt}, " + ", ".join(pers),
                        "JNE y ONPE, compilados por el INEI (Brechas de Género 2025, anexo), hasta 2022." + extra +
                        " El sexo de las autoridades de 2026 se identifica por el nombre. Registros completos, sin "
                        "error muestral; con pocos cargos por departamento, un caso cambia mucho el porcentaje."))
    figuras.append(("mapa_participacion_gobiernos_regionales", mapa_gobiernos_regionales(pd_dep, geo),
                    "Gobernadoras regionales por departamento, 2019–2022, 2023–2026 y 2027–2030",
                    "2019–2022 y 2023–2026: informes de cumplimiento de la LIO (MIMP, 2018, p. 26, y 2022, p. 103; "
                    "fuente primaria: JNE). 2027–2030: dos fórmulas más votadas según la ONPE, con sexo por nombre; "
                    "gana en primera vuelta la fórmula con no menos del 30 % de los votos válidos (artículo 5 de la "
                    f"Ley N.° 27683, modificado por la Ley N.° 29470). {NOTA_2026}, así que el resultado no cambia. "
                    "Lima Metropolitana no elige gobernador: la Municipalidad Metropolitana de Lima ejerce las "
                    "funciones del gobierno regional."))
    pd.DataFrame(figuras, columns=["clave", "archivo", "titulo", "nota"]).to_csv(f"{FIG}/indice_figuras.csv",
                                                                                  index=False, encoding="utf-8")
    pp = pd_dep.pivot_table(index=["cargo", "ambito", "departamento"], columns="periodo",
                            values="porcentaje_mujeres").reset_index()
    columnas = ["encuesta", "nombre", "indicador", "anio", "ambito", "departamento", "valor_a", "ee_a", "cv_a", "n_a",
                "valor_b", "ee_b", "cv_b", "n_b", "valor", "ee", "cv", "significativa", "condicion"]
    hojas = {"Resumen por indicador": r,
             "Brechas por ámbito": t[columnas].rename(columns={
                 "valor_a": "mujeres", "valor_b": "hombres", "valor": "brecha o nivel", "nombre": "indicador (nombre)",
                 "departamento": "ámbito (nombre)"}).round(3),
             "Cambio por ámbito": c.round(3),
             "Participación política": pp,
             "Validación con el INEI": v,
             "Validación informalidad": pd.read_csv(f"{D}/validacion_informal_enaho.csv").rename(columns={
                 "anio": "año", "ocupinf": "ocupinf del INEI (%)", "regla_inei": "regla del INEI reconstruida (%)",
                 "diferencia": "diferencia (puntos)", "concordancia": "concordancia individual ponderada (%)",
                 "discrepancias": "casos discrepantes", "casos": "ocupados residentes"}),
             "Diseño muestral": diseno(),
             "Criterios": criterios()}
    escribir(hojas)
    pd.set_option("display.width", 250)
    print(r[["Encuesta", "Clave", "Años comparados", "Ámbitos con estimación (último año)", "Confiables (último año)",
             "Cambios significativos", "Mapa"]].to_string(index=False))
    print(f"validación: {(v['Dentro del IC 95 %'] == 'Sí').sum()} de {len(v)} dentro del IC 95 %")
    print(f"{len(figuras)} figuras en {FIG}")


if __name__ == "__main__":
    main()
