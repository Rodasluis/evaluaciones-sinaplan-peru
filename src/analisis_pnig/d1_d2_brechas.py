"""D1 y D2: consolidación de brechas, manifestaciones del problema y situación futura deseada.

Lee las salidas de d1_brechas_enaho.py, d1_brechas_endes.py, d1_brechas_enares.py, d1_brechas_enut.py,
d1_brechas_epen.py y d1_brechas_ena.py, y la serie de participación política del JNE y la ONPE
(d1_brechas/participacion_politica_jne.csv), y arma evaluaciones/pnig_2026/04_analisis/D1_D2_brechas_situacion_futura.xlsx
con estas hojas:
  - «Evolución de las brechas» (D1): por indicador, brecha en el año del diagnóstico, en 2019 y en el último año, cambio
    2019–último año con su error estándar y clasificación;
  - «Brechas por año»: serie completa por sexo con error estándar;
  - «ENAHO y EPEN»: brechas laborales de 2023 a 2025 con ambas encuestas (contraste; la serie de seguimiento es la
    ENAHO, con el empleo informal de 2024 y 2025 reconstruido con la regla del INEI);
  - «Participación política» (D2): congresistas, alcaldesas y regidoras por periodo, con verificación de cada
    porcentaje en la página del INEI que lo publica;
  - «Otras encuestas»: inventario de las encuestas exploradas para medir brechas, con su uso y veredicto
    (d1_brechas/inventario_encuestas_brechas.csv);
  - «Manifestaciones ENDES»: indicadores de mujeres 2019–2025, contraste con los valores oficiales y valores 2025;
  - «Situación futura deseada» (D2-H1): cada enunciado de la p. 20 del documento de la PNIG con su causa o efecto, su
    OP, la correspondencia con el indicador del OP, el último valor y la proyección lineal al 2030;
  - «Validación», «Criterios» y «Microdatos».

Regla de clasificación de una brecha (equipo evaluador): se compara la brecha (mujeres menos hombres; razón para
ingresos y horas) entre 2019 y el último año disponible. Si el cambio es significativo al 95 % y la brecha se acerca a
la paridad, «se reduce»; si se aleja, «se amplía»; si no es significativo, «sin cambio significativo». Además se indica
si en el último año la brecha sigue siendo significativa («persiste»).

Valor al 2030 (D2-H1): extrapolación lineal del ritmo anual observado entre la línea de base de la situación futura
deseada y el último valor; es un cálculo mecánico, no un pronóstico. Para congresistas y alcaldesas, la composición
vigente en 2030 la fijan las elecciones de 2026: se usa la composición electa cuando está disponible y, si no, la
proyección lineal hasta esa elección. «Probable» si la proyección alcanza la
meta, «improbable» si no la alcanza, «no evaluable» si falta línea de base o el indicador del OP mide otra cosa.

Uso (desde la raíz del repositorio, después de los seis scripts por encuesta):
    uv run --no-project --with pandas --with openpyxl python -I -X utf8 src/analisis_pnig/d1_d2_brechas.py
"""
import os
import re
import subprocess
import sys

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import svy  # noqa: E402
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

D = f"{EVAL}/04_analisis/d1_brechas"
SALIDA = f"{EVAL}/04_analisis/D1_D2_brechas_situacion_futura.xlsx"
PNIG = "docs/fuentes/pnig/01_normativa/2019_mimp_ds-008-2019-mimp_pnig.pdf"
PSI = f"{EVAL}/04_analisis/REV_IND_base_indicadores.xlsx"
GRD = f"{EVAL}/04_analisis/REV_DOC_estado_implementacion_servicios.xlsx"
REGISTRO = "docs/fuentes/descargas_microdatos-inei_2026-11-01.tsv"
INEI_BRECHAS = "docs/fuentes/pnig/04_seguimiento/fuentes_indicadores/2025_inei_brechas-de-genero-2025.pdf"
ONPE_AE3 = "docs/fuentes/pnig/04_seguimiento/fuentes_indicadores/2026_onpe_analitica-electoral-03_resultados-eg-2026.pdf"

ROLES = "Causa: asignación desigual de roles"
ECON = "Efecto: vulneración de los derechos económicos y sociales"
PATRONES = "Causa: patrones socioculturales discriminatorios"
NORMAS = "Causa: normas y cultura institucional"
# indicador: (causa o efecto, medida de la brecha, año base del diagnóstico, cita del diagnóstico)
MAPA = {
    "horas_no_remunerado": (ROLES, "razon", 2010, "M 39 h 28 min y H 15 h 54 min semanales (ENUT 2010); p. 17"),
    "tasa_actividad": (ROLES, "diferencia", 2017, "M 64,0 % y H 81,0 % (2017); p. 26"),
    "tfnr": (ROLES, "diferencia", 2016, "Mujeres 17,0 %; p. 26"),
    "trabajo_hogar": (ROLES, "diferencia", 2016, "Mujeres 5,1 %, hombres casi nulo; p. 26"),
    "sin_ingresos": (ECON, "diferencia", 2016, "M 31,5 % y H 12,0 % (2016); p. 26"),
    "ingreso_laboral": (ECON, "razon", 2017, "Las mujeres ganaron 29,3 % menos que los hombres (2017); p. 26"),
    "informal": (ECON, "diferencia", 2016, "75,1 % de las mujeres ocupadas (2016); p. 26"),
    "afiliado_pension": (ECON, "diferencia", 2017, ""),
    "tasa_desempleo": (ECON, "diferencia", 2017, "5,4 % de las mujeres urbanas (2017); p. 26"),
    "analfabetismo": (ECON, "diferencia", 2017, ""),
    "nini_15_29": (ECON, "diferencia", 2017, "M 24,0 % y H 11,8 % (2017); p. 24"),
    "internet": (ECON, "diferencia", 2016, "M 42,6 % y H 48,3 % (2016); p. 19"),
    "ingreso_publico": (NORMAS, "razon", 2017, "Brecha salarial de 16 % en el servicio civil (SERVIR, 2018); p. 26"),
    "tolerancia_10": (PATRONES, "diferencia", 2019, "Índice de tolerancia social 54,8 % (ENARES 2015); p. 18"),
    "credito_solicito": (ECON, "diferencia", 2012, "Poco acceso a recursos productivos y créditos; p. 19"),
}
# La superficie de las parcelas (diagnóstico, p. 19: 1,8 ha agrícolas frente a 3 ha, CENAGRO 2012) no se clasifica: su
# promedio oscila entre años por parcelas muy grandes y la mediana no tiene error estándar. Va en «Brechas por año».
# brechas que continúan la serie de otra encuesta: se muestran aparte y no se cuentan dos veces (ninguna desde que el
# empleo informal de la ENAHO 2024 y 2025 se reconstruye con la regla del INEI; la EPEN queda como contraste)
CONTINUA = {}


def leer_brechas():
    """Todas las series por sexo; las que tienen entrada en MAPA pasan a la clasificación."""
    e = pd.read_csv(f"{D}/brechas_enaho.csv")
    u = pd.read_csv(f"{D}/uso_tiempo_enut.csv")
    r = pd.read_csv(f"{D}/tolerancia_enares.csv")
    r = r[r.indicador == "tolerancia_10"]
    p = pd.read_csv(f"{D}/brechas_epen.csv")
    a = pd.read_csv(f"{D}/brechas_ena.csv")
    cols = ["encuesta", "indicador", "etiqueta", "anio", "valor_a", "ee_a", "valor_b", "ee_b", "diferencia",
            "ee_diferencia", "razon", "ee_razon", "nota"]
    for df in (e, u, r, p, a):
        for c in cols:
            if c not in df.columns:
                df[c] = np.nan
    t = pd.concat([e[cols], u[cols], r[cols], p[cols], a[cols]], ignore_index=True)
    return t[t.valor_a.notna()]


def clasificar(t):
    filas = []
    for ind, g in t[t.indicador.isin(MAPA)].groupby("indicador", sort=False):
        causa, medida, anio_base, cita = MAPA[ind]
        g = g.set_index("anio")
        valor = "razon" if medida == "razon" else "diferencia"
        ee = "ee_razon" if medida == "razon" else "ee_diferencia"
        paridad = 1.0 if medida == "razon" else 0.0
        conee = g[g[ee].notna()].index  # los valores publicados sin error estándar no sirven de punto de comparación
        inicio = 2019 if 2019 in conee else conee.min()
        fin = conee.max()
        if inicio == fin:
            inicio = conee.min()
        v0, e0, v1, e1 = g.loc[inicio, valor], g.loc[inicio, ee], g.loc[fin, valor], g.loc[fin, ee]
        dif, ee_dif, sig = svy.cambio(v0, e0, v1, e1)
        acerca = abs(v1 - paridad) < abs(v0 - paridad)
        if not sig:
            clase = "Sin cambio significativo"
        else:
            clase = "Se reduce" if acerca else "Se amplía"
        persiste = abs(v1 - paridad) > svy.Z95 * e1
        nota = "; ".join(sorted({str(x) for x in g.nota.dropna()}))
        if ind == "horas_no_remunerado":
            nota = "Instrumentos distintos en 2010 y 2024 (lista frente a diario): la comparación es indicativa"
        filas.append({
            "Indicador": g.etiqueta.iloc[0], "Clave": ind, "Brecha del diagnóstico": CONTINUA.get(ind, ind),
            "Encuesta": g.encuesta.iloc[0], "Causa o efecto": causa,
            "Medida de la brecha": "Razón mujeres/hombres" if medida == "razon" else "Mujeres menos hombres (pp)",
            "Diagnóstico de la PNIG": cita,
            f"Brecha en el año base": round(g.loc[anio_base, valor], 2) if anio_base in g.index else np.nan,
            "Año base": anio_base if anio_base in g.index else np.nan,
            "Año inicial": inicio, "Brecha inicial": round(v0, 3 if medida == "razon" else 2),
            "Año final": fin, "Brecha final": round(v1, 3 if medida == "razon" else 2),
            "Mujeres (año final)": round(g.loc[fin, "valor_a"], 2), "Hombres (año final)": round(g.loc[fin, "valor_b"], 2),
            "Cambio": round(dif, 3), "EE del cambio": round(ee_dif, 3), "Clasificación": clase,
            "Persiste en el año final": "Sí" if persiste else "No", "Nota": nota})
    return pd.DataFrame(filas)


def texto_pagina(pdf, pag):
    out = subprocess.run(["pdftotext", "-enc", "UTF-8", "-f", str(pag), "-l", str(pag), pdf, "-"],
                         capture_output=True, check=True).stdout.decode("utf-8")
    out = out.replace("ﬁ", "fi").replace("ﬂ", "fl")
    return re.sub(r"\s+", " ", out)


def participacion_politica():
    """Serie del JNE y la ONPE; cada porcentaje se busca en la página del documento que lo publica (INEI u ONPE)."""
    pp = pd.read_csv(f"{D}/participacion_politica_jne.csv")
    paginas = {}
    for _, r in pp[pp.verificacion == "Verificado en el documento"].iterrows():
        pdf = ONPE_AE3 if r.fuente.startswith("ONPE") else INEI_BRECHAS
        pag = int(str(r.pagina).split()[0].rstrip(","))
        if (pdf, pag) not in paginas:
            paginas[(pdf, pag)] = texto_pagina(pdf, pag)
        # la ONPE escribe los decimales con punto y el INEI con coma
        cifra = f"{r.porcentaje_mujeres:.1f}" if pdf == ONPE_AE3 else fmt(r.porcentaje_mujeres)
        assert r.porcentaje_mujeres == 0 or cifra in paginas[(pdf, pag)], \
            f"{r.cargo} {r.periodo}: {cifra} no consta en p. {pag} de {pdf}"
        if pd.notna(r.mujeres) and pd.notna(r.total):
            calc = round(100 * float(r.mujeres) / float(r.total), 1)
            assert calc == r.porcentaje_mujeres, f"{r.cargo} {r.periodo}: {calc} frente a {r.porcentaje_mujeres}"
    return pp


def situacion_futura(endes, pp):
    s = pd.read_csv(f"{D}/situacion_futura_deseada.csv", dtype={"anio_base_fsd": "Int64"})
    pag = texto_pagina(PNIG, 20)
    faltan = [c for c in s.cita_pnig if re.sub(r"\s+", " ", c) not in pag]
    assert not faltan, f"citas que no constan en la p. 20 del documento de la PNIG: {faltan}"
    base = pd.read_excel(PSI, sheet_name="Base REV_IND.1").set_index("Código")
    enaho = pd.read_csv(f"{D}/brechas_enaho.csv")
    ing = enaho[enaho.indicador == "ingreso_laboral"].set_index("anio")
    grd = pd.read_excel(GRD, sheet_name="Estado por servicio").set_index("Código")
    e = endes.set_index(["indicador", "anio"])
    cong = pp[pp.cargo.str.startswith("Congreso") & ~pp.cargo.str.contains("ONPE")].set_index("periodo")
    alc = pp[pp.cargo.str.startswith("Alcaldías")].set_index("periodo")
    serie_alc = alc[alc.porcentaje_mujeres.notna()].porcentaje_mujeres
    serie_alc_ant = serie_alc.drop("2027/2030", errors="ignore")
    salto_max = serie_alc_ant.diff().max()
    electo = {  # composición vigente en 2030: (valor, verificado, método)
        "SFD4a": (cong.loc["2026/2031", "porcentaje_mujeres"], True,
                  "Composición electa en la EG 2026, vigente de 2026 a 2031: 70 de 190 (ONPE, Analítica Electoral "
                  "n.° 3, Tabla 3, p. 22; Senado 16 de 60 y Cámara de Diputados 54 de 130)"),
        "SFD4b": (alc.loc["2027/2030", "porcentaje_mujeres"], True,
                  f"Composición electa en las ERM 2026, vigente de 2027 a 2030: {int(alc.loc['2027/2030', 'mujeres'])}"
                  f" alcaldesas de {int(alc.loc['2027/2030', 'total'])} ganadores con sexo determinado (resultados "
                  f"preliminares de la ONPE, con 99,85 % de actas provinciales y 99,81 % de distritales contabilizadas "
                  f"al 2026-11-02, pendientes de proclamación del JNE; sexo por nombre, con los casos dudosos revisados "
                  f"a mano). Las alcaldías no tienen segunda vuelta. Llegar a 20 % exigía subir "
                  f"{fmt(20 - serie_alc_ant.iloc[-1])} puntos en una elección; el mayor aumento entre dos elecciones desde "
                  f"1983 fue de {fmt(salto_max)} puntos"),
    }
    ultimo = {
        "SFD1": (float(base.loc["IOP.06.01", "Valor 2024"]), 2024, "IOP.06.01, valor oficial 2024"),
        "SFD2": (float(base.loc["IOP.05.01", "Valor 2025"]), 2025,
                 "IOP.05.01, valor oficial 2025 (entidades públicas, no ministerios)"),
        "SFD3": (float(e.loc[("violencia_fisica_o_sexual_12m", 2025), "estimacion"]), 2025,
                 "ENDES 2025, cálculo del equipo evaluador (oficial 2024: 7,5 %)"),
        "SFD4a": (cong.loc["2021/2026", "porcentaje_mujeres"], 2021,
                  "JNE y ONPE, compilado por el INEI (Brechas de Género 2025, p. 56): 52 de 130 congresistas en "
                  "funciones en 2021–2026; la meta se alcanzó en ese periodo. Electas en la EG 2021: 37,7 % (ONPE, "
                  "Analítica Electoral n.° 3, p. 21)"),
        "SFD4b": (alc.loc["2023/2026", "porcentaje_mujeres"], 2022,
                  "JNE y ONPE, compilado por el INEI (Brechas de Género 2025, p. 58): 102 alcaldesas en 2023–2026, "
                  "igual a la cifra del JNE (El Peruano, 8-3-2023)"),
        "SFD5": (float(100 * ing.loc[2025, "razon"]), 2025,
                 f"ENAHO 2025, razón del ingreso laboral mensual (2017: {100 * ing.loc[2017, 'razon']:.1f} %)"),
        "SFD6": (np.nan, np.nan, f"Servicio 4.1.1: {grd.loc['4.1.1', 'Estado asignado (preliminar)'].lower()} "
                                 f"(REV_DOC)"),
        "SFD7": (float(e.loc[("embarazo_adolescente", 2025), "estimacion"]), 2025,
                 "ENDES 2025, cálculo del equipo evaluador (oficial 2024: 8,4 %)"),
    }
    filas = []
    for _, r in s.iterrows():
        v, a, fuente = ultimo[r.codigo]
        proy, clase, metodo, verificado = np.nan, "No evaluable", "", "Sí"
        directa = r.codigo in electo  # el enunciado se mide directamente con la serie del JNE
        if pd.notna(v) and pd.notna(r.linea_base_fsd) and (r.correspondencia_indicador != "Distinta" or directa):
            ritmo = (v - r.linea_base_fsd) / (a - int(r.anio_base_fsd))
            proy, metodo = v + ritmo * (2030 - a), "Lineal"
            if directa:
                valor, ok, metodo = electo[r.codigo]
                proy = valor if pd.notna(valor) else v + ritmo * (2026 - a)
                verificado = "Sí" if ok else "No"
            logra = proy <= r.meta_2030_fsd if r.sentido == "Descendente" else proy >= r.meta_2030_fsd
            clase = ("Probable" if logra else "Improbable") + ("" if verificado == "Sí" else " [POR VERIFICAR]")
        elif r.codigo == "SFD5" and pd.notna(v):
            clase = "No evaluable: la línea de base no se reproduce"
        filas.append({"Código": r.codigo, "Enunciado (cita literal)": r.cita_pnig, "Página": r.pagina,
                      "Causa o efecto del modelo": f"{r.plano_modelo}: {r.causa_o_efecto}", "OP": r.op,
                      "Indicador del OP": r.indicador_op, "Correspondencia con el indicador del OP":
                          r.correspondencia_indicador, "Sustento": r.sustento_correspondencia,
                      "Línea de base": r.linea_base_fsd, "Año base": r.anio_base_fsd, "Meta 2030": r.meta_2030_fsd,
                      "Último valor": round(v, 2) if pd.notna(v) else np.nan, "Año": a, "Fuente del último valor": fuente,
                      "Valor 2030": round(proy, 2) if pd.notna(proy) else np.nan, "Método del valor 2030": metodo,
                      "Valor 2030 verificado": verificado if metodo else "", "Alcanzabilidad": clase})
    return pd.DataFrame(filas)


def validacion():
    v = pd.read_csv(f"{D}/validacion_enaho_diagnostico.csv")
    v = v[v.calculado.notna()].assign(encuesta="ENAHO")
    n = pd.read_csv(f"{D}/manifestaciones_endes.csv")
    n = n[n.oficial.notna()]
    n = pd.DataFrame({"encuesta": "ENDES", "indicador": n.indicador, "anio": n.anio, "sexo": "Mujer",
                      "referencia": n.oficial, "origen": "Valor oficial del reporte de seguimiento (REV_IND)",
                      "calculado": n.estimacion.round(2), "diferencia": n.diferencia_oficial.round(2),
                      "dentro_ic95": n.dentro_ic95})
    u = pd.read_csv(f"{D}/uso_tiempo_enut.csv")
    u = u[u.anio == 2010]
    u = pd.DataFrame({"encuesta": "ENUT", "indicador": "horas_no_remunerado", "anio": 2010, "sexo": ["Mujer", "Hombre"],
                      "referencia": [u.diagnostico_mujer.iloc[0], u.diagnostico_hombre.iloc[0]],
                      "origen": "Diagnóstico de la PNIG (Freyre y López, 2011)",
                      "calculado": [round(u.valor_a.iloc[0], 2), round(u.valor_b.iloc[0], 2)]})
    u["diferencia"] = (u.calculado - u.referencia).round(2)
    r = pd.read_csv(f"{D}/tolerancia_enares.csv")
    r = pd.DataFrame({"encuesta": "ENARES", "indicador": r.indicador, "anio": r.anio, "sexo": "Total",
                      "referencia": r.oficial_IOP0601, "origen": "IOP.06.01 oficial (no se reproduce; ver Criterios)",
                      "calculado": r.total.round(2), "diferencia": (r.total - r.oficial_IOP0601).round(2),
                      "dentro_ic95": False})
    p = pd.read_csv(f"{D}/validacion_epen.csv").assign(encuesta="EPEN")
    a = pd.read_csv(f"{D}/validacion_ena.csv").assign(encuesta="ENA")
    return pd.concat([v, n, u, r, p, a], ignore_index=True)


def enaho_epen(t):
    """Brechas laborales comparables de 2023 a 2025 con ambas encuestas (diferencia M − H; razón para ingresos)."""
    pares = {"tasa_actividad": "tasa_actividad_epen", "informal": "informal_epen", "afiliado_pension":
             "afiliado_pension_epen", "ingreso_laboral": "ingreso_laboral_epen"}
    filas = []
    for enaho, epen in pares.items():
        col = "razon" if enaho == "ingreso_laboral" else "diferencia"
        for anio in (2023, 2024, 2025):
            fila = {"Indicador": t.loc[t.indicador == epen, "etiqueta"].iloc[0], "Año": anio,
                    "Medida de la brecha": "Razón M/H" if col == "razon" else "M − H (pp)"}
            for enc, clave in (("ENAHO", enaho), ("EPEN", epen)):
                g = t[(t.indicador == clave) & (t.anio == anio)]
                fila[f"Mujeres {enc}"] = round(g.valor_a.iloc[0], 2) if len(g) else np.nan
                fila[f"Hombres {enc}"] = round(g.valor_b.iloc[0], 2) if len(g) else np.nan
                fila[f"Brecha {enc}"] = round(g[col].iloc[0], 3) if len(g) else np.nan
            filas.append(fila)
    return pd.DataFrame(filas)


def fmt(x):
    return f"{x:.1f}".replace(".", ",")


def criterios():
    return pd.DataFrame([
        ["Equipo evaluador", "Fuentes", "ENAHO 2016–2025 (sin 2018), ENDES 2019–2025, ENARES 2019 y 2024, ENUT 2010 y "
         "2024, EPEN 2023–2025 y ENA 2019 y 2021–2025. Microdatos del portal del INEI; los descargados en esta "
         "evaluación tienen su hash en " + REGISTRO + ". Participación política: JNE y ONPE, compilados por el INEI "
         "(Brechas de Género 2025)."],
        ["Equipo evaluador", "Estimación", "Estimador de razón ponderado con linealización de Taylor (svy.py), "
         "validado contra el repositorio «Indicadores ENDES» del usuario (mismo error estándar a tres decimales)."],
        ["Equipo evaluador", "Validación", "Cada definición se contrastó con la cifra del diagnóstico de la PNIG, con la "
         "serie oficial del INEI o con el valor oficial del reporte de seguimiento (hoja «Validación»)."],
        ["Equipo evaluador", "Clasificación de la brecha", "Cambio 2019–último año significativo al 95 % y en "
         "dirección a la paridad: se reduce; en dirección contraria: se amplía; no significativo: sin cambio "
         "significativo."],
        ["Equipo evaluador", "Año 2020", "ENAHO y ENDES 2020 se levantaron en buena parte por teléfono; sus valores "
         "se muestran, pero no se usan como punto de comparación."],
        ["Equipo evaluador", "Empleo informal", "La ENAHO no publica la variable ocupinf en 2024 y 2025, porque el INEI "
         "pasó la medición a la EPEN. Se reconstruye con la regla del INEI para la ENAHO (Producción y Empleo Informal "
         "en el Perú, Cuenta Satélite de la Economía Informal 2022–2024, anexo A2.2, pp. 87–90), que reproduce ocupinf "
         "en 2019–2023 con una concordancia individual de 99,97 % o más y tasas por sexo a 0,03 puntos o menos "
         "(d1_brechas/validacion_informal_enaho.csv). Así la brecha se sigue con la ENAHO de 2019 a 2025. La EPEN "
         "2023–2025 queda como contraste (hoja «ENAHO y EPEN»): mide niveles unos 3 puntos menores y una brecha "
         "similar; su base anual 2022 está dañada en el portal."],
        ["Equipo evaluador", "Productoras agropecuarias", "ENA, productores personas naturales (fila del productor en "
         "el capítulo 1100). Crédito: solicitó algún crédito en los últimos 12 meses. Título: sobre los propietarios, "
         "como el INEI. Reproduce las ocho cifras del INEI para 2024. El estrato geográfico no se publica: los errores "
         "estándar son aproximados."],
        ["Equipo evaluador", "Participación política", "Serie por periodo del JNE y la ONPE compilada por el INEI "
         "(cuadros 4.1, 4.3 y 4.4 del anexo) y serie de electas de la ONPE (Analítica Electoral n.° 3, Gráfico 13 y "
         "Tabla 3); cada porcentaje se recalcula con los casos y se busca en la página del documento. Las dos series "
         "del Congreso difieren por definición: el INEI informa la composición en funciones, actualizada a agosto de "
         "2025 (p. 56), y la ONPE las electas al proclamarse los resultados (37,7 % frente a 40,0 % en 2021). Las bases "
         "de Infogob se piden con un formulario de datos personales y captcha, que debe llenar el usuario."],
        ["Equipo evaluador", "Tolerancia social", "Ni la regla de la sintaxis (al menos una respuesta tolerante) ni "
         "ningún umbral reproducen el índice oficial de IOP.06.01. Se usa un índice comparable con las 10 afirmaciones "
         "comunes a 2019 y 2024 solo para la brecha por sexo y la dirección del cambio."],
        ["Equipo evaluador", "Uso del tiempo", "ENUT 2010 (lista de actividades) y 2024 (diario de 10 minutos) no son "
         "comparables en nivel; la razón mujeres/hombres se compara de forma indicativa."],
        ["Equipo evaluador", "Jóvenes que no estudian ni están ocupados", "Definición estándar; no reproduce la cifra "
         "del diagnóstico (24,0 % y 11,8 %), cuya definición no se publica."],
        ["Equipo evaluador", "Valor al 2030", "Lineal, con el ritmo observado entre la línea de base de la "
         "situación futura deseada y el último valor. No es un pronóstico. Para congresistas y alcaldesas se usa la "
         "composición elegida en 2026, que es la vigente en 2030, o, si aún no está disponible, la proyección lineal "
         "hasta esa elección."],
        ["Equipo evaluador", "Valores 2025 de la ENDES", "Cálculo del equipo evaluador con microdatos; no son valores "
         "reportados por las entidades. Los de 2019 a 2024 reproducen los oficiales."],
    ], columns=["Plano", "Elemento", "Regla"])


def escribir(hojas):
    with pd.ExcelWriter(SALIDA, engine="openpyxl") as w:
        for nombre, df in hojas.items():
            df.to_excel(w, sheet_name=nombre, index=False)
    wb = load_workbook(SALIDA)
    azul = PatternFill("solid", fgColor="1F3864")
    colores = {"Se reduce": "C6EFCE", "Se amplía": "F8CBAD", "Sin cambio significativo": "FFEB9C",
               "Probable": "C6EFCE", "Improbable": "F8CBAD", "Improbable [POR VERIFICAR]": "F8CBAD"}
    for ws in wb.worksheets:
        for c in ws[1]:
            c.fill, c.font = azul, Font(bold=True, color="FFFFFF")
            c.alignment = Alignment(wrap_text=True, vertical="center")
        for col in ws.columns:
            largo = max(len(str(c.value or "")) for c in col[:50])
            ws.column_dimensions[col[0].column_letter].width = min(max(10, largo * 0.9), 60)
        for fila in ws.iter_rows(min_row=2):
            for c in fila:
                c.alignment = Alignment(wrap_text=True, vertical="top")
                if isinstance(c.value, str) and c.value in colores:
                    c.fill = PatternFill("solid", fgColor=colores[c.value])
        ws.freeze_panes = "B2"
    wb.save(SALIDA)


if __name__ == "__main__":
    t = leer_brechas()
    evol = clasificar(t)
    endes = pd.read_csv(f"{D}/manifestaciones_endes.csv")
    pp = participacion_politica()
    sfd = situacion_futura(endes, pp)
    serie = t.rename(columns={"valor_a": "Mujeres", "ee_a": "EE mujeres", "valor_b": "Hombres", "ee_b": "EE hombres",
                              "diferencia": "Diferencia (M − H)", "ee_diferencia": "EE diferencia",
                              "razon": "Razón M/H", "ee_razon": "EE razón"})
    man = endes[["indicador", "etiqueta", "anio", "estimacion", "ee", "ic_inf", "ic_sup", "n", "codigo_op", "oficial",
                 "diferencia_oficial", "referencia_diagnostico"]]
    micro = pd.read_csv(REGISTRO, sep="\t")
    hojas = {"Evolución de las brechas": evol, "Brechas por año": serie, "ENAHO y EPEN": enaho_epen(t),
             "Manifestaciones ENDES": man, "Situación futura deseada": sfd, "Participación política": pp,
             "Otras encuestas": pd.read_csv(f"{D}/inventario_encuestas_brechas.csv"),
             "Validación": validacion(), "Criterios": criterios(), "Microdatos": micro}
    escribir(hojas)
    pd.set_option("display.width", 220)
    print(evol[["Clave", "Año inicial", "Brecha inicial", "Año final", "Brecha final", "Cambio", "EE del cambio",
                "Clasificación", "Persiste en el año final"]].to_string(index=False))
    print(sfd[["Código", "Correspondencia con el indicador del OP", "Línea de base", "Meta 2030", "Último valor", "Año",
               "Valor 2030", "Alcanzabilidad"]].to_string(index=False))
