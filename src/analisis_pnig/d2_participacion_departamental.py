"""D2: participación política de las mujeres por departamento (congresistas, alcaldesas y regidoras).

Extrae del anexo de «Perú: Brechas de Género, 2025» del INEI (fuente primaria: ONPE y JNE) los cuadros por departamento:
  - 4.2, parlamentarios según departamento de representación, periodos 2011/2016 a 2021/2026 (pp. 222 y 223);
  - 4.3, alcaldesas y alcaldes elegidos por distritos, periodos 1983/1986 a 2023/2026 (pp. 224 y 225);
  - 4.4, regidoras y regidores elegidos por distritos, mismos periodos (pp. 226 y 227).
Lee el texto plano de cada página (pdftotext -raw), donde cada ámbito ocupa una línea y le siguen las filas «Mujeres» y
«Hombres» («-» es cero). Comprueba que la suma de los ámbitos iguale la fila «Nacional» en cada periodo y sexo, y que
el total nacional coincida con la serie ya verificada en participacion_politica_jne.csv. Lima Metropolitana, Región
Lima y el Callao se conservan como ámbitos separados, como los publica el INEI; antes de 2007 los cuadros municipales
no desagregan Lima y esos periodos quedan sin dato para ambos ámbitos.

Agrega, si existe, d1_brechas/autoridades_2026_onpe.csv (diputados y senadores del distrito múltiple 2026/2031,
alcaldías y gobiernos regionales 2027/2030 por ámbito; ver src/fuentes/onpe_resultados_2026.js y
src/analisis_pnig/d2_autoridades_2026.py), y las gobernadoras de 2019/2022 y 2023/2026 por ámbito de los informes de
cumplimiento de la LIO del MIMP (2018, p. 26: ninguna; 2022, p. 103: Lima y Moquegua). Lima Metropolitana no elige
gobernador: su gobierno regional lo ejerce la Municipalidad Metropolitana de Lima.

Son registros administrativos completos, sin error muestral.

Uso (desde la raíz del repositorio):
    uv run --no-project --with pandas python -I -X utf8 src/analisis_pnig/d2_participacion_departamental.py
Salida: evaluaciones/pnig_2026/04_analisis/d1_brechas/participacion_politica_departamental.csv
"""
import os
import re
import subprocess
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import svy  # noqa: E402
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

PDF = "docs/fuentes/pnig/04_seguimiento/fuentes_indicadores/2025_inei_brechas-de-genero-2025.pdf"
D = f"{EVAL}/04_analisis/d1_brechas"
CUADROS = [  # (cargo como en participacion_politica_jne.csv, cuadro, páginas)
    ("Congreso de la República", "4.2", (222, 223)),
    ("Alcaldías (provinciales y distritales)", "4.3", (224, 225)),
    ("Regidurías", "4.4", (226, 227)),
]
# ámbitos que no son un departamento: Lima se divide en Lima Metropolitana y Región Lima (svy.AMBITOS)
OTROS = {"lima metropolitana": "LM", "region lima": "RL", "prov. const. del callao": "07",
         "residentes en el extranjero": "extranjero", "nacional": "nacional"}


def texto(pag):
    return subprocess.run(["pdftotext", "-enc", "UTF-8", "-raw", "-f", str(pag), "-l", str(pag), PDF, "-"],
                          capture_output=True, check=True).stdout.decode("utf-8")


def ambito(linea):
    """Código del ámbito si la línea empieza con su nombre (con o sin cifras y llamadas «1/»), si no None."""
    nombre = re.sub(r"(\s+\d/)?(\s+[\d-]+)*\s*$", "", linea).strip()
    if not nombre or nombre.startswith(("Mujeres", "Hombres")):
        return None
    clave = re.sub(r"\s+", " ", nombre).lower()
    clave = clave.translate(str.maketrans("áéíóú", "aeiou"))
    if clave in OTROS:
        return nombre, OTROS[clave]
    try:
        return nombre, svy.codigo_departamento(nombre)
    except KeyError:
        return None


def numeros(linea):
    return [0 if x == "-" else int(x) for x in linea.split()[1:]]


def extraer(cargo, cuadro, paginas):
    lineas = [l.strip() for p in paginas for l in texto(p).splitlines()]
    periodos = re.findall(r"(\d{4}) ?/ ?\n?(\d{4})", "\n".join(lineas))
    periodos = list(dict.fromkeys(f"{a}/{b}" for a, b in periodos))
    filas, actual = [], None
    for l in lineas:
        if l.startswith(("Mujeres", "Hombres")) and actual:
            v = numeros(l)
            assert len(v) == len(periodos), (cuadro, actual, l)
            sexo = "mujeres" if l.startswith("Mujeres") else "hombres"
            for per, x in zip(periodos, v):
                filas.append({"cargo": cargo, "cuadro": cuadro, "periodo": per, "ambito": actual[0],
                              "ccdd": actual[1], "sexo": sexo, "casos": x})
        else:
            a = ambito(l)
            if a:
                actual = a
    t = pd.DataFrame(filas).pivot_table(index=["cargo", "cuadro", "periodo", "ambito", "ccdd"], columns="sexo",
                                        values="casos", aggfunc="sum").reset_index()
    t.columns.name = None
    return t


def main():
    serie = pd.read_csv(f"{D}/participacion_politica_jne.csv")
    todo = []
    for cargo, cuadro, paginas in CUADROS:
        t = extraer(cargo, cuadro, paginas)
        nac = t[t.ccdd == "nacional"].set_index("periodo")
        # antes de 2007 los cuadros municipales no desagregan Lima («-» en Lima Metropolitana y Región Lima): el total
        # nacional sí la incluye, así que en esos periodos Lima queda sin dato y la suma no se exige
        lima = t[t.ccdd.isin(["LM", "RL"])].groupby("periodo")[["mujeres", "hombres"]].sum().sum(axis=1)
        sin_lima = lima[lima == 0].index
        t.loc[t.ccdd.isin(["LM", "RL"]) & t.periodo.isin(sin_lima), ["mujeres", "hombres"]] = float("nan")
        resto = t[t.ccdd != "nacional"].groupby("periodo")[["mujeres", "hombres"]].sum()
        dif = (resto - nac[["mujeres", "hombres"]]).drop(sin_lima).abs().sum().sum()
        assert dif == 0, f"cuadro {cuadro}: la suma de los ámbitos no iguala el total nacional ({dif})"
        ref = serie[serie.cargo == cargo].set_index("periodo")
        for per, r in nac.iterrows():
            assert (r.mujeres, r.hombres) == (ref.loc[per, "mujeres"], ref.loc[per, "hombres"]), (cuadro, per)
        print(f"cuadro {cuadro}: {t.ambito.nunique() - 1} ámbitos, {len(nac)} periodos; suma = total nacional "
              f"(Lima sin desagregar en {len(sin_lima)} periodos)")
        todo.append(t[t.ccdd != "nacional"])
    t = pd.concat(todo, ignore_index=True)
    t = t[t.ccdd != "extranjero"]  # 2 congresistas 2021/2026 de la circunscripción de peruanos en el extranjero
    dep = t.groupby(["cargo", "cuadro", "periodo", "ccdd"], as_index=False)[["mujeres", "hombres"]].sum(min_count=1)
    dep = dep.rename(columns={"ccdd": "ambito"})
    dep["departamento"] = dep.ambito.map(svy.AMBITOS)
    dep["total"] = dep.mujeres + dep.hombres
    dep["porcentaje_mujeres"] = (100 * dep.mujeres / dep.total).round(1)
    dep["fuente"] = "INEI, Perú: Brechas de Género 2025, anexo (fuente primaria: ONPE y JNE)"
    dep["paginas"] = dep.cuadro.map({c: f"{p[0]} y {p[1]}" for _, c, p in CUADROS})
    dep["nota"] = ""
    dep.loc[dep.ambito.isin(["LM", "RL"]) & dep.total.isna(), "nota"] = "El cuadro no desagrega Lima en este periodo"
    regiones = [c for c in svy.DEPARTAMENTOS if c != "15"] + ["RL"]   # 25 circunscripciones regionales
    lio = pd.DataFrame([{"cargo": "Gobiernos regionales", "cuadro": f"LIO {a}", "periodo": per, "ambito": c,
                         "mujeres": int(c in electas), "hombres": int(c not in electas),
                         "fuente": f"MIMP, Informe de cumplimiento de la LIO {a} (fuente primaria: JNE)",
                         "paginas": pag, "nota": ""}
                        for per, a, pag, electas in (("2019/2022", 2018, "26", ()), ("2023/2026", 2022, "103",
                                                                                     ("RL", "18")))
                        for c in regiones])
    assert len(lio) == 50 and lio.mujeres.sum() == 2
    lio["departamento"] = lio.ambito.map(svy.AMBITOS)
    lio["total"] = 1
    lio["porcentaje_mujeres"] = 100.0 * lio.mujeres
    dep = pd.concat([dep, lio], ignore_index=True)
    onpe = f"{D}/autoridades_2026_onpe.csv"
    if os.path.exists(onpe):
        n = pd.read_csv(onpe, dtype={"ambito": str})
        n["departamento"] = n.ambito.map(svy.AMBITOS)
        n["total"] = n.mujeres + n.hombres
        n["porcentaje_mujeres"] = (100 * n.mujeres / n.total).round(1)
        dep = pd.concat([dep, n], ignore_index=True)
    dep = dep[["cargo", "cuadro", "periodo", "ambito", "departamento", "total", "mujeres", "hombres",
               "porcentaje_mujeres", "fuente", "paginas", "nota"]]
    dep.to_csv(f"{D}/participacion_politica_departamental.csv", index=False, encoding="utf-8")
    print(dep.groupby(["cargo", "periodo"]).ambito.nunique().unstack(0).to_string())


if __name__ == "__main__":
    main()
