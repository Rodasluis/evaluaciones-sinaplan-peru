"""D2: mujeres electas en 2026 por ámbito (Congreso de la EG 2026; alcaldías y gobiernos regionales de las ERM 2026).

Diputados y senadores 2026–2031: listas oficiales de las Resoluciones N.° 1411-2026-JNE y 1412-2026-JNE
(proclamación, 19-6-2026), por circunscripción y organización política. El sexo se identifica por el primer nombre con
las mismas reglas que src/fuentes/onpe_resultados_2026.js (decisión del usuario, 2026-11-02) y se valida contra la
Tabla 3 de la Analítica Electoral 3 de la ONPE (p. 22): mujeres y hombres electos por organización en cada cámara. Si
no cuadra, el script se detiene. El Senado tiene 30 escaños en el distrito único nacional y 30 en el distrito múltiple
(cuatro en Lima Metropolitana y uno en cada otra circunscripción); solo el múltiple se asigna a un ámbito.

Alcaldías 2027–2030: conteos por ámbito que produjo onpe_resultados_2026.js en el navegador (ganador por provincia y
distrito según el conteo de la ONPE; elección en una sola vuelta), con los ganadores sin sexo determinado revisados a
mano (03_datos/raw/onpe_erm2026/). Los que siguen sin dato quedan como no determinados y no entran en el porcentaje.

Gobiernos regionales 2027–2030: las dos fórmulas más votadas de cada circunscripción y su porcentaje de votos válidos
(03_datos/raw/onpe_erm2026/gobernadores_2026-11-02.csv). Gana en primera vuelta la fórmula con no menos del 30 % de los
votos válidos; si ninguna llega, hay segunda elección entre las dos más votadas (artículo 5 de la Ley N.° 27683,
modificado por la Ley N.° 29470; el JNE la fijó para el 6-12-2026). Con el sexo de las dos fórmulas se obtiene el
número mínimo y máximo de gobernadoras: si las dos las encabezan hombres, el ámbito no tendrá gobernadora.

Los resultados de las ERM 2026 son preliminares: conteo de la ONPE (actas contabilizadas en NOTA_2026), pendiente de
proclamación del JNE.

Ningún nombre se guarda: la salida solo tiene conteos.

Uso (desde la raíz del repositorio):
    uv run --no-project --with pandas python -I -X utf8 src/analisis_pnig/d2_autoridades_2026.py
Salida: evaluaciones/pnig_2026/04_analisis/d1_brechas/autoridades_2026_onpe.csv (la lee d2_participacion_departamental.py)
"""
import os
import re
import subprocess
import sys
import unicodedata

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import svy  # noqa: E402
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

FI = "docs/fuentes/pnig/04_seguimiento/fuentes_indicadores"
RES = f"{FI}/2026_jne_resolucion-1411-2026-jne_proclamacion-diputados.pdf"
RES_SENADO = f"{FI}/2026_jne_resolucion-1412-2026-jne_proclamacion-senadores.pdf"
ERM = f"{EVAL}/03_datos/raw/onpe_erm2026/conteo_sexo_alcaldias_2026-11-02_revisado.csv"
GOB = f"{EVAL}/03_datos/raw/onpe_erm2026/gobernadores_2026-11-02.csv"
# Nota única de los resultados de las ERM 2026 (la repiten las tablas, las figuras y el informe)
NOTA_2026 = ("Resultados preliminares de la ONPE (corte del 2026-11-02: 99,83 % de actas contabilizadas en la elección "
             "regional, 99,85 % en la provincial y 99,81 % en la distrital), pendientes de proclamación del JNE. La "
             "segunda elección regional del 6 de diciembre de 2026 define al gobernador en 15 circunscripciones; en "
             "todas, las dos fórmulas más votadas están encabezadas por hombres")
SALIDA = f"{EVAL}/04_analisis/d1_brechas/autoridades_2026_onpe.csv"
PARTIDOS = ("FUERZA POPULAR", "JUNTOS POR EL PERU", "RENOVACION POPULAR", "PARTIDO DEL BUEN GOBIERNO",
            "PARTIDO CIVICO OBRAS", "AHORA NACION - AN")
# ONPE, Analítica Electoral 3, Tabla 3 (p. 22): diputadas y diputados electos por organización
TABLA3 = {"FUERZA POPULAR": (20, 21), "JUNTOS POR EL PERU": (15, 17), "RENOVACION POPULAR": (4, 11),
          "PARTIDO DEL BUEN GOBIERNO": (10, 8), "PARTIDO CIVICO OBRAS": (3, 11), "AHORA NACION - AN": (2, 8)}
# Misma tabla, senadoras y senadores; y circunscripciones del distrito múltiple que eligieron senadora (p. 22)
TABLA3_SENADO = {"FUERZA POPULAR": (5, 17), "JUNTOS POR EL PERU": (1, 13), "RENOVACION POPULAR": (3, 5),
                 "PARTIDO DEL BUEN GOBIERNO": (3, 4), "PARTIDO CIVICO OBRAS": (2, 3), "AHORA NACION - AN": (2, 2)}
SENADORAS_MULTIPLE = {"LIMA METROPOLITANA": 2, "ANCASH": 1, "PIURA": 1, "TACNA": 1}

# Reglas de sexo por nombre: las mismas de src/fuentes/onpe_resultados_2026.js
FEM = set(("CARMEN ROSARIO MARISOL RUTH BEATRIZ ISABEL INES RAQUEL SOLEDAD MERCEDES DOLORES NIEVES LUZ ESTHER ESTER "
           "MIRIAM MIRIAN NOEMI MARLENY MARLENI MARLENE NELLY GLADYS ROCIO CONSUELO PILAR MARIBEL ELIZABETH ELISABETH "
           "SHIRLEY KAREN LISBETH LIZBETH LIZ YENNY JENNY YENI ROSMERY JUDITH JUDIT EDITH MILAGROS ANGELES IRIS ABIGAIL "
           "LOURDES MARY BETTY NANCY EVELYN EVELING JAKELIN JAQUELINE JACQUELINE KATHERINE LEILY LEINY HEYDE DEICE EILEN "
           "ELENITH FLOR MARGOT ANALI LEYDI KELLY SUSAN YANET JANET MAGALY ARACELY ANABEL DORIS LUCY CLARIBEL PILY "
           "ROSSMERY MIRTHA GRISELDA MERCY NATALY NATALI CATHERIN CATHERINE HAYDEE JEENNY JENIFER YENIFER MADY MERY "
           "NARY NATHALY PIERANGELI YULI KIM").split())
# primeros nombres masculinos que pueden ir seguidos de un segundo nombre femenino (José María, Wilmer Dolores)
MASC_SEG = set("JOSE JUAN WILMER".split())
MASC_A = set("JOSHUA LUCA BAUTISTA GARCIA RUNTSMA".split())
AMBIG = set("ROSARIO GUADALUPE TRINIDAD CRUZ ASUNCION ACENCION CONCEPCION ADIS ESLI ESLY NIVE EALY ARLEN ARLA SANTA "
            "GABY".split())


def norm(s):
    return unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().upper().strip()


def sexo(nombre):
    t = norm(nombre).split()

    def una(x):
        if not x:
            return "?"
        if x == "MARIA":
            return "F"
        if x in AMBIG:
            return "?"
        if x in FEM:
            return "F"
        return "F" if x.endswith("A") and x not in MASC_A else "M"

    if t[0] in ("JOSE", "JUAN"):
        return "M"
    s = una(t[0])
    if s == "M" and len(t) >= 4 and t[0] not in MASC_SEG and una(t[1]) == "F":
        return "F"   # primer nombre poco frecuente y segundo nombre femenino (Mery Eliana, Yuli Liliana)
    if s != "?":
        return s
    return "?" if len(t) <= 3 else una(t[1])


def diputados():
    texto = subprocess.run(["pdftotext", "-enc", "UTF-8", "-raw", RES, "-"], capture_output=True,
                           check=True).stdout.decode("utf-8")
    # une las líneas cortadas: encabezados de circunscripción sin «escaños» y nombres sin la cifra de votos
    lineas, crudas = [], [x.strip() for x in texto.splitlines() if x.strip()]
    i = 0
    while i < len(crudas):
        l = crudas[i]
        while i + 1 < len(crudas) and ((l.startswith("3.") and "CIRCUNSCRIPCIÓN" in l and "escaño" not in l) or
                                       re.fullmatch(r"\d+\. [A-ZÁÉÍÓÚÑÜ ]+", l)):  # nombre en mayúsculas sin votos
            i += 1
            l = f"{l} {crudas[i]}"
        lineas.append(l)
        i += 1
    filas, circ, escanos, partido = [], None, {}, None
    for l in lineas:
        m = re.match(r"3\.\d+\. CIRCUNSCRIPCIÓN ELECTORAL DE (.+?) \((\d+) escaños?\)", l)
        if m:
            circ = norm(m.group(1))
            escanos[circ] = int(m.group(2))
            partido = None
            continue
        m = re.fullmatch(r"(.+?) (\d+)", norm(l))
        if m and m.group(1) in PARTIDOS:
            partido = m.group(1)
            continue
        m = re.fullmatch(r"\d+\. (.+?) \d+", l)
        if m and circ and partido:
            filas.append({"circunscripcion": circ, "partido": partido, "sexo": sexo(m.group(1))})
    d = pd.DataFrame(filas)
    por_circ = d.groupby("circunscripcion").size()
    dif = por_circ.reindex(list(escanos)).fillna(0) - pd.Series(escanos)
    assert (dif == 0).all() and len(d) == 130, f"no coincide: {dif[dif != 0].to_dict()}, {len(d)} filas, {sum(escanos.values())} escaños, {len(escanos)} circunscripciones"
    tabla = d.pivot_table(index="partido", columns="sexo", aggfunc="size", fill_value=0)
    for p, (f, m) in TABLA3.items():
        assert (tabla.loc[p, "F"], tabla.loc[p, "M"]) == (f, m), f"{p}: {tabla.loc[p].to_dict()} frente a Tabla 3 {(f, m)}"
    print(f"diputados: 130 en {len(escanos)} circunscripciones; sexo por nombre = Tabla 3 de la ONPE en los 6 partidos")
    return d


def senadores():
    texto = subprocess.run(["pdftotext", "-enc", "UTF-8", "-raw", RES_SENADO, "-"], capture_output=True,
                           check=True).stdout.decode("utf-8")
    # el PDF es la edición de El Peruano: la parte resolutiva va de «3. PROCLAMAR» al «4. DISPONER» que le sigue
    ini = texto.index("3. PROCLAMAR como senadores")
    parte = re.sub(r"\s+", " ", texto[ini:texto.index("4. DISPONER", ini)])
    partidos = "|".join(re.escape(x) for x in PARTIDOS)
    patron = re.compile(r"(?P<unico>DISTRITO ELECTORAL ÚNICO NACIONAL)|(?P<multiple>DISTRITO ELECTORAL MÚLTIPLE)"
                        r"|\d+\) CIRCUNSCRIPCIÓN ELECTORAL DE (?P<circ>.+?) VOTOS PREFERENCIALES"
                        rf"|(?P<partido>{partidos})(?: \d+(?!\.))?"
                        r"|\d+\. (?P<nombre>[A-ZÁÉÍÓÚÑÜ][A-ZÁÉÍÓÚÑÜ ]*?) \d{3,}")
    filas, distrito, circ, partido = [], None, None, None
    for m in patron.finditer(parte):
        if m["unico"] or m["multiple"]:
            distrito, circ, partido = ("único" if m["unico"] else "múltiple"), None, None
        elif m["circ"]:
            circ, partido = norm(m["circ"]), None
        elif m["partido"]:
            partido = norm(m["partido"])
        elif m["nombre"] and partido:
            filas.append({"distrito": distrito, "circunscripcion": circ if distrito == "múltiple" else "NACIONAL",
                          "partido": partido, "sexo": sexo(m["nombre"])})
    d = pd.DataFrame(filas)
    por_dist = d.groupby("distrito").size().to_dict()
    mult = d[d.distrito == "múltiple"].groupby("circunscripcion").size()
    assert por_dist == {"único": 30, "múltiple": 30}, por_dist
    assert len(mult) == 27 and mult["LIMA METROPOLITANA"] == 4 and (mult.drop("LIMA METROPOLITANA") == 1).all(), mult
    tabla = d.pivot_table(index="partido", columns="sexo", aggfunc="size", fill_value=0)
    for p, (f, m) in TABLA3_SENADO.items():
        assert (tabla.loc[p, "F"], tabla.loc[p, "M"]) == (f, m), f"{p}: {tabla.loc[p].to_dict()} frente a {(f, m)}"
    sm = d[(d.distrito == "múltiple") & (d.sexo == "F")].groupby("circunscripcion").size().to_dict()
    assert sm == SENADORAS_MULTIPLE, sm
    print("senadores: 30 del distrito único y 30 del múltiple en 27 circunscripciones; sexo por nombre = Tabla 3 de la "
          "ONPE en los 6 partidos")
    return d


def gobernadores():
    """Mínimo y máximo de gobernadoras 2027–2030 por circunscripción, con la regla del 30 % y la segunda elección."""
    g = pd.read_csv(GOB)
    assert len(g) == 25 and g[["sexo_primera_formula", "sexo_segunda_formula"]].isin(["F", "M"]).all().all()
    primera = g.pct_validos_primera >= 30
    f1, f2 = g.sexo_primera_formula == "F", g.sexo_segunda_formula == "F"
    g["minimo"] = (primera & f1) | (~primera & f1 & f2)
    g["maximo"] = (primera & f1) | (~primera & (f1 | f2))
    par = g.sexo_primera_formula + g.sexo_segunda_formula
    g["estado"] = [("Elegido en primera vuelta" if s == "M" else "Elegida en primera vuelta") if pv else
                   {"MM": "Segunda elección entre dos fórmulas encabezadas por hombres",
                    "FF": "Segunda elección entre dos fórmulas encabezadas por mujeres"}.get(
                       pp, "Segunda elección entre una fórmula encabezada por una mujer y otra por un hombre")
                   for pv, s, pp in zip(primera, g.sexo_primera_formula, par)]
    print(f"gobernadores 2026: {int(primera.sum())} en primera vuelta y {int((~primera).sum())} en segunda elección; "
          f"gobernadoras entre {int(g.minimo.sum())} y {int(g.maximo.sum())}")
    return g


def ambito_circ(c):
    return {"LIMA": "LM", "LIMA METROPOLITANA": "LM", "LIMA PROVINCIAS": "RL", "CALLAO": "07"}.get(c) or (
        None if "EXTRANJERO" in c else svy.codigo_departamento(c))


def main():
    d = diputados()
    d["ambito"] = d.circunscripcion.map(ambito_circ)
    d = d[d.ambito.notna()]
    dip = d.pivot_table(index="ambito", columns="sexo", aggfunc="size", fill_value=0).reset_index()
    dip = pd.DataFrame({"cargo": "Congreso de la República", "cuadro": "JNE 1411-2026", "periodo": "2026/2031",
                        "ambito": dip.ambito, "mujeres": dip.F, "hombres": dip.M,
                        "fuente": "JNE, Resolución N.° 1411-2026-JNE (proclamación de diputados), sexo por nombre "
                                  "validado con la ONPE (Analítica Electoral 3, Tabla 3)",
                        "paginas": "", "nota": "Cámara de Diputados (130 escaños por circunscripción; los 2 de "
                                              "peruanos en el extranjero no se mapean)"})
    s = senadores()
    s = s[s.distrito == "múltiple"].copy()
    s["ambito"] = s.circunscripcion.map(ambito_circ)
    s = s[s.ambito.notna()]
    sen = s.pivot_table(index="ambito", columns="sexo", aggfunc="size", fill_value=0).reset_index()
    sen = pd.DataFrame({"cargo": "Senado, distrito múltiple", "cuadro": "JNE 1412-2026", "periodo": "2026/2031",
                        "ambito": sen.ambito, "mujeres": sen.F, "hombres": sen.M,
                        "fuente": "JNE, Resolución N.° 1412-2026-JNE (proclamación de senadores), sexo por nombre "
                                  "validado con la ONPE (Analítica Electoral 3, Tabla 3)",
                        "paginas": "", "nota": "30 escaños del distrito múltiple (Lima Metropolitana 4; el de peruanos "
                                              "en el extranjero no se mapea). Los 30 del distrito único nacional no "
                                              "tienen ámbito"})
    g = gobernadores()
    g["ambito"] = g.ambito_onpe.map(lambda x: {"LIMA": "RL", "CALLAO": "07"}.get(x) or svy.codigo_departamento(x))
    assert g.ambito.notna().all()
    gob = pd.DataFrame({"cargo": "Gobiernos regionales", "cuadro": "ONPE ERM 2026", "periodo": "2027/2030",
                        "ambito": g.ambito, "mujeres": g.minimo.astype(int), "hombres": (~g.maximo).astype(int),
                        "fuente": "ONPE, resultados de las ERM 2026 (elección regional), dos fórmulas más votadas; "
                                  "sexo por nombre",
                        "paginas": "", "nota": g.estado + ". " + NOTA_2026})
    e = pd.read_csv(ERM)
    e["ambito"] = e.ambito_onpe.map(lambda x: x if x in ("LM", "RL") else ("07" if x == "CALLAO" else
                                                                            svy.codigo_departamento(x)))
    a = e.groupby("ambito", as_index=False)[["mujeres", "hombres", "ambiguo", "sin_nombre"]].sum()
    tot = a[["mujeres", "hombres", "ambiguo", "sin_nombre"]].sum()
    print(f"alcaldías 2026: {int(tot.sum())} ganadores; {int(tot.mujeres)} mujeres, {int(tot.hombres)} hombres, "
          f"{int(tot.ambiguo + tot.sin_nombre)} no determinados")
    alc = pd.DataFrame({"cargo": "Alcaldías (provinciales y distritales)", "cuadro": "ONPE ERM 2026",
                        "periodo": "2027/2030", "ambito": a.ambito, "mujeres": a.mujeres, "hombres": a.hombres,
                        "fuente": "ONPE, resultados de las ERM 2026, ganador por provincia y distrito; sexo por "
                                  "nombre, con los casos dudosos revisados a mano",
                        "paginas": "",
                        "nota": [(f"Sin determinar: {int(x)}. " if x else "") + NOTA_2026
                                 for x in a.ambiguo + a.sin_nombre]})
    out = pd.concat([dip, sen, alc, gob], ignore_index=True)
    out.to_csv(SALIDA, index=False, encoding="utf-8")
    print(out.groupby("cargo")[["mujeres", "hombres"]].sum())


if __name__ == "__main__":
    main()
