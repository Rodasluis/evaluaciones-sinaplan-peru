"""REV_IND.4: valor 2025 de los dos indicadores del OP3 sin valor oficial (IOP.03.02 e IOP.03.01).

IOP.03.02, porcentaje de servidoras civiles funcionarias y directivas. SERVIR lo mide con la ENAHO. La definición que
reproduce su serie es: residentes habituales ((p204 = 1 y p205 = 2) o (p204 = 2 y p206 = 1)), ocupados (ocu500 = 1) que
trabajan para la administración pública (p510 = 2) y cuya ocupación principal está en el gran grupo 1 de la CIUO-88
(p505 de 100 a 199), con el factor fac500a. Regla de aceptación, fijada antes de calcular: el método se acepta si
reproduce a ±0,5 puntos los valores que SERVIR publicó con la ENAHO de 2016 a 2019 y de 2021 (SERVIR, 2023, Gráfico
28, p. 31; la misma serie del Gráfico 12, p. 17). SERVIR (2025 y 2026) publica 2023 y 2024 con una «metodología
actualizada» (INEI, s. f.) que los microdatos públicos no permiten reproducir (traen el factor con proyecciones del
CPV-2007); esas diferencias se informan sin forzarlas. Estimador con diseño muestral complejo (svy.py).

IOP.03.01, índice de empoderamiento político. Método del Acta N.° 04-2022 (no ratificada), que es el que usan los
reportes desde 2022: promedio de cuatro cocientes mujeres/hombres: A parlamentarias electas, B regidoras electas,
C alcaldesas electas y D candidatas inscritas en las elecciones regionales y municipales. Cada año toma la composición
de las últimas elecciones (el Congreso elegido y las ERM cuyo periodo está en curso), por eso la serie oficial repite
el valor entre elecciones. Fuentes: tablas del acta (pp. 1 a 4) para el Congreso 2016, 2020 y 2021 y la ERM 2018;
INEI, Brechas de Género 2025, cuadros 4.3 y 4.4 (participacion_politica_jne.csv) para B y C de las ERM 2018 y 2022;
JNE, Perfil Electoral N.° 4 (2022), p. 5 del PDF, para D de la ERM 2022 (43,0 % de candidaturas inscritas son
mujeres; el JNE no publica el conteo). El cálculo debe reproducir los valores oficiales de 2019 a 2022 a cuatro
decimales y el de 2023 (0,58) a dos. Las revocatorias y vacancias, que el acta menciona como supuesto, no se
incorporan.

Planos: las fichas y las metas son de la PNIG; la reproducción, la regla de aceptación y los valores 2024 y 2025 son
cálculo del equipo evaluador y no sustituyen al reporte oficial.

Uso (desde la raíz del repositorio):
    uv run --no-project --with pandas --with pyreadstat --with openpyxl python -I -X utf8 src/analisis_pnig/rev_ind_4_recalculo_op3.py
Salida: evaluaciones/pnig_2026/04_analisis/REV_IND_4_recalculo_OP3.xlsx
"""
import os
import sys

import numpy as np
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import d1_brechas_enaho as E  # noqa: E402
import svy  # noqa: E402
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

A = f"{EVAL}/04_analisis"
SALIDA = f"{A}/REV_IND_4_recalculo_OP3.xlsx"
PSI = f"{A}/REV_IND_base_indicadores.xlsx"
JNE = f"{A}/d1_brechas/participacion_politica_jne.csv"
ANIOS = range(2016, 2026)
TOLERANCIA = 0.5
# Serie de SERVIR con la ENAHO: año: (valor, fuente, mismo método que 2016–2021)
SERVIR = {
    **{a: (v, "SERVIR (2023), Gráfico 28, p. 31", True)
       for a, v in {2016: 32.9, 2017: 34.3, 2018: 28.5, 2019: 40.0, 2021: 34.4}.items()},
    2023: (39.0, "SERVIR (2025), Gráfico 10, p. 19; metodología actualizada aplicada al periodo 2023", False),
    2024: (44.1, "SERVIR (2026), Gráfico 12, p. 22; metodología actualizada aplicada al periodo 2024", False),
}
ACTA = "Acta N.° 04-2022 (no ratificada), anexo de precisiones"
# Composiciones electorales: clave: (mujeres, hombres, fuente). D de la ERM 2022 en porcentaje (el JNE no da conteos).
CONGRESO = {
    2016: (37, 93, f"{ACTA}, componente A, p. 1"),
    2020: (34, 96, f"{ACTA}, componente A, p. 1"),
    2021: (49, 81, f"{ACTA}, componente A, p. 1 (electas al proclamarse; ONPE, 37,7 %)"),
}
CANDIDATAS = {
    2018: (41849, 63893, f"{ACTA}, componente D, p. 4"),
    2022: (43.0, 57.0, "JNE, Perfil Electoral N.° 4 (2022), p. 5 del PDF: 43,0 % de candidaturas inscritas"),
}
ACTA_BC_2018 = {"B": (3186, 7468), "C": (90, 1772)}  # componentes B y C de la ERM 2018 en el acta (pp. 2 y 3)
VIGENTE = {  # año: (Congreso elegido, ERM con periodo en curso)
    2019: (2016, 2018), 2020: (2020, 2018), 2021: (2021, 2018), 2022: (2021, 2018),
    2023: (2021, 2022), 2024: (2021, 2022), 2025: (2021, 2022),
}


def serie_servidoras():
    filas = []
    for a in ANIOS:
        df = E.leer(a, "500", E.LLAVE + ["ubigeo", "estrato", "p204", "p205", "p206", "p207", "ocu500", "p505", "p510",
                                         "fac500a"])
        if df is None:
            raise SystemExit(f"falta la ENAHO {a}, módulo 500")
        res = ((df.p204 == 1) & (df.p205 == 2)) | ((df.p204 == 2) & (df.p206 == 1))
        p505 = pd.to_numeric(df.p505, errors="coerce")
        assert p505.dropna().between(1, 999).all(), f"{a}: p505 fuera del rango de la CIUO-88 de tres dígitos"
        univ = (res & (df.ocu500 == 1) & (df.p510 == 2) & p505.between(100, 199)).to_numpy()
        d = svy.Diseno(df, peso="fac500a", estrato="estrato", psu="conglome")
        r = svy.estimar(d, (df.p207 == 2).astype(float), universo=univ, escala=100)
        v, fuente, comparable = SERVIR.get(a, (np.nan, "", None))
        filas.append({"Año": a, "Estimación (%)": r["estimacion"], "EE": r["ee"], "IC inferior": r["ic_inf"],
                      "IC superior": r["ic_sup"], "CV (%)": r["cv"], "n": r["n"], "SERVIR (%)": v,
                      "Diferencia": r["estimacion"] - v if pd.notna(v) else np.nan, "Fuente SERVIR": fuente,
                      "Mismo método que SERVIR 2016–2021": {True: "Sí", False: "No (metodología actualizada)",
                                                            None: ""}[comparable]})
    t = pd.DataFrame(filas)
    val = t[t["Mismo método que SERVIR 2016–2021"] == "Sí"]
    t["Reproduce"] = np.where(t["Mismo método que SERVIR 2016–2021"] == "Sí",
                              np.where(t.Diferencia.abs() <= TOLERANCIA, "Sí", "No"), "")
    acepta = bool((val.Diferencia.abs() <= TOLERANCIA).all())
    return t, acepta


def indice_politico(base):
    j = pd.read_csv(JNE)

    def cuenta(cargo, periodo):
        r = j[(j.cargo == cargo) & (j.periodo == periodo)].iloc[0]
        return int(r.mujeres), int(r.hombres), f"INEI, Brechas de Género 2025, {r.cuadro_o_grafico}, p. {r.pagina}"

    erm = {2018: {"B": cuenta("Regidurías", "2019/2022"), "C": cuenta("Alcaldías (provinciales y distritales)", "2019/2022")},
           2022: {"B": cuenta("Regidurías", "2023/2026"), "C": cuenta("Alcaldías (provinciales y distritales)", "2023/2026")}}
    for k, v in ACTA_BC_2018.items():  # el acta y el INEI coinciden en la ERM 2018
        assert erm[2018][k][:2] == v, f"componente {k} de la ERM 2018 distinto entre el acta y el INEI"
    filas = []
    for anio, (cong, e) in VIGENTE.items():
        comp = {"A": CONGRESO[cong], "B": erm[e]["B"], "C": erm[e]["C"], "D": CANDIDATAS[e]}
        fila = {"Año": anio, "Congreso": cong, "ERM": e}
        for k, (m, h, f) in comp.items():
            fila |= {f"{k} mujeres": m, f"{k} hombres": h, f"{k} cociente": m / h, f"{k} fuente": f}
        fila["Índice"] = np.mean([fila[f"{k} cociente"] for k in "ABCD"])
        oficial = pd.to_numeric(base.loc["IOP.03.01", f"Valor {anio}"], errors="coerce")
        meta = pd.to_numeric(base.loc["IOP.03.01", f"Meta {anio}"], errors="coerce")
        fila |= {"Valor oficial": oficial, "Diferencia": fila["Índice"] - oficial if pd.notna(oficial) else np.nan,
                 "Meta": meta, "Avance (%)": 100 * fila["Índice"] / meta}
        filas.append(fila)
    t = pd.DataFrame(filas)
    o = t.set_index("Año")
    for a in (2019, 2020, 2021, 2022):
        assert abs(o.loc[a, "Diferencia"]) < 5e-5, f"IOP.03.01 {a}: no reproduce el valor oficial a cuatro decimales"
    assert abs(o.loc[2023, "Diferencia"]) < 5e-3, "IOP.03.01 2023: no reproduce 0,58 a dos decimales"
    assert o.loc[[2024, 2025], "Valor oficial"].isna().all(), "IOP.03.01 ya tiene valor oficial en 2024 o 2025"
    return t


def main():
    base = pd.read_excel(PSI, sheet_name="Base REV_IND.1").set_index("Código")
    serie, acepta = serie_servidoras()
    print(serie[["Año", "Estimación (%)", "IC inferior", "IC superior", "n", "SERVIR (%)", "Diferencia",
                 "Reproduce"]].round(2).to_string(index=False))
    print("IOP.03.02: método aceptado" if acepta else "IOP.03.02: el método NO cumple la regla; queda como aproximación")
    ind = indice_politico(base)
    print(ind[["Año", "Congreso", "ERM", "A cociente", "B cociente", "C cociente", "D cociente", "Índice",
               "Valor oficial", "Meta", "Avance (%)"]].round(4).to_string(index=False))

    s25 = serie.set_index("Año").loc[2025]
    i25 = ind.set_index("Año").loc[2025]
    m02 = float(base.loc["IOP.03.02", "Meta 2025"])
    res = pd.DataFrame([
        {"Código": "IOP.03.01", "Estimación": i25["Índice"], "IC inferior": i25["Índice"],
         "IC superior": i25["Índice"], "Meta 2025": i25["Meta"], "Avance (%)": i25["Avance (%)"],
         "Método": "registros electorales con el método del acta", "Tipo": "Recalculado",
         "Sustento": "Reproduce los valores oficiales de 2019 a 2023; sin elecciones nacionales ni regionales entre "
                     "2023 y 2025, el valor es el de 2023"},
        {"Código": "IOP.03.02", "Estimación": s25["Estimación (%)"], "IC inferior": s25["IC inferior"],
         "IC superior": s25["IC superior"], "Meta 2025": m02, "Avance (%)": 100 * s25["Estimación (%)"] / m02,
         "Método": "ENAHO 2025 con el método que reproduce la serie de SERVIR",
         "Tipo": "Estimado" if acepta else "Aproximación",
         "Sustento": f"n = {int(s25['n'])}; reproduce SERVIR 2016–2019 y 2021 a ±{TOLERANCIA} puntos"
                     if acepta else "El método no cumple la regla de aceptación"},
    ])
    criterios = pd.DataFrame([
        ["IOP.03.02", "Definición", "Residentes habituales ocupados de la administración pública (p510 = 2) con "
         "ocupación principal en el gran grupo 1 de la CIUO-88 (p505 de 100 a 199); fac500a; estrato y conglomerado.",
         "Equipo evaluador"],
        ["IOP.03.02", "Regla de aceptación", f"Diferencia con SERVIR de ±{TOLERANCIA} puntos o menos en 2016–2019 y "
         "2021, años publicados con el mismo método.", "Equipo evaluador"],
        ["IOP.03.02", "SERVIR 2023 y 2024", "Publicados con una metodología actualizada (INEI, s. f.) que los "
         "microdatos públicos no permiten reproducir; se informan las diferencias.", "SERVIR (2025, 2026)"],
        ["IOP.03.01", "Método", "Promedio de los cocientes mujeres/hombres de parlamentarias electas, regidoras "
         "electas, alcaldesas electas y candidatas inscritas en las ERM.", "Acta N.° 04-2022 (no ratificada)"],
        ["IOP.03.01", "Año", "Composición de las últimas elecciones: Congreso elegido y ERM con periodo en curso; "
         "sin revocatorias ni vacancias.", "Equipo evaluador, según la serie oficial"],
        ["IOP.03.01", "Reproducción", "Valores oficiales de 2019 a 2022 a cuatro decimales y 0,58 en 2023 a dos.",
         "Equipo evaluador"],
        ["Ambos", "Nivel", "El nivel frente a la meta se asigna en sintesis_hallazgos.py con la Tabla 9 de la nota.",
         "Nota metodológica, Tabla 9"],
        ["Ambos", "Plano", "Cálculo del equipo evaluador; no sustituye al reporte oficial.", "Equipo evaluador"],
    ], columns=["Indicador", "Elemento", "Regla", "Plano"])
    with pd.ExcelWriter(SALIDA) as w:
        res.to_excel(w, sheet_name="Resultado 2025", index=False)
        serie.to_excel(w, sheet_name="IOP.03.02 serie", index=False)
        ind.to_excel(w, sheet_name="IOP.03.01 componentes", index=False)
        criterios.to_excel(w, sheet_name="Criterios", index=False)
    print(res[["Código", "Estimación", "IC inferior", "IC superior", "Meta 2025", "Avance (%)", "Tipo"]].round(3)
          .to_string(index=False))
    print("Guardado:", SALIDA)


if __name__ == "__main__":
    main()
