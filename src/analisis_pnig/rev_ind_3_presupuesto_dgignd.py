"""REV_IND.3: ejecución presupuestal de la actividad 5001072 «Promoción de la mujer» remitida por la DGIGND (2019–2026),
conciliada con los datos abiertos del gasto del MEF que usa el análisis de la conducción (ronda 34, pedido del usuario).

Fuente: docs/fuentes/pnig/10_presupuesto/2026_mimp-dgignd_ejecucion-presupuestal-2019-2026_actividad-5001072.xlsx
(PIM y ejecución por finalidad, genérica y específica de gasto; sin datos personales).

Cada finalidad se asigna al órgano cuya meta la ejecuta, con las mismas agrupaciones que la hoja «Conducción por órgano»
de REV_IND_3_verificacion_presupuestal.xlsx (datos abiertos del MEF por meta). La conciliación compara, por órgano y año,
la ejecución del archivo de la DGIGND con el devengado del MEF; la composición reparte la ejecución de cada órgano de la
conducción entre personal y obligaciones sociales, bienes y servicios, y otros gastos. El año 2026 solo es contexto: está
fuera del periodo evaluado y su ejecución es parcial.

Uso (desde la raíz del repositorio):
    uv run --no-project --with openpyxl --with pandas python src/analisis_pnig/rev_ind_3_presupuesto_dgignd.py
Salida: evaluaciones/pnig_2026/04_analisis/REV_IND_3_presupuesto_dgignd.xlsx
"""
import pandas as pd
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
from nucleo.config import insumo  # noqa: E402
from nucleo.presupuesto import conciliar  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

ORIGEN = "docs/fuentes/pnig/10_presupuesto/2026_mimp-dgignd_ejecucion-presupuestal-2019-2026_actividad-5001072.xlsx"
MEF = f"{EVAL}/04_analisis/REV_IND_3_verificacion_presupuestal.xlsx"
SALIDA = f"{EVAL}/04_analisis/REV_IND_3_presupuesto_dgignd.xlsx"
ANIOS = range(2019, 2027)
# finalidad (código de 7 dígitos) -> órgano, con las agrupaciones de «Conducción por órgano» (MEF)
_PRE = insumo("presupuesto_finalidades.yaml")   # evaluaciones/<id>/02_insumos (Fase 2)
ORGANO = _PRE["organo_por_finalidad"]
FILA_MEF = _PRE["fila_mef"]
CONDUCCION = tuple(_PRE["conduccion"])


def leer():
    v = pd.read_excel(ORIGEN, header=None)
    anios, campos = v.iloc[0], v.iloc[1]
    cols = [str(campos[i]).strip() if i < 12 else
            f"{'PIM' if str(campos[i]).startswith('PIM') else 'Ejecución'} {int(anios[i])}" for i in range(v.shape[1])]
    d = v.iloc[2:].copy()
    d.columns = cols
    d = d[d["activ_obra_accinv"].astype(str).str.startswith("5001072")].copy()
    num = [c for c in cols if c.startswith(("PIM ", "Ejecución "))]
    d[num] = d[num].apply(pd.to_numeric, errors="coerce").fillna(0)
    d["Finalidad"] = d.finalidad.astype(str).str.strip()
    d["Código"] = d.Finalidad.str[:7]
    d["Órgano"] = d["Código"].map(ORGANO).fillna("Otros órganos (finalidades de 2026)")
    gen = d["GENERICA\nDETALLE"].astype(str) if "GENERICA\nDETALLE" in d else d.iloc[:, 10].astype(str)
    esp = d["ESPECIFICA DE DETALLE"].astype(str)
    # el CAS pasó en 2024 de bienes y servicios (2.3.2.8) a personal (2.1.1.13): se agrupa con el personal para que el
    # cambio de clasificador no parezca un cambio del gasto
    d["Genérica"] = [("Personal y CAS" if g.startswith("2.1") or "CONTRATO ADMINISTRATIVO DE SERVICIOS" in e.upper() else
                      "Locación de servicios de personas naturales" if e.replace(" ", "").startswith("2.3.29.1") else
                      "Otros bienes y servicios" if g.startswith("2.3") else "Otros gastos") for g, e in zip(gen, esp)]
    return d, num


def main():
    d, num = leer()
    por_fin = d.groupby(["Código", "Finalidad", "Órgano"])[num].sum().reset_index()
    # conciliación con el MEF (devengado por órgano y año)
    mef = pd.read_excel(MEF, sheet_name="Conducción por órgano").set_index("Órgano")
    conc = conciliar(
        (({"Órgano": org, "Año": a}, d.loc[d["Órgano"] == org, f"Ejecución {a}"].sum(),
          mef.loc[fila_mef, f"{a} Devengado"] if fila_mef in mef.index else float("nan"))
         for org, fila_mef in FILA_MEF.items() for a in range(2019, 2026)),
        "Ejecución (DGIGND, S/)", "Devengado (MEF, S/)")
    # composición del gasto de la conducción
    comp = (d[d["Órgano"].isin(CONDUCCION)].groupby("Genérica")[[f"Ejecución {a}" for a in ANIOS]].sum())
    comp_pct = (100 * comp / comp.sum()).round(1)
    tot = d.groupby("Órgano")[num].sum()
    criterios = pd.DataFrame([
        ["Fuente", ORIGEN + " (remitida por la DGIGND; copia registrada en "
                            "docs/fuentes/movimientos_dgignd-conaig_2026-11-10.tsv)"],
        ["Órgano de cada finalidad", "; ".join(f"{k}: {v}" for k, v in ORGANO.items())],
        ["Conciliación", "Ejecución del archivo de la DGIGND frente al devengado de la hoja «Conducción por órgano» de "
                         "REV_IND_3_verificacion_presupuestal.xlsx (datos abiertos del MEF por meta), 2019–2025; "
                         "coincide si la diferencia es menor de S/ 1."],
        ["Composición", "Personal y CAS: genérica 2.1 y contrato administrativo de servicios (que pasó de la "
                        "específica 2.3.2.8 a la 2.1.1.13 en 2024); locación de servicios de personas naturales: "
                        "2.3.2.9.1; otros bienes y servicios: resto de la genérica 2.3; otros gastos: las demás. "
                        "Conducción: DGIGND, DPIGND y DPPDM."],
        ["2026", "Fuera del periodo evaluado y con ejecución parcial: solo contexto."],
    ], columns=["Elemento", "Criterio"])
    with pd.ExcelWriter(SALIDA, engine="openpyxl") as xw:
        por_fin.to_excel(xw, sheet_name="Por finalidad", index=False)
        tot.reset_index().to_excel(xw, sheet_name="Por órgano", index=False)
        conc.to_excel(xw, sheet_name="Conciliación MEF", index=False)
        comp.reset_index().to_excel(xw, sheet_name="Composición conducción", index=False)
        comp_pct.reset_index().to_excel(xw, sheet_name="Composición (%)", index=False)
        criterios.to_excel(xw, sheet_name="Criterios", index=False)
    print(f"{len(d)} filas; conciliación: {int((conc.Coincide == 'Sí').sum())} de {len(conc)} órgano-año coinciden")
    print(conc[conc.Coincide == "No"].to_string() if (conc.Coincide == "No").any() else "sin diferencias")
    print(comp_pct.to_string())


if __name__ == "__main__":
    main()
