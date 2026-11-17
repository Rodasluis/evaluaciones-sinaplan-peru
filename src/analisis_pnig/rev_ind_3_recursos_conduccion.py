"""REV_IND.3: recursos presupuestales del MIMP vinculados a la conducción de la PNIG, 2019–2025 (pregunta I3).

Fuente: Consulta Amigable del MEF, pliego 039 (MIMP), descargada con src/fuentes/descargar_consulta_amigable.py en
03_datos/raw/consulta_amigable/ (registro: docs/fuentes/descargas_consulta-amigable_2026-10-29.tsv). Totales por unidad
ejecutora y, para la 001 (Administración Nivel Central, donde están la DGIGND y la OGMEPGD), detalle por categoría
presupuestal, producto y actividad.

Planos:
  Nota metodológica (REV_IND.3): las actividades vinculadas a la conducción de la Política se identifican por su
    denominación y se contrastan con el POI del MIMP; el criterio se declara.
  Equipo evaluador: criterio de vinculación (VINCULACION). La Consulta Amigable no muestra el gasto por órgano: la
    actividad 5001072 «Promoción de la mujer» agrupa la promoción de la mujer en la Administración Central. La parte de
    cada órgano (DGIGND, DGTEG/DTEG, DGCVG) se separa por meta con los datos abiertos del MEF en
    rev_ind_3_verificacion_presupuestal.py (hoja «Conducción por meta»).
  Alcance: descriptivo (pregunta I3, eficiencia de alcance acotado); no se relaciona el gasto con resultados.

Servicios del OP1: las categorías 1002 (productos específicos para reducción de la violencia contra la mujer, 2021–2025) y
0080 (lucha contra la violencia familiar, 2019–2023) se descargan por entidad, producto y actividad con
src/fuentes/descargar_consulta_amigable_categorias.py. CORRESPONDENCIA asigna cada actividad del Gobierno Nacional a un
servicio de la PNIG por la entidad y la denominación (equipo evaluador): directa, probable o sin servicio equivalente.

Uso (desde la raíz del repositorio):
    uv run --no-project --with openpyxl --with pandas python src/analisis_pnig/rev_ind_3_recursos_conduccion.py
"""
import glob

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

RAIZ = EVAL
ENTRADA = f"{RAIZ}/03_datos/raw/consulta_amigable/*_mef_consulta-amigable_pliego-039.csv"
SALIDA = f"{RAIZ}/04_analisis/REV_IND_3_recursos_conduccion.xlsx"
POI = "docs/fuentes/planeamiento/poi_mimp/"

# Vinculación con la conducción de la PNIG (equipo evaluador): (categoría, actividad o None para toda la categoría)
VINCULACION = {
    ("9002", "5001072"): ("Directa por denominación",
                          "«Promoción de la mujer» corresponde a la finalidad de la PNIG. Agrupa la promoción de la mujer de "
                          "la Administración Central; la parte de cada órgano se separa por meta con los datos abiertos del "
                          "MEF."),
    ("1002", None): ("Indirecta: OP1 (violencia)",
                     "Productos específicos para la reducción de la violencia contra la mujer (programa presupuestal "
                     "orientado a resultados); se vincula con el objetivo prioritario 1, no con la conducción."),
    ("0080", None): ("Indirecta: OP1 (violencia)",
                     "Programa presupuestal 0080 Lucha contra la violencia familiar, ejecutado en la Administración Central "
                     "en 2019 y 2020; se vincula con el objetivo prioritario 1, no con la conducción."),
    ("9001", None): ("Soporte institucional",
                     "Acciones centrales (planeamiento y presupuesto, conducción y orientación superior, gestión "
                     "administrativa): sostienen a todos los órganos, incluida la conducción, sin asignación atribuible."),
}
# Contenido de los documentos del POI descargados (revisión del equipo, 2026-10-29)
POI_DOCS = [
    (2019, "2019_mimp_poi-anual-2019.pdf", "POI anual 2019 (Anexo B-3, formulado en mayo de 2018)",
     "Escaneo con OCR desde el 2026-10-29. Aprobado por R.M. N.° 173-2018-MIMP (número indicado por el usuario; poco "
     "legible en el escaneo), antes del PIA. Centros de costo 03.02 DGIGND (S/ 3 229 240) y 03.03 DGTEG (S/ 2 466 404), "
     "p. 19."),
    (2020, "2020_mimp_poi-anual-2020_modificado-version-2.pdf", "POI anual 2020 modificado, versión 2",
     "Programación física y financiera por centro de costo, pero sin bloque de la DGIGND en esta versión."),
    (2021, "2021_mimp_poi-anual-2021_modificado-version-1.pdf", "POI anual 2021 modificado, versión 1",
     "Centro de costo 03.02 DGIGND con la actividad operativa AOI00108700017 y su programación financiera mensual."),
    (2022, "2022_mimp_poi-anual-2022_modificado-version-1.pdf", "POI anual 2022 modificado, versión 1",
     "Centro de costo 03.02 DGIGND con nueve actividades operativas (AOI00108700017 y AOI00108700406 a 413) y su "
     "programación financiera."),
    (2023, "2023_mimp_poi-anual-2023_modificado-version-2.pdf", "POI anual 2023 modificado, versión 2",
     "Centros de costo DGIGND (AOI00108700525) y DPIGND (AOI00108700527 a 529) con programación financiera."),
    (2024, "2024_mimp_poi-anual-2024_actualizado-version-3.pdf", "POI anual 2024 actualizado, versión 3",
     "DGIGND (AOI00108700577) y DPIGND (AOI00108700527, 686 a 689). El documento trae dos montos distintos para la "
     "AOI00108700577 (resumen y detalle): requiere conciliación."),
    (2025, "2025_mimp_poi-anual-2025_actualizado-version-3.pdf", "POI anual 2025 actualizado, versión 3",
     "DPIGND (AOI00108700527, 676 y 677) y DGIGND (AOI00108700577) con programación financiera mensual."),
]
MONTOS = ["pia", "pim", "certificacion", "compromiso_anual", "devengado", "girado"]
CATEGORIAS = f"{RAIZ}/03_datos/raw/consulta_amigable/categorias/*_mef_consulta-amigable_categoria-*.csv"
ACTIVIDAD = f"{RAIZ}/03_datos/raw/consulta_amigable/*_mef_consulta-amigable_pliego-039_actividad-5001072.csv"
# Correspondencia con los servicios de la PNIG (equipo evaluador): (categoría, pliego, actividad) → (servicio, grado)
CORRESPONDENCIA = {
    ("1002", "004", "5006349"): ("1.1.5 Otorgamiento de medidas de protección (PJ)", "Directa"),
    ("1002", "022", "5006349"): ("1.1.6 Investigación en fiscalías especializadas (MPFN)", "Directa"),
    ("1002", "006", "5006349"): ("1.1.7 Defensa pública especializada (MINJUSDH)", "Directa"),
    ("1002", "007", "5006349"): ("1.1.4 Registro y seguimiento de medidas de protección (MININTER)", "Directa"),
    ("1002", "007", "5006350"): ("1.1.4 Registro y seguimiento de medidas de protección (MININTER)", "Probable"),
    ("1002", "010", "5006342"): ("1.2.4 Convivencia escolar y prevención de la violencia (MINEDU)", "Directa"),
    ("1002", "010", "5006352"): ("1.2.4 Convivencia escolar y prevención de la violencia (MINEDU)", "Directa"),
    ("1002", "011", "5006375"): ("1.1.1 Atención a niñas, niños y adolescentes víctimas: MAMIS (MINSA)", "Probable"),
    ("1002", "011", "5006383"): ("1.1.1 Atención a niñas, niños y adolescentes víctimas: MAMIS (MINSA)", "Probable"),
    ("1002", "039", "5006357"): ("1.1.3 Atención integral en los CEM (AURORA)", "Directa"),
    ("1002", "039", "5006343"): ("1.2.2 Redes de hombres (AURORA)", "Directa"),
    ("1002", "039", "5006344"): ("1.2.2 Redes de hombres (AURORA)", "Directa"),
    ("1002", "039", "5006348"): ("1.2.1 Estrategia comunitaria de prevención (AURORA)", "Probable"),
    ("0080", "039", "5006193"): ("1.1.3 Atención integral en los CEM (AURORA)", "Directa"),
    ("0080", "039", "5003452"): ("1.2.3 Información para la prevención con enfoque intercultural (AURORA)", "Probable"),
    ("0080", "039", "5004136"): ("1.2.3 Información para la prevención con enfoque intercultural (AURORA)", "Probable"),
    ("0080", "039", "5003455"): ("1.2.2 Redes de hombres (AURORA)", "Directa"),
    ("0080", "039", "5006195"): ("1.2.1 Estrategia comunitaria de prevención (AURORA)", "Probable"),
    ("0080", "004", "5006257"): ("1.1.5 Otorgamiento de medidas de protección (PJ)", "Directa"),
    ("0080", "022", "5006257"): ("1.1.6 Investigación en fiscalías especializadas (MPFN)", "Directa"),
    ("0080", "007", "5006257"): ("1.1.4 Registro y seguimiento de medidas de protección (MININTER)", "Directa"),
    ("0080", "006", "5006258"): ("1.1.7 Defensa pública especializada (MINJUSDH)", "Directa"),
    ("0080", "010", "5006259"): ("1.2.4 Convivencia escolar y prevención de la violencia (MINEDU)", "Directa"),
    ("0080", "011", "5006260"): ("1.1.2 Salud mental comunitaria: CSMC (MINSA)", "Directa"),
    ("0080", "011", "5006262"): ("1.1.1 Atención a niñas, niños y adolescentes víctimas: MAMIS (MINSA)", "Directa"),
}


def vinculacion(cat, act):
    v = VINCULACION.get((cat, act)) or VINCULACION.get((cat, None))
    return v if v else ("Sin vinculación", "")


d = pd.concat([pd.read_csv(f) for f in sorted(glob.glob(ENTRADA))], ignore_index=True)
ue = d[d.nivel == "unidad_ejecutora"].copy()
act = d[d.nivel == "actividad"].copy()
act["cod_cat"] = act["categoria"].str[:4]
act["cod_act"] = act["actividad"].str.split(":").str[0].str.strip()
act[["vinculacion", "sustento"]] = act.apply(lambda r: pd.Series(vinculacion(r.cod_cat, r.cod_act)), axis=1)


def ejecucion(df):
    df["ejecucion_%"] = (df["devengado"] / df["pim"] * 100).round(1).where(df["pim"] > 0)
    return df


pliego = ejecucion(ue.groupby(["anio", "unidad_ejecutora"], as_index=False)[MONTOS].sum())
total = ejecucion(ue.groupby("anio", as_index=False)[MONTOS].sum())
total.insert(1, "unidad_ejecutora", "Pliego 039: total")
pliego = pd.concat([pliego, total]).sort_values(["anio", "unidad_ejecutora"])
por_cat = ejecucion(act.groupby(["anio", "categoria"], as_index=False)[MONTOS].sum())
detalle = ejecucion(act[["anio", "categoria", "producto", "actividad", "vinculacion", "sustento"] + MONTOS].copy())
resumen = ejecucion(act.groupby(["vinculacion", "anio"], as_index=False)[MONTOS].sum())
ue001 = ue[ue.unidad_ejecutora.str.startswith("001")].set_index("anio")["pim"]
resumen["participacion_pim_ue001_%"] = (resumen["pim"] / resumen["anio"].map(ue001) * 100).round(2)
directa = resumen[resumen.vinculacion == "Directa por denominación"][["anio", "pim", "devengado", "ejecucion_%",
                                                                         "participacion_pim_ue001_%"]]
# Servicios del OP1 en las categorías 1002 y 0080
cat = pd.concat([pd.read_csv(f, dtype={"categoria": str}) for f in sorted(glob.glob(CATEGORIAS))], ignore_index=True)
cat["cod_pliego"] = cat["pliego"].fillna("").str.split(":").str[0].str.strip()
cat["cod_act"] = cat["actividad"].str.split(":").str[0].str.strip()
cat[["servicio_pnig", "grado"]] = cat.apply(
    lambda r: pd.Series(CORRESPONDENCIA.get((r.categoria, r.cod_pliego, r.cod_act),
                                            ("Sin servicio PNIG equivalente" if r.nivel == "Gobierno Nacional"
                                             else "Gobiernos subnacionales (sin desagregar)", ""))), axis=1)
cat_det = ejecucion(cat[["anio", "categoria", "nivel", "pliego", "producto", "actividad", "servicio_pnig", "grado"]
                        + MONTOS].copy())
servicios = ejecucion(cat.groupby(["servicio_pnig", "anio"], as_index=False)[MONTOS].sum())
servicios_pim = servicios.pivot_table(index="servicio_pnig", columns="anio", values="pim", aggfunc="sum").reset_index()
# Conducción: actividad 5001072 por fuente y genérica
archivos_5001072 = sorted(glob.glob(ACTIVIDAD))
detalle_5001072 = ejecucion(pd.concat([pd.read_csv(f) for f in archivos_5001072], ignore_index=True)
                            if archivos_5001072 else pd.DataFrame(columns=["anio", "dimension", "codigo", "denominacion"]
                                                                   + MONTOS))
cuadre = pd.DataFrame([
    ("Unidad de comparación", "El POI no indica la actividad presupuestal de cada actividad operativa: no hay cuadre "
     "actividad por actividad con la Consulta Amigable.", "POI 2019–2025 (revisión de texto)"),
    ("Versiones", "Las versiones modificadas del POI son programaciones intermedias; el PIM de la Consulta Amigable es el "
     "final del año. La versión comparable es el POI consistente con el PIA: en 2021 y 2022 cuadra con el PIA del MEF por "
     "meta (DGIGND exacto; DGTEG con S/ 2 y S/ 3 de diferencia).",
     "Colección del POI del MIMP; REV_IND_3_verificacion_presupuestal.xlsx, hoja «POI apertura vs MEF»"),
    ("Total 2023", "La ejecución financiera por acción estratégica del reporte de seguimiento 2023 suma S/ 751,6 millones, "
     "88,8 % del devengado del pliego en la Consulta Amigable (S/ 846,1 millones).",
     POI + "2024_mimp_reporte-seguimiento-anual-poi-2023.pdf"),
    ("Donación AECID", "El POI 2023 indica que la actividad operativa de la DGIGND incorpora la donación de la AECID; la "
     "Consulta Amigable registra en 2023 S/ 2,23 millones de PIM por donaciones y transferencias en la actividad 5001072.",
     POI + "2023_mimp_poi-anual-2023_modificado-version-2.pdf; hoja «Conducción 5001072»"),
    ("DGIGND 2021", "AOI00108700017, única actividad operativa de la DGIGND en el POI 2021: S/ 2 398 676 programados (suma "
     "de los doce meses), p. 12 del PDF; equivale a 25,8 % del PIM de la actividad 5001072 de ese año.",
     POI + "2021_mimp_poi-anual-2021_modificado-version-1.pdf"),
], columns=["Aspecto", "Resultado", "Fuente"])
poi = pd.DataFrame(POI_DOCS, columns=["Año", "Archivo", "Documento", "Contenido para la conducción de la PNIG"])
poi["Archivo"] = POI + poi["Archivo"]
criterios = pd.DataFrame([
    ("Fuente", "Consulta Amigable del MEF, pliego 039; montos en soles; avance = devengado / PIM. Fecha de consulta y de "
               "actualización del MEF en el registro de descargas.", "src/fuentes/descargar_consulta_amigable.py"),
    ("Unidad de análisis", "Unidad ejecutora 001 (Administración Nivel Central), que comprende a la DGIGND, la DPIGND y la "
                           "OGMEPGD. INABIF y AURORA se muestran solo como contexto del pliego.", "Equipo evaluador"),
    ("Vinculación", "Directa por denominación: actividad cuyo nombre corresponde a la finalidad de la PNIG. Indirecta: "
                    "financia servicios o seguimiento del OP1. Soporte institucional: acciones centrales sin asignación "
                    "atribuible. Sin vinculación: el resto.", "Equipo evaluador; nota metodológica, Anexo 3 (REV_IND.3)"),
    ("Servicios del OP1", "Correspondencia por entidad y denominación de la actividad: directa si la actividad y la entidad "
                          "coinciden con el servicio; probable si coinciden en parte; sin servicio PNIG equivalente si "
                          "financia algo que la PNIG no define como servicio. Los gobiernos subnacionales no se desagregan.",
     "Equipo evaluador"),
    ("Límite", "La Consulta Amigable no separa el gasto por órgano. La parte de la conducción se separa por meta con los "
               "datos abiertos del MEF (REV_IND_3_verificacion_presupuestal.xlsx), sin transcribir montos del POI.",
     "Equipo evaluador"),
], columns=["Elemento", "Regla", "Fuente"])

with pd.ExcelWriter(SALIDA, engine="openpyxl") as xw:
    for hoja, df in [("Vinculada directa", directa), ("Conducción 5001072", detalle_5001072),
                     ("Resumen por vinculación", resumen), ("Servicios OP1 por año", servicios_pim),
                     ("Servicios OP1", servicios), ("Correspondencia 1002 y 0080", cat_det), ("Cuadre POI", cuadre),
                     ("Pliego por UE", pliego), ("UE 001 por categoría", por_cat), ("Actividades UE 001", detalle),
                     ("POI", poi), ("Criterios", criterios)]:
        df.to_excel(xw, sheet_name=hoja, index=False)
        ws = xw.sheets[hoja]
        ws.freeze_panes = "B2"
        for c in ws[1]:
            c.font = Font(bold=True)
            c.alignment = Alignment(wrap_text=True, vertical="top")
        for col in ws.columns:
            ancho = max(len(str(c.value or "")) for c in col[:60])
            ws.column_dimensions[col[0].column_letter].width = min(max(ancho, 9), 60)
            if col[0].value in MONTOS or isinstance(col[0].value, int):
                for c in col[1:]:
                    c.number_format = "#,##0"
            if col[0].value == "vinculacion":
                for c in col[1:]:
                    if c.value == "Directa por denominación":
                        c.fill = PatternFill("solid", fgColor="C6EFCE")

print(f"OK {SALIDA}: {len(detalle)} actividad-año de la UE 001")
print(directa.assign(pim=directa.pim.round(-3), devengado=directa.devengado.round(-3)).to_string(index=False))
print(resumen.pivot_table(index="vinculacion", columns="anio", values="pim", aggfunc="sum").div(1e6).round(2).to_string())
print(servicios_pim.set_index("servicio_pnig").div(1e6).round(1).to_string())
if len(detalle_5001072):
    print(detalle_5001072.groupby(["dimension", "denominacion", "anio"]).pim.sum().div(1e6).round(2).unstack().to_string())
