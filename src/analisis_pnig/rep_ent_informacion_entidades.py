"""REP_ENT: ficha de información reportada por las entidades (nota metodológica v14, Anexo 5).

Productos (04_analisis/REP_ENT_informacion_entidades.xlsx):
  - Panel por indicador 2023–2025 (bloque B), con el contraste del valor y la meta del formulario frente al consolidado
    de la OMEP y los reportes de cumplimiento que ya hizo REV_IND (hoja «Serie anual»), y la lista de discrepancias;
  - Matriz por servicio 2019–2025 (bloques A y C): sección del informe LIO, páginas, registros del formulario, montos
    reportados y verificados en el MEF (REV_IND.3) y segmentos codificables;
  - Codificación de los campos abiertos (bloques B, D, E y F) con el libro de códigos común (libro_codigos_pnig.csv) y
    sus frecuencias por pregunta, fuente y servicio.

Insumos: 03_datos/procesado/rep_ent_segmentos.csv (src/analisis_pnig/rep_ent_segmentos.py), la codificación en
04_analisis/REP_ENT_codificacion.csv (Claude, primera codificación, validada por el equipo evaluador), el resultado de la
doble codificación de los hallazgos críticos (04_analisis/REP_ENT_doble_codificacion_resultado.csv, de
rep_ent_doble_codificacion.py), REV_DOC y REV_IND.

Uso (desde la raíz del repositorio):
    uv run --no-project --with openpyxl --with pandas python src/analisis_pnig/rep_ent_informacion_entidades.py
"""
import os
import re
import warnings

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

warnings.filterwarnings("ignore", category=UserWarning)
RAIZ = EVAL
PANEL = f"{RAIZ}/03_datos/procesado/panel_formulario_seguimiento_pnig_2023_2025.csv"
COBERTURA = f"{RAIZ}/03_datos/procesado/cobertura_lio_servicios_2019_2025.csv"
SECCIONES = f"{RAIZ}/03_datos/procesado/secciones_lio_2019_2025.csv"
SEGMENTOS = f"{RAIZ}/03_datos/procesado/rep_ent_segmentos.csv"
CODIFICACION = f"{RAIZ}/04_analisis/REP_ENT_codificacion.csv"
DOBLE = f"{RAIZ}/04_analisis/REP_ENT_doble_codificacion_resultado.csv"
LIBRO = f"{RAIZ}/04_analisis/libro_codigos_pnig.csv"
PSI05 = f"{RAIZ}/04_analisis/REV_IND_base_indicadores.xlsx"
VERIF = f"{RAIZ}/04_analisis/REV_IND_3_verificacion_presupuestal.xlsx"
GRD = f"{RAIZ}/04_analisis/REV_DOC_estado_implementacion_servicios.xlsx"
SALIDA = f"{RAIZ}/04_analisis/REP_ENT_informacion_entidades.xlsx"
SENTIDOS = {"Favorable", "Dificultad", "Mixto", "Propuesta", "Neutro"}
MIN_LINEAS = 15                        # mismo umbral que REV_DOC para una sección sustantiva del LIO

libro = pd.read_csv(LIBRO)
seg = pd.read_csv(SEGMENTOS, dtype={"servicio": str})
cod = pd.read_csv(CODIFICACION, dtype=str)

# Controles de la codificación
faltan = set(cod.id) - set(seg.id)
assert not faltan, f"segmentos codificados que no existen: {sorted(faltan)[:5]}"
assert cod.id.is_unique, "segmento codificado dos veces"
codigos = cod.codigos.str.split(r";\s*").explode().str.strip()
ajenos = set(codigos) - set(libro.codigo)
assert not ajenos, f"códigos fuera del libro: {ajenos}"
assert set(cod.sentido.dropna()) <= SENTIDOS, set(cod.sentido.dropna()) - SENTIDOS

# Bloque B: panel por indicador con el contraste de REV_IND
panel = pd.read_csv(PANEL, dtype={"servicio": str})
panel["Código"] = panel.id.str.replace("PNIG_", "", regex=False)
serie = pd.read_excel(PSI05, sheet_name="Serie anual")
contraste = serie[["Código", "Año", "Meta", "Valor", "Fuente del valor", "Valor consolidado", "Valor RC",
                   "Contraste valor formulario", "Contraste meta formulario", "Nota sobre el formulario"]].rename(
    columns={"Meta": "Meta verificada (REV_IND)", "Valor": "Valor adoptado (REV_IND)"})
panel = panel.merge(contraste, left_on=["Código", "anio"], right_on=["Código", "Año"], how="left").drop(columns="Año")
cols_panel = ["anio", "Código", "nivel", "op", "servicio", "indicador", "entidad", "organo", "dispone_informacion",
              "motivo_no_reporte", "valor", "meta", "avance_pct", "cumplimiento", "Meta verificada (REV_IND)",
              "Valor adoptado (REV_IND)", "Fuente del valor", "Valor consolidado", "Valor RC",
              "Contraste valor formulario", "Contraste meta formulario", "Nota sobre el formulario"]
panel = panel[cols_panel].rename(columns={"anio": "Año", "nivel": "Nivel", "op": "OP", "servicio": "Servicio",
                                          "indicador": "Indicador", "entidad": "Entidad", "organo": "Órgano",
                                          "dispone_informacion": "Dispone de información",
                                          "motivo_no_reporte": "Motivo de no reporte", "valor": "Valor (formulario)",
                                          "meta": "Meta (formulario)", "avance_pct": "Avance (formulario, %)",
                                          "cumplimiento": "Nivel de cumplimiento (formulario)"})
SIN_DISCREPANCIA = ["Coincide", "Ambos sin valor"]
discrepancias = panel[~panel["Contraste valor formulario"].fillna("Coincide").isin(SIN_DISCREPANCIA)
                      | ~panel["Contraste meta formulario"].fillna("Coincide").isin(SIN_DISCREPANCIA)]

# Bloques A y C: matriz por servicio y año
grd = pd.read_excel(GRD, sheet_name="Estado por servicio", dtype=str)
cob = pd.read_csv(COBERTURA, dtype={"servicio": str})
sec = pd.read_csv(SECCIONES, dtype={"servicio": str})[["anio", "servicio", "pagina_inicio", "pagina_fin"]]
cif = pd.read_excel(VERIF, sheet_name="Cifras verificadas", dtype={"Servicio": str})
cif["verificada"] = cif.Resultado.ne("Sin coincidencia")
montos = cif[cif.Servicio.notna()].groupby(["Servicio", "Año"]).agg(
    montos=("Resultado", "size"), verificados=("verificada", "sum")).reset_index()
form = pd.read_csv(PANEL, dtype={"servicio": str})
form = form[form.servicio.notna()].groupby(["servicio", "anio"]).agg(
    registros=("id", "size"), cumplimiento=("cumplimiento", lambda s: "; ".join(sorted(set(s.dropna()))))).reset_index()
n_seg = seg[seg.servicio.notna()].groupby(["servicio", "anio"]).size().rename("segmentos").reset_index()

m = cob.merge(sec, on=["anio", "servicio"], how="left")
m = m.merge(montos, left_on=["servicio", "anio"], right_on=["Servicio", "Año"], how="left").drop(columns=["Servicio", "Año"])
m = m.merge(form, on=["servicio", "anio"], how="left").merge(n_seg, on=["servicio", "anio"], how="left")
m = m.merge(grd[["Código", "Servicio", "Entidad u órgano responsable", "Protocolo aprobado"]],
            left_on="servicio", right_on="Código", how="left")


def seccion(r):
    if r.seccion != "Sí":
        return "No"
    return "Sí" if r.lineas >= MIN_LINEAS and r.sin_info == 0 else "Mínima"


m["Sección LIO"] = m.apply(seccion, axis=1)
m["Páginas LIO"] = m.apply(lambda r: f"{int(r.pagina_inicio)}–{int(r.pagina_fin)}" if pd.notna(r.pagina_inicio) else None,
                           axis=1)
clave = lambda c: c.map(lambda x: tuple(int(i) for i in x.split(".")))
matriz = m.assign(montos=m.montos.fillna(0).astype(int), verificados=m.verificados.fillna(0).astype(int),
                  registros=m.registros.fillna(0).astype(int), segmentos=m.segmentos.fillna(0).astype(int))
matriz = matriz.sort_values(["servicio", "anio"], key=lambda c: clave(c) if c.name == "servicio" else c)[[
    "servicio", "Servicio", "Entidad u órgano responsable", "Protocolo aprobado", "anio", "Sección LIO", "Páginas LIO",
    "lineas", "cobertura", "territorio", "articulacion", "dificultad", "montos", "verificados", "registros",
    "cumplimiento", "segmentos"]].rename(columns={
        "servicio": "Código", "anio": "Año", "lineas": "Líneas de la sección", "cobertura": "Menciones de cobertura",
        "territorio": "Menciones territoriales", "articulacion": "Menciones de articulación",
        "dificultad": "Menciones de dificultad", "montos": "Montos reportados (REV_IND.3)",
        "verificados": "Montos verificados en el MEF", "registros": "Registros del formulario",
        "cumplimiento": "Nivel de cumplimiento (formulario)", "segmentos": "Segmentos codificables (D, E, F)"})

# Codificación y frecuencias
codif = seg.merge(cod, on="id", how="left")
# Doble codificación: los candidatos a hallazgo crítico toman los códigos, el sentido y la atribución adjudicados y
# quedan como hallazgo crítico confirmado o descartado.
codif["atribucion_adjudicada"] = ""
if os.path.exists(DOBLE):
    res = pd.read_csv(DOBLE, dtype=str).fillna("").set_index("id")
    en = codif.id.isin(res.index)
    for col, origen in (("codigos", "codigos"), ("sentido", "sentido"), ("nota", "nota"), ("atribucion_adjudicada", "atribucion")):
        codif.loc[en, col] = codif.loc[en, "id"].map(res[origen])
    codif.loc[en, "hallazgo_critico"] = codif.loc[en, "id"].map(res.critico).map({"Sí": "Confirmado", "No": "Descartado"})
    codif.loc[en, "codificador"] = "Claude (1.ª) y agente ciego (2.ª); adjudicación del equipo evaluador"
    codif.loc[en, "validacion"] = "Doble codificación adjudicada (2026-10-31): " + codif.loc[en, "id"].map(res.origen)
codif["estado"] = codif.codigos.notna().map({True: "Codificado", False: "Pendiente"})


def atribuir(r):
    """Servicio u objeto al que corresponde el segmento. Las secciones del LIO a veces incluyen texto de otro servicio, del
    objetivo prioritario, de la CONAIG o de las conclusiones; la nota de codificación lo indica con «Corresponde a …»."""
    propio = r.servicio if isinstance(r.servicio, str) else ("CONAIG" if "CONAIG" in r.fuente else r.indicador)
    if r.atribucion_adjudicada:  # la adjudicación de la doble codificación prevalece sobre la nota
        return propio if r.atribucion_adjudicada == "Sin reatribuir" else r.atribucion_adjudicada
    nota = r.nota if isinstance(r.nota, str) else ""
    if m := re.match(r"Corresponde a (\d\.\d\.\d)", nota):
        return m.group(1)
    if m := re.match(r"Corresponde a (OP\d)", nota):
        return m.group(1)
    for clave, destino in (("Corresponde a la CONAIG", "CONAIG"), ("Corresponde a la conclusión", "Conclusiones del informe"),
                           ("Corresponde a la recomendación", "Conclusiones del informe"),
                           ("Corresponde a otro servicio", "Sin atribuir")):
        if nota.startswith(clave):
            return destino
    return propio


codif["atribuido a"] = codif.apply(atribuir, axis=1)
larga = codif[codif.codigos.notna()].assign(codigo=lambda d: d.codigos.str.split(r";\s*")).explode("codigo")
larga = larga.merge(libro[["codigo", "pregunta", "nombre"]], on="codigo", how="left")
larga["tipo de fuente"] = larga.fuente.str.extract(r"^(Formulario|Informe LIO|.*CONAIG)")[0].str.replace(
    r".*CONAIG", "Informes de la CONAIG", regex=True)
por_pregunta = larga.pivot_table(index=["pregunta", "codigo", "nombre"], columns="tipo de fuente", values="id",
                                 aggfunc="count", fill_value=0).reset_index()
por_servicio = larga.pivot_table(index="codigo", columns="atribuido a", values="id",
                                                         aggfunc="count", fill_value=0).reset_index()
avance = codif.groupby("fuente").estado.value_counts().unstack(fill_value=0).reset_index()

criterios = pd.DataFrame([
    ("Unidad de codificación", "Formulario: cada campo abierto con texto. Informes LIO: cada párrafo de la sección del "
     "servicio con palabras de los bloques D, E o F. Informes de la CONAIG: cada párrafo con palabras de dificultades, "
     "recomendaciones, sesiones o acuerdos (rep_ent_segmentos.py).", "Equipo evaluador"),
    ("Preselección", "Los párrafos que no contienen esas palabras no se codifican; su contenido es de ejecución (bloque C) "
     "y se resume en la matriz por servicio. Es una reducción del material de lectura, no un muestreo.", "Equipo evaluador"),
    ("Libro de códigos", "Deductivo, derivado de las preguntas D5, I1, I4 e I5 (nota, numeral 9); común con ENTREV, que "
     "añade I3 y D7. C.EJEC marca los segmentos sin factor; EMER, el contenido no previsto, para revisión.",
     "Nota metodológica (numeral 9)"),
    ("Sentido", "Favorable, dificultad, mixto (ambos), propuesta (recomendación o ajuste) o neutro (sin valencia).",
     "Equipo evaluador"),
    ("Codificación", "Primera codificación de Claude con nota de sustento; validación del equipo evaluador y doble "
     "codificación de los hallazgos críticos (columna «hallazgo_critico»: «Candidato» lo propone la primera "
     "codificación).", "Nota metodológica (numeral 9); decisión del usuario, 2026-10-30"),
    ("Doble codificación", "Segunda codificación ciega de los candidatos por un agente nuevo, concordancia y adjudicación "
     "de las discrepancias por el equipo evaluador (REP_ENT_doble_codificacion.xlsx, hoja «Acta»). Los candidatos quedan "
     "como «Confirmado» o «Descartado» con los códigos, el sentido y la atribución adjudicados.",
     "Nota metodológica (numeral 9); decisión del usuario, 2026-10-31"),
    ("Contraste del panel", "Valor y meta del formulario frente al consolidado de la OMEP y los reportes de cumplimiento "
     "(REV_IND, hoja «Serie anual»); las discrepancias se documentan y no se corrigen sin sustento en la ficha técnica.",
     "Nota metodológica (Anexo 5)"),
    ("Sección LIO", f"Sí: {MIN_LINEAS} líneas o más sin mención de falta de información; Mínima: más breve o con esa "
     "mención; No: sin sección ubicada (mismo criterio que REV_DOC).", "Equipo evaluador"),
    ("Atribución", "La sección de un servicio en el informe LIO puede incluir texto de otro servicio, del objetivo "
     "prioritario, de la CONAIG o de las conclusiones (sobre todo la última sección de cada objetivo). La codificación lo "
     "anota con «Corresponde a …» y la columna «atribuido a» reasigna el segmento en las frecuencias por servicio.",
     "Equipo evaluador"),
], columns=["Elemento", "Criterio", "Plano"])

hojas = [("Panel por indicador", panel), ("Discrepancias del panel", discrepancias), ("Matriz por servicio", matriz),
         ("Codificación", codif[["id", "fuente", "anio", "pagina", "servicio", "atribuido a", "indicador", "entidad", "campo",
                                 "bloque", "texto", "codigos", "sentido", "nota", "codificador", "fecha", "lote", "validacion",
                                 "hallazgo_critico", "estado"]]),
         ("Hallazgos críticos", codif[codif.hallazgo_critico == "Confirmado"][[
             "id", "fuente", "anio", "pagina", "atribuido a", "entidad", "texto", "codigos", "sentido", "nota"]]),
         ("Avance de la codificación", avance), ("Códigos por pregunta", por_pregunta),
         ("Códigos por servicio", por_servicio), ("Libro de códigos", libro), ("Criterios", criterios)]
with pd.ExcelWriter(SALIDA, engine="openpyxl") as xw:
    for nombre, df in hojas:
        df.to_excel(xw, sheet_name=nombre, index=False)
        ws = xw.sheets[nombre]
        ws.freeze_panes = "B2"
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F3864")
            c.alignment = Alignment(wrap_text=True, vertical="top")
        for col in ws.columns:
            titulo = str(col[0].value)
            ancho = 80 if titulo in ("texto", "Criterio", "definicion") else 40 if titulo in (
                "nota", "Indicador", "Servicio", "incluye", "excluye", "Motivo de no reporte") else 14
            ws.column_dimensions[col[0].column_letter].width = ancho
            if ancho >= 40:
                for c in col[1:]:
                    c.alignment = Alignment(wrap_text=True, vertical="top")
print(avance.to_string(index=False))
print("discrepancias del panel:", len(discrepancias), "| filas de la matriz:", len(matriz))
