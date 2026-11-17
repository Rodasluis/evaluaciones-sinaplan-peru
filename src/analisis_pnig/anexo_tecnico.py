"""Anexo técnico de la evaluación de cierre de la PNIG (2019–2025): arma la carpeta con los documentos técnicos que
acompañan el informe y permiten entenderlo y replicarlo (ronda 36, pedido del usuario del 2026-11-11).

Contenido (subcarpetas numeradas):
  00_indice                    índice del anexo (Word y PDF) y lista de archivos con su SHA-256
  01_informe_y_presentacion    informe de cierre (PDF de Word), presentaciones de resultados (completa y resumida) e infografía
  02_nota_metodologica         nota metodológica vigente (PDF), con las fichas de los instrumentos
  03_instrumentos              cuestionarios y guías aplicadas de las entrevistas (PDF) y libro de códigos
  04_matrices_y_libros         libros de análisis con su nombre publicado y hojas sin siglas (ADR-010)
  05_entrevistas               transcripciones anonimizadas y registro codificado, sin nombres (ADR-007)
  06_rastro                    rastro de hallazgos, tablas y figuras, y correspondencia de nombres
  07_fuentes                   documentos de respaldo (Anexo 10) e inventario del corpus
  08_figuras                   figuras del informe en PNG, con su número
  09_codigo                    scripts de procesamiento (rutas locales reemplazadas por %USERPROFILE%)

La carpeta no se versiona (.gitignore): se regenera con este script después del informe (que deja el PDF de Word en
%TEMP% y 06_informe/indice_figuras.csv) y de la presentación de resultados.

Uso (desde la raíz del repositorio):
    uv run --no-project --with pandas --with openpyxl --with python-docx --with pillow --with pyyaml python -X utf8
        src/analisis_pnig/anexo_tecnico.py
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl
import pandas as pd

sys.path.insert(0, os.path.dirname(__file__))
import informe_referencias as R  # noqa: E402
import sintesis_hallazgos as S  # noqa: E402
from docx_mimp import Documento, lista_es  # noqa: E402
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

BASE = Path(EVAL)
RAIZ = BASE / "07_anexo_tecnico"
NOMBRE = "Anexo_tecnico_PNIG_2019-2025"   # corto: las rutas de Windows no pasan de 260 caracteres
SALIDA = RAIZ / NOMBRE
from nucleo.anexo import libro_publicado as _libro_publicado, publico as _publico, tabla_xlsx  # noqa: E402
from nucleo.fuentes import a_pdf, sha256, slug  # noqa: E402,F401
PDF_INFORME = Path(tempfile.gettempdir()) / "informe_cierre_pnig_word.pdf"
INFORME = BASE / "06_informe/2026_mimp-ogmepgd_informe-evaluacion-cierre-pnig.docx"
PRESENTACION = BASE / "06_informe/2026_mimp-ogmepgd_presentacion-resultados-evaluacion-cierre-pnig.pptx"
RESUMIDA = BASE / "06_informe/2026_mimp-ogmepgd_presentacion-resumida-resultados-evaluacion-cierre-pnig.pptx"
INFOGRAFIA = BASE / "06_informe/2026_mimp-ogmepgd_infografia-evaluacion-cierre-pnig.pdf"
NOTA = BASE / "00_encargo/nota_metodologica/Nota Metodologica Evaluación PNIG - evaluación de cierre.docx"
ENTREV = BASE / "05_entrevistas"
ANALISIS = BASE / "04_analisis"
INFORME_DIR = BASE / "06_informe"

# instrumentos: código interno (en los scripts y el rastro) y nombre en el informe (ADR-009)
INSTRUMENTOS = [("REV_DOC", "Revisión documental"), ("ENTREV", "Entrevistas al nivel conductor"),
                ("REP_ENT", "Información reportada por las entidades"), ("REV_ALI", "Coherencia y alineamiento"),
                ("REV_IND", "Indicadores y recursos"), ("REV_EVI", "Revisión de evidencia")]

DESCRIPCION = {
    "Brechas_situacion_futura.xlsx": "Evolución nacional de las brechas del diagnóstico, manifestaciones medidas con la "
                                     "ENDES y correspondencia de la situación futura deseada con los indicadores "
                                     "(numerales 1.3, 3.1 y 5.4).",
    "Brechas_departamentales.xlsx": "Brechas por departamento con su coeficiente de variación y su cambio entre 2019 y "
                                    "2025, participación política y validaciones con el INEI (numeral 3.1 y Anexo 6).",
    "Enfoques_transversales.xlsx": "Lista de verificación de los siete enfoques transversales y su presencia en los "
                                   "protocolos (numeral 3.7).",
    "Coherencia_alineamiento.xlsx": "Coherencia interna, alineamiento vertical y horizontal, correspondencia con la Ley "
                                    "N.° 32535 y los compromisos internacionales, y correspondencia conceptual "
                                    "(numerales 1.4, 1.7, 3.2, 3.3 y 3.5).",
    "Contraste_matriz_incompatibilidades.xlsx": "Contraste de la clasificación de los servicios según el nuevo marco con la matriz de incompatibilidades de la DPIGND: regla, tabla cruzada, kappa y discrepancias (numeral 3.2).",
    "Estado_implementacion_servicios.xlsx": "Estado de implementación de los 52 servicios, con sus componentes, su "
                                            "ejecución por año y los criterios de la escala (numeral 4.1 y Anexo 3).",
    "Matriz_evidencia.xlsx": "Estrategias de búsqueda, flujo de selección, matriz de evidencia por estudio y por tipo "
                             "de intervención, y referencias (numeral 3.4 y Anexo 4).",
    "Base_indicadores.xlsx": "Indicadores de objetivo prioritario y de servicio 2019–2025, con metas, valores y nivel "
                             "de avance (numerales 3.6 y 5.1 a 5.3, y Anexo 5).",
    "Calidad_fichas.xlsx": "Atributos de las 62 fichas técnicas de los indicadores e indicadores del PEMIG por tipo "
                           "(numeral 3.6).",
    "Recursos_conduccion.xlsx": "Recursos de la conducción de la Política por meta, órgano y año (numeral 4.2).",
    "Verificacion_presupuestal.xlsx": "Montos reportados por las entidades frente a los datos abiertos del MEF y "
                                      "devengado de la conducción por órgano (numerales 4.1 y 4.2).",
    "Presupuesto_DGIGND.xlsx": "Ejecución de la actividad 5001072 remitida por la DGIGND, conciliada con el MEF, y "
                               "composición del gasto de la conducción (numeral 4.2).",
    "Recalculo_indicadores_OP3.xlsx": "Recálculo de los indicadores del OP3 para 2025 (numeral 5.2).",
    "Informacion_entidades.xlsx": "Sistematización de lo reportado por las entidades en el formulario de seguimiento "
                                  "2023–2025 y en los informes LIO 2019–2025, con su codificación (numerales 4.1, 4.4 "
                                  "a 4.6).",
    "Doble_codificacion.xlsx": "Segunda codificación ciega de los segmentos críticos y su concordancia (numeral 2.6).",
    "Seguimiento_recomendaciones.xlsx": "Seguimiento de las recomendaciones de las evaluaciones anteriores (numeral 4.7 "
                                        "y Anexo 7).",
    "Matriz_triangulacion.xlsx": "Matriz de triangulación por hallazgo, cobertura de los instrumentos por pregunta y "
                                 "valoración de las preguntas (numeral 2.5 y Anexos 1 y 2).",
    "Registro_entrevistas.xlsx": "Registro codificado de las entrevistas al nivel conductor, sin nombres (Anexo 8).",
    "Validacion_empleo_informal_ENAHO.csv": "Validación del empleo informal de la ENAHO 2024 y 2025 reconstruido con "
                                            "la regla del INEI (numeral 3.1).",
}

CARPETAS = [
    ("00_indice", "Este índice y la lista de archivos con su SHA-256.", "Todo el informe"),
    ("01_informe_y_presentacion", "Informe de evaluación de cierre (PDF), presentación de resultados completa y resumida (PowerPoint) e infografía (PDF).",
     "Documento principal"),
    ("02_nota_metodologica", "Nota metodológica vigente, con las fichas de los seis instrumentos.", "Numeral 2"),
    ("03_instrumentos", "Cuestionarios y guías aplicadas de las entrevistas y libro de códigos común de la "
                        "codificación temática.", "Numeral 2.3 y Anexo 8"),
    ("04_matrices_y_libros", "Libros de análisis que sustentan los hallazgos, con sus criterios.", "Anexos 1 a 7"),
    ("05_entrevistas", "Transcripciones anonimizadas y registro codificado de las entrevistas.", "Anexo 8"),
    ("06_rastro", "Rastro técnico de cada hallazgo, tabla y figura, y correspondencia entre los nombres del "
                  "repositorio y los publicados.", "Anexo 2 y notas al pie"),
    ("07_fuentes", "Documentos de respaldo con su enlace o vía de consulta e inventario del corpus documental.",
     "Anexo 10 y referencias"),
    ("08_figuras", "Figuras del informe en PNG, con el número que llevan en el informe.", "Todas las figuras"),
    ("09_codigo", "Scripts de Python que generan las matrices, las figuras, el informe y este anexo.",
     "Reproducibilidad"),
]

ORDEN_CODIGO = [  # (grupo, prefijos o archivos) en el orden en que se ejecutan
    ("Fuentes y corpus", ["src/fuentes/"]),
    ("Información reportada por las entidades", ["src/insumos_seguimiento/", "rep_ent_"]),
    ("Brechas y situación futura", ["svy.py", "d1_", "d2_"]),
    ("Coherencia y alineamiento", ["rev_ali_"]),
    ("Indicadores y recursos", ["rev_ind_"]),
    ("Revisión documental", ["rev_doc_"]),
    ("Revisión de evidencia", ["rev_evi_"]),
    ("Enfoques y recomendaciones anteriores", ["d6_", "seguimiento_recomendaciones"]),
    ("Entrevistas", ["entrev_"]),
    ("Síntesis y triangulación", ["sintesis_"]),
    ("Informe, presentación y anexo", ["docx_mimp", "figuras_informe", "informe_", "actualizar_campos_word",
                                       "src/nucleo/", "presentacion_resultados", "anexo_tecnico"]),
]














def main():
    assert SALIDA.parent == RAIZ and SALIDA.name == NOMBRE
    if SALIDA.exists():
        shutil.rmtree(SALIDA)
    d = {c: SALIDA / c for c, _, _ in CARPETAS}
    for ruta in d.values():
        ruta.mkdir(parents=True)
    archivos = []          # (carpeta, archivo, descripción, origen)

    def registrar(carpeta, ruta, descripcion, origen=""):
        # ruta relativa a la subcarpeta (el código conserva sus carpetas src/…)
        archivos.append((carpeta, Path(ruta).relative_to(d[carpeta]).as_posix(), descripcion,
                         str(origen).replace("\\", "/")))

    # 01 informe y presentación
    assert PDF_INFORME.exists(), "falta el PDF de Word del informe: corra informe_cierre.py"
    destino = d["01_informe_y_presentacion"] / "2026_mimp-ogmepgd_informe-evaluacion-cierre-pnig.pdf"
    shutil.copy2(PDF_INFORME, destino)
    registrar("01_informe_y_presentacion", destino, "Informe de evaluación de cierre (PDF exportado por Word).", INFORME)
    destino = d["01_informe_y_presentacion"] / PRESENTACION.name
    shutil.copy2(PRESENTACION, destino)
    registrar("01_informe_y_presentacion", destino, "Presentación de resultados para la validación.", PRESENTACION)
    for origen, desc in ((RESUMIDA, "Presentación resumida de los resultados (ocho láminas)."),
                         (INFOGRAFIA, "Infografía de los resultados (dos caras A4).")):
        destino = d["01_informe_y_presentacion"] / origen.name
        shutil.copy2(origen, destino)
        registrar("01_informe_y_presentacion", destino, desc, origen)

    # 02 nota metodológica
    tmp = Path(tempfile.mkdtemp(prefix="anexo_"))
    copia = tmp / "nota.docx"
    shutil.copy2(NOTA, copia)
    a_pdf([copia], tmp)
    destino = d["02_nota_metodologica"] / "2026_mimp-ogmepgd_nota-metodologica-evaluacion-cierre-pnig.pdf"
    shutil.move(tmp / "nota.pdf", destino)
    registrar("02_nota_metodologica", destino, "Nota metodológica vigente: preguntas, criterios, instrumentos y "
              "plan de análisis.", NOTA)

    # 03 instrumentos
    guias = sorted(ENTREV.glob("ENTREV_guia_*.docx"))
    a_pdf(guias, tmp)
    for g in guias:
        org = g.stem.split("_")[-1]
        destino = d["03_instrumentos"] / f"Guia_aplicada_entrevista_{org}.pdf"
        shutil.move(tmp / f"{g.stem}.pdf", destino)
        registrar("03_instrumentos", destino, f"Guía aplicada de la entrevista a la {org}: preguntas, formulación, "
                  f"minuto y evidencia de contraste.", g)
    for c in sorted(ENTREV.glob("ENTREV_cuestionario_*.pdf")):
        org = c.stem.split("_")[-1]
        destino = d["03_instrumentos"] / f"Cuestionario_entrevista_{org}.pdf"
        shutil.copy2(c, destino)
        registrar("03_instrumentos", destino, f"Cuestionario aplicado en la entrevista a la {org}.", c)
    codigos = pd.read_csv(ANALISIS / "libro_codigos_pnig.csv")
    destino = d["03_instrumentos"] / "Libro_codigos.xlsx"
    tabla_xlsx(codigos, destino, "Libro de códigos")
    registrar("03_instrumentos", destino, "Libro de códigos común de la codificación de las entidades y de las "
              "entrevistas.", ANALISIS / "libro_codigos_pnig.csv")

    # 04 libros de análisis y 05 entrevistas
    for rep, pub in S.LIBROS_PUBLICOS.items():
        origen = next((r for r in (ANALISIS / rep, ENTREV / rep, ANALISIS / "d1_brechas" / rep) if r.exists()), None)
        assert origen is not None, f"falta el libro {rep}"
        carpeta = "05_entrevistas" if rep.startswith("ENTREV") else "04_matrices_y_libros"
        destino = d[carpeta] / pub
        if rep.endswith(".xlsx"):
            _libro_publicado(origen, destino, S.hoja_publica)
        else:
            shutil.copy2(origen, destino)
        registrar(carpeta, destino, DESCRIPCION[pub], origen)
    for t in sorted((ENTREV / "transcripciones").glob("ENTREV_transcripcion_*.txt")):
        org = t.stem.split("_")[-1]
        destino = d["05_entrevistas"] / f"Transcripcion_anonimizada_{org}.txt"
        shutil.copy2(t, destino)
        registrar("05_entrevistas", destino, f"Transcripción anonimizada de la entrevista a la {org}.", t)

    # 06 rastro
    rh = pd.read_csv(INFORME_DIR / "rastro_hallazgos.csv")
    rh = rh.rename(columns={"hallazgo": "Hallazgo", "pregunta": "Pregunta", "confianza": "Confianza",
                            "fuentes_independientes": "Fuentes independientes",
                            "respaldo": "Respaldo (nombres publicados)", "rastro": "Rastro en el repositorio",
                            "nota": "Nota"})
    destino = d["06_rastro"] / "Rastro_hallazgos.xlsx"
    tabla_xlsx(rh[["Hallazgo", "Pregunta", "Confianza", "Fuentes independientes", "Respaldo (nombres publicados)",
                   "Nota", "Rastro en el repositorio"]], destino, "Hallazgos")
    registrar("06_rastro", destino, "Cada hallazgo con su confianza, sus fuentes independientes y el libro y la hoja "
              "que lo respaldan.", INFORME_DIR / "rastro_hallazgos.csv")
    rt = pd.read_csv(INFORME_DIR / "rastro_tablas_figuras.csv")
    rt["Rastro (nombres publicados)"] = rt.rastro.map(lambda t: _publico(t, S.LIBROS_PUBLICOS, S.HOJAS_PUBLICAS))
    rt = rt.rename(columns={"tipo": "Tipo", "numero": "Número", "seccion": "Numeral", "leyenda": "Leyenda",
                            "rastro": "Rastro en el repositorio"})
    destino = d["06_rastro"] / "Rastro_tablas_figuras.xlsx"
    tabla_xlsx(rt[["Tipo", "Número", "Numeral", "Leyenda", "Rastro (nombres publicados)", "Rastro en el repositorio"]],
               destino, "Tablas y figuras")
    registrar("06_rastro", destino, "Cada tabla y figura del informe con el libro y la hoja de donde sale.",
              INFORME_DIR / "rastro_tablas_figuras.csv")
    cl = pd.read_csv(INFORME_DIR / "correspondencia_libros.csv")
    destino = d["06_rastro"] / "Correspondencia_libros.xlsx"
    tabla_xlsx(cl.rename(columns={"libro_repositorio": "Libro en el repositorio", "libro_publicado": "Libro publicado",
                                  "hoja_repositorio": "Hoja en el repositorio", "hoja_publicada": "Hoja publicada"}),
               destino, "Correspondencia")
    registrar("06_rastro", destino, "Nombres de los libros y hojas en el repositorio y en este anexo.",
              INFORME_DIR / "correspondencia_libros.csv")

    # 07 fuentes
    docs = [x for x in R.documentos_respaldo("docs/fuentes/registro_fuentes.md", "docs/fuentes/inventario_fuentes.csv",
                                             "docs/fuentes/enlaces_publicos_2026-11-06.tsv")
            if not S.INTERNOS.search(x["documento"] + " " + x["norma"])]
    dr = pd.DataFrame([{"Grupo": re.sub(r"^\d+[a-z]?\.\s*", "", x["seccion"]), "Documento": x["documento"],
                        "Norma o versión": x["norma"], "Fecha": x["fecha"],
                        "Enlace o vía de consulta": x["url"] or x["via"]} for x in docs])
    destino = d["07_fuentes"] / "Documentos_de_respaldo.xlsx"
    tabla_xlsx(dr, destino, "Documentos de respaldo")
    registrar("07_fuentes", destino, "Documentos que respaldan los hallazgos, con su enlace público o vía de "
              "consulta (Anexo 10 del informe).", "docs/fuentes/registro_fuentes.md")
    inv = pd.read_csv("docs/fuentes/inventario_fuentes.csv")
    inv = inv[["ruta", "anio", "organismo", "titulo_corto", "formato", "paginas", "sha256", "acceso", "enlace_publico"]]
    destino = d["07_fuentes"] / "Inventario_fuentes.xlsx"
    tabla_xlsx(inv.rename(columns={"ruta": "Ruta en el repositorio", "anio": "Año", "organismo": "Organismo",
                                   "titulo_corto": "Título corto", "formato": "Formato", "paginas": "Páginas",
                                   "sha256": "SHA-256", "acceso": "Acceso", "enlace_publico": "Enlace público"}),
               destino, "Inventario")
    registrar("07_fuentes", destino, f"Inventario del corpus documental ({len(inv)} documentos) con su SHA-256, "
              "acceso y enlace público.", "docs/fuentes/inventario_fuentes.csv")

    # 08 figuras
    fi = pd.read_csv(INFORME_DIR / "indice_figuras.csv")
    for _, f in fi.iterrows():
        origen = Path(f.archivo)
        destino = d["08_figuras"] / f"Figura_{int(f.numero):02d}_{slug(str(f.leyenda))}{origen.suffix}"
        shutil.copy2(origen, destino)
        registrar("08_figuras", destino, f"Figura {int(f.numero)}. {f.leyenda} ({f.seccion}).", origen)

    # 09 código, con las rutas locales del equipo reemplazadas
    fuentes = (sorted(Path("src/analisis_pnig").glob("*.py")) + sorted(Path("src/analisis_pnig").glob("*.ps1")) +
               sorted(Path("src/insumos_seguimiento").glob("*.py")) + sorted(Path("src/fuentes").glob("*.py")) +
               sorted(Path("src/nucleo").rglob("*.py")) + [Path("config/portada.yaml"), Path("config/entorno.yaml")])
    home = str(Path.home())
    for f in fuentes:
        destino = d["09_codigo"] / f.as_posix()
        destino.parent.mkdir(parents=True, exist_ok=True)
        texto = f.read_text(encoding="utf-8")
        for var in (home, home.replace("\\", "/"), home.replace("\\", "\\\\")):
            texto = texto.replace(var, "%USERPROFILE%")
        destino.write_text(texto, encoding="utf-8")
        registrar("09_codigo", destino, primera_linea(texto), f)
    leeme = leeme_codigo(fuentes)
    (d["09_codigo"] / "LEEME.txt").write_text(leeme, encoding="utf-8")
    registrar("09_codigo", d["09_codigo"] / "LEEME.txt", "Orden de ejecución, requisitos y descripción de cada "
              "script.")

    # 00 índice: lista de archivos con hash y documento índice
    lista = pd.DataFrame([{"Carpeta": c, "Archivo": a, "Descripción": desc, "Origen en el repositorio": o,
                           "SHA-256": sha256(SALIDA / c / a)} for c, a, desc, o in archivos])
    destino = d["00_indice"] / "Lista_de_archivos.xlsx"
    tabla_xlsx(lista, destino, "Archivos")
    indice(d["00_indice"], archivos, lista)
    a_pdf([d["00_indice"] / "Indice_del_anexo_tecnico.docx"], d["00_indice"])
    shutil.rmtree(tmp, ignore_errors=True)
    print(f"Anexo técnico: {SALIDA} ({len(archivos)} archivos y el índice)")
    return SALIDA


def primera_linea(texto):
    m = re.search(r'^(?:\ufeff)?\s*(?:"""|#)\s*(.+)', texto)
    return (m.group(1).strip().rstrip('"') if m else "")[:200]


def leeme_codigo(fuentes):
    lineas = ["Código de procesamiento de la evaluación de cierre de la PNIG (2019–2025)", "",
              "Requisitos: Python 3.11 o superior con pandas, openpyxl, python-docx, python-pptx, pillow, pyyaml,",
              "pymupdf y lxml; LibreOffice y Microsoft Word (para terminar el informe); Google Chrome sin ventana",
              "(para las figuras). Los scripts se ejecutan desde la raíz del repositorio; cada uno documenta su uso",
              "en el encabezado.", "",
              "Datos: los microdatos del INEI y los datos abiertos del MEF no se incluyen; se descargan de sus",
              "portales con los scripts de src/fuentes. Las rutas locales del equipo se reemplazaron por",
              "%USERPROFILE%; pueden cambiarse con las variables de entorno que indica cada script.", "",
              "Códigos de los instrumentos en los nombres de archivo:"]
    lineas += [f"  {c}: {n}" for c, n in INSTRUMENTOS]
    lineas += ["", "Orden de ejecución:"]
    usados = set()
    for grupo, claves in ORDEN_CODIGO:
        sel = [f for f in fuentes if f not in usados and any(k in f.as_posix() for k in claves)]
        if not sel:
            continue
        lineas += ["", grupo]
        for f in sel:
            usados.add(f)
            lineas.append(f"  {f.as_posix()}: {primera_linea(f.read_text(encoding='utf-8'))}")
    resto = [f for f in fuentes if f not in usados]
    if resto:
        lineas += ["", "Apoyo"] + [f"  {f.as_posix()}" for f in resto]
    return "\n".join(lineas) + "\n"


def indice(carpeta, archivos, lista):
    doc = Documento("Anexo técnico · Evaluación de cierre de la PNIG (2019–2025)",
                    titulo_doc="Anexo técnico de la evaluación de cierre de la PNIG (2019–2025)", informe=True)
    doc.titulo("Anexo técnico", "Evaluación de cierre de la Política Nacional de Igualdad de Género (2019–2025)",
               "Oficina General de Monitoreo, Evaluación de Políticas y Gestión Descentralizada · Noviembre de 2026")
    doc.h2("1. Propósito")
    doc.p("El informe de evaluación de cierre es autocontenido: cada hallazgo lleva su respaldo en una nota al pie y el "
          "Anexo 2 del informe presenta las fuentes de cada uno. Este anexo técnico entrega los materiales que permiten "
          "verificar esos respaldos y replicar los cálculos: los libros de análisis, los instrumentos aplicados, las "
          "entrevistas anonimizadas, el rastro de cada hallazgo, tabla y figura, el registro de las fuentes y el código "
          "de procesamiento.")
    doc.h2("2. Contenido")
    doc.p("La carpeta del anexo se organiza en diez subcarpetas, numeradas en el orden en que conviene consultarlas.")
    n_por = pd.Series([c for c, _, _, _ in archivos]).value_counts()
    doc.tabla(["Carpeta", "Contenido", "Relación con el informe", "Archivos"],
              [[c, t, r, str(int(n_por.get(c, 0)) + (3 if c == "00_indice" else 0))] for c, t, r in CARPETAS],
              [2300, 3900, 1900, 900], size=16)
    doc.h2("3. Libros de análisis")
    doc.p("Los libros llevan el nombre con el que los cita el informe. Sus hojas se renombraron sin las siglas de los "
          "instrumentos, de modo que cada cita «libro, hoja “…”» del Anexo 2 del informe corresponde a una hoja del "
          "libro entregado. Cada libro incluye una hoja de criterios cuando su clasificación depende de reglas.")
    doc.tabla(["Libro", "Contenido y relación con el informe"],
              [[a, desc] for c, a, desc, _ in archivos if c in ("04_matrices_y_libros", "05_entrevistas")
               and a.endswith((".xlsx", ".csv"))], [3000, 6000], size=16)
    doc.h2("4. Cómo seguir un hallazgo hasta su fuente")
    for paso in (
            "La nota al pie del hallazgo en el informe indica su identificador (por ejemplo, I1-H1), su nivel de "
            "confianza y las fuentes que lo respaldan.",
            "El Anexo 2 del informe presenta, para cada hallazgo, el libro y la hoja de cada evidencia.",
            "En 06_rastro, el libro Rastro_hallazgos.xlsx repite ese respaldo con los nombres publicados y con los del "
            "repositorio; Rastro_tablas_figuras.xlsx hace lo mismo para cada tabla y figura.",
            "En 04_matrices_y_libros se abre el libro y la hoja citados; las fuentes primarias se identifican en "
            "07_fuentes, con su enlace público o su vía de consulta."):
        doc.vineta(paso)
    doc.h2("5. Instrumentos y códigos")
    doc.p("El informe nombra los instrumentos por su nombre; los scripts, los libros de trabajo y el rastro usan un "
          "código corto. La correspondencia es la siguiente.")
    doc.tabla(["Código", "Instrumento"], [[c, n] for c, n in INSTRUMENTOS], [2000, 7000], size=16)
    doc.h2("6. Datos personales")
    doc.p("El anexo no contiene datos personales. Las transcripciones de las entrevistas están anonimizadas y "
          "atribuyen lo dicho al órgano informante; las grabaciones, las versiones con nombres y las notas en bruto no "
          "se entregan, conforme a la Ley N.° 29733, Ley de Protección de Datos Personales. La información reportada "
          "por las entidades se usa sin datos de las personas que la registraron.")
    doc.h2("7. Reproducción de los cálculos")
    doc.p("La subcarpeta 09_codigo contiene los scripts que generan las matrices, las figuras, el informe, la "
          "presentación y este anexo, con un archivo LEEME.txt que indica el orden de ejecución y los requisitos. Los "
          "microdatos del INEI y los datos abiertos del MEF no se incluyen por su tamaño; se descargan de sus portales. "
          "En los scripts, las rutas locales del equipo se reemplazaron por %USERPROFILE%.")
    doc.h2("8. Integridad de los archivos")
    doc.p(f"La lista de los {len(lista)} archivos del anexo, además de este índice, con su descripción, su origen en "
          f"el repositorio de la evaluación y su huella SHA-256, está en 00_indice/Lista_de_archivos.xlsx. Las subcarpetas con más "
          f"archivos son {lista_es([f'{c} ({int(n)})' for c, n in n_por.head(3).items()])}.")
    doc.guardar(carpeta / "Indice_del_anexo_tecnico.docx", comillas_inglesas=True)


if __name__ == "__main__":
    main()
