"""REP_ENT: unidades de codificación (segmentos) de lo que las entidades reportaron al MIMP sobre la PNIG.

Fuentes y unidad:
  - formulario de seguimiento de la OMEP 2023–2025 (panel armonizado): cada campo abierto con texto de cada registro;
  - informes LIO 2019–2025 (secciones exportadas por src/insumos_seguimiento/cobertura_lio.py): cada párrafo de la
    sección del servicio que contiene palabras de los bloques D, E o F (factores, articulación, sostenibilidad);
  - informes de gobernanza de la CONAIG I a IV: cada párrafo con palabras del bloque E.
La preselección por palabras clave reduce la lectura; los párrafos que no la superan quedan en la sección y se resumen
en el bloque C de la matriz por servicio. Cada segmento conserva su página para citarlo.

Uso (desde la raíz del repositorio):
    uv run --no-project --with pandas python src/analisis_pnig/rep_ent_segmentos.py
"""
import re
import subprocess
from pathlib import Path

import pandas as pd
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

RAIZ = Path(EVAL)
PANEL = RAIZ / "03_datos/procesado/panel_formulario_seguimiento_pnig_2023_2025.csv"
SECCIONES = RAIZ / "03_datos/procesado/secciones_lio_2019_2025.csv"
CONAIG = Path("docs/fuentes/pnig/09_conaig_gobernanza")
SALIDA = RAIZ / "03_datos/procesado/rep_ent_segmentos.csv"

# Campos abiertos del formulario: (columna del panel, rótulo, bloque de REP_ENT)
CAMPOS = [
    ("motivo_no_reporte", "Motivo de no reporte", "B"),
    ("falla_arriba", "Causa de la posible falla hacia arriba", "B"),
    ("falla_abajo", "Causa de la posible falla hacia abajo", "B"),
    ("factores_favorables", "Factores favorables", "D"),
    ("factores_dificultad", "Factores de dificultad", "D"),
    ("buenas_practicas", "Buenas prácticas", "D"),
    ("medidas", "Medidas adoptadas o propuestas", "D"),
    ("dificultad_coordinacion", "Dificultad de coordinación", "E"),
    ("causa_dificultad", "Causa de la dificultad de coordinación", "E"),
    ("contexto", "Contexto", "F"),
    ("efectos_no_previstos", "Efectos no previstos", "F"),
    ("oportunidades_mejora", "Oportunidades de mejora", "F"),
]
# Respuestas sin contenido codificable, incluidas las de sí o no de las preguntas cerradas que se guardaron como texto
VACIOS = re.compile(r"^\s*(-+|\.|s[ií]|no|ninguno|ninguna|n/?a|no aplica|sin comentarios?|no se presentaron?\.?)\s*$", re.I)
NADA = re.compile(r"^(ninguno|ninguna|no se (ha|han) identificado|no se registra|no aplica|sin informaci[oó]n)\.?$|"
                  r"no requiere ser completad", re.I)
# Datos de personas: el protocolo de REP_ENT excluye los datos de quienes registraron o figuran como contacto
PERSONALES = [(re.compile(r"(contacto del MIMP\s*:?\s*)[^)\n]+", re.I), r"\1[nombre omitido]"),
              (re.compile(r"[\w.+-]+@[\w-]+\.[\w.]+"), "[correo omitido]"),
              (re.compile(r"(?<!\d)9\d{2}\s?\d{3}\s?\d{3}(?!\d)"), "[teléfono omitido]")]


def vacio(v):
    """Sin contenido: vacío, sí o no, o un campo de contexto cuyas tres partes dicen «ninguno» o equivalente."""
    t = re.sub(r"Contexto (econ[oó]mico|pol[ií]tico|social)\s*:", " ", v, flags=re.I)
    partes = [re.sub(r"^[\s\-–•.:]+|[\s.]+$", "", x) for x in t.split("|")]
    return all(len(x) < 4 or VACIOS.match(x) or NADA.search(x) for x in partes)


def anonimizar(v):
    for patron, reemplazo in PERSONALES:
        v = patron.sub(reemplazo, v)
    return v
# Palabras de los bloques D, E y F. Se excluyen a propósito términos que aparecen en casi todo párrafo descriptivo
# («presupuesto», «personal», «fortalecimiento», «mejora»): con ellos la preselección triplicaba los párrafos sin
# añadir factores (prueba del 2026-10-30: 3278 párrafos frente a 1073).
CLAVES_LIO = re.compile(
    r"dificult|limitac|obstacul|restriccion|insuficien|carenc|desafio|retos? (para|en|de)|articulac|coordinac|conaig|"
    r"comision multisectorial|recomend|propuesta|sostenib|pandemia|covid|emergencia sanitaria|"
    r"no (remitio|reporto|brindo|envio|presento)|rotacion|brecha de personal|factores?|lecciones|buenas? practicas?", re.I)
CLAVES_CONAIG = re.compile(
    r"dificult|limitac|obstacul|desafio|recomend|propuesta de mejora|propone|conclusi|leccion|buenas? practicas?|"
    r"sesion(es)? (ordinaria|extraordinaria|de la)|acuerdos? (de|adoptad)|rectoria|nudos? critico|factores?", re.I)
MIN_CARACTERES = 120


def sin_tildes(t):
    import unicodedata
    return unicodedata.normalize("NFKD", t).encode("ascii", "ignore").decode()


def parrafos(texto, pagina):
    """Divide un texto de pdftotext -layout en párrafos y devuelve (página, párrafo) uniendo los cortes de línea."""
    for bloque in re.split(r"\n\s*\n", texto):
        m = re.findall(r"\[\[p\. (\d+)\]\]", bloque)
        p0 = pagina
        if m:
            pagina = int(m[-1])
            # si el bloque empieza con el marcador, el párrafo es de la página nueva
            if bloque.strip().startswith("[[p."):
                p0 = int(m[0])
        limpio = re.sub(r"\[\[p\. \d+\]\]", " ", bloque)
        limpio = re.sub(r"(\w)-\s*\n\s*(\w)", r"\1\2", limpio)
        limpio = re.sub(r"\s+", " ", limpio).strip()
        if limpio:
            yield p0, limpio


filas = []
# Formulario
panel = pd.read_csv(PANEL, dtype=str)
for _, r in panel.iterrows():
    for col, rotulo, bloque in CAMPOS:
        v = r.get(col)
        if isinstance(v, str) and v.strip() and not vacio(v):
            filas.append({"id": f"FORM-{r.anio}-{r.id.replace('PNIG_', '')}-{col}", "fuente": f"Formulario {r.anio}",
                          "anio": int(r.anio), "pagina": None, "servicio": r.servicio, "indicador": r.id.replace("PNIG_", ""),
                          "entidad": r.entidad, "campo": rotulo, "bloque": bloque, "texto": anonimizar(v.strip())})
# Informes LIO
sec = pd.read_csv(SECCIONES)
for _, s in sec.iterrows():
    n = 0
    for k, (pag, par) in enumerate(parrafos(s.texto, int(s.pagina_inicio))):
        if k == 0 or len(par) < MIN_CARACTERES or not CLAVES_LIO.search(sin_tildes(par)):
            continue                      # el primer párrafo es el encabezado del servicio
        n += 1
        filas.append({"id": f"LIO-{s.anio}-{s.servicio}-{n:02d}", "fuente": f"Informe LIO {s.anio}", "anio": int(s.anio),
                      "pagina": pag, "servicio": s.servicio, "indicador": None, "entidad": None,
                      "campo": "Sección del servicio", "bloque": "D/E/F", "texto": anonimizar(par)})
# Informes de gobernanza de la CONAIG
for pdf in sorted(CONAIG.glob("*.pdf")):
    num = re.search(r"_(i{1,3}|iv)-informe", pdf.name).group(1).upper()
    anio = int(pdf.name[:4])
    n = 0
    todo = subprocess.run(["pdftotext", "-enc", "UTF-8", "-layout", str(pdf), "-"],
                          capture_output=True).stdout.decode("utf-8", "ignore")
    for p, txt in enumerate(todo.split("\f"), start=1):    # pdftotext separa las páginas con «\f»
        for _, par in parrafos(txt, p):
            if len(par) < MIN_CARACTERES or not CLAVES_CONAIG.search(sin_tildes(par)):
                continue
            n += 1
            filas.append({"id": f"CONAIG-{anio}-{n:03d}", "fuente": f"{num} Informe de gobernanza de la CONAIG", "anio": anio,
                          "pagina": p, "servicio": None, "indicador": None, "entidad": None,
                          "campo": "Informe de gobernanza", "bloque": "E", "texto": anonimizar(par)})

seg = pd.DataFrame(filas)
assert seg.id.is_unique
seg.to_csv(SALIDA, index=False, encoding="utf-8-sig")
print(seg.groupby("fuente").size().to_string())
print("total", len(seg))
