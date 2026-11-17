"""Huellas de los productos del piloto de la PNIG, para comprobar que la Fase 2 no los cambia.

- Libros de análisis (.xlsx): una huella por hoja, del contenido como texto (no del archivo, que guarda fechas).
- CSV de análisis y del informe: huella del contenido con saltos de línea normalizados.
- Informe: texto de los párrafos, las tablas y las notas al pie, guardado completo para ver las diferencias.
- Figuras: huella del HTML de cada una (el PNG depende del motor de Chrome).
- Nota, presentaciones y logos: huella del archivo, porque no se generan con código.

Uso: python tests/referencia.py capturar   (reescribe tests/referencia_pnig/ con el estado actual)

En la versión pública, tests/referencia_pnig/excluidos_publico.txt lista los archivos que no se publican (nota,
registros de entrevistas, logos); solo si ese archivo existe se admiten como ausentes (ADR-014).
"""
import hashlib
import json
import sys
import zipfile
from pathlib import Path

import pandas as pd

RAIZ = Path(__file__).resolve().parents[1]
EVAL = RAIZ / "evaluaciones/pnig_2026"
REF = Path(__file__).resolve().parent / "referencia_pnig"
INFORME = EVAL / "06_informe/2026_mimp-ogmepgd_informe-evaluacion-cierre-pnig.docx"
ARCHIVOS_FIJOS = [  # no los genera el código: deben quedar intactos
    EVAL / "00_encargo/nota_metodologica/Nota Metodologica Evaluación PNIG - evaluación de cierre.docx",
    EVAL / "00_encargo/nota_metodologica/Nota Metodologica Evaluación PNIG - presentación.pptx",
    EVAL / "00_encargo/nota_metodologica/Nota Metodologica Evaluación PNIG - presentación resumida.pptx",
    RAIZ / "assets/logos/logo_institucion.png",
    RAIZ / "assets/logos/logo_gobierno.png",
]


CRLF = "\r\n"


def _h(datos):
    return hashlib.sha256(datos if isinstance(datos, bytes) else datos.encode("utf-8")).hexdigest()


def excluidos():
    """Archivos de la referencia que la versión pública no lleva; en el repositorio privado, ninguno."""
    ruta = REF / "excluidos_publico.txt"
    if not ruta.exists():
        return set()
    return {x.strip() for x in ruta.read_text(encoding="utf-8").splitlines() if x.strip() and not x.startswith("#")}


def _presente(r):
    return r.relative_to(RAIZ).as_posix() not in excluidos()


def huellas_libros():
    libros = sorted((EVAL / "04_analisis").glob("*.xlsx")) + [EVAL / "05_entrevistas/ENTREV_registro_codificacion.xlsx"]
    libros = [x for x in libros if _presente(x)]
    salida = {}
    for libro in libros:
        hojas = pd.read_excel(libro, sheet_name=None, header=None, dtype=str, keep_default_na=False, engine="openpyxl")
        # los saltos de línea dentro de las celdas se normalizan: git puede convertir LF en CRLF en los CSV de entrada
        salida[libro.relative_to(RAIZ).as_posix()] = {
            h: [int(df.shape[0]), int(df.shape[1]),
                _h(df.replace(CRLF, "\n", regex=True).to_csv(index=False, header=False))] for h, df in hojas.items()}
    return salida


def huellas_csv():
    rutas = (sorted((EVAL / "04_analisis").rglob("*.csv")) + sorted((EVAL / "06_informe").glob("*.csv")) +
             sorted((EVAL / "05_entrevistas").glob("*.csv")))
    return {r.relative_to(RAIZ).as_posix(): _h(r.read_bytes().replace(b"\r\n", b"\n")) for r in rutas}


def huellas_figuras():
    return {r.name: _h(r.read_bytes().replace(b"\r\n", b"\n"))
            for r in sorted((EVAL / "06_informe/figuras").glob("*.html"))}


def huellas_fijos():
    return {r.relative_to(RAIZ).as_posix(): _h(r.read_bytes()) for r in ARCHIVOS_FIJOS if _presente(r)}


def texto_informe(ruta=INFORME):
    """Párrafos y celdas del cuerpo, en orden, y las notas al pie (sin campos de índice ni números de página)."""
    import docx
    d = docx.Document(ruta)
    lineas = []
    cuerpo = d.element.body
    for el in cuerpo.iterchildren():
        if el.tag.endswith("}p"):
            t = "".join(x.text or "" for x in el.iter() if x.tag.endswith("}t"))
            lineas.append(t)
        elif el.tag.endswith("}tbl"):
            for tr in el.iter():
                if tr.tag.endswith("}tr"):
                    celdas = []
                    for tc in tr.iterchildren():
                        if tc.tag.endswith("}tc"):
                            celdas.append(" ¶ ".join("".join(x.text or "" for x in p.iter() if x.tag.endswith("}t"))
                                                     for p in tc.iter() if p.tag.endswith("}p")))
                    lineas.append(" | ".join(celdas))
    with zipfile.ZipFile(ruta) as z:
        if "word/footnotes.xml" in z.namelist():
            from lxml import etree
            raiz = etree.fromstring(z.read("word/footnotes.xml"))
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            for fn in raiz.findall("w:footnote", ns):
                t = "".join(fn.itertext()).strip()
                if t:
                    lineas.append(f"[nota {fn.get('{%s}id' % ns['w'])}] {t}")
    # los índices llevan números de página, que dependen de Word: se quitan los tabuladores y lo que sigue
    return "\n".join(l.split("\t")[0] for l in lineas)


def capturar():
    REF.mkdir(exist_ok=True)
    datos = {"libros": huellas_libros(), "csv": huellas_csv(), "figuras": huellas_figuras(), "fijos": huellas_fijos()}
    (REF / "huellas.json").write_text(json.dumps(datos, ensure_ascii=False, indent=1, sort_keys=True), encoding="utf-8")
    (REF / "informe_texto.txt").write_text(texto_informe(), encoding="utf-8")
    print(f"Referencia capturada: {len(datos['libros'])} libros, {len(datos['csv'])} CSV, {len(datos['figuras'])} "
          f"figuras, {len(datos['fijos'])} archivos fijos e informe ({INFORME.name})")


if __name__ == "__main__" and sys.argv[1:] == ["capturar"]:
    capturar()
