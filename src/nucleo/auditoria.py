"""Auditorías de los productos y del repositorio (núcleo).

Informe (Word):
  - leyendas_sin_analisis(): toda tabla y figura del cuerpo tiene después un párrafo propio de análisis;
  - estilo(): guiones explicativos, preguntas retóricas fuera de las de evaluación, comillas « » sin convertir y
    lectura de figuras por su disposición («Cada fila es…», «a la izquierda…»);
  - terminos(): términos internos (regex de la evaluación) en el cuerpo y las notas al pie.
Archivos (antes de un commit o de una entrega):
  - archivos_sensibles(): nombres de las categorías de la regla 6 de CLAUDE.md (CV, órdenes de servicio, recibos,
    contactos, puntos focales, credenciales), archivos de más de 5 MB y rutas locales con el nombre de usuario;
  - datos_personales(): correos, DNI y celulares en textos y libros.

Uso:
    uv run python src/nucleo/auditoria.py informe ruta/al/informe.docx
    uv run python src/nucleo/auditoria.py commit         # archivos preparados para el commit (lo usa .githooks/pre-commit)
    uv run python src/nucleo/auditoria.py entrega carpeta/ # carpeta que se va a entregar (anexo técnico, versión pública)
"""
import os
import re
import subprocess
import sys
from pathlib import Path

SENSIBLES = re.compile(r"(^|[/_\-\s])(cv|curr[ií]cul|orden(es)?[_\- ]de[_\- ]servicio|recibo|contacto|punto[s]?[_\- ]focal|"
                       r"directorio[_\- ]destinatario|credencial|contrase|password|\.env$|token|semilla)", re.I)
MAX_BYTES = 5_000_000
CORREO = re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+")
# DNI y celular solo junto a su rótulo: los números de 8 o 9 dígitos de las bases (poblaciones, montos) no son datos
# personales
DNI = re.compile(r"\bDNI\b\D{0,12}(\d{8})(?!\d)", re.I)
CELULAR = re.compile(r"\b(?:cel(?:ular)?|tel(?:éfono)?|m[oó]vil|whatsapp)\b\D{0,12}(9\d{8})(?!\d)", re.I)
# celular escrito como cadena exacta en el código o en un CSV, sin rótulo: "9xxxxxxxx"
CELULAR_CADENA = re.compile(r"""(?<![\w])(["'])9\d{8}\1""")
COLUMNA_PERSONAL = re.compile(r"^(dni|correo|e-?mail|tel[eé]fono|celular|nombres? y apellidos|apellidos?)$", re.I)
RETORICA = re.compile(r"(^|[.!?]\s+)¿[^?]+\?")
DISPOSICION = re.compile(r"^Cada (fila|columna|barra|panel|tarjeta) |\ba la izquierda\b|\bse recorre de arriba\b")
GUION = re.compile(r" — | - ")
TEXTO = {".py", ".md", ".txt", ".csv", ".tsv", ".yaml", ".yml", ".json", ".ps1", ".html"}


def _cuerpo(doc):
    """(estilo, texto) de los párrafos del cuerpo hasta los anexos, en orden, con las tablas como «[tabla]»."""
    from docx.text.paragraph import Paragraph
    salida, en_cuerpo = [], False
    for el in doc.element.body.iterchildren():
        if el.tag.endswith("}p"):
            p = Paragraph(el, doc)
            est = p.style.name if p.style is not None else ""
            t = p.text.strip()
            if est.startswith("Heading 1") and re.match(r"^[1-9]\. ", t):
                en_cuerpo = True
            if est.startswith("Heading 1") and (t.startswith("Anexo") or t.startswith("Referencias")):
                break
            if en_cuerpo:
                salida.append((est, t))
        elif el.tag.endswith("}tbl") and en_cuerpo:
            salida.append(("tabla", "[tabla]"))
    return salida


def leyendas_sin_analisis(ruta, min_caracteres=150):
    import docx
    pend, falta, n = None, [], 0
    for est, t in _cuerpo(docx.Document(ruta)):
        if re.match(r"^(Figura|Tabla) \d+\. ", t):
            if pend:
                falta.append(pend)
            pend, n = t[:90], n + 1
        elif pend and len(t) > min_caracteres and not t.startswith(("Fuente", "Nota")):
            pend = None
    if pend:
        falta.append(pend)
    return n, falta


def estilo(ruta, excepciones=("Pregunta ",)):
    import docx
    hallados = []
    for est, t in _cuerpo(docx.Document(ruta)):
        if not t or t.startswith(excepciones) or est == "tabla":
            continue
        for nombre, rx in (("pregunta retórica", RETORICA), ("lectura por disposición", DISPOSICION),
                           ("guion explicativo", GUION), ("comillas « »", re.compile(r"[«»]"))):
            m = rx.search(t)
            if m:
                hallados.append((nombre, t[max(0, m.start() - 40):m.end() + 40]))
    return hallados


def terminos(ruta, internos):
    import docx
    doc = docx.Document(ruta)
    malos = [t[max(0, m.start() - 40):m.end() + 30] for _, t in _cuerpo(doc) for m in internos.finditer(t)]
    for parte in doc.part.package.parts:
        if str(parte.partname).endswith("footnotes.xml"):
            txt = re.sub(r"<[^>]+>", "", parte.blob.decode("utf-8"))
            malos += [txt[max(0, m.start() - 40):m.end() + 30] for m in internos.finditer(txt)]
    return malos


def archivos_sensibles(rutas, usuario=None, max_bytes=MAX_BYTES, sin_revisar_rutas=()):
    """Problemas de una lista de archivos: nombre sensible, tamaño (si max_bytes) y rutas locales con el nombre de
    usuario (salvo en sin_revisar_rutas: CLAUDE.md nombra la carpeta administrativa a propósito)."""
    usuario = usuario or os.environ.get("USERNAME") or os.environ.get("USER") or ""
    marcas = [f"C:/Users/{usuario}", f"C:\\Users\\{usuario}", f"C:\\\\Users\\\\{usuario}"] if usuario else []
    problemas = []
    for r in map(Path, rutas):
        if SENSIBLES.search(r.as_posix()):
            problemas.append((str(r), "nombre de una categoría sensible (regla 6 de CLAUDE.md)"))
        if not r.exists():
            continue
        if max_bytes and r.stat().st_size > max_bytes:
            problemas.append((str(r), f"pesa {r.stat().st_size / 1e6:.1f} MB (más de 5 MB)"))
        if r.suffix.lower() in TEXTO and marcas and r.as_posix() not in sin_revisar_rutas:
            try:
                texto = r.read_text(encoding="utf-8", errors="ignore")
            except OSError:
                continue
            if any(m in texto for m in marcas):
                problemas.append((str(r), "contiene una ruta local con el nombre de usuario"))
    return problemas


def datos_personales(rutas):
    """Correos, DNI y celulares (junto a su rótulo) en textos y en las celdas de texto de los libros de Excel, y
    columnas con nombre de dato personal (DNI, correo, teléfono, nombres y apellidos)."""
    hallados = []
    for r in map(Path, rutas):
        textos = []
        if r.suffix.lower() in TEXTO | {".txt"}:
            textos = [r.read_text(encoding="utf-8", errors="ignore")]
        elif r.suffix.lower() == ".xlsx":
            import openpyxl
            wb = openpyxl.load_workbook(r, read_only=True, data_only=True)
            for ws in wb.worksheets:
                filas = ws.iter_rows(values_only=True)
                cabecera = next(filas, ()) or ()
                hallados += [(str(r), "columna", f"{ws.title}: {c}") for c in cabecera
                             if isinstance(c, str) and COLUMNA_PERSONAL.match(c.strip())]
                textos.append(" ".join(c for fila in filas for c in fila if isinstance(c, str)))
        for t in textos:
            for nombre, rx in (("correo", CORREO), ("DNI", DNI), ("celular", CELULAR), ("celular", CELULAR_CADENA)):
                hallados += [(str(r), nombre, m.group(0)) for m in rx.finditer(t)]
    return hallados


def _preparados():
    salida = subprocess.run(["git", "diff", "--cached", "--name-only", "--diff-filter=ACMR"], capture_output=True,
                            text=True, check=True).stdout
    return [x for x in salida.splitlines() if x.strip()]


def main(argv):
    if not argv:
        print(__doc__)
        return 0
    modo = argv[0]
    if modo == "informe":
        ruta = argv[1]
        n, falta = leyendas_sin_analisis(ruta)
        est = estilo(ruta)
        print(f"{n} leyendas en el cuerpo; {len(falta)} sin párrafo de análisis propio")
        for f in falta:
            print("  sin análisis:", f)
        for nombre, frag in est:
            print(f"  {'aviso, ' if nombre == 'guion explicativo' else ''}{nombre}: …{frag}…")
        # los guiones son avisos: los nombres oficiales de algunos servicios los llevan
        return 1 if falta or any(n != "guion explicativo" for n, _ in est) else 0
    if modo == "commit":
        # el tamaño solo avisa: el informe y las fuentes oficiales pesan más de 5 MB; bloquean las categorías
        # sensibles y las rutas con el nombre de usuario
        todos = archivos_sensibles(_preparados(), sin_revisar_rutas=("CLAUDE.md",))
        problemas = [(r, m) for r, m in todos if " MB " not in m]
        # datos personales en los textos preparados: celulares y DNI bloquean; los correos solo avisan (hay correos
        # institucionales públicos en las fuentes)
        textos = [r for r in _preparados() if Path(r).suffix.lower() in TEXTO and Path(r).exists()]
        for r, tipo, _valor in datos_personales(textos):
            motivo = f"contiene un {tipo}" if tipo != "correo" else "contiene un correo"
            todos.append((r, motivo))
            if tipo != "correo":
                problemas.append((r, motivo))
        for r, motivo in todos:
            print(f"{'BLOQUEADO' if (r, motivo) in problemas else 'AVISO'} {r}: {motivo}")
        if problemas:
            print("Saque esos archivos del commit (git restore --staged <archivo>) o muévalos fuera del repositorio.")
        return 1 if problemas else 0
    if modo == "entrega":
        rutas = [p for p in Path(argv[1]).rglob("*") if p.is_file()]
        problemas = archivos_sensibles(rutas, max_bytes=None)   # en una entrega no importa el tamaño
        personales = datos_personales(rutas)
        for r, motivo in problemas:
            print(f"REVISAR {r}: {motivo}")
        for r, tipo, valor in personales[:50]:
            print(f"DATO PERSONAL {r}: {tipo} {valor}")
        print(f"{len(rutas)} archivos revisados; {len(problemas)} problemas; {len(personales)} posibles datos personales")
        return 1 if problemas or personales else 0
    print(__doc__)
    return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
