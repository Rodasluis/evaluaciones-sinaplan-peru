"""Ubicación de citas literales en sus documentos (generalizada de src/analisis_pnig/rev_ali_coherencia_alineamiento.py).

Una cita se busca normalizada (ligaduras, cortes de palabra al fin de línea, mayúsculas y espacios) en cada página de
un PDF (pdftotext en dos lecturas: con disposición y en orden de lectura), en cada párrafo de un .docx o en un .txt
(las páginas se separan con salto de página, \\f). Devuelve las páginas o párrafos donde aparece.
"""
import re
import shutil
import subprocess
import unicodedata
from functools import lru_cache
from pathlib import Path


def norm(texto):
    t = unicodedata.normalize("NFKC", str(texto)).replace("­", "")
    t = re.sub(r"(\w)-\s*\n\s*(\w)", r"\1\2", t)
    t = t.replace("“", '"').replace("”", '"').replace("«", '"').replace("»", '"').replace("’", "'")
    return re.sub(r"\s+", " ", t).strip().lower()


@lru_cache(maxsize=None)
def partes(ruta):
    """Texto normalizado de cada página (PDF, .txt) o párrafo (.docx), como tuplas de lecturas alternativas."""
    ruta = Path(ruta)
    suf = ruta.suffix.lower()
    if suf == ".pdf":
        if not shutil.which("pdftotext"):
            raise RuntimeError("pdftotext no está disponible para ubicar citas en PDF")
        lecturas = []
        for modo in (["-layout"], ["-raw"]):
            txt = subprocess.run(["pdftotext", "-enc", "UTF-8", *modo, str(ruta), "-"], capture_output=True).stdout
            lecturas.append([norm(p) for p in txt.decode("utf-8", "ignore").split("\f")])
        return tuple(zip(*lecturas))
    if suf == ".docx":
        import docx
        return tuple((norm(p.text),) for p in docx.Document(str(ruta)).paragraphs)
    if suf in (".txt", ".md"):
        return tuple((norm(p),) for p in ruta.read_text(encoding="utf-8").split("\f"))
    raise ValueError(f"formato sin lector de citas: {ruta.name}")


def ubicar(cita, ruta):
    """Páginas (PDF, .txt) o párrafos (.docx) donde aparece la cita, numerados desde 1."""
    c = norm(cita)
    return [i + 1 for i, lecturas in enumerate(partes(str(ruta))) if any(c in x for x in lecturas)]
