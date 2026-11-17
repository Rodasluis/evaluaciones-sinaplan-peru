"""Fuentes y archivos: huella SHA-256, nombres de archivo seguros y conversión a PDF con LibreOffice."""
import hashlib
import re
import subprocess
import unicodedata
from pathlib import Path

from .entorno import herramienta


def sha256(ruta):
    """Huella SHA-256 de un archivo, leída por bloques."""
    h = hashlib.sha256()
    with open(ruta, "rb") as f:
        for bloque in iter(lambda: f.read(1 << 20), b""):
            h.update(bloque)
    return h.hexdigest()


def slug(texto, n=60):
    """Texto en minúsculas, sin tildes y con guiones, para nombres de archivo («Figura 1. Hitos…» -> «figura-1-hitos»)."""
    t = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]+", "-", t).strip("-")[:n].rstrip("-")


def a_pdf(archivos, destino):
    """Convierte a PDF con LibreOffice, llamando a soffice.exe directamente (con PowerShell se pierden las comillas de
    las rutas con espacios)."""
    destino = Path(destino)
    destino.mkdir(parents=True, exist_ok=True)
    subprocess.run([herramienta("soffice"), "--headless", "--convert-to", "pdf", "--outdir", str(destino)]
                   + [str(a) for a in archivos], check=True, capture_output=True)
