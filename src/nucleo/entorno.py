"""Rutas de las herramientas externas (Chrome, LibreOffice) y de los datos que viven fuera del repositorio (microdatos
del INEI, repositorio «Indicadores ENDES», carpeta administrativa), para que el repositorio funcione en otros equipos
sin tocar el código y sin rutas personales versionadas.

Prioridad: variable de entorno con el nombre de la clave en mayúsculas (CHROME, SOFFICE, MICRODATOS_INEI, REPO_ENDES,
CARPETA_ADMINISTRATIVA) > config/entorno.local.yaml (propio de cada equipo, no se versiona) > config/entorno.yaml
(versionado, sin rutas personales) > valor por defecto de Windows (solo herramientas).

Uso:
    from nucleo.entorno import herramienta, ruta_externa
    herramienta("chrome")                 # ruta de chrome.exe
    ruta_externa("microdatos_inei")       # carpeta de los microdatos, o None si el equipo no la tiene configurada
    contacto()["correo"]                  # especialista de contacto de los documentos (oficios, guías, memorándums)
"""
import os
from functools import lru_cache
from pathlib import Path

from .config import RAIZ_REPO

POR_DEFECTO = {
    "chrome": r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    "soffice": r"C:\Program Files\LibreOffice\program\soffice.exe",
}
ARCHIVOS = ("config/entorno.local.yaml", "config/entorno.yaml")   # el local tiene prioridad


@lru_cache(maxsize=None)
def _seccion(nombre):
    """Sección «herramientas» o «rutas», con el archivo local sobre el versionado."""
    import yaml
    datos = {}
    for rel in reversed(ARCHIVOS):
        ruta = RAIZ_REPO / rel
        if ruta.exists():
            with open(ruta, encoding="utf-8") as f:
                datos.update({k: v for k, v in ((yaml.safe_load(f) or {}).get(nombre) or {}).items() if v})
    return datos


def herramienta(nombre):
    """Ruta de la herramienta: «chrome» o «soffice»."""
    valor = os.environ.get(nombre.upper()) or _seccion("herramientas").get(nombre) or POR_DEFECTO[nombre]
    return str(Path(os.path.expandvars(valor)))


def ruta_externa(clave):
    """Carpeta de datos fuera del repositorio («microdatos_inei», «repo_endes», «carpeta_administrativa»), o None si
    no está configurada en este equipo."""
    valor = os.environ.get(clave.upper()) or _seccion("rutas").get(clave)
    return str(Path(os.path.expandvars(valor))) if valor else None


def contacto():
    """Especialista de contacto que figura en los documentos de la evaluación (nombre, correo y celular). Son datos
    personales: van solo en config/entorno.local.yaml (regla 6 de CLAUDE.md). Si faltan, el documento lleva
    [POR COMPLETAR] en su lugar."""
    datos = _seccion("contacto")
    return {c: datos.get(c) or f"[POR COMPLETAR: {c} de contacto]" for c in ("nombre", "correo", "celular")}
