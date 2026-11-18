"""Configuración de una evaluación: lee evaluaciones/<id>/config.yaml y entrega sus rutas y parámetros.

Uso:
    from nucleo.config import cargar
    cfg = cargar()                 # la evaluación de la variable EVALUACION, o pnig_2026
    cfg.raiz                       # "evaluaciones/pnig_2026" (relativa a la raíz del repositorio)
    cfg.ruta("analisis")           # "evaluaciones/pnig_2026/04_analisis"
    cfg.datos["politica"]["sigla"] # cualquier dato del config.yaml

Las rutas son relativas a la raíz del repositorio, con barras normales, porque los scripts se ejecutan desde ella y
algunas quedan escritas en los productos (hojas de criterios, rastro).
"""
import os
from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path

RAIZ_REPO = Path(__file__).resolve().parents[2]
POR_DEFECTO = "pnig_2026"
# subcarpetas de toda evaluación (plantillas/nueva_evaluacion)
CARPETAS = {
    "encargo": "00_encargo",
    "insumos": "02_insumos",
    "datos_raw": "03_datos/raw",
    "datos_procesado": "03_datos/procesado",
    "analisis": "04_analisis",
    "entrevistas": "05_entrevistas",
    "informe": "06_informe",
    "anexo": "07_anexo_tecnico",
}


@dataclass(frozen=True)
class Config:
    id: str
    raiz: str
    datos: dict = field(default_factory=dict, hash=False, compare=False)
    repo: Path = field(default=RAIZ_REPO, compare=False)   # raíz del repositorio (otra en las pruebas)

    def ruta(self, carpeta, *partes):
        """Ruta de una subcarpeta de la evaluación (clave de CARPETAS o nombre literal) y, opcionalmente, un archivo."""
        base = f"{self.raiz}/{CARPETAS.get(carpeta, carpeta)}"
        return "/".join([base, *partes]) if partes else base

    def abs(self, carpeta, *partes):
        """Ruta absoluta de ruta()."""
        return self.repo / self.ruta(carpeta, *partes)

    def __getitem__(self, clave):
        return self.datos[clave]

    def get(self, clave, defecto=None):
        return self.datos.get(clave, defecto)


def insumo(nombre, evaluacion=None):
    """Parámetros del caso guardados en evaluaciones/<id>/02_insumos/ (YAML o CSV): ajustes, verificaciones, catálogos
    y datos que no son código (Fase 2, etapa 3)."""
    ruta = RAIZ_REPO / cargar(evaluacion).ruta("insumos", nombre)
    if ruta.suffix == ".csv":
        import pandas as pd
        return pd.read_csv(ruta, dtype=str).fillna("")
    import yaml
    with open(ruta, encoding="utf-8") as f:
        return yaml.safe_load(f)


@lru_cache(maxsize=None)
def cargar(evaluacion=None):
    return cargar_desde(RAIZ_REPO, evaluacion or os.environ.get("EVALUACION") or POR_DEFECTO)


def cargar_desde(repo, ev):
    """Configuración de evaluaciones/<ev>/ bajo otra raíz (sin caché; la usan las pruebas y el generador)."""
    repo = Path(repo)
    raiz = f"evaluaciones/{ev}"
    archivo = repo / raiz / "config.yaml"
    if not archivo.exists():
        raise FileNotFoundError(f"No existe {archivo}: cree la evaluación con src/nucleo/nueva_evaluacion.py")
    import yaml
    with open(archivo, encoding="utf-8") as f:
        datos = yaml.safe_load(f) or {}
    if datos.get("id", ev) != ev:
        raise ValueError(f"{archivo}: el id ({datos.get('id')}) no coincide con la carpeta ({ev})")
    return Config(id=ev, raiz=raiz, datos=datos, repo=repo)
