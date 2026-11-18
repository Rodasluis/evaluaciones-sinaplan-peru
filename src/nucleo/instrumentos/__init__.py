"""Instrumentos genéricos de evaluación (Fase 4, ADR-017).

Cada instrumento lee entradas estándar en CSV (evaluaciones/<id>/02_insumos/ y, para las entrevistas,
05_entrevistas/), las valida y escribe su libro de análisis en 04_analisis/, con una hoja «Criterios» que separa lo que
exige el CEPLAN, lo que aporta la literatura y lo que decide el equipo evaluador. Las reglas vienen del núcleo
(escalas, evidencia, codificación, presupuesto) y sus parámetros, de config.yaml.

Uso (desde la raíz del repositorio):
    uv run python src/nucleo/instrumentos/correr.py <id> [instrumento ...] [--validar]

Los esquemas de las entradas están en plantillas/nueva_evaluacion/evaluacion/02_insumos/LEEME.md.
"""
from . import entrev, rep_ent, rev_ali, rev_doc, rev_evi, rev_ind

# orden de ejecución: la revisión de evidencia antes de la coherencia, que lee su plausibilidad
INSTRUMENTOS = {"REV_DOC": rev_doc, "REV_IND": rev_ind, "REV_EVI": rev_evi, "REV_ALI": rev_ali, "REP_ENT": rep_ent,
                "ENTREV": entrev}
# entradas comunes: la coherencia usa los servicios y los indicadores; la evidencia, los servicios
DEPENDE = {"REV_ALI": ["REV_DOC", "REV_IND"], "REV_EVI": ["REV_DOC"], "REP_ENT": ["REV_DOC"]}


def activos(cfg):
    """Instrumentos de la evaluación que tienen versión genérica, en el orden del núcleo."""
    modalidad = cfg.get("modalidad") or {}
    lista = [k for k in INSTRUMENTOS if k in (cfg.get("instrumentos") or {})]
    if not modalidad.get("entrevistas", True):
        lista = [k for k in lista if k != "ENTREV"]
    if not (cfg.get("escalas") or {}).get("implementacion"):
        lista = [k for k in lista if k != "REV_DOC"]   # el estado de implementación no aplica (p. ej., diseño)
    return lista


def plantillas_entrada(cfg, instrumentos):
    """{(carpeta, archivo): columnas} de las entradas de los instrumentos dados y de las que estos usan."""
    usados = list(dict.fromkeys(d for k in instrumentos if k in INSTRUMENTOS for d in DEPENDE.get(k, []) + [k]))
    salida = {}
    for k in usados:
        for carpeta, nombre, columnas in INSTRUMENTOS[k].entradas(cfg):
            salida.setdefault((carpeta, nombre), columnas)
    return salida
