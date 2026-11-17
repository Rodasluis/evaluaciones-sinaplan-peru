"""Piezas comunes de la capa de evaluación: configuración, hallazgos y rutas (plantilla de src/nucleo/nueva_evaluacion.py).

La evaluación es la de la carpeta: src/analisis_<id>/ trabaja sobre evaluaciones/<id>/.
"""
import os
import re
import sys
from pathlib import Path

AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI.parent))   # src/, para importar el núcleo
from nucleo import hallazgos as NH  # noqa: E402
from nucleo.config import RAIZ_REPO, cargar  # noqa: E402

ID = AQUI.name.removeprefix("analisis_")
os.chdir(RAIZ_REPO)                     # las rutas de la configuración son relativas a la raíz del repositorio
CFG = cargar(ID)
# términos internos que no van en el texto publicable: códigos de instrumento e identificadores de hallazgo (ADR-010)
INTERNOS = re.compile(r"\b(?:" + "|".join(map(re.escape, CFG["instrumentos"])) + r")\b|\b[A-Z]\d+-H\d+\b")
ASPECTOS = ("Implementación", "Diseño", "Seguimiento y evaluación")   # Tabla 5 de la guía (p. 57)


def _tipologia():
    import yaml
    with open(RAIZ_REPO / "tipologias" / CFG["tipologia"] / "estructura.yaml", encoding="utf-8") as f:
        return yaml.safe_load(f)["nombre"]


TIPOLOGIA = _tipologia()   # «Evaluación de cierre», «Evaluación de diseño», …


def ruta(carpeta, *partes):
    return RAIZ_REPO / CFG.ruta(carpeta, *partes)


def leer_hallazgos():
    """preguntas (valoración y respuesta), hallazgos (con sus evidencias como en nucleo.hallazgos) y recomendaciones de
    04_analisis/hallazgos.yaml."""
    import yaml
    with open(ruta("analisis", "hallazgos.yaml"), encoding="utf-8") as f:
        datos = yaml.safe_load(f) or {}
    campos = ("aspecto", "enunciado", "estado", "detalle", "salvedad", "nota")
    H = []
    for h in datos.get("hallazgos") or []:
        h = {**{c: "" for c in campos}, **h}
        h = {k: ("" if v is None else v) for k, v in h.items()}
        h["evidencias"] = [NH.evidencia(e["plano"], e["instrumento"], e["texto"], e["fuente"], e["base"],
                                        int(e.get("n_fuentes", 1))) for e in (h.get("evidencias") or [])]
        H.append(h)
    P = {k: {"valoracion": (v or {}).get("valoracion") or "", "respuesta": (v or {}).get("respuesta") or ""}
         for k, v in (datos.get("preguntas") or {}).items()}
    return P, H, list(datos.get("recomendaciones") or [])


def contar(h):
    return NH.fuentes_independientes(h)


def verificar(P, H, R):
    """Reglas del núcleo (nucleo.hallazgos.verificar) y las de esta capa: valoraciones de la escala, instrumentos de la
    configuración y recomendaciones con aspecto de la Tabla 5 y hallazgos existentes."""
    preguntas = CFG["preguntas"]
    NH.verificar(H, preguntas, contar, internos=INTERNOS)
    escala = CFG["escalas"]["valoracion"]
    for k, v in P.items():
        assert k in preguntas, f"valoración de una pregunta que no está en config.yaml: {k}"
        assert v["valoracion"] in escala + [""], f"{k}: valoración fuera de la escala ({v['valoracion']})"
    for h in H:
        for e in h["evidencias"]:
            assert e["instrumento"] in CFG["instrumentos"], f"{h['id']}: instrumento {e['instrumento']} sin definir"
            assert e["instrumento"] in preguntas[h["pregunta"]]["instrumentos"], (
                f"{h['id']}: {e['instrumento']} no es instrumento de la pregunta {h['pregunta']} en config.yaml")
    ids = {h["id"] for h in H}
    for n, r in enumerate(R, 1):
        assert r.get("aspecto") in ASPECTOS, f"recomendación {n}: aspecto fuera de la Tabla 5"
        for c in ("recomendacion", "responsable", "plazo"):
            assert r.get(c), f"recomendación {n}: falta {c}"
        assert set(r.get("hallazgos") or []) <= ids, f"recomendación {n}: hallazgo inexistente"
