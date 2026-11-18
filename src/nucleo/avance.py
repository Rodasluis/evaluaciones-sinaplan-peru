"""Avance de una evaluación por etapas (Fase 4, ADR-017): lo usa el comando /evaluar de Claude Code.

Cada etapa tiene sus agentes, lo que el usuario aprueba y sus requisitos previos. El estado se guarda en
evaluaciones/<id>/avance.yaml: pendiente, en_curso, por_aprobar, aprobada u omitida. Una etapa queda omitida si no
aplica a la modalidad de la evaluación (config.yaml, sección modalidad): el recojo primario sin entrevistas ni
formulario, la validación si solo se entrega el informe final. Solo el usuario aprueba (CLAUDE.md, regla 5): el
comando registra la aprobación cuando el usuario la da en el chat.

Uso (desde la raíz del repositorio):
    uv run python src/nucleo/avance.py <id>                          estado de las etapas y la siguiente
    uv run python src/nucleo/avance.py <id> marcar <etapa> <estado> ["nota"]
"""
import sys
from dataclasses import dataclass
from datetime import date
from pathlib import Path

import yaml

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from nucleo.config import RAIZ_REPO, cargar_desde  # noqa: E402

ESTADOS = ("pendiente", "en_curso", "por_aprobar", "aprobada", "omitida")
ARCHIVO = "avance.yaml"


@dataclass(frozen=True)
class Etapa:
    clave: str
    nombre: str
    agentes: tuple
    aprueba: str


ETAPAS = (
    Etapa("diseno", "Diseño", ("disenador-de-evaluacion",),
          "preguntas adaptadas, alcance y matriz pregunta–instrumento–fuente"),
    Etapa("fuentes", "Fuentes", ("recolector-de-fuentes",), "lista de fuentes registradas y lo que falta conseguir"),
    Etapa("politica", "Lectura de la política", ("lector-de-la-politica",),
          "catálogo de servicios, indicadores y cadena de resultados"),
    Etapa("instrumentos", "Instrumentos", ("analista-documental", "analista-de-indicadores", "analista-de-alineamiento",
                                           "analista-de-evidencia", "programador-de-instrumentos"),
          "reglas, umbrales y libros de análisis"),
    Etapa("recojo", "Recojo de información primaria", ("analista-de-entrevistas",),
          "guías de entrevista y formulario antes de aplicarlos; luego, el registro codificado"),
    Etapa("hallazgos", "Hallazgos", ("redactor-de-hallazgos",), "hallazgos, valoraciones y recomendaciones"),
    Etapa("informe", "Informe", ("revisor-de-informe",), "versión del informe"),
    Etapa("validacion", "Validación con el órgano rector", (), "respuestas a las observaciones y versión final"),
    Etapa("entrega", "Entrega", ("auditor-de-publicacion",), "envío del informe y del anexo técnico (acción externa)"),
)
POR_CLAVE = {e.clave: e for e in ETAPAS}


def modalidad(cfg):
    m = cfg.get("modalidad") or {}
    return {"entrevistas": bool(m.get("entrevistas", True)),
            "formulario_entidades": bool(m.get("formulario_entidades", False)),
            "documentos_internos": bool(m.get("documentos_internos", False)),
            "validacion": m.get("validacion", "organo_rector"),
            "personalizacion": list(m.get("personalizacion") or [])}


def aplica(cfg, clave):
    m = modalidad(cfg)
    if clave == "recojo":
        return m["entrevistas"] or m["formulario_entidades"]
    if clave == "validacion":
        return m["validacion"] == "organo_rector"
    return True


def requisitos(cfg, clave):
    """Lo que falta para empezar una etapa, además de la aprobación de las anteriores."""
    falta = []
    pol = (cfg.get("politica") or {}).get("documento") or ""
    if clave == "politica" and not (pol and (cfg.repo / pol).is_file()):
        falta.append("config.yaml: politica.documento con la ruta del documento de la política en docs/fuentes/")
    if clave == "instrumentos" and not cfg.abs("insumos", "servicios.csv").is_file():
        falta.append("02_insumos/servicios.csv")
    if clave in ("recojo", "hallazgos") and not list(cfg.abs("analisis").glob("*.xlsx")):
        falta.append("libros de análisis en 04_analisis/")
    if clave == "informe" and not cfg.abs("analisis", "matriz_triangulacion_preguntas.xlsx").is_file():
        falta.append("04_analisis/matriz_triangulacion_preguntas.xlsx (sintesis.py)")
    informe = (cfg.get("productos") or {}).get("informe", "")
    if clave in ("validacion", "entrega") and not (informe and cfg.abs("informe", informe.split("/", 1)[-1]).is_file()):
        falta.append(f"{informe or 'informe'} (informe.py)")
    return falta


def leer(cfg):
    archivo = cfg.repo / cfg.raiz / ARCHIVO
    if not archivo.exists():
        return iniciar(cfg)
    with open(archivo, encoding="utf-8") as f:
        return yaml.safe_load(f) or {"etapas": {}}


def guardar(cfg, datos):
    archivo = cfg.repo / cfg.raiz / ARCHIVO
    cabecera = ("# Avance de la evaluación por etapas (src/nucleo/avance.py, comando /evaluar). Estados: "
                + ", ".join(ESTADOS) + ".\n# La aprobación de cada etapa es del usuario (CLAUDE.md, regla 5).\n")
    archivo.write_text(cabecera + yaml.safe_dump(datos, allow_unicode=True, sort_keys=False, width=110),
                       encoding="utf-8")


def iniciar(cfg):
    datos = {"etapas": {e.clave: {"estado": "pendiente" if aplica(cfg, e.clave) else "omitida", "fecha": "", "nota": ""}
                        for e in ETAPAS}}
    guardar(cfg, datos)
    return datos


def siguiente(cfg, datos=None):
    """Primera etapa que no está aprobada ni omitida, o None si la evaluación terminó."""
    datos = datos or leer(cfg)
    return next((e for e in ETAPAS if datos["etapas"][e.clave]["estado"] not in ("aprobada", "omitida")), None)


def marcar(cfg, clave, estado, nota=""):
    """Cambia el estado de una etapa. Solo se aprueba lo que está por aprobar, y no se empieza una etapa con una
    anterior sin aprobar."""
    if clave not in POR_CLAVE or estado not in ESTADOS:
        raise ValueError(f"etapa ({', '.join(POR_CLAVE)}) o estado ({', '.join(ESTADOS)}) no válidos")
    datos = leer(cfg)
    actual = datos["etapas"][clave]["estado"]
    if estado == "aprobada" and actual != "por_aprobar":
        raise ValueError(f"{clave}: solo se aprueba una etapa por aprobar (está {actual})")
    if estado in ("en_curso", "por_aprobar"):
        previas = [e.clave for e in ETAPAS[:[e.clave for e in ETAPAS].index(clave)]
                   if datos["etapas"][e.clave]["estado"] not in ("aprobada", "omitida")]
        if previas:
            raise ValueError(f"{clave}: faltan aprobar {', '.join(previas)}")
        if estado == "en_curso" and (falta := requisitos(cfg, clave)):
            raise ValueError(f"{clave}: faltan {'; '.join(falta)}")
    datos["etapas"][clave] = {"estado": estado, "fecha": date.today().isoformat(), "nota": nota}
    guardar(cfg, datos)
    return datos


def resumen(cfg):
    datos = leer(cfg)
    lineas = [f"{i}. {e.nombre}: {datos['etapas'][e.clave]['estado']}"
              + (f" ({datos['etapas'][e.clave]['fecha']})" if datos["etapas"][e.clave]["fecha"] else "")
              for i, e in enumerate(ETAPAS, 1)]
    s = siguiente(cfg, datos)
    if s:
        falta = requisitos(cfg, s.clave)
        lineas.append(f"Siguiente: {s.nombre} (agentes: {', '.join(s.agentes) or 'ninguno'}; aprueba el usuario: "
                      f"{s.aprueba})" + (f". Falta: {'; '.join(falta)}" if falta else ""))
    else:
        lineas.append("Evaluación terminada.")
    return "\n".join(lineas)


def main(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    if not argv:
        print(__doc__)
        return 2
    cfg = cargar_desde(RAIZ_REPO, argv[0])
    if len(argv) >= 4 and argv[1] == "marcar":
        try:
            marcar(cfg, argv[2], argv[3], argv[4] if len(argv) > 4 else "")
        except ValueError as e:
            print(e)
            return 1
    print(resumen(cfg))
    return 0


if __name__ == "__main__":
    sys.exit(main())
