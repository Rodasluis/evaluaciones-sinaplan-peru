"""Crea una evaluación nueva desde la plantilla (plantillas/nueva_evaluacion/) y la tipología elegida (tipologias/).

Uso (desde la raíz del repositorio):
    uv run python src/nucleo/nueva_evaluacion.py <id> --tipologia cierre --titulo "Evaluación de cierre de la …" \
        --politica "Política Nacional de …" --sigla PNX --inicio 2020 --fin 2025 [--organismo mimp-ogmepgd]

Crea:
    evaluaciones/<id>/          carpetas 00 a 07, config.yaml, portada.yaml y 04_analisis/hallazgos.yaml
    src/analisis_<id>/          capa de evaluación mínima (sintesis.py e informe.py) sobre el núcleo

config.yaml lleva las preguntas tipo, los instrumentos sugeridos, las escalas y la estructura del informe de la
tipología: son un punto de partida que el equipo evaluador ajusta. No sobrescribe una evaluación existente.
"""
import argparse
import re
import shutil
import sys
from datetime import date
from pathlib import Path
from string import Template

import yaml

RAIZ_REPO = Path(__file__).resolve().parents[2]
PLANTILLA = RAIZ_REPO / "plantillas" / "nueva_evaluacion"
TIPOLOGIAS = RAIZ_REPO / "tipologias"
ID_VALIDO = re.compile(r"^[a-z][a-z0-9_]{2,40}$")
# nombres de los instrumentos (ADR-009); la tipología dice cuáles se sugieren
INSTRUMENTOS = {
    "REV_DOC": "Revisión documental",
    "ENTREV": "Entrevistas",
    "REP_ENT": "Información reportada por las entidades",
    "REV_ALI": "Coherencia y alineamiento",
    "REV_IND": "Indicadores y recursos",
    "REV_EVI": "Revisión de evidencia",
}
# semáforo de avance del piloto (reportes de cumplimiento de la PNIG); se confirma con el reporte de la política
AVANCE_PILOTO = [{"desde": 150, "nivel": "Alto con posible falla de planeación"}, {"desde": 95, "nivel": "Alto"},
                 {"desde": 75, "nivel": "Medio"}, {"desde": 15, "nivel": "Bajo"},
                 {"desde": float("-inf"), "nivel": "Bajo con posible falla de planeación"}]
MESES = ("enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre",
         "noviembre", "diciembre")


def tipologias():
    return sorted(p.parent.name for p in TIPOLOGIAS.glob("*/estructura.yaml"))


def leer_tipologia(tipo):
    with open(TIPOLOGIAS / tipo / "estructura.yaml", encoding="utf-8") as f:
        return yaml.safe_load(f)


def _bloque(datos, sangria=2):
    """YAML de un bloque, sangrado para insertarlo bajo una clave de la plantilla."""
    texto = yaml.safe_dump(datos, allow_unicode=True, sort_keys=False, width=110, default_flow_style=False)
    return "\n".join((" " * sangria + ln) if ln else ln for ln in texto.rstrip("\n").split("\n"))


def escalas(t, inicio, fin):
    """Escalas de config.yaml a partir de las de la tipología; los años de cierre son los dos últimos del periodo."""
    e = t.get("escalas", {})
    salida = {"valoracion": list(e.get("valoracion", ["Satisfactorio", "Parcial", "Insuficiente"])),
              "confianza": list(e.get("confianza", ["Sólido", "Con respaldo parcial", "Indicio"]))}
    if "No determinado" not in salida["confianza"]:
        salida["confianza"].append("No determinado")
    if "implementacion" in e:
        imp = e["implementacion"]
        salida["implementacion"] = {"componentes": list(imp["componentes"]),
                                    "minimo_componentes": int(imp["minimo_componentes"]),
                                    "anios_cierre": [max(inicio, fin - 1), fin]}
    for clave in ("alineamiento", "solidez_evidencia", "recomendaciones_previas"):
        if clave in e:
            salida[clave] = list(e[clave])
    salida["avance"] = AVANCE_PILOTO
    return salida


def config_texto(a, t):
    preguntas = {k: {"criterio": q["criterio"], "texto": q["texto"], "instrumentos": list(q["instrumentos"])}
                 for k, q in t["preguntas_tipo"].items()}
    informe = [{k: v for k, v in s.items()} for s in t["informe"]]
    raiz = f"evaluaciones/{a.id}"
    sigla = a.sigla.lower()
    valores = {
        "id": a.id, "tipologia": a.tipologia, "estado_tipologia": t["estado"], "titulo": a.titulo,
        "politica": a.politica, "sigla": a.sigla, "inicio": a.inicio, "fin": a.fin, "raiz": raiz,
        "instrumentos": _bloque({k: INSTRUMENTOS[k] for k in t["instrumentos_sugeridos"]}),
        "preguntas": _bloque(preguntas), "escalas": _bloque(escalas(t, a.inicio, a.fin)), "informe": _bloque(informe),
        "informe_docx": f"06_informe/{date.today().year}_{a.organismo}_informe-evaluacion-{a.tipologia}-{sigla}.docx",
        "anexo": f"07_anexo_tecnico/Anexo_tecnico_{a.sigla}_{a.inicio}-{a.fin}",
    }
    with open(PLANTILLA / "config.yaml.plantilla", encoding="utf-8") as f:
        return Template(f.read()).substitute(valores)


def portada_texto(a, t):
    """portada.yaml de la evaluación: la del repositorio (config/portada.yaml) con la etiqueta, el título y la fecha."""
    hoy = date.today()
    texto = (RAIZ_REPO / "config" / "portada.yaml").read_text(encoding="utf-8")
    cambios = {"etiqueta": f"Informe de {t['nombre'][0].lower()}{t['nombre'][1:]}",
               "titulo": f"{a.politica} ({a.inicio}–{a.fin})", "lugar_fecha": f"Lima, {MESES[hoy.month - 1]} de {hoy.year}"}
    for clave, valor in cambios.items():
        texto, n = re.subn(rf'^(  {clave}: )"[^"\n]*"', lambda m: m.group(1) + '"' + valor.replace('"', "'") + '"',
                           texto, count=1, flags=re.M)
        assert n == 1, f"config/portada.yaml sin la clave {clave}"
    return texto


def hallazgos_texto(a, t):
    """04_analisis/hallazgos.yaml: un hallazgo pendiente por pregunta, para que la síntesis y el informe corran desde
    el primer día."""
    preguntas = {k: {"valoracion": "", "respuesta": ""} for k in t["preguntas_tipo"]}
    hallazgos = [{"id": f"{k}-H1", "pregunta": k, "aspecto": "", "enunciado": f"[POR COMPLETAR] hallazgo de la pregunta {k}",
                  "confianza": "No determinado", "estado": "Análisis pendiente", "detalle": "", "salvedad": "",
                  "nota": "", "evidencias": []} for k in t["preguntas_tipo"]]
    with open(PLANTILLA / "hallazgos.yaml.plantilla", encoding="utf-8") as f:
        cabecera = Template(f.read()).substitute({"id": a.id})
    return cabecera + "preguntas:\n" + _bloque(preguntas) + "\nhallazgos:\n" + _bloque(hallazgos) + "\nrecomendaciones: []\n"


def crear(a, raiz_repo=RAIZ_REPO):
    t = leer_tipologia(a.tipologia)
    destino = raiz_repo / "evaluaciones" / a.id
    capa = raiz_repo / "src" / f"analisis_{a.id}"
    for ruta in (destino, capa):
        if ruta.exists():
            raise FileExistsError(f"Ya existe {ruta}: elija otro id o borre la carpeta a mano")
    shutil.copytree(PLANTILLA / "evaluacion", destino)
    shutil.copytree(PLANTILLA / "analisis", capa)
    (destino / "config.yaml").write_text(config_texto(a, t), encoding="utf-8")
    (destino / "portada.yaml").write_text(portada_texto(a, t), encoding="utf-8")
    (destino / "04_analisis" / "hallazgos.yaml").write_text(hallazgos_texto(a, t), encoding="utf-8")
    yaml.safe_load((destino / "config.yaml").read_text(encoding="utf-8"))   # valida el YAML generado
    return destino, capa, t


def argumentos(argv):
    p = argparse.ArgumentParser(description="Crea una evaluación nueva desde la plantilla y una tipología.")
    p.add_argument("id", help="identificador en minúsculas, p. ej. pnx_2027")
    p.add_argument("--tipologia", required=True, choices=tipologias())
    p.add_argument("--titulo", required=True)
    p.add_argument("--politica", required=True, help="nombre de la política o plan")
    p.add_argument("--sigla", required=True)
    p.add_argument("--inicio", type=int, required=True)
    p.add_argument("--fin", type=int, required=True)
    p.add_argument("--organismo", default="mimp-ogmepgd", help="organismo del nombre de archivo (AAAA_organismo_…)")
    a = p.parse_args(argv)
    if not ID_VALIDO.match(a.id):
        p.error("id: minúsculas, dígitos y guion bajo, de 3 a 41 caracteres, empezando por letra")
    if a.fin < a.inicio:
        p.error("--fin anterior a --inicio")
    return a


def main(argv=None):
    a = argumentos(sys.argv[1:] if argv is None else argv)
    destino, capa, t = crear(a)
    rel = lambda p: p.relative_to(RAIZ_REPO).as_posix()  # noqa: E731
    print(f"Evaluación creada: {rel(destino)} y {rel(capa)} (tipología {a.tipologia}, {t['estado']}).")
    print("Siguientes pasos (plantillas/nueva_evaluacion/LEEME.md):")
    print(f"  1. Completar {rel(destino)}/config.yaml y ajustar las preguntas tipo a la política.")
    print(f"  2. Registrar las fuentes en docs/fuentes/ y los parámetros del caso en {rel(destino)}/02_insumos/.")
    print(f"  3. Escribir los hallazgos en {rel(destino)}/04_analisis/hallazgos.yaml.")
    print(f"  4. EVALUACION={a.id} uv run python {rel(capa)}/sintesis.py y luego {rel(capa)}/informe.py")


if __name__ == "__main__":
    main()
