"""Generador de evaluaciones nuevas (src/nucleo/nueva_evaluacion.py) con las cuatro tipologías, en una carpeta
temporal: config.yaml válido, preguntas vinculadas con instrumentos definidos, secciones de análisis con preguntas y
hallazgos pendientes que pasan la verificación del núcleo."""
import argparse

import pytest
import yaml

from nucleo import hallazgos as NH
from nucleo import nueva_evaluacion as NE


def _args(tipo):
    return argparse.Namespace(id=f"prueba_{tipo}", tipologia=tipo, titulo=f"Evaluación de {tipo} de prueba",
                              politica="Política Nacional de Prueba", sigla="PNP", inicio=2021, fin=2024,
                              organismo="mimp-ogmepgd")


@pytest.mark.parametrize("tipo", NE.tipologias())
def test_crear_evaluacion(tipo, tmp_path):
    destino, capa, t = NE.crear(_args(tipo), raiz_repo=tmp_path)
    cfg = yaml.safe_load((destino / "config.yaml").read_text(encoding="utf-8"))
    assert cfg["id"] == f"prueba_{tipo}" and cfg["tipologia"] == tipo
    assert set(cfg["preguntas"]) == set(t["preguntas_tipo"])
    for k, q in cfg["preguntas"].items():
        assert q["instrumentos"] and set(q["instrumentos"]) <= set(cfg["instrumentos"]), k
    prefijos = [str(s["preguntas"]) for s in cfg["informe"] if "preguntas" in s]
    assert prefijos and all(any(k.startswith(p) for k in cfg["preguntas"]) for p in prefijos)
    if "implementacion" in cfg["escalas"]:
        assert cfg["escalas"]["implementacion"]["anios_cierre"] == [2023, 2024]
    assert "No determinado" in cfg["escalas"]["confianza"]
    h = yaml.safe_load((destino / "04_analisis" / "hallazgos.yaml").read_text(encoding="utf-8"))
    H = [{**x, "evidencias": []} for x in h["hallazgos"]]
    NH.verificar(H, cfg["preguntas"], NH.fuentes_independientes)
    portada = yaml.safe_load((destino / "portada.yaml").read_text(encoding="utf-8"))["portada"]
    assert portada["titulo"] == "Política Nacional de Prueba (2021–2024)"
    assert portada["etiqueta"].startswith("Informe de evaluación de")
    assert {p.name for p in capa.glob("*.py")} == {"comun.py", "sintesis.py", "informe.py"}
    for carpeta in ("00_encargo", "02_insumos", "03_datos/raw", "04_analisis", "05_entrevistas", "06_informe"):
        assert (destino / carpeta).is_dir(), carpeta


def test_no_sobrescribe(tmp_path):
    NE.crear(_args("cierre"), raiz_repo=tmp_path)
    with pytest.raises(FileExistsError):
        NE.crear(_args("cierre"), raiz_repo=tmp_path)


def test_id_invalido():
    with pytest.raises(SystemExit):
        NE.argumentos(["Mal-Id", "--tipologia", "cierre", "--titulo", "x", "--politica", "x", "--sigla", "X",
                       "--inicio", "2020", "--fin", "2025"])


def _con(tipo, **modalidad):
    return argparse.Namespace(**vars(_args(tipo)), **modalidad)


def test_modalidad_sin_entrevistas(tmp_path):
    destino, _, _ = NE.crear(_con("cierre", sin_entrevistas=True, validacion="informe_final",
                                  personalizar="REV_IND"), raiz_repo=tmp_path)
    cfg = yaml.safe_load((destino / "config.yaml").read_text(encoding="utf-8"))
    assert cfg["modalidad"] == {"entrevistas": False, "formulario_entidades": False, "documentos_internos": False,
                                "validacion": "informe_final", "personalizacion": ["REV_IND"]}
    assert "ENTREV" not in cfg["instrumentos"]
    assert all("ENTREV" not in q["instrumentos"] and q["instrumentos"] for q in cfg["preguntas"].values())
    assert not (destino / "05_entrevistas" / "entrev_preguntas.csv").exists()
    assert (destino / "02_insumos" / "rev_doc_componentes.csv").read_text(encoding="utf-8").startswith(
        "codigo,protocolo,protocolo_fuente,asignacion")
    avance = yaml.safe_load((destino / "avance.yaml").read_text(encoding="utf-8"))["etapas"]
    assert avance["validacion"]["estado"] == "omitida" and avance["recojo"]["estado"] == "omitida"


def test_pregunta_sin_instrumentos_recibe_revision_documental():
    t = NE.leer_tipologia("cierre")
    t["preguntas_tipo"]["I5"]["instrumentos"] = ["ENTREV"]
    _, instrumentos, preguntas, avisos = NE.preparar(_con("cierre", sin_entrevistas=True), t)
    assert preguntas["I5"]["instrumentos"] == ["REV_DOC"] and "REV_DOC" in instrumentos
    assert any("I5" in a for a in avisos)


def test_personalizar_instrumento_inexistente():
    with pytest.raises(ValueError, match="ENTREV"):
        NE.preparar(_con("cierre", sin_entrevistas=True, personalizar="ENTREV"), NE.leer_tipologia("cierre"))


def test_modalidad_con_entrevistas_y_formulario(tmp_path):
    destino, _, _ = NE.crear(_con("implementacion", formulario_entidades=True, documentos_internos=True),
                             raiz_repo=tmp_path)
    cfg = yaml.safe_load((destino / "config.yaml").read_text(encoding="utf-8"))
    assert cfg["modalidad"]["formulario_entidades"] and cfg["modalidad"]["documentos_internos"]
    assert {"ENTREV", "REP_ENT"} <= set(cfg["instrumentos"])
    for nombre in ("entrev_preguntas.csv", "entrev_segmentos.csv", "entrev_codificacion.csv"):
        assert (destino / "05_entrevistas" / nombre).is_file(), nombre
    assert (destino / "02_insumos" / "libro_codigos.csv").is_file()
