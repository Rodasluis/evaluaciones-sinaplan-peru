"""Avance por etapas (src/nucleo/avance.py): etapas que se omiten según la modalidad, requisitos previos y regla de
aprobación (solo se aprueba lo que está por aprobar, y en orden)."""
import argparse

import pytest

from nucleo import avance
from nucleo import nueva_evaluacion as NE
from nucleo.config import cargar_desde


def _crear(tmp_path, **modalidad):
    a = argparse.Namespace(id="prueba_avance", tipologia="cierre", titulo="Evaluación de prueba", politica="Política",
                           sigla="PP", inicio=2021, fin=2024, organismo="prueba", **modalidad)
    NE.crear(a, raiz_repo=tmp_path)
    return cargar_desde(tmp_path, "prueba_avance")


def test_etapas_segun_modalidad(tmp_path):
    cfg = _crear(tmp_path, sin_entrevistas=True, validacion="informe_final")
    estados = {k: v["estado"] for k, v in avance.leer(cfg)["etapas"].items()}
    assert estados["recojo"] == "omitida" and estados["validacion"] == "omitida"
    assert estados["diseno"] == "pendiente" and avance.siguiente(cfg).clave == "diseno"
    cfg2 = _crear(tmp_path / "otra", formulario_entidades=True, sin_entrevistas=True)
    assert avance.aplica(cfg2, "recojo") and avance.aplica(cfg2, "validacion")


def test_aprobacion_en_orden(tmp_path):
    cfg = _crear(tmp_path)
    with pytest.raises(ValueError, match="solo se aprueba una etapa por aprobar"):
        avance.marcar(cfg, "diseno", "aprobada")
    with pytest.raises(ValueError, match="faltan aprobar diseno"):
        avance.marcar(cfg, "fuentes", "en_curso")
    avance.marcar(cfg, "diseno", "en_curso")
    avance.marcar(cfg, "diseno", "por_aprobar", "propuesta de diseño lista")
    avance.marcar(cfg, "diseno", "aprobada", "aprobada por el usuario")
    assert avance.siguiente(cfg).clave == "fuentes"
    avance.marcar(cfg, "fuentes", "en_curso")
    avance.marcar(cfg, "fuentes", "por_aprobar")
    avance.marcar(cfg, "fuentes", "aprobada")
    with pytest.raises(ValueError, match="politica.documento"):
        avance.marcar(cfg, "politica", "en_curso")
    assert "Siguiente: Lectura de la política" in avance.resumen(cfg)
