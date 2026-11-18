"""Instrumentos genéricos (src/nucleo/instrumentos/) sobre el caso ficticio de tests/datos/evaluacion_ficticia/.

La evaluación se crea con el generador en una carpeta temporal y recibe las entradas del caso. Se comprueban los
libros esperados, los errores de validación (columna faltante, ajuste sin sustento, cita no ubicada), las guías de
entrevista, el formulario a las entidades y la corrida completa con correr.py.
"""
import argparse
import shutil
from pathlib import Path

import pandas as pd
import pytest

from nucleo import nueva_evaluacion as NE
from nucleo.config import cargar_desde
from nucleo.instrumentos import activos, base, correr, entrev, rep_ent, rev_ali, rev_doc, rev_evi, rev_ind

CASO = Path(__file__).resolve().parent / "datos" / "evaluacion_ficticia"
ID = "prueba_ficticia"


def crear(tmp_path, **modalidad):
    a = argparse.Namespace(id=ID, tipologia="cierre", titulo="Evaluación de cierre de la Política Nacional Ficticia "
                           "(2021–2024)", politica="Política Nacional Ficticia", sigla="PNF", inicio=2021, fin=2024,
                           organismo="prueba", **modalidad)
    destino, _, _ = NE.crear(a, raiz_repo=tmp_path)
    for carpeta in ("02_insumos", "05_entrevistas"):
        for f in (CASO / carpeta).glob("*.csv"):
            shutil.copy(f, destino / carpeta / f.name)
    docs = tmp_path / "docs" / "fuentes" / ID
    docs.mkdir(parents=True)
    shutil.copy(CASO / "documentos" / "plan_nacional_ficticio.txt", docs)
    return cargar_desde(tmp_path, ID)


@pytest.fixture
def cfg(tmp_path):
    return crear(tmp_path, formulario_entidades=True)


def hoja(r, nombre):
    return pd.read_excel(r["libro"], sheet_name=nombre, dtype=str).fillna("")


def test_revision_documental(cfg):
    r = rev_doc.correr(cfg)
    assert (r["Implementado"], r["Implementado parcialmente"], r["No implementado"]) == (2, 3, 1)
    e = hoja(r, "Estado por servicio").set_index("Código")
    assert e.loc["3.1.1", "Clasificación por criterios"] == "Sin información suficiente"
    assert e.loc["3.1.1", "Estado asignado"] == "Implementado parcialmente"
    assert e.loc["1.2.1", "Estado asignado"] == "No implementado"
    res = hoja(r, "Resumen por objetivo").set_index("Objetivo prioritario")
    assert res.loc["Total", "Servicios"] == "6" and res.loc["OP1", "Implementados (%)"] == "33.3"
    assert "Criterios" in pd.ExcelFile(r["libro"]).sheet_names


def test_indicadores(cfg):
    r = rev_ind.correr(cfg)
    assert r["con_avance"] == 6 and r["fichas"] == 2
    a = hoja(r, "Avance por año").set_index(["Código", "Año"])
    assert a.loc[("IOP1", "2021"), "Nivel de avance"] == "Alto"
    assert a.loc[("IOP1", "2023"), "Motivo sin avance"] == "Sin valor"
    assert a.loc[("IOP1", "2024"), "Motivo sin avance"] == "Sin registro"
    assert a.loc[("IOP2", "2022"), "Nivel de avance"] == "Alto con posible falla de planeación"
    assert a.loc[("IOP2", "2023"), "Advertencia"] != ""
    assert a.loc[("IOP3", "2021"), "Motivo sin avance"] == "Sin sentido definido"
    assert a.loc[("ISS111", "2023"), "Nivel de avance"] == "Bajo con posible falla de planeación"
    d = hoja(r, "Dictamen de las fichas").set_index("Código").Dictamen
    assert d.to_dict() == {"IOP1": "Con observaciones", "IOP3": "Con observación crítica"}
    c = hoja(r, "Conciliación")
    assert list(c.Coincide) == ["Sí", "No"]


def test_evidencia(cfg):
    r = rev_evi.correr(cfg)
    m = hoja(r, "Matriz por tipo").set_index("Tipo").Plausibilidad
    assert m.to_dict() == {"T01": "Sólida, favorable", "T02": "Moderada, mixta (evidencia parcial)",
                           "T03": "Limitada, favorable"}
    assert len(hoja(r, "Plausibilidad por servicio")) == 6


def test_coherencia_y_alineamiento(cfg):
    rev_doc.correr(cfg)
    rev_evi.correr(cfg)
    r = rev_ali.correr(cfg)
    assert (r["servicios_completos"], r["citas_ubicadas"], r["citas_sin_verificar"]) == (2, 2, 1)
    s = hoja(r, "Coherencia por servicio").set_index("Código")
    assert s.loc["1.2.1", "Vacíos detectados"] == "Sin indicador vigente; Sin causa ni efecto del modelo"
    assert s.loc["1.1.1", "Plausibilidad causal"] == "Sólida, favorable"
    assert s.loc["1.2.1", "Estado de implementación"] == "No implementado"
    lin = hoja(r, "Coherencia por lineamiento").set_index("Lineamiento").Valoración
    assert lin["3.2. Evaluar la gestión"] == "Rota" and lin["1.1. Ampliar la oferta"] == "Completa"
    a = hoja(r, "Alineamiento por objetivo")
    fila = a[(a.Tipo == "Vertical") & (a["Objetivo prioritario"] == "OP1")].iloc[0]
    assert fila["Valoración"] == "1 directa, 0 indirecta, 0 sin correspondencia"


def test_codificaciones(cfg):
    r = rep_ent.correr(cfg)
    f = hoja(r, "Frecuencias por código").set_index("Código").Segmentos
    assert f.to_dict() == {"REC": "2", "ART": "1", "NOR": "1"}
    assert hoja(r, "Por servicio").set_index("Código").loc["1.1.2", "Hallazgos críticos"] == "1"
    e = entrev.correr(cfg)
    assert e["roles"] == 2 and e["acuerdo_exacto"] == 75.0
    assert len(hoja(e, "Contraste de codificaciones")) == 4


def test_guias_y_formulario(cfg):
    guias = entrev.guias(cfg)
    assert sorted(g.name for g in guias) == ["ENTREV_guia_direccion-conductora.docx",
                                             "ENTREV_guia_oficina-de-planeamiento.docx"]
    import docx
    texto = "\n".join(p.text for p in docx.Document(str(guias[0])).paragraphs)
    assert "Ley N.° 29733" in texto and "P01." in texto
    forms = rep_ent.formulario(cfg)
    assert len(forms) == 3
    # respuesta simulada de la Entidad Ficticia A
    respuesta = cfg.abs("datos_raw", "formularios", "respuesta_a.xlsx")
    respuesta.parent.mkdir(parents=True, exist_ok=True)
    a = next(f for f in forms if "entidad-ficticia-a" in f.name)
    hojas = pd.read_excel(a, sheet_name=None, dtype=str)
    hojas["Servicios"] = hojas["Servicios"].fillna("")
    hojas["Servicios"].loc[0, "dificultades"] = "Falta de personal para las visitas"
    with pd.ExcelWriter(respuesta, engine="openpyxl") as w:
        for nombre, df in hojas.items():
            df.to_excel(w, sheet_name=nombre, index=False)
    n_seg, _ = rep_ent.consolidar(cfg)
    assert n_seg == 1
    # el segmento nuevo entra a la información de las entidades y hay que codificarlo antes de correr el instrumento
    with pytest.raises(base.ErrorInsumo, match="segmentos sin codificar: F-entidad-ficticia-a-1.1.1-dificultades"):
        rep_ent.validar(cfg)


def test_errores_de_validacion(tmp_path):
    cfg = crear(tmp_path)
    comp = cfg.abs("insumos", "rev_doc_componentes.csv")
    df = pd.read_csv(comp, dtype=str).fillna("")
    df.loc[0, "ajuste"] = "No implementado"
    df.to_csv(comp, index=False)
    with pytest.raises(base.ErrorInsumo, match="requiere sustento_ajuste"):
        rev_doc.validar(cfg)
    serv = cfg.abs("insumos", "servicios.csv")
    pd.read_csv(serv, dtype=str).drop(columns="servicio").to_csv(serv, index=False)
    with pytest.raises(base.ErrorInsumo, match="faltan las columnas servicio"):
        rev_doc.validar(cfg)
    ali = cfg.abs("insumos", "alineamiento.csv")
    df = pd.read_csv(ali, dtype=str).fillna("")
    df.loc[0, "cita"] = "una cita que no está en el plan"
    df.to_csv(ali, index=False)
    shutil.copy(CASO / "02_insumos" / "servicios.csv", serv)
    with pytest.raises(base.ErrorInsumo, match="la cita no aparece"):
        rev_ali.validar(cfg)


def test_corrida_completa_y_modalidad(tmp_path):
    cfg = crear(tmp_path, sin_entrevistas=True)
    assert "ENTREV" not in activos(cfg)
    salida = []
    r = correr.correr(cfg, salida=salida.append)
    assert set(r) == {"REV_DOC", "REV_IND", "REV_EVI", "REV_ALI", "REP_ENT"}
    assert not any(isinstance(x, Exception) for x in r.values()), salida
    for modulo in (rev_doc, rev_ind, rev_evi, rev_ali, rep_ent):
        assert cfg.abs("analisis", modulo.LIBRO).is_file()


def test_sintesis_e_informe_con_resultados(tmp_path):
    """Corrida completa: instrumentos, síntesis e informe de la capa de plantilla, sin Word."""
    import os
    import subprocess
    import sys

    import docx

    from nucleo.config import RAIZ_REPO
    cfg = crear(tmp_path, formulario_entidades=True)
    r = correr.correr(cfg, salida=lambda *_: None)
    assert not any(isinstance(x, Exception) for x in r.values())
    entorno = {**os.environ, "PYTHONPATH": str(RAIZ_REPO / "src"), "PYTHONUTF8": "1"}
    capa = tmp_path / "src" / f"analisis_{ID}"
    for script in ("sintesis.py", "informe.py"):
        p = subprocess.run([sys.executable, str(capa / script)], cwd=tmp_path, env=entorno, capture_output=True,
                           text=True, encoding="utf-8")
        assert p.returncode == 0, p.stdout + p.stderr
    informe = next(cfg.abs("informe").glob("*.docx"))
    texto = "\n".join(par.text for par in docx.Document(str(informe)).paragraphs)
    assert "Anexo 2. Resultados de los instrumentos" in texto
    assert ("De los 6 servicios de la política, 2 se implementaron, 3 se implementaron parcialmente, 1 no se "
            "implementó") in texto
    assert "Estado_implementacion_servicios.xlsx" in texto


def test_busqueda_academica_sin_red(cfg):
    """Estrategias del YAML, registro de búsquedas y marca de duplicados, con respuestas simuladas de Crossref."""
    import json

    from nucleo import busqueda_academica as BA
    cfg.abs("insumos", "rev_evi_estrategias.yaml").write_text(
        "revisados: {Campbell: 2}\nestrategias:\n"
        "  - {id: ACA-001, tipos: T01, base: Campbell, cadena: 'ficticio acceso'}\n"
        "  - {id: ACA-002, tipos: T02, base: 3ie, cadena: 'ficticio capacitación'}\n", encoding="utf-8")
    items = [{"DOI": "10.0000/FICTICIO.1", "title": ["Estudio ficticio uno"], "type": "journal-article"},
             {"DOI": "10.0000/ficticio.2", "title": ["Estudio ficticio dos"], "type": "journal-article"}]

    # 3ie devuelve la misma obra que Campbell con otro DOI (duplicado por título) y una obra nueva
    otros = [{"DOI": "10.1111/otro.1", "title": ["Estudio ficticio uno"], "type": "report"},
             {"DOI": "10.1111/otro.3", "title": ["Estudio ficticio tres"], "type": "report"}]

    class Falso(BA.Cliente):
        def pedir(self, url, nombre, url_registro=None):
            self.descargas.append((url, nombre, "0" * 64, self.fecha, "simulada"))
            lista = otros if "prefix" in url else items
            return json.dumps({"message": {"total-results": 42, "items": lista}}).encode()

    filas, regs = BA.buscar(cfg, Falso(cfg.abs("datos_raw", "x"), "2026-01-01"), fecha="2026-01-01")
    assert [f["total"] for f in filas] == [42, 42] and len(regs) == 4
    assert [r["duplicado_de"] for r in regs] == ["", "", "DOI:10.0000/ficticio.1", ""]
    b = pd.read_csv(cfg.abs("insumos", "rev_evi_busquedas.csv"), dtype=str)
    assert list(b.id) == ["ACA-001", "ACA-002"] and b.fecha.iloc[0] == "2026-01-01"
    assert (cfg.repo / "docs" / "fuentes" / f"descargas_{ID}_rev-evi-busqueda_2026-01-01.tsv").is_file()
    assert BA.componentes("PubMed", '"Domestic Violence"[mh] AND advocacy[tiab]')[:2] == (["Domestic Violence"],
                                                                                         ["advocacy"])
    rev_evi.validar(cfg)   # el registro de búsquedas pasa la validación de la revisión de evidencia
