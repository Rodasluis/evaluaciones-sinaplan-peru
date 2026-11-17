"""Pruebas de los módulos del núcleo (src/nucleo)."""
import re

import pandas as pd
import pytest

from nucleo import codificacion, escalas, evidencia, hallazgos, presupuesto
from nucleo.auditoria import CELULAR, CELULAR_CADENA, DNI, archivos_sensibles
from nucleo.config import CARPETAS, cargar


def test_config_piloto():
    cfg = cargar("pnig_2026")
    assert cfg.raiz == "evaluaciones/pnig_2026"
    assert cfg.ruta("analisis") == "evaluaciones/pnig_2026/04_analisis"
    assert cfg.ruta("informe", "x.docx") == "evaluaciones/pnig_2026/06_informe/x.docx"
    assert cfg["politica"]["sigla"] == "PNIG" and len(cfg["preguntas"]) == 12
    assert set(CARPETAS) >= {"encargo", "insumos", "analisis", "informe", "anexo"}
    with pytest.raises(FileNotFoundError):
        cargar("no_existe_2099")


@pytest.mark.parametrize("faltan,reciente,periodo,nada,esperado", [
    ([], True, True, False, escalas.IMPLEMENTADO),
    (["protocolo"], True, True, False, escalas.IMPLEMENTADO),            # basta con dos de tres (ADR-012)
    (["protocolo", "reporte de indicador"], True, True, False, escalas.PARCIAL),
    ([], False, True, False, escalas.PARCIAL),                            # interrumpido antes del cierre
    (["protocolo", "asignación identificable", "reporte de indicador"], False, False, True, escalas.NO_IMPLEMENTADO),
    (["protocolo"], False, False, False, escalas.SIN_INFO),
])
def test_escala_implementacion(faltan, reciente, periodo, nada, esperado):
    assert escalas.clasificar_implementacion(faltan, reciente, periodo, nada) == esperado


def test_escala_con_tres_de_tres():
    assert escalas.clasificar_implementacion(["protocolo"], True, True, False, minimo_componentes=3) == escalas.PARCIAL


def test_avance_y_nivel():
    assert escalas.avance("Ascendente", 50, 60) == (120.0, None)
    assert escalas.avance("Descendente", 10, 8) == (125.0, None)
    assert escalas.avance(None, 10, 8)[1] == "Sin sentido definido"
    assert escalas.avance("Ascendente", 0, 5)[1].startswith("No interpretable")
    assert escalas.nivel(160) == "Alto con posible falla de planeación"
    assert escalas.nivel(96) == "Alto" and escalas.nivel(80) == "Medio" and escalas.nivel(20) == "Bajo"
    assert escalas.nivel(5) == "Bajo con posible falla de planeación"


def test_codificacion():
    assert codificacion.kappa(["a", "b", "a", "b"], ["a", "b", "a", "b"]) == 1.0
    assert codificacion.kappa(["a", "a"], ["a", "a"]) is None
    assert codificacion.kappa(["a", "b", "a", "b"], ["a", "a", "b", "b"]) == 0.0
    assert codificacion.jaccard({"x", "y"}, {"y", "z"}) == pytest.approx(1 / 3)
    assert codificacion.acuerdo([1, 2, 3, 4], [1, 2, 0, 4]) == 75


def _estudios(filas):
    return pd.DataFrame(filas, columns=["comparable", "diseno", "direccion", "calidad"])


def test_solidez_evidencia():
    rs = _estudios([(True, "Revisión sistemática", "Favorable", "Moderada")])
    assert evidencia.solidez(rs) == ("Sólida", "favorable")
    dos = _estudios([(True, "Experimental", "Favorable", ""), (True, "Cuasi experimental", "Nulo", "")])
    assert evidencia.solidez(dos) == ("Moderada", "mixta")
    otros = _estudios([(True, "Descriptivo", "Favorable", "")])
    assert evidencia.solidez(otros)[0] == "Limitada"
    assert evidencia.solidez(_estudios([]).astype({"comparable": bool})) == ("Sin evidencia identificada", "")
    assert evidencia.plausibilidad("Moderada", "mixta", "Parcial") == "Moderada, mixta (evidencia parcial)"


def _hallazgo(confianza, bases, pregunta="P1"):
    return {"id": f"{pregunta}-H1", "pregunta": pregunta, "aspecto": "a", "enunciado": "Texto sin guiones.",
            "evidencias": [hallazgos.evidencia("Documental", "REV_DOC", "e", "f", b) for b in bases],
            "confianza": confianza, "estado": "Validado", "nota": "", "salvedad": "", "detalle": "Breve."}


def test_hallazgos_reglas():
    h = _hallazgo("Sólido", ["Informe A", "Informe B", "Entrevista (OGX)"])
    assert hallazgos.fuentes_independientes(h, no_independiente="Entrevista (OGX)") == 2
    hallazgos.verificar([h], {"P1": None}, lambda x: hallazgos.fuentes_independientes(x))
    with pytest.raises(AssertionError):
        hallazgos.verificar([_hallazgo("Sólido", ["Informe A"])], {"P1": None}, hallazgos.fuentes_independientes)
    with pytest.raises(AssertionError):
        hallazgos.verificar([h], {"P1": None, "P2": None}, hallazgos.fuentes_independientes)  # pregunta sin hallazgo
    malo = _hallazgo("Indicio", ["Informe A"])
    malo["enunciado"] = "Texto - con guion explicativo."
    with pytest.raises(AssertionError):
        hallazgos.verificar([malo], {"P1": None}, hallazgos.fuentes_independientes)


def test_hallazgos_origen():
    """Dos documentos que reproducen el dato de un mismo productor cuentan como un solo origen (ADR-015)."""
    h = _hallazgo("Indicio", ["Reporte X", "Consolidado X"])
    assert hallazgos.fuentes_independientes(h) == 2   # sin mapa, cada fuente es su propio origen
    mapa = {"Reporte X": "entidad X", "Consolidado X": "entidad X"}
    assert hallazgos.fuentes_independientes(h, origen=lambda f: mapa.get(f, f)) == 1
    h["evidencias"].append(hallazgos.evidencia("Cuantitativa", "REV_IND", "e", "f", "Encuesta Y", origen="INEI"))
    assert hallazgos.fuentes_independientes(h, origen=lambda f: mapa.get(f, f)) == 2


def test_rechazar_cambios():
    from nucleo.rechazar_cambios import procesar
    w = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    xml = (f'<w:document {w}><w:body><w:p>'
           '<w:r><w:rPr><w:b/><w:rPrChange w:id="1"><w:rPr/></w:rPrChange></w:rPr><w:t xml:space="preserve">Uno </w:t></w:r>'
           '<w:ins w:id="2"><w:r><w:t>nuevo</w:t></w:r></w:ins>'
           '<w:del w:id="3"><w:r><w:delText>viejo</w:delText></w:r></w:del>'
           '</w:p></w:body></w:document>').encode("utf-8")
    from collections import Counter
    conteo = Counter()
    salida = procesar(xml, conteo).decode("utf-8")
    assert "nuevo" not in salida and "<w:t>viejo</w:t>" in salida and "Change" not in salida and "<w:b/>" in salida
    assert conteo == Counter({"inserciones rechazadas": 1, "eliminaciones rechazadas": 1,
                              "cambios de formato aceptados": 1})
    marca = (f'<w:document {w}><w:body><w:p><w:pPr><w:rPr><w:ins w:id="4"/></w:rPr></w:pPr></w:p></w:body>'
             '</w:document>').encode("utf-8")
    with pytest.raises(NotImplementedError):
        procesar(marca, Counter())


def test_hojas_y_libros_publicados():
    assert hallazgos.hoja_publica("REV_ALI.1 X", {"REV_ALI.1 X": "Coherencia"}) == "Coherencia"
    with pytest.raises(AssertionError):
        hallazgos.hoja_publica("Hoja REV_DOC", {})
    with pytest.raises(AssertionError):
        hallazgos.hoja_publica("x", {"x": "a" * 32})
    assert hallazgos.libro_publico("ruta/REV.xlsx", {"REV.xlsx": "Pub.xlsx"}) == "Pub.xlsx"


def test_conciliacion():
    c = presupuesto.conciliar([({"Órgano": "A", "Año": 2020}, 100.4, 100.0), ({"Órgano": "A", "Año": 2021}, 0, 0),
                               ({"Órgano": "B", "Año": 2020}, 50, 70)], "Fuente 1", "Fuente 2")
    assert list(c.Coincide) == ["Sí", "No"] and list(c["Diferencia (S/)"]) == [0.4, -20.0]


def test_auditoria_patrones(tmp_path):
    # datos sintéticos armados por partes, para que el propio archivo no active el control previo al commit
    dni, cel = "4012" + "3456", "98765" + "4321"
    assert DNI.search(f"DNI N.° {dni}") and not DNI.search(f"población {dni}")
    assert CELULAR.search(f"celular: {cel}") and not CELULAR.search(f"monto {cel}")
    assert CELULAR_CADENA.search(f'CONTACTO = ["{cel}"]') and not CELULAR_CADENA.search(f'ID = "a{cel}b"')
    sensible = tmp_path / "orden_de_servicio_2026.pdf"
    sensible.write_text("x")
    ruta = tmp_path / "script.py"
    ruta.write_text('RUTA = "C:/Users/alguien/Downloads"')
    problemas = dict(archivos_sensibles([sensible, ruta], usuario="alguien"))
    assert str(sensible) in problemas and str(ruta) in problemas
    assert re.search("categoría sensible", problemas[str(sensible)])


def test_ruta_externa(monkeypatch):
    from nucleo import entorno
    monkeypatch.setenv("REPO_PRUEBA_FASE2", "C:/datos/prueba")
    assert entorno.ruta_externa("repo_prueba_fase2").replace("\\", "/") == "C:/datos/prueba"
    monkeypatch.delenv("REPO_PRUEBA_FASE2")
    assert entorno.ruta_externa("clave_que_no_existe") is None
    assert entorno.herramienta("chrome").lower().endswith("chrome.exe")


def test_auditoria_commit_rutas(tmp_path, monkeypatch):
    from nucleo import auditoria
    monkeypatch.setenv("USERNAME", "usuaria_prueba")
    con_ruta = tmp_path / "script.py"
    con_ruta.write_text('R = "C:/Users/usuaria_prueba/datos"\n', encoding="utf-8")
    limpio = tmp_path / "otro.py"
    limpio.write_text("R = 1\n", encoding="utf-8")
    motivos = [m for _, m in auditoria.archivos_sensibles([con_ruta, limpio, tmp_path / "cv_postulante.pdf"])]
    assert any("ruta local" in m for m in motivos) and any("categoría sensible" in m for m in motivos)
    assert len(motivos) == 2


def test_contacto(monkeypatch):
    from nucleo import entorno
    monkeypatch.setattr(entorno, "_seccion", lambda nombre: {"nombre": "Especialista de prueba"} if nombre == "contacto"
                        else {})
    c = entorno.contacto()
    assert c["nombre"] == "Especialista de prueba" and c["correo"].startswith("[POR COMPLETAR")
