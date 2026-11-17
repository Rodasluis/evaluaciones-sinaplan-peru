"""Los productos del piloto de la PNIG no cambian con la Fase 2 (comparación con tests/referencia_pnig/).

Si un cambio es intencional, se revisa la diferencia y se vuelve a capturar con: python tests/referencia.py capturar
"""
import difflib
import json

import pandas as pd
import pytest

import referencia as R

REF = json.loads((R.REF / "huellas.json").read_text(encoding="utf-8"))


def _diferencias(actual, esperado):
    faltan = sorted(set(esperado) - set(actual))
    sobran = sorted(set(actual) - set(esperado))
    cambian = sorted(k for k in set(actual) & set(esperado) if actual[k] != esperado[k])
    return faltan, sobran, cambian


@pytest.mark.parametrize("grupo,funcion", [("libros", R.huellas_libros), ("csv", R.huellas_csv),
                                           ("figuras", R.huellas_figuras), ("fijos", R.huellas_fijos)])
def test_huellas(grupo, funcion):
    esperado = {k: v for k, v in REF[grupo].items() if k not in R.excluidos()}   # versión pública: ver referencia.py
    faltan, sobran, cambian = _diferencias(funcion(), esperado)
    detalle = []
    if grupo == "libros":
        actual = funcion()
        for libro in cambian:
            f2, s2, c2 = _diferencias(actual[libro], esperado[libro])
            detalle.append(f"{libro}: hojas que faltan {f2}, nuevas {s2}, distintas {c2}")
    assert not (faltan or sobran or cambian), (f"{grupo}: faltan {faltan}; nuevos {sobran}; distintos {cambian}\n"
                                               + "\n".join(detalle))


def test_texto_informe():
    esperado = (R.REF / "informe_texto.txt").read_text(encoding="utf-8").splitlines()
    actual = R.texto_informe().splitlines()
    dif = list(difflib.unified_diff(esperado, actual, "referencia", "actual", n=0, lineterm=""))
    assert not dif, "el texto del informe cambió:\n" + "\n".join(dif[:60])


def test_cifras_clave():
    """Cifras clave del informe (desde la revisión institucional del 13-11-2026, ADR-015): estados de los servicios,
    conciliación presupuestal, confianza de los hallazgos por origen independiente y valoraciones."""
    a = R.EVAL / "04_analisis"
    est = pd.read_excel(a / "REV_DOC_estado_implementacion_servicios.xlsx", sheet_name="Estado por servicio")
    assert est["Estado asignado (preliminar)"].value_counts().to_dict() == {
        "Implementado parcialmente": 26, "Implementado": 24, "No implementado": 2}
    conc = pd.read_excel(a / "REV_IND_3_presupuesto_dgignd.xlsx", sheet_name="Conciliación MEF")
    assert (conc.Coincide == "Sí").sum() == 30 and len(conc) == 30
    tri = pd.read_excel(a / "matriz_triangulacion_preguntas.xlsx", sheet_name="Matriz de triangulación")
    assert tri["Confianza propuesta"].value_counts().to_dict() == {"Sólido": 24, "Con respaldo parcial": 10,
                                                                  "Indicio": 9}
    val = pd.read_excel(a / "matriz_triangulacion_preguntas.xlsx", sheet_name="Valoración por pregunta")
    assert val["Valoración"].value_counts().to_dict() == {"Parcial": 11, "Insuficiente": 1}
    fig = pd.read_csv(R.EVAL / "06_informe/indice_figuras.csv")
    assert len(fig) == 59
