"""Pruebas de las funciones que pasan al núcleo: formato, comillas, hojas publicadas y estimación con diseño muestral."""
import numpy as np
import pandas as pd
import pytest

import docx_mimp as F
import sintesis_hallazgos as S
import svy


def test_formato_numeros_y_montos():
    assert F.fmt_num(3314210.4) == "3 314 210"
    assert F.fmt_num(57.14, 1) == "57,1"
    assert F.monto(3_314_210) == "3,31 millones de soles"
    assert F.monto(845_210.4) == "845 210 soles"
    assert F.millones(2_230_000) == "2,23"
    assert F.millones(0) == "s. m."
    assert F.pct(57.14) == "57,1 %"
    assert F.lista_es(["a"]) == "a" and F.lista_es(["a", "b", "c"]) == "a, b y c"


def test_comillas_con_anidamiento():
    assert F.comillas("la hoja «Criterios»")[0] == "la hoja “Criterios”"
    assert F.comillas("dijo «la «Política» no»")[0] == "dijo “la ‘Política’ no”"
    # el nivel pasa de un fragmento al siguiente del mismo párrafo
    t1, n = F.comillas("abre «aquí")
    t2, n2 = F.comillas(" y cierra» luego", n)
    assert (t1, t2, n2) == ("abre “aquí", " y cierra” luego", 0)


def test_hojas_publicadas():
    for hoja, publicada in S.HOJAS_PUBLICAS.items():
        assert S.hoja_publica(hoja) == publicada
        assert len(publicada) <= 31
    assert S.hoja_publica("Criterios") == "Criterios"
    with pytest.raises(AssertionError):
        S.hoja_publica("Hoja REV_DOC sin nombre")


def _encuesta(n_estratos=6, conglomerados=8, personas=12, semilla=7):
    rng = np.random.default_rng(semilla)
    filas = []
    for h in range(n_estratos):
        for c in range(conglomerados):
            efecto = rng.normal(0, 0.05)
            for _ in range(personas):
                sexo = rng.choice(["Mujer", "Hombre"])
                p = (0.55 if sexo == "Mujer" else 0.75) + efecto
                filas.append({"estrato": h, "psu": f"{h}-{c}", "peso": rng.uniform(50, 150), "sexo": sexo,
                              "y": float(rng.random() < p)})
    return pd.DataFrame(filas)


def test_svy_estimar_y_brecha():
    df = _encuesta()
    d = svy.Diseno(df, peso="peso", estrato="estrato", psu="psu")
    r = svy.estimar(d, df.y)
    ponderada = (df.peso * df.y).sum() / df.peso.sum()
    assert r["estimacion"] == pytest.approx(ponderada)
    assert 0 < r["ee"] < 0.05 and r["n"] == len(df)
    assert r["ic_inf"] < r["estimacion"] < r["ic_sup"]
    b = svy.brecha(d, df.y, df.sexo, "Mujer", "Hombre", escala=100)
    assert b["diferencia"] == pytest.approx(b["valor_a"] - b["valor_b"])
    assert b["diferencia"] < 0 and b["ee_diferencia"] > 0
    dif, ee, sig = svy.cambio(50, 1.0, 54, 1.2)
    assert dif == 4 and sig and ee == pytest.approx((1 + 1.44) ** 0.5)


def test_svy_ambitos():
    assert svy.ambito_ubigeo("150101") == "LM" and svy.ambito_ubigeo("150501") == "RL"
    assert svy.ambito_ubigeo("080101") == "08"
    assert svy.codigo_departamento("  Huanuco") == "10"
