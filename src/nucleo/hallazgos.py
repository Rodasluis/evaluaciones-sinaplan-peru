"""Hallazgos de una evaluación: modelo de evidencia, fuentes independientes, reglas de confianza y nombres publicados de
libros y hojas para el rastro (núcleo; extraído de src/analisis_pnig/sintesis_hallazgos.py).

Un hallazgo es un dict con id, pregunta, aspecto, enunciado, evidencias, confianza, estado, nota, salvedad y detalle.
Cada evidencia es un dict de evidencia(). La capa de cada evaluación escribe sus hallazgos y llama a verificar().

Reglas de confianza (nota metodológica del piloto, Tabla 9):
  Sólido: al menos dos fuentes de origen independiente coinciden.
  Con respaldo parcial: las fuentes coinciden, con discrepancias que se explican (también dos o más orígenes).
  Indicio: un solo origen independiente.
Desde la revisión institucional del piloto (2026-11-13), la independencia se cuenta por el origen primario del dato y no
por el documento: dos documentos que reproducen el dato de un mismo productor (p. ej., el informe LIO y el formulario,
ambos reportados por las entidades) cuentan como uno.
"""
import re

CONFIANZA = ("Sólido", "Con respaldo parcial", "Indicio", "No determinado")
VALORACION = ("Satisfactorio", "Parcial", "Insuficiente")
GUION = re.compile(r" — | - ")


def evidencia(plano, instrumento, texto, fuente, base, n_fuentes=1, origen=None):
    """n_fuentes > 1 cuando la evidencia agrega varias fuentes primarias independientes (p. ej., una matriz de
    evidencia con varios estudios). origen: origen primario del dato cuando no se deduce de la base (p. ej., un
    recálculo que reproduce el dato oficial con la misma encuesta)."""
    e = {"plano": plano, "instrumento": instrumento, "texto": texto, "fuente": fuente, "base": base,
         "n_fuentes": n_fuentes}
    if origen:
        e["origen"] = origen
    return e


def fuentes_independientes(h, no_independiente=None, partes=lambda base: [base], origen=lambda parte: parte):
    """Orígenes primarios distintos que respaldan el hallazgo, sin contar la base no independiente; una entrevista
    cuenta como un origen; una evidencia que agrega varias fuentes primarias cuenta por cada una (n_fuentes).
    partes: función que separa una base compuesta en sus fuentes («ENAHO; INEI…» -> [«ENAHO», «INEI…»]);
    origen: función fuente -> origen primario (por defecto, cada fuente es su propio origen); el campo «origen» de una
    evidencia tiene prioridad."""
    origenes = set()
    for e in h["evidencias"]:
        if e["base"] == no_independiente:
            continue
        if e.get("origen"):
            origenes.add(e["origen"])
            continue
        fuentes = [e["base"]] if e["base"].startswith("Entrevista") else partes(e["base"])
        origenes |= {origen(f) for f in fuentes}
    return max(len(origenes), max((e["n_fuentes"] for e in h["evidencias"]), default=0))


def verificar(H, preguntas, contar, internos=None, max_palabras_detalle=60):
    """Comprueba la consistencia de los hallazgos (falla con AssertionError):
    identificadores únicos, pregunta y confianza válidas, dos fuentes para sólido o con respaldo parcial, evidencia
    salvo en los pendientes, sin guiones explicativos, sin términos internos en lo publicable (internos: regex),
    detalle breve y al menos un hallazgo por pregunta. contar: función h -> número de fuentes independientes."""
    ids = [h["id"] for h in H]
    assert len(ids) == len(set(ids)), "identificadores de hallazgo repetidos"
    for h in H:
        assert h["pregunta"] in preguntas, h["id"]
        assert h["confianza"] in CONFIANZA, h["id"]
        n = contar(h)
        if h["confianza"] in ("Sólido", "Con respaldo parcial"):
            assert n >= 2, f"{h['id']}: {h['confianza']} requiere dos fuentes independientes"
        if h["confianza"] == "No determinado":
            assert h["estado"] == "Análisis pendiente" and not h["evidencias"], h["id"]
        else:
            assert h["evidencias"], f"{h['id']}: hallazgo sin evidencia"
        for txt in [h["enunciado"]] + [e["texto"] for e in h["evidencias"]]:
            assert not GUION.search(txt), f"{h['id']}: guion explicativo"
        for campo in ("enunciado", "detalle", "salvedad", "nota"):
            txt = h[campo] + h.get("nota_tabla", "") * (campo == "nota")
            if internos is not None:
                assert not internos.search(txt), f"{h['id']}, {campo}: término interno «{internos.search(txt).group(0)}»"
            assert not GUION.search(txt), f"{h['id']}, {campo}: guion explicativo"
        palabras = len(h["detalle"].split())
        assert palabras <= max_palabras_detalle, (f"{h['id']}: detalle de {palabras} palabras; redactar uno de 1 a 3 "
                                                  f"oraciones")
    for p in preguntas:
        assert any(h["pregunta"] == p for h in H), f"pregunta sin hallazgo ni pendiente: {p}"


def hoja_publica(hoja, tabla, siglas=r"REV_|REP_|ENTREV"):
    """Nombre publicado de una hoja: el de la tabla si lleva sigla; si no, el mismo. Sin siglas de instrumento y con 31
    caracteres o menos (límite de Excel), para que las citas del informe coincidan con las hojas entregadas."""
    h = tabla.get(hoja, hoja)
    assert not re.search(siglas, h), f"hoja con sigla sin nombre publicado: {hoja}"
    assert len(h) <= 31, f"nombre de hoja de más de 31 caracteres: {h}"
    return h


def libro_publico(libro, tabla):
    """Nombre publicado de un libro de análisis (tabla: nombre en el repositorio -> nombre publicado)."""
    nombre = libro.split("/")[-1]
    assert nombre in tabla, f"libro sin nombre publicado: {nombre}"
    return tabla[nombre]
