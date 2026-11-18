"""Segmentos codificados de las entrevistas y de la información de las entidades: validación, frecuencias y
concordancia de la doble codificación (nucleo.codificacion).

Los segmentos se guardan sin códigos, para dárselos tal cual al agente codificador-ciego; la primera codificación va en
un archivo aparte y la segunda, si la hay, en otro. El libro de códigos es común (02_insumos/libro_codigos.csv).
"""
import pandas as pd

from ..codificacion import acuerdo, jaccard, kappa
from . import base

SENTIDOS = ("Facilitador", "Obstáculo", "Neutro")
LIBRO_CODIGOS = ["codigo", "nombre", "definicion"]
CODIFICACION = ["id_segmento", "codigos", "sentido"]


def leer_libro(cfg, p):
    lib = base.leer(cfg, "libro_codigos.csv", LIBRO_CODIGOS, p)
    if lib is not None:
        base.unicos(lib, "codigo", "libro_codigos.csv", p)
        for c in LIBRO_CODIGOS:
            base.obligatorio(lib, c, "libro_codigos.csv", p)
    return lib


def leer_codificacion(cfg, carpeta, nombre, seg, lib, p, segunda=False):
    cols = ["id_segmento", "codigos"] if segunda else CODIFICACION
    cod = base.leer(cfg, nombre, cols, p, carpeta=carpeta, opcional=segunda)
    if cod is None or seg is None or lib is None:
        return cod
    base.unicos(cod, "id_segmento", nombre, p)
    base.existentes(cod, "id_segmento", seg.id_segmento, nombre, p)
    base.existentes(cod, "codigos", lib.codigo, nombre, p, separador=";")
    if not segunda:
        base.en_escala(cod, "sentido", SENTIDOS, nombre, p, vacio=True)
        sin = sorted(set(seg.id_segmento) - set(cod.id_segmento))
        if sin:
            p(f"{nombre}: segmentos sin codificar: {', '.join(sin[:10])}" + (" …" if len(sin) > 10 else ""))
    return cod


def frecuencias(seg, cod, lib, grupo):
    """Segmentos por código y número de valores distintos de grupo (rol o entidad) que lo mencionan."""
    larga = seg.merge(cod, on="id_segmento").assign(codigo=lambda d: d.codigos.map(base.lista)).explode("codigo")
    larga = larga[larga.codigo.notna() & (larga.codigo != "")]
    n = larga.groupby("codigo").agg(Segmentos=("id_segmento", "nunique"), Grupos=(grupo, "nunique"))
    f = lib.set_index("codigo")[["nombre", "definicion"]].join(n, how="left").fillna({"Segmentos": 0, "Grupos": 0})
    f = f.astype({"Segmentos": int, "Grupos": int}).sort_values("Segmentos", ascending=False).reset_index()
    return larga, f.rename(columns={"codigo": "Código", "nombre": "Nombre", "definicion": "Definición"})


def concordancia(cod, seg2, lib):
    """Acuerdo entre dos codificaciones: Jaccard medio, acuerdo exacto y kappa por código (presencia o ausencia)."""
    a = cod.set_index("id_segmento").codigos.map(lambda x: frozenset(base.lista(x)))
    b = seg2.set_index("id_segmento").codigos.map(lambda x: frozenset(base.lista(x)))
    comunes = a.index.intersection(b.index)
    a, b = a[comunes], b[comunes]
    filas = [{"Código": "Todos", "Segmentos": len(comunes),
              "Acuerdo exacto (%)": round(acuerdo(a, b), 1),
              "Jaccard medio": round(sum(jaccard(x, y) for x, y in zip(a, b)) / len(comunes), 2) if len(comunes) else None,
              "Kappa": None}]
    for c in lib.codigo:
        x, y = [c in s for s in a], [c in s for s in b]
        if any(x) or any(y):
            filas.append({"Código": c, "Segmentos": int(sum(x) + sum(y)), "Acuerdo exacto (%)": round(acuerdo(x, y), 1),
                          "Jaccard medio": None, "Kappa": kappa(x, y)})
    detalle = pd.DataFrame({"id_segmento": comunes, "Primera codificación": ["; ".join(sorted(s)) for s in a],
                            "Segunda codificación": ["; ".join(sorted(s)) for s in b],
                            "Coinciden": ["Sí" if x == y else "No" for x, y in zip(a, b)]})
    return pd.DataFrame(filas), detalle


CRITERIOS = [
    ("Codificación", "Libro de códigos común con definición, inclusión y exclusión; un segmento puede llevar más de un "
                     "código.", "Literatura (análisis temático)"),
    ("Doble codificación", "La segunda codificación la hace el agente codificador-ciego sin ver la primera; se mide el "
                           "acuerdo exacto, el índice de Jaccard y el kappa de Cohen por código. El equipo evaluador "
                           "resuelve todas las discrepancias y revisa una muestra de las coincidencias.",
     "Equipo evaluador"),
    ("Datos personales", "Los segmentos identifican a las personas por su cargo o rol, nunca por su nombre.",
     "Ley N.° 29733"),
]
