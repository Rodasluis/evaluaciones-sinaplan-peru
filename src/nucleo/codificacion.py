"""Codificación temática: concordancia entre dos codificaciones (doble codificación ciega).

La segunda codificación la hace un codificador que no ve la primera (agente codificador-ciego); estas funciones miden
el acuerdo. El equipo evaluador resuelve todas las discrepancias y revisa una muestra de las coincidencias.
"""
import pandas as pd


def kappa(x, y):
    """Kappa de Cohen para dos vectores de categorías; None si el acuerdo esperado es 1."""
    x, y = pd.Series(list(x)), pd.Series(list(y))
    po = (x == y).mean()
    pe = sum(x.value_counts(normalize=True).get(k, 0) * y.value_counts(normalize=True).get(k, 0)
             for k in set(x) | set(y))
    return None if pe == 1 else round((po - pe) / (1 - pe), 2)


def jaccard(a, b):
    """Índice de Jaccard entre dos conjuntos de códigos (1 si ambos están vacíos)."""
    a, b = set(a), set(b)
    return 1.0 if not a and not b else len(a & b) / len(a | b)


def acuerdo(x, y):
    """Porcentaje de acuerdo simple entre dos vectores de categorías."""
    x, y = list(x), list(y)
    return 100 * sum(i == j for i, j in zip(x, y)) / len(x) if x else float("nan")
