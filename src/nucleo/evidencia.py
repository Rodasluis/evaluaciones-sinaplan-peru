"""Revisión de evidencia: solidez y dirección de la evidencia de un tipo de intervención (regla del piloto, Anexo 6 de
la nota de la PNIG) y su plausibilidad.

Sólida: al menos una revisión sistemática favorable cuya evidencia no sea de calidad baja, o dos o más evaluaciones
primarias comparables, todas favorables. Moderada: alguna revisión sistemática o evaluación primaria comparable.
Limitada: solo otros estudios. Solo cuentan los estudios de contextos comparables (columna comparable).
"""
import pandas as pd

PRIMARIOS_POR_DEFECTO = ("Experimental", "Cuasi experimental")


def solidez(g, primarios=PRIMARIOS_POR_DEFECTO):
    """g: estudios incluidos de un tipo, con las columnas comparable, diseno, direccion y calidad."""
    comp = g[g.comparable]
    rs = comp[comp.diseno == "Revisión sistemática"]
    evals = comp[comp.diseno.isin(primarios)]
    rs_fav = rs[(rs.direccion == "Favorable") & ~rs.calidad.str.match(r"(Muy b|B)aja")]
    if len(rs_fav) or (len(evals) >= 2 and (evals.direccion == "Favorable").all()):
        nivel = "Sólida"
    elif len(evals) or len(rs):
        nivel = "Moderada"
    elif len(g):
        nivel = "Limitada"
    else:
        return "Sin evidencia identificada", ""
    base = pd.concat([rs, evals]) if len(rs) + len(evals) else g
    dirs = base.direccion[base.direccion != "Sin estimación"]
    if dirs.empty:
        direccion = "sin estimación de efectos"
    elif (dirs == "Favorable").all():
        direccion = "favorable"
    elif dirs.isin(["Nulo", "Desfavorable"]).all():
        direccion = "nula o desfavorable"
    else:
        direccion = "mixta"
    return nivel, direccion


def plausibilidad(nivel, direccion, pertinencia):
    if nivel == "Sin evidencia identificada":
        return nivel
    texto = f"{nivel}, {direccion}" if direccion else nivel
    return texto + (" (evidencia parcial)" if pertinencia == "Parcial" else "")
