"""Revisión de evidencia (REV_EVI): solidez de la evidencia por tipo de intervención y plausibilidad causal de cada
servicio.

Entradas (02_insumos/):
    servicios.csv        la de la revisión documental
    rev_evi_tipos.csv    tipo, nombre, servicios («1.1.1; 1.1.3»), causa_efecto, resultado_esperado, pertinencia
                         (Directa o Parcial). Cada servicio pertenece a un solo tipo.
    rev_evi_estudios.csv id, tipos («T01; T03»), referencia, enlace, diseno, ambito, anio, calidad, direccion,
                         incluido (Sí o No), comparable (Sí o No); opcionales resultado_principal, verificacion y
                         motivo_exclusion (obligatorio si incluido = No)
    rev_evi_busquedas.csv (opcional) id, base, tipos, cadena, fecha, total, revisados; opcionales filtros, url,
                         seleccionados e incluidos. La escribe nucleo/busqueda_academica.py o el analista.
Regla: nucleo.evidencia.solidez sobre los estudios incluidos de cada tipo (solo cuentan los comparables para la
solidez moderada o sólida) y nucleo.evidencia.plausibilidad.
Salida: 04_analisis/REV_EVI_evidencia.xlsx.
"""
import pandas as pd

from ..evidencia import plausibilidad, solidez
from . import base, rev_doc

NOMBRE = "Revisión de evidencia"
LIBRO = "REV_EVI_evidencia.xlsx"
PUBLICO = "Revision_de_evidencia.xlsx"
DISENOS = ("Revisión sistemática", "Revisión narrativa", "Experimental", "Cuasi experimental", "Descriptivo")
PRIMARIOS = ("Experimental", "Cuasi experimental")
DIRECCIONES = ("Favorable", "Mixto", "Nulo", "Desfavorable", "Sin estimación")
TIPOS = ["tipo", "nombre", "servicios", "causa_efecto", "resultado_esperado", "pertinencia"]
ESTUDIOS = ["id", "tipos", "referencia", "enlace", "diseno", "ambito", "anio", "calidad", "direccion", "incluido",
            "comparable"]
BUSQUEDAS = ["id", "base", "tipos", "cadena", "fecha", "total", "revisados"]


def entradas(cfg):
    return [("insumos", "rev_evi_tipos.csv", TIPOS),
            ("insumos", "rev_evi_estudios.csv", ESTUDIOS[:2] + ["referencia", "enlace", "verificacion", "diseno", "ambito",
                                                              "anio", "calidad", "resultado_principal", "direccion",
                                                              "incluido", "comparable", "motivo_exclusion"]),
            ("insumos", "rev_evi_busquedas.csv", BUSQUEDAS + ["filtros", "seleccionados", "incluidos", "url"])]


def validar(cfg):
    p = base.Problemas("REV_EVI")
    s = rev_doc.leer_servicios(cfg, p)
    tip = base.leer(cfg, "rev_evi_tipos.csv", TIPOS, p)
    est = base.leer(cfg, "rev_evi_estudios.csv", ESTUDIOS, p)
    bus = base.leer(cfg, "rev_evi_busquedas.csv", BUSQUEDAS, p, opcional=True)
    if tip is not None:
        base.unicos(tip, "tipo", "rev_evi_tipos.csv", p)
        base.en_escala(tip, "pertinencia", ("Directa", "Parcial"), "rev_evi_tipos.csv", p)
        if s is not None:
            base.existentes(tip, "servicios", s.codigo, "rev_evi_tipos.csv", p, separador=";")
            asignados = pd.Series([c for x in tip.servicios for c in base.lista(x)])
            dobles = sorted(asignados[asignados.duplicated()].unique())
            if dobles:
                p(f"rev_evi_tipos.csv: servicios en más de un tipo: {', '.join(dobles)}")
            sin_tipo = sorted(set(s.codigo) - set(asignados))
            if sin_tipo:
                p(f"rev_evi_tipos.csv: servicios sin tipo de intervención: {', '.join(sin_tipo)}")
    if est is not None:
        n = "rev_evi_estudios.csv"
        base.unicos(est, "id", n, p)
        base.en_escala(est, "diseno", DISENOS, n, p)
        base.en_escala(est, "direccion", DIRECCIONES, n, p)
        base.en_escala(est, "incluido", base.SI_NO, n, p)
        base.en_escala(est, "comparable", base.SI_NO, n, p)
        base.obligatorio(est, "enlace", n, p)
        if tip is not None:
            base.existentes(est, "tipos", tip.tipo, n, p, separador=";")
        if "motivo_exclusion" not in est.columns:
            est["motivo_exclusion"] = ""
        for i, (inc, mot) in enumerate(zip(est.incluido, est.motivo_exclusion), start=2):
            if inc == "No" and not mot:
                p(f"{n}, fila {i}: un estudio excluido requiere motivo_exclusion")
        for c in ("resultado_principal", "verificacion"):
            if c not in est.columns:
                est[c] = ""
    if bus is not None and tip is not None:
        base.existentes(bus, "tipos", tip.tipo, "rev_evi_busquedas.csv", p, separador=";")
    p.revisar()
    return s, tip, est, bus


def por_tipo(tip, est):
    larga = est.assign(tipo=est.tipos.map(base.lista)).explode("tipo")
    larga = larga.assign(comparable=larga.comparable == "Sí")
    filas = []
    for t in tip.itertuples():
        g = larga[larga.tipo == t.tipo]
        inc = g[g.incluido == "Sí"]
        nivel, direccion = solidez(inc, PRIMARIOS)
        filas.append({"Tipo": t.tipo, "Tipo de intervención": t.nombre, "Servicios": t.servicios,
                      "Causa o efecto que atiende": t.causa_efecto, "Resultado esperado": t.resultado_esperado,
                      "Estudios incluidos": ", ".join(inc.id), "Estudios excluidos": ", ".join(g[g.incluido == "No"].id),
                      "Solidez": nivel, "Dirección": direccion, "Pertinencia del resultado": t.pertinencia,
                      "Plausibilidad": plausibilidad(nivel, direccion, t.pertinencia)})
    return pd.DataFrame(filas)


def correr(cfg):
    s, tip, est, bus = validar(cfg)
    matriz = por_tipo(tip, est)
    de_tipo = {c: r for r in matriz.to_dict("records") for c in base.lista(r["Servicios"])}
    columnas = ["Tipo", "Tipo de intervención", "Solidez", "Dirección", "Plausibilidad"]
    serv = pd.DataFrame([{"Código": sv.codigo, "Servicio": sv.servicio, **{c: de_tipo[sv.codigo][c] for c in columnas}}
                         for sv in s.itertuples()])
    estudios = est.rename(columns={"id": "Id", "tipos": "Tipos", "referencia": "Referencia", "enlace": "Enlace",
                                   "diseno": "Diseño", "ambito": "Ámbito", "anio": "Año", "calidad": "Calidad",
                                   "resultado_principal": "Resultado principal", "direccion": "Dirección",
                                   "incluido": "Incluido", "comparable": "Comparable",
                                   "motivo_exclusion": "Motivo de exclusión", "verificacion": "Verificación"})
    hojas = {"Matriz por tipo": matriz, "Plausibilidad por servicio": serv, "Estudios": estudios}
    if bus is not None:
        hojas["Búsquedas"] = bus
    crit = [
        ("Objeto", "Evidencia sobre la efectividad de cada tipo de intervención de la política y plausibilidad de que "
                   "sus servicios actúen sobre las causas o efectos del problema.",
         "CEPLAN (Guía de políticas nacionales: alternativas de solución basadas en evidencia)"),
        ("Sólida", "Al menos una revisión sistemática favorable cuya evidencia no sea de calidad baja, o dos o más "
                   "evaluaciones experimentales o cuasi experimentales comparables, todas favorables.", "Literatura"),
        ("Moderada", "Alguna revisión sistemática o evaluación primaria comparable, sin cumplir lo anterior.",
         "Literatura"),
        ("Limitada", "Solo otros estudios o contextos poco comparables.", "Literatura"),
        ("Inclusión y comparabilidad", "Las decide el equipo para cada estudio, con el motivo de cada exclusión.",
         "Equipo evaluador"),
        ("Búsqueda", "Estrategias con la cadena exacta por base, fecha, total de resultados y registros revisados.",
         "Literatura (búsquedas replicables)"),
    ]
    destino = base.ruta(cfg, "analisis", LIBRO)
    base.escribir_libro(destino, hojas, crit)
    return {"libro": destino, "tipos": len(matriz), "estudios": len(est),
            **{f"solidez_{k.lower()}": int(v) for k, v in matriz.Solidez.value_counts().items()}}
