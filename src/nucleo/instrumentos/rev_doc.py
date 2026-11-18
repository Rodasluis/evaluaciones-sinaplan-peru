"""Revisión documental (REV_DOC): estado de implementación de los servicios de la política.

Entradas (02_insumos/):
    servicios.csv             codigo, op, lineamiento, servicio; opcionales objetivo_prioritario y responsable
    rev_doc_componentes.csv   codigo; una columna por componente de escalas.implementacion (Sí o No) con su
                              <columna>_fuente; ejecucion_<AAAA> por año del periodo (Sí, No o Sin información) con
                              ejecucion_fuente; ajuste (estado asignado por lectura cualitativa, opcional) y
                              sustento_ajuste (obligatorio si hay ajuste)
Regla (nucleo.escalas.clasificar_implementacion, parámetros de config.yaml):
    Implementado: al menos minimo_componentes de los componentes y ejecución en alguno de los años de cierre.
    Implementado parcialmente: ejecución en algún año del periodo, o valor reportado del indicador del servicio.
    No implementado: todos los años del periodo sin ejecución.
    Sin información suficiente: lo demás.
Salida: 04_analisis/REV_DOC_estado_servicios.xlsx (estado por servicio, resumen por objetivo, criterios).
"""
import pandas as pd

from ..escalas import IMPLEMENTADO, NO_IMPLEMENTADO, PARCIAL, SIN_INFO, clasificar_implementacion
from ..fuentes import slug
from . import base

NOMBRE = "Revisión documental"
LIBRO = "REV_DOC_estado_servicios.xlsx"
PUBLICO = "Estado_implementacion_servicios.xlsx"
ESTADOS = (IMPLEMENTADO, PARCIAL, NO_IMPLEMENTADO, SIN_INFO)
EJECUCION = ("Sí", "No", "Sin información")
SERVICIOS = ["codigo", "op", "lineamiento", "servicio"]
# columna de cada componente en rev_doc_componentes.csv; un componente nuevo usa su nombre sin tildes
COLUMNA = {"protocolo": "protocolo", "asignación identificable": "asignacion", "reporte de indicador": "reporte_indicador"}


def columna(componente):
    return COLUMNA.get(componente, slug(componente).replace("-", "_"))


def escala(cfg):
    e = (cfg.get("escalas") or {}).get("implementacion")
    if not e:
        raise base.ErrorInsumo("REV_DOC: config.yaml no tiene escalas.implementacion (componentes, minimo_componentes "
                               "y anios_cierre)")
    return {"componentes": list(e["componentes"]), "minimo": int(e["minimo_componentes"]),
            "cierre": [int(a) for a in e["anios_cierre"]]}


def cabecera(cfg):
    """Columnas de rev_doc_componentes.csv para esta evaluación (las usa el generador de evaluaciones)."""
    e = escala(cfg)
    cols = ["codigo"]
    for c in e["componentes"]:
        cols += [columna(c), f"{columna(c)}_fuente"]
    return cols + [f"ejecucion_{a}" for a in base.anios(cfg)] + ["ejecucion_fuente", "ajuste", "sustento_ajuste"]


def entradas(cfg):
    """Archivos de entrada con todas sus columnas (cabeceras que escribe el generador de evaluaciones)."""
    salida = [("insumos", "servicios.csv", ["codigo", "op", "objetivo_prioritario", "lineamiento", "servicio",
                                             "responsable"])]
    if (cfg.get("escalas") or {}).get("implementacion"):
        salida.append(("insumos", "rev_doc_componentes.csv", cabecera(cfg)))
    return salida


def leer_servicios(cfg, problemas):
    s = base.leer(cfg, "servicios.csv", SERVICIOS, problemas)
    if s is not None:
        base.unicos(s, "codigo", "servicios.csv", problemas)
        for c in SERVICIOS:
            base.obligatorio(s, c, "servicios.csv", problemas)
        for c in ("objetivo_prioritario", "responsable"):
            if c not in s.columns:
                s[c] = ""
    return s


def validar(cfg):
    p = base.Problemas("REV_DOC")
    e = escala(cfg)
    s = leer_servicios(cfg, p)
    comp = base.leer(cfg, "rev_doc_componentes.csv", cabecera(cfg)[:1] + [columna(c) for c in e["componentes"]]
                     + [f"ejecucion_{a}" for a in base.anios(cfg)], p)
    if s is not None and comp is not None:
        n = "rev_doc_componentes.csv"
        base.unicos(comp, "codigo", n, p)
        base.existentes(comp, "codigo", s.codigo, n, p)
        faltan = sorted(set(s.codigo) - set(comp.codigo))
        if faltan:
            p(f"{n}: servicios sin registro: {', '.join(faltan)}")
        for c in e["componentes"]:
            base.en_escala(comp, columna(c), base.SI_NO, n, p)
        for a in base.anios(cfg):
            base.en_escala(comp, f"ejecucion_{a}", EJECUCION, n, p, vacio=True)
        for col in ("ajuste", "sustento_ajuste"):
            if col not in comp.columns:
                comp[col] = ""
        base.en_escala(comp, "ajuste", ESTADOS, n, p, vacio=True)
        for i, (aj, su) in enumerate(zip(comp.ajuste, comp.sustento_ajuste), start=2):
            if aj and not su:
                p(f"{n}, fila {i}: el ajuste «{aj}» requiere sustento_ajuste")
    p.revisar()
    return s, comp, e


def clasificar(r, e, periodo):
    """Clasificación por criterios de un registro de rev_doc_componentes.csv."""
    faltan = [c for c in e["componentes"] if r[columna(c)] != "Sí"]
    ej = {a: (r[f"ejecucion_{a}"] or "Sin información") for a in periodo}
    reporte = [c for c in e["componentes"] if "reporte" in c]
    reciente = any(ej[a] == "Sí" for a in e["cierre"])
    en_periodo = any(v == "Sí" for v in ej.values()) or any(r[columna(c)] == "Sí" for c in reporte)
    sin_ejecucion = all(v == "No" for v in ej.values())
    return faltan, clasificar_implementacion(faltan, reciente, en_periodo, sin_ejecucion, e["minimo"],
                                             len(e["componentes"]))


def correr(cfg):
    s, comp, e = validar(cfg)
    periodo = base.anios(cfg)
    por_codigo = comp.set_index("codigo")
    filas = []
    for _, sv in s.iterrows():
        r = por_codigo.loc[sv.codigo]
        faltan, automatico = clasificar(r, e, periodo)
        fila = {"Código": sv.codigo, "Objetivo prioritario": sv.op, "Lineamiento": sv.lineamiento,
                "Servicio": sv.servicio, "Responsable": sv.responsable}
        for c in e["componentes"]:
            fila[c[0].upper() + c[1:]] = r[columna(c)]
            fila[f"Fuente: {c}"] = r.get(f"{columna(c)}_fuente", "")
        fila |= {f"Ejecución {a}": r[f"ejecucion_{a}"] or "Sin información" for a in periodo}
        fila |= {"Fuente de la ejecución": r.get("ejecucion_fuente", ""),
                 "Componentes faltantes": ", ".join(faltan), "Clasificación por criterios": automatico,
                 "Ajuste por lectura cualitativa": r.ajuste, "Sustento del ajuste": r.sustento_ajuste,
                 "Estado asignado": r.ajuste or automatico}
        filas.append(fila)
    estado = pd.DataFrame(filas)
    resumen = (estado.groupby("Objetivo prioritario", sort=False)["Estado asignado"].value_counts().unstack(fill_value=0)
               .reindex(columns=list(ESTADOS), fill_value=0))
    resumen.insert(0, "Servicios", resumen.sum(axis=1))
    resumen.loc["Total"] = resumen.sum()
    resumen["Implementados (%)"] = (100 * resumen[IMPLEMENTADO] / resumen["Servicios"]).round(1)
    resumen = resumen.reset_index()
    n = len(e["componentes"])
    crit = [
        ("Objeto", "Estado de implementación de cada servicio de la política en el periodo "
                   f"{periodo[0]}–{periodo[-1]}.", "CEPLAN (Guía de seguimiento y evaluación, p. 73: implementación)"),
        ("Componentes", f"{', '.join(e['componentes'])}; cada uno con su fuente.", "Equipo evaluador (config.yaml)"),
        ("Implementado", f"Al menos {e['minimo']} de los {n} componentes y ejecución reportada en "
                         f"{' o '.join(map(str, e['cierre']))}.", "Equipo evaluador (config.yaml)"),
        ("Implementado parcialmente", "Ejecución reportada en algún año del periodo, o valor reportado del indicador del "
                                      "servicio, sin cumplir la condición anterior.", "Equipo evaluador (config.yaml)"),
        ("No implementado", "Todos los años del periodo registrados sin ejecución.", "Equipo evaluador (config.yaml)"),
        ("Sin información suficiente", "Años sin información que impiden aplicar las reglas anteriores.",
         "Equipo evaluador (config.yaml)"),
        ("Ajuste por lectura cualitativa", "Cambia el estado solo con sustento escrito; la clasificación por criterios "
                                           "se conserva para el contraste.", "Equipo evaluador"),
    ]
    destino = base.ruta(cfg, "analisis", LIBRO)
    base.escribir_libro(destino, {"Estado por servicio": estado, "Resumen por objetivo": resumen}, crit)
    cuenta = estado["Estado asignado"].value_counts().to_dict()
    return {"libro": destino, "servicios": len(estado), **{k: int(cuenta.get(k, 0)) for k in ESTADOS}}
