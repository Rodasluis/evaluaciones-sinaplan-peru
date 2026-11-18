"""Indicadores y recursos (REV_IND): avance de los indicadores frente a sus metas, calidad de las fichas y ejecución
presupuestal.

Entradas (02_insumos/):
    indicadores.csv          codigo, nivel (Objetivo prioritario o Servicio), op, nombre, sentido (Ascendente,
                             Descendente o vacío); opcionales servicio, linea_base, anio_base, responsable y vigente
    rev_ind_valores.csv      codigo, anio, meta, valor, fuente; opcionales pagina, comparable (Sí o No) y nota
    rev_ind_fichas.csv       (opcional) codigo, atributo (clave de tipologias/comun/atributos_ficha.yaml), cumple (Sí,
                             Parcial o No), severidad (Crítica, Mayor, Menor o vacío), observacion
    rev_ind_presupuesto.csv  (opcional) organo, anio, pia, pim, devengado, fuente; opcionales monto_otra_fuente y
                             otra_fuente, para la conciliación
Reglas: nucleo.escalas.avance (razón sin acotar según el sentido) y nucleo.escalas.nivel con escalas.avance de
config.yaml; nucleo.presupuesto.conciliar. Cada indicador se evalúa en todos los años del periodo: un año sin registro
queda «Sin meta» o «Sin valor», como hallazgo.
Salida: 04_analisis/REV_IND_indicadores_recursos.xlsx.
"""
import pandas as pd
import yaml

from ..config import RAIZ_REPO
from ..escalas import avance, nivel
from ..presupuesto import conciliar
from . import base

NOMBRE = "Indicadores y recursos"
LIBRO = "REV_IND_indicadores_recursos.xlsx"
PUBLICO = "Indicadores_y_recursos.xlsx"
NIVELES = ("Objetivo prioritario", "Servicio")
SENTIDOS = ("Ascendente", "Descendente")
CUMPLE = ("Sí", "Parcial", "No")
SEVERIDAD = ("Crítica", "Mayor", "Menor")
INDICADORES = ["codigo", "nivel", "op", "nombre", "sentido"]
VALORES = ["codigo", "anio", "meta", "valor", "fuente"]
FICHAS = ["codigo", "atributo", "cumple", "severidad", "observacion"]
PRESUPUESTO = ["organo", "anio", "pia", "pim", "devengado", "fuente"]


def entradas(cfg):
    return [("insumos", "indicadores.csv", INDICADORES[:3] + ["servicio", "nombre", "sentido", "linea_base", "anio_base",
                                                               "responsable", "vigente"]),
            ("insumos", "rev_ind_valores.csv", VALORES + ["pagina", "comparable", "nota"]),
            ("insumos", "rev_ind_fichas.csv", FICHAS),
            ("insumos", "rev_ind_presupuesto.csv", PRESUPUESTO + ["monto_otra_fuente", "otra_fuente"])]


def atributos():
    with open(RAIZ_REPO / "tipologias" / "comun" / "atributos_ficha.yaml", encoding="utf-8") as f:
        return yaml.safe_load(f)


def escala_avance(cfg):
    e = (cfg.get("escalas") or {}).get("avance")
    if not e:
        raise base.ErrorInsumo("REV_IND: config.yaml no tiene escalas.avance ([{desde, nivel}, …] de mayor a menor)")
    return [(float(x["desde"]), x["nivel"]) for x in e]


def leer_indicadores(cfg, p):
    ind = base.leer(cfg, "indicadores.csv", INDICADORES, p)
    if ind is not None:
        n = "indicadores.csv"
        base.unicos(ind, "codigo", n, p)
        base.en_escala(ind, "nivel", NIVELES, n, p)
        base.en_escala(ind, "sentido", SENTIDOS, n, p, vacio=True)
        for c in ("servicio", "linea_base", "anio_base", "responsable", "vigente"):
            if c not in ind.columns:
                ind[c] = "Sí" if c == "vigente" else ""
        base.en_escala(ind, "vigente", base.SI_NO, n, p)
    return ind


def validar(cfg):
    p = base.Problemas("REV_IND")
    esc = escala_avance(cfg)
    ind = leer_indicadores(cfg, p)
    val = base.leer(cfg, "rev_ind_valores.csv", VALORES, p)
    fic = base.leer(cfg, "rev_ind_fichas.csv", FICHAS, p, opcional=True)
    pre = base.leer(cfg, "rev_ind_presupuesto.csv", PRESUPUESTO, p, opcional=True)
    periodo = base.anios(cfg)
    if ind is not None and val is not None:
        n = "rev_ind_valores.csv"
        base.existentes(val, "codigo", ind.codigo, n, p)
        base.en_escala(val, "anio", [str(a) for a in periodo], n, p)
        rep = val[val.duplicated(["codigo", "anio"])]
        for _, r in rep.iterrows():
            p(f"{n}: {r.codigo} tiene más de un registro para {r.anio}")
        for c in ("pagina", "comparable", "nota"):
            if c not in val.columns:
                val[c] = "Sí" if c == "comparable" else ""
        base.en_escala(val, "comparable", base.SI_NO, n, p, vacio=True)
        val = base.numeros(val, ["meta", "valor"], n, p)
        for i, (m, v, f) in enumerate(zip(val.meta, val.valor, val.fuente), start=2):
            if (m is not None or v is not None) and not f:
                p(f"{n}, fila {i}: meta o valor sin fuente")
    if fic is not None and ind is not None:
        n = "rev_ind_fichas.csv"
        base.existentes(fic, "codigo", ind.codigo, n, p)
        base.en_escala(fic, "atributo", [a["clave"] for a in atributos()["atributos"]], n, p)
        base.en_escala(fic, "cumple", CUMPLE, n, p)
        base.en_escala(fic, "severidad", SEVERIDAD, n, p, vacio=True)
        for i, (c, s) in enumerate(zip(fic.cumple, fic.severidad), start=2):
            if c != "Sí" and not s:
                p(f"{n}, fila {i}: un atributo que no se cumple del todo requiere severidad")
    if pre is not None:
        n = "rev_ind_presupuesto.csv"
        base.en_escala(pre, "anio", [str(a) for a in periodo], n, p)
        pre = base.numeros(pre, ["pia", "pim", "devengado", "monto_otra_fuente"], n, p)
    p.revisar()
    return ind, val, fic, pre, esc


def serie(ind, val, periodo, esc):
    """Una fila por indicador y año del periodo, con su avance y su nivel."""
    reg = {(r.codigo, int(r.anio)): r for r in val.itertuples()}
    filas = []
    for i in ind.itertuples():
        for a in periodo:
            r = reg.get((i.codigo, a))
            meta, valor = (base.vacio(r.meta), base.vacio(r.valor)) if r is not None else (None, None)
            av, motivo = avance(i.sentido, meta, valor)
            no_comparable = r is not None and r.comparable == "No"
            filas.append({
                "Código": i.codigo, "Nivel": i.nivel, "Objetivo prioritario": i.op, "Servicio": i.servicio,
                "Indicador": i.nombre, "Año": a, "Sentido": i.sentido, "Meta": meta, "Valor": valor,
                "Avance (%)": round(av, 1) if av is not None else None,
                "Nivel de avance": nivel(av, esc) if av is not None else None,
                "Motivo sin avance": ("Sin registro" if r is None and i.sentido else motivo) or None,
                "Advertencia": "Fuente no comparable con la de la meta" if no_comparable else None,
                "Fuente": r.fuente if r is not None else None, "Página": (r.pagina or None) if r is not None else None,
                "Nota": (r.nota or None) if r is not None else None})
    return pd.DataFrame(filas)


def correr(cfg):
    ind, val, fic, pre, esc = validar(cfg)
    periodo = base.anios(cfg)
    larga = serie(ind, val, periodo, esc)
    ancha = larga.pivot_table(index="Código", columns="Año", values=["Meta", "Valor", "Avance (%)"], aggfunc="first",
                              dropna=False)
    ancha.columns = [f"{m} {a}" for m, a in ancha.columns]
    orden = [f"{m} {a}" for m in ("Meta", "Valor", "Avance (%)") for a in periodo]
    ancha = ind.set_index("codigo")[["nivel", "op", "nombre", "sentido"]].join(ancha.reindex(columns=orden))
    ancha = ancha.reset_index().rename(columns={"codigo": "Código", "nivel": "Nivel", "op": "Objetivo prioritario",
                                                "nombre": "Indicador", "sentido": "Sentido"})
    etiquetas = [n for _, n in esc]
    resumen = (larga.assign(Clase=larga["Nivel de avance"].fillna("Sin avance calculable"))
               .groupby(["Nivel", "Año"]).Clase.value_counts().unstack(fill_value=0)
               .reindex(columns=etiquetas + ["Sin avance calculable"], fill_value=0).reset_index())
    motivos = (larga[larga["Motivo sin avance"].notna()].groupby(["Nivel", "Año"])["Motivo sin avance"]
               .value_counts().unstack(fill_value=0).reset_index())
    hojas = {"Avance por año": larga, "Avance por indicador": ancha, "Resumen por año": resumen,
             "Sin avance por motivo": motivos}
    crit = [
        ("Objeto", "Avance anual de los indicadores de la política frente a sus metas, en el periodo "
                   f"{periodo[0]}–{periodo[-1]}.", "CEPLAN (Guía de seguimiento y evaluación: eficacia)"),
        ("Razón de avance", "Ascendente: valor entre meta por 100; descendente: meta entre valor por 100; sin acotar.",
         "Equipo evaluador (nucleo.escalas.avance)"),
        ("Niveles de avance", "; ".join(f"{n}: desde {d:g} %" for d, n in esc) + ".", "Equipo evaluador (config.yaml)"),
        ("Sin avance calculable", "Sin sentido definido, sin meta, sin valor, sin registro o meta o valor no "
                                  "interpretables; se consigna como hallazgo y no se estima.", "Equipo evaluador"),
    ]
    salida = {"indicadores": len(ind), "con_avance": int(larga["Avance (%)"].notna().sum())}
    if fic is not None:
        cat = {a["clave"]: a for a in atributos()["atributos"]}
        fic = fic.assign(Atributo=fic.atributo.map(lambda k: cat[k]["atributo"]),
                         Sustento=fic.atributo.map(lambda k: cat[k]["sustento"]))
        dictamen = []
        for c, g in fic.groupby("codigo", sort=False):
            no = g[g.cumple != "Sí"]
            critica = (no.severidad == "Crítica").any()
            dictamen.append({"Código": c, "Atributos revisados": len(g), "Atributos con observación": len(no),
                             "Observaciones críticas": int((no.severidad == "Crítica").sum()),
                             "Dictamen": "Con observación crítica" if critica else
                             ("Con observaciones" if len(no) else "Conforme")})
        hojas["Calidad de las fichas"] = fic.rename(columns={"codigo": "Código", "cumple": "Cumple",
                                                             "severidad": "Severidad", "observacion": "Observación"})[
            ["Código", "Atributo", "Cumple", "Severidad", "Observación", "Sustento"]]
        hojas["Dictamen de las fichas"] = pd.DataFrame(dictamen)
        crit.append(("Calidad de las fichas", f"Atributos y páginas de la {atributos()['guia']} "
                                              "(tipologias/comun/atributos_ficha.yaml). Dictamen: con observación "
                                              "crítica, con observaciones o conforme.", "CEPLAN (con página)"))
        salida["fichas"] = len(dictamen)
    if pre is not None:
        ej = pre.assign(**{"Ejecución sobre el PIM (%)": [round(100 * d / m, 1) if m else None
                                                          for d, m in zip(pre.devengado, pre.pim)]})
        hojas["Ejecución presupuestal"] = ej.rename(columns={"organo": "Órgano", "anio": "Año", "pia": "PIA (S/)",
                                                            "pim": "PIM (S/)", "devengado": "Devengado (S/)",
                                                            "fuente": "Fuente"})
        crit.append(("Ejecución presupuestal", "Devengado entre presupuesto institucional modificado (PIM) por 100.",
                     "Equipo evaluador"))
        if "monto_otra_fuente" in pre.columns and pre.monto_otra_fuente.notna().any():
            otra = pre.get("otra_fuente", pd.Series([""] * len(pre)))
            filas = [({"Órgano": r.organo, "Año": r.anio, "Otra fuente": o}, r.devengado, r.monto_otra_fuente)
                     for r, o in zip(pre.itertuples(), otra) if r.monto_otra_fuente is not None]
            hojas["Conciliación"] = conciliar(filas, "Devengado (S/)", "Monto de la otra fuente (S/)")
            crit.append(("Conciliación", "Coincide si la diferencia absoluta es menor que un sol.",
                         "Equipo evaluador (nucleo.presupuesto.conciliar)"))
    destino = base.ruta(cfg, "analisis", LIBRO)
    base.escribir_libro(destino, hojas, crit)
    return {"libro": destino, **salida}
