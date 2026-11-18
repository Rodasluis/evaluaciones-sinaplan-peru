"""Diagramas y gráficos del informe de evaluación de cierre de la PNIG (2019–2025).

Cada figura se arma como una página HTML (CSS y SVG en línea, sin dependencias externas) a partir de los libros de
análisis y de las cifras de sintesis_hallazgos.py, y se renderiza a PNG con Chrome sin interfaz (escala 3, unos 450
ppp a 16 cm). Los HTML quedan junto a los PNG como fuente editable; informe_cierre.py inserta los PNG y la presentación
de resultados los reutiliza. Las figuras no llevan título dentro de la imagen: la leyenda numerada va en el informe.

Estilo (guía dataviz): Arial, como el cuerpo del informe; fondo blanco; los OP con los seis primeros colores de la
paleta categórica validada (validate_palette.js sobre #ffffff: pasa; tres colores bajo 3:1 de contraste, por lo que
cada marca de OP lleva su etiqueta visible); el estado de implementación con la paleta de estado, siempre con texto;
la correspondencia de alineamiento con una rampa ordinal azul y gris para «sin correspondencia».

Ronda 27 (pedido del usuario): sin barras laterales ni bordes de color en las cajas, que delataban el HTML; la categoría
se lee por el relleno de la caja o por un punto de color. Ronda 30: a pedido del usuario, los colores vuelven a ser los
originales (rellenos plenos, con texto blanco sobre los oscuros); las barras no vuelven.

Salida: evaluaciones/pnig_2026/06_informe/figuras/<clave>.html y .png, e indice_figuras_informe.csv (clave, archivo,
título, fuente). Chrome corre con un perfil temporal, para no tocar el del usuario.

Uso (desde la raíz del repositorio):
    uv run --no-project --with pandas --with openpyxl --with python-docx python src/analisis_pnig/figuras_informe.py
"""
import html
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import pandas as pd

sys.path.insert(0, os.path.dirname(__file__))
import sintesis_hallazgos as S  # noqa: E402
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
from nucleo.config import insumo  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

SALIDA = Path(f"{EVAL}/06_informe/figuras")
A = Path(f"{EVAL}/04_analisis")
MAC = A / "REV_ALI_coherencia_alineamiento.xlsx"
GRD = A / "REV_DOC_estado_implementacion_servicios.xlsx"
SERVICIOS = Path(f"{EVAL}/02_insumos/servicios.csv")   # catálogo de servicios
PNIG = Path("docs/fuentes/pnig/01_normativa/2019_mimp_ds-008-2019-mimp_pnig.pdf")

from nucleo.figuras import (AZUL, CHROME, CSS, ESCALA, ESTADO, GRILLA, TENUE, TINTA, TINTA2,  # noqa: E402,F401
                            _chrome, e, leyenda, mezcla, pagina, renderizar)

# paleta de la PNIG
OP = {"OP1": "#2a78d6", "OP2": "#eb6834", "OP3": "#1baf7a", "OP4": "#eda100", "OP5": "#e87ba4", "OP6": "#008300"}
# azul medio para bloques y barras grandes de las figuras de indicadores y de actores (ronda 33: el azul oscuro se veía
# muy fuerte); con texto blanco, contraste de 5,8:1
AZUL_MEDIO = "#4A6698"
CORRESP = {"Directa": "#1c5cab", "Indirecta": "#86b6ef", "Sin correspondencia": "#f0efec"}



# ---------------------------------------------------------------------------------------------------------------
# datos comunes
def datos():
    C = S.cifras()
    srv = pd.read_csv(SERVICIOS, dtype=str).fillna("")
    grd = pd.read_excel(GRD, sheet_name="Estado por servicio", dtype={"Código": str}).set_index("Código")
    obj = pd.read_excel(MAC, sheet_name="REV_ALI.1 Por objetivo").set_index("OP")
    mod = pd.read_excel(MAC, sheet_name="REV_ALI.1 Modelo causal")
    ps = pd.read_excel(MAC, sheet_name="REV_ALI.1 Por servicio", dtype={"Código": str})
    lin = pd.read_excel(MAC, sheet_name="REV_ALI.1 Por lineamiento", dtype=str)
    nombres_lin = {re.match(r"\s*(\d\.\d)", x).group(1): re.sub(r"^\s*\d\.\d\.?\s*", "", x).strip()
                   for x in lin.Lineamiento}
    return {"C": C, "srv": srv, "grd": grd, "obj": obj, "mod": mod, "ps": ps, "lin": nombres_lin}


def op_corto(D, op):
    return D["obj"].loc[op, "Objetivo prioritario"]


# ---------------------------------------------------------------------------------------------------------------
# F2. Modelo del problema público
def problema(D):
    mod = D["mod"]
    causas, efectos = mod[mod.Tipo == "Causa"], mod[mod.Tipo == "Efecto"]

    def tarjeta(r, fondo):
        ops = " ".join(f'<span class="op" style="background:{OP[o]}">{o}</span>' for o in r.OP.split(", "))
        return (f'<div class="t" style="background:{fondo}">'
                f'<div class="n">{e(r["Causa o efecto del modelo"])}</div>'
                f'<div class="m">{r.Servicios} servicio{"s" * (r.Servicios != 1)} · {ops}</div></div>')

    cuerpo = (
        '<div class="g">'
        '<div class="c"><div class="col-t">Causas</div>'
        + "".join(tarjeta(r, "#eef2f8") for _, r in causas.iterrows()) +
        '</div><div class="fl">&#10140;</div>'
        f'<div class="p"><div class="col-t" style="color:#fff;opacity:.8">Problema público</div>'
        f'<div class="pt">Discriminación estructural contra las mujeres</div></div>'
        '<div class="fl">&#10140;</div>'
        '<div class="c"><div class="col-t">Efectos</div>'
        + "".join(tarjeta(r, "#f7f1ee") for _, r in efectos.iterrows()) +
        "</div></div>"
        '<div class="nota">Número de servicios de la PNIG vinculados a cada causa o efecto en la matriz de consistencia '
        "(Anexo 5 de la PNIG) y objetivos prioritarios a los que pertenecen.</div>")
    css = f"""
.g{{display:grid;grid-template-columns:1fr 34px 210px 34px 1fr;align-items:center;gap:6px}}
.c{{display:flex;flex-direction:column;gap:9px}}
.t{{border-radius:8px;padding:10px 12px}}
.n{{font-size:13.5px;line-height:1.3;color:{TINTA}}}
.m{{font-size:11.5px;color:{TINTA2};margin-top:6px;display:flex;gap:4px;align-items:center;flex-wrap:wrap}}
.op{{color:#fff;font-weight:bold;font-size:10.5px;border-radius:4px;padding:1px 5px}}
.fl{{font-size:26px;color:{TENUE};text-align:center}}
.p{{background:{AZUL};border-radius:12px;padding:22px 16px;text-align:center}}
.pt{{color:#fff;font-size:19px;font-weight:bold;line-height:1.3;margin-top:8px}}
"""
    return pagina(cuerpo, 1000, css), 1000


# ---------------------------------------------------------------------------------------------------------------
# F3. Cadena de valor: causas y efectos -> OP -> lineamientos, servicios, indicadores y situación futura deseada
def cadena(D):
    ps, obj, C = D["ps"], D["obj"], D["C"]
    enlaces = {}
    for _, r in ps.iterrows():
        for nodo in filter(None, (x.strip() for x in str(r["Causa o efecto del modelo (Anexo 5)"]).split(";"))):
            if nodo != "nan":
                enlaces[(nodo, r.OP)] = enlaces.get((nodo, r.OP), 0) + 1
    mod = D["mod"]
    nodos = list(mod["Causa o efecto del modelo"])
    tipo = dict(zip(mod["Causa o efecto del modelo"], mod.Tipo))
    ops = list(OP)
    n_ind = C["op_tabla"].groupby("OP").size()
    n_sfd = C["sfd"].groupby("OP").size()
    W, H, top = 1000, 640, 40
    x0, w0, x1, w1 = 0, 250, 330, 170
    gap = 10

    def bloques(claves, totales, alto_min):
        disp = H - top - gap * (len(claves) - 1)
        tot = sum(totales[k] for k in claves)
        altos = {k: max(alto_min, disp * totales[k] / tot) for k in claves}
        f = disp / sum(altos.values())
        y, pos = top, {}
        for k in claves:
            pos[k] = (y, altos[k] * f)
            y += altos[k] * f + gap
        return pos

    tot_nodo = {n: sum(v for (a, _), v in enlaces.items() if a == n) for n in nodos}
    tot_op = {o: sum(v for (_, b), v in enlaces.items() if b == o) for o in ops}
    pn, po = bloques(nodos, tot_nodo, 52), bloques(ops, {o: 1 for o in ops}, 0)
    # cintas
    off_n = {n: 0.0 for n in nodos}
    off_o = {o: 0.0 for o in ops}
    cintas = []
    for n in nodos:
        for o in ops:
            v = enlaces.get((n, o), 0)
            if not v:
                continue
            yn, hn = pn[n]
            yo, ho = po[o]
            an, ao = hn * v / tot_nodo[n], ho * v / tot_op[o]
            a1, a2 = yn + off_n[n], yo + off_o[o]
            off_n[n] += an
            off_o[o] += ao
            xa, xb = x0 + w0, x1
            xm = (xa + xb) / 2
            cintas.append(f'<path d="M{xa},{a1} C{xm},{a1} {xm},{a2} {xb},{a2} L{xb},{a2 + ao} C{xm},{a2 + ao} '
                          f'{xm},{a1 + an} {xa},{a1 + an} Z" fill="{OP[o]}" fill-opacity="0.32"/>')
    svg = f'<svg width="{W}" height="{H}" style="position:absolute;left:0;top:0">{"".join(cintas)}</svg>'
    cajas = []
    for n in nodos:
        y, h = pn[n]
        fondo = "#eef2f8" if tipo[n] == "Causa" else "#f7f1ee"
        cajas.append(f'<div class="nd" style="top:{y}px;height:{h}px;width:{w0}px;left:{x0}px;background:'
                     f'{fondo}"><span class="k">{e(tipo[n])}</span> {e(n)}</div>')
    for o in ops:
        y, h = po[o]
        cajas.append(f'<div class="op" style="top:{y}px;height:{h}px;width:{w1}px;left:{x1}px;background:{OP[o]}">'
                     f'<b>{o}</b><span class="on">{e(op_corto(D, o))}</span></div>')
        filas = [("Lineamientos", int(obj.loc[o, "Lineamientos"])),
                 ("Servicios", int(obj.loc[o, "Servicios"])),
                 ("con indicador vigente", int(obj.loc[o, "Servicios con indicador vigente"])),
                 ("Indicadores del OP", int(n_ind.get(o, 0))),
                 ("Situación futura deseada", int(n_sfd.get(o, 0)))]
        celdas = "".join(f'<div class="cel" style="background:{mezcla(OP[o], 0.10)}"><b>{v}</b><span>{e(t)}</span></div>'
                         for t, v in filas)
        cajas.append(f'<div class="lane" style="top:{y}px;height:{h}px;left:{x1 + w1 + 14}px;width:'
                     f'{W - x1 - w1 - 14}px">{celdas}</div>')
    cab = (f'<div class="hd" style="left:{x0}px;width:{w0}px">Causas y efectos del problema</div>'
           f'<div class="hd" style="left:{x1}px;width:{w1}px">Objetivos prioritarios</div>'
           f'<div class="hd" style="left:{x1 + w1 + 14}px;width:{W - x1 - w1 - 14}px">Despliegue: lineamientos, '
           f'servicios e indicadores</div>')
    cuerpo = (f'<div style="position:relative;height:{H}px">{svg}<div style="position:absolute;left:0;top:0">'
              f'{cab}{"".join(cajas)}</div></div>'
              '<div class="nota">El ancho de cada cinta es el número de servicios que vinculan una causa o un efecto con '
              'el objetivo prioritario (matriz de consistencia, Anexo 5 de la PNIG). «Con indicador vigente»: servicios con '
              'indicador aprobado en su protocolo. «Situación futura deseada»: enunciados de la PNIG vinculados al OP.'
              '</div>')
    css = f"""
.hd{{position:absolute;top:8px;font-size:11px;font-weight:bold;letter-spacing:.04em;text-transform:uppercase;color:{TINTA2}}}
.nd{{position:absolute;border-radius:6px;padding:6px 9px;font-size:11.5px;line-height:1.25;
     display:flex;align-items:center;overflow:hidden}}
.nd .k{{display:none}}
.op{{position:absolute;border-radius:8px;color:#fff;padding:7px 10px;font-size:11px;line-height:1.22;display:flex;
     flex-direction:column;justify-content:center;gap:3px}}
.op b{{font-size:15px}}
.on{{display:-webkit-box;-webkit-line-clamp:4;-webkit-box-orient:vertical;overflow:hidden}}
.lane{{position:absolute;display:grid;grid-template-columns:repeat(5,1fr);gap:6px}}
.cel{{border-radius:6px;display:flex;flex-direction:column;justify-content:center;align-items:center;text-align:center;
      padding:4px}}
.cel b{{font-size:22px;color:{TINTA}}}
.cel span{{font-size:10.5px;color:{TINTA2};line-height:1.2;margin-top:2px}}
"""
    return pagina(cuerpo, W + 44, css), W + 44


# ---------------------------------------------------------------------------------------------------------------
# F14. Mapa de los 52 servicios por OP y lineamiento con su estado de implementación
def mapa_servicios(D, ops):
    srv, grd = D["srv"], D["grd"]
    srv = srv.assign(lin=srv.codigo_corto.str.rsplit(".", n=1).str[0])
    bloques = []
    for o in ops:
        s = srv[srv.op == f"OP.0{o[-1]}"]
        filas = []
        for lin, g in s.groupby("lin", sort=False):
            nombre_lin = D["lin"][lin]
            tiles = []
            for _, r in g.iterrows():
                st = grd.loc[r.codigo_corto, "Estado asignado (preliminar)"]
                prot = grd.loc[r.codigo_corto, "Protocolo aprobado"] == "Sí"
                ind = "ISS" in r.indicadores
                nombre = re.sub(r"\s+", " ", r.servicio).strip()
                badges = (f'<span class="bd {"on" if prot else ""}">Protocolo</span>'
                          f'<span class="bd {"on" if ind else ""}">Indicador</span>')
                tiles.append(f'<div class="tl"><div class="cd"><span class="es" style="background:{ESTADO[st]}"></span>'
                             f'{e(r.codigo_corto)}</div><div class="nm">{e(nombre)}</div>'
                             f'<div class="bs">{badges}</div></div>')
            filas.append((lin, nombre_lin, tiles))
        # un OP con un solo servicio por lineamiento (OP5) va en columnas, para no dejar la figura vacía a la derecha
        if len(ops) == 1 and max(len(t) for _, _, t in filas) == 1 and len(filas) > 2:
            # dos filas comunes (nombre del lineamiento y tarjeta), para que las tarjetas queden alineadas
            cuerpo_op = (f'<div class="lc" style="grid-template-columns:repeat({len(filas)},1fr)">'
                         + "".join(f'<div class="lt">{"<b>" + e(lin) + "</b> " + e(nl)}</div>{t[0]}'
                                   for lin, nl, t in filas) + '</div>')
        else:
            cuerpo_op = "".join(f'<div class="ln"><div class="lt"><b>{e(lin)}</b> {e(nl)}</div>'
                                f'<div class="ts">{"".join(t)}</div></div>' for lin, nl, t in filas)
        bloques.append(f'<div class="ob"><div class="oh" style="background:{OP[o]}"><b>{o}</b> {e(op_corto(D, o))}'
                       f'</div>{cuerpo_op}</div>')
    ley = leyenda([(k, v) for k, v in ESTADO.items()])
    cuerpo = ("".join(bloques) + ley +
              '<div class="nota">Punto junto al código: estado de implementación asignado en la revisión documental. «Protocolo» resaltado: '
              'protocolo aprobado al 31-12-2025; «Indicador» resaltado: indicador vigente.'
              + (' El lineamiento 3.2 no tuvo servicio.' if 'OP3' in ops else '') + '</div>')
    css = f"""
.ob{{margin-bottom:12px}}
.oh{{color:#fff;font-size:13px;border-radius:6px;padding:6px 10px}}
.oh b{{margin-right:6px}}
.ln{{display:grid;grid-template-columns:170px 1fr;gap:8px;margin-top:6px;align-items:start}}
.lt{{font-size:11px;color:{TINTA2};line-height:1.25;padding-top:4px}}
.lt b{{color:{TINTA}}}
.ts{{display:grid;grid-template-columns:repeat(5,1fr);gap:6px}}
.lc{{display:grid;grid-template-rows:auto auto;grid-auto-flow:column;column-gap:8px;row-gap:5px;margin-top:6px}}
.tl{{background:#f7f7f4;border-radius:6px;padding:6px 7px;min-height:86px;display:flex;flex-direction:column}}
.cd{{font-size:12px;font-weight:bold;display:flex;align-items:center;gap:5px}}
.es{{width:10px;height:10px;border-radius:50%;flex:none}}
.nm{{font-size:10px;line-height:1.22;color:{TINTA2};margin-top:2px;display:-webkit-box;-webkit-line-clamp:3;
     -webkit-box-orient:vertical;overflow:hidden;max-height:3.66em}}
.bs{{display:flex;gap:4px;margin-top:auto;padding-top:4px}}
.bd{{font-size:9px;border:1px solid {GRILLA};color:{TENUE};border-radius:3px;padding:0 4px}}
.bd.on{{background:{AZUL};border-color:{AZUL};color:#fff}}
"""
    return pagina(cuerpo, 1000, css), 1000


# ---------------------------------------------------------------------------------------------------------------
# F1. Línea de tiempo normativa e institucional (fechas del registro de fuentes)
HITOS = [(x["fecha"], x["categoria"], x["hito"], x["instrumento"]) for x in insumo("hitos.yaml")]
# ronda 33 (pedido del usuario): sin los hitos de seguimiento y evaluación ni los de la transición de 2026
CAT_HITO = {"Marco normativo": AZUL, "Implementación": "#127a73"}
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "setiembre", "octubre", "noviembre",
         "diciembre"]


def _fecha_txt(f):
    p = f.split("-")
    return f"{int(p[0])} de {MESES[int(p[1]) - 1]} de {p[2]}" if len(p) == 3 else f"{MESES[int(p[0]) - 1]} de {p[1]}"


def linea_tiempo(D):
    filas, anio_prev = [], None
    for i, (f, cat, hito, inst) in enumerate(HITOS):
        anio = f[-4:]
        marca = f'<div class="an">{anio}</div>' if anio != anio_prev else '<div class="an"></div>'
        anio_prev = anio
        lado = "izq" if i % 2 == 0 else "der"
        tarjeta = (f'<div class="tc"><div class="fe">{_fecha_txt(f)}</div>'
                   f'<div class="hi">{e(hito)}</div><div class="in">{e(inst)}</div></div>')
        filas.append(f'<div class="fl">{tarjeta if lado == "izq" else "<div></div>"}'
                     f'<div class="eje">{marca}<span class="pt" style="background:{CAT_HITO[cat]}"></span></div>'
                     f'{tarjeta if lado == "der" else "<div></div>"}</div>')
    cuerpo = f'<div class="tl">{"".join(filas)}</div>' + leyenda(list(CAT_HITO.items()))
    css = f"""
.tl{{position:relative}}
.tl:before{{content:"";position:absolute;left:50%;top:0;bottom:0;width:2px;background:{GRILLA};transform:translateX(-1px)}}
.fl{{display:grid;grid-template-columns:1fr 70px 1fr;align-items:center;margin-bottom:-14px}}
.eje{{position:relative;height:100%;display:flex;flex-direction:column;align-items:center;justify-content:center}}
.an{{font-size:12px;font-weight:bold;color:{TINTA2};background:#fff;padding:1px 4px;position:relative;z-index:1}}
.pt{{width:12px;height:12px;border-radius:50%;border:2px solid #fff;position:relative;z-index:1;margin-top:2px}}
.tc{{background:#f7f7f4;border-radius:7px;padding:7px 11px;margin:3px 0}}
.fl:nth-child(even) .tc{{text-align:right}}
.fe{{font-size:10.5px;color:{TENUE}}}
.hi{{font-size:12.5px;line-height:1.25;margin-top:1px}}
.in{{font-size:10.5px;color:{TINTA2};margin-top:2px}}
"""
    return pagina(cuerpo, 900, css), 900


# ---------------------------------------------------------------------------------------------------------------
# F4. Indicadores: de objetivo, vigentes de servicio y del PEMIG por tipo
TIPO_IND = {"Cobertura": AZUL_MEDIO, "Calidad": "#6250d6", "Otro": "#c3c2b7"}
TINTA_IND = {"Cobertura": "#fff", "Calidad": "#fff", "Otro": TINTA}


def indicadores(D):
    C, srv = D["C"], D["srv"]
    tp = pd.read_excel(A / "REV_IND_2_calidad_fichas.xlsx", sheet_name="Tipo PEMIG", dtype={"servicio": str})
    vig = srv.assign(n=srv.indicadores.str.count(r"ISS\.")).groupby("op").n.sum()
    n_op = C["op_tabla"].groupby("OP").size()
    tiles = (f'<div class="st"><b>{C["op_ind"]}</b><span>indicadores de objetivo prioritario</span></div>'
             f'<div class="st"><b>{int(vig.sum())}</b><span>indicadores vigentes de servicio, aprobados con los '
             f'protocolos</span></div><div class="st"><b>{len(tp)}</b><span>indicadores de servicio del PEMIG '
             f'(su anexo 5), que no se siguieron</span></div>')
    maximo = tp.groupby("OP").size().max()
    filas = []
    for o in OP:
        g = tp[tp.OP == o].tipo.value_counts()
        seg = "".join(f'<span class="sg" style="width:{100 * g.get(t, 0) / maximo}%;background:{c};color:{TINTA_IND[t]}">'
                      f'{g.get(t, 0) if g.get(t, 0) >= 2 else ""}</span>' for t, c in TIPO_IND.items())
        filas.append(f'<div class="fr"><span class="opc" style="background:{OP[o]}">{o}</span>'
                     f'<div class="br">{seg}</div><div class="nm"><b>{int(n_op.get(o, 0))}</b> de objetivo · '
                     f'<b>{int(vig.get(f"OP.0{o[-1]}", 0))}</b> vigentes</div></div>')
    cuerpo = (f'<div class="sts">{tiles}</div><div class="col-t" style="margin:18px 0 8px">Indicadores de servicio '
              f'del PEMIG por tipo y objetivo prioritario</div>{"".join(filas)}'
              + leyenda([(f"{t} ({int((tp.tipo == t).sum())})", c) for t, c in TIPO_IND.items()]) +
              '<div class="nota">Tipo asignado a cada indicador con las definiciones del Recuadro 2 de la guía de '
              'indicadores del CEPLAN (pp. 61 y 62). «Otro»: participación de las mujeres entre los beneficiarios, '
              'paridad, resultado en la población o gestión de la información.</div>')
    css = f"""
.sts{{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}}
.st{{background:#f2f4f8;border-radius:10px;padding:14px 16px;display:flex;gap:12px;align-items:center}}
.st b{{font-size:38px;color:{AZUL_MEDIO}}}
.st span{{font-size:12.5px;color:{TINTA2};line-height:1.3}}
.fr{{display:grid;grid-template-columns:46px 1fr 190px;gap:10px;align-items:center;margin-bottom:7px}}
.opc{{color:#fff;font-weight:bold;font-size:11.5px;border-radius:4px;text-align:center;padding:3px 0}}
.br{{display:flex;gap:2px;height:24px}}
.sg{{display:flex;align-items:center;justify-content:center;color:#fff;font-size:11.5px;font-weight:bold;
     border-radius:3px}}
.nm{{font-size:12px;color:{TINTA2}}}
.nm b{{color:{TINTA}}}
"""
    return pagina(cuerpo, 1000, css), 1000


# ---------------------------------------------------------------------------------------------------------------
# F5. Mapa de actores
def actores(D):
    grd = D["grd"]
    cuenta, por_op = {}, {}
    for cod, r in grd.iterrows():
        for ent in (x.strip() for x in str(r["Entidad u órgano responsable"]).split(";")):
            base = "MIMP" if ent.startswith("MIMP") else ent
            cuenta[base] = cuenta.get(base, 0) + 1
            o = f"OP{int(r.OP[-2:])}"
            por_op.setdefault(o, {}).setdefault(base, 0)
            por_op[o][base] += 1
    organos_mimp = sorted({org.strip() for v in grd["Entidad u órgano responsable"] for x in str(v).split(";")
                           if x.strip().startswith("MIMP (")
                           for org in re.split(r",| y ", re.sub(r"^MIMP \((.*)\)$", r"\1", x.strip()))})
    cond = [("DGIGND", "Dirección General de Igualdad de Género y No Discriminación: conduce la Política; su "
                       "Dirección de Políticas (DPIGND) formula, diseña y brinda asistencia técnica"),
            ("OGPPM", "Oficina General de Planeamiento, Presupuesto y Modernización: planeamiento y punto focal "
                      "ante el CEPLAN"),
            ("OGMEPGD", "Oficina General de Monitoreo, Evaluación de Políticas y Gestión Descentralizada: seguimiento "
                        "y evaluación (OMEP)")]
    conductores = "".join(f'<div class="cd"><b>{s}</b><span>{e(t)}</span></div>' for s, t in cond)
    columnas = []
    for o in OP:
        ents = sorted(por_op.get(o, {}).items(), key=lambda x: (-x[1], x[0]))
        chips = "".join(f'<div class="ch"><span>{e(n)}</span><b>{v}</b></div>' for n, v in ents)
        columnas.append(f'<div class="oc"><div class="oh" style="background:{OP[o]}">{o}</div>{chips}</div>')
    cuerpo = (
        '<div class="top">'
        f'<div class="bx ex"><div class="col-t">Rector del Sinaplan</div><b>CEPLAN</b><span>Guías de seguimiento y '
        'evaluación, aplicativo de seguimiento, opinión técnica sobre la Política</span></div>'
        f'<div class="bx mimp"><div class="col-t" style="color:#fff;opacity:.85">Rectoría y conducción: MIMP</div>'
        f'<div class="cds">{conductores}</div></div>'
        f'<div class="bx ex"><div class="col-t">Articulación multisectorial</div><b>CONAIG</b><span>Comisión '
        'Multisectorial de Alto Nivel (D.S. N.° 018-2021-MIMP): seguimiento de la implementación e informes anuales '
        'de gobernanza</span></div></div>'
        f'<div class="col-t" style="margin:16px 0 8px">Entidades responsables de los servicios, por objetivo '
        f'prioritario (número de servicios a su cargo)</div><div class="ops">{"".join(columnas)}</div>'
        f'<div class="nota">{len(cuenta)} entidades responsables de los {len(grd)} servicios; un servicio con dos '
        f'proveedores cuenta para ambos. MIMP agrupa a sus órganos y programas: {e(", ".join(organos_mimp))}. El INEI '
        'es además la fuente de las encuestas de los indicadores de objetivo.</div>')
    css = f"""
.top{{display:grid;grid-template-columns:1fr 2.1fr 1fr;gap:12px;align-items:stretch}}
.bx{{border-radius:10px;padding:12px 14px;display:flex;flex-direction:column;gap:5px}}
.ex{{background:#f2f4f8}}
.ex b{{font-size:20px;color:{AZUL_MEDIO}}}
.ex span{{font-size:11.5px;color:{TINTA2};line-height:1.3}}
.mimp{{background:{AZUL_MEDIO};color:#fff}}
.cds{{display:flex;flex-direction:column;gap:6px;margin-top:3px}}
.cd{{display:grid;grid-template-columns:70px 1fr;gap:8px;background:rgba(255,255,255,.1);border-radius:6px;
     padding:6px 8px;align-items:center}}
.cd b{{font-size:13px}}
.cd span{{font-size:11px;line-height:1.28;opacity:.92}}
.ops{{display:grid;grid-template-columns:repeat(6,1fr);gap:8px;align-items:start}}
.oh{{color:#fff;font-weight:bold;font-size:12.5px;text-align:center;border-radius:5px;padding:4px 0;margin-bottom:5px}}
.ch{{display:flex;justify-content:space-between;gap:4px;background:#f7f7f4;border-radius:5px;padding:4px 7px;
     margin-bottom:4px;font-size:11px}}
.ch b{{color:{AZUL}}}
"""
    return pagina(cuerpo, 1000, css), 1000


# ---------------------------------------------------------------------------------------------------------------
# F6 y F7. Matrices de alineamiento (OP × instrumento)
def _tipo_corresp(celda):
    t = str(celda)
    if "(directa" in t or t.strip().endswith("Directa"):
        return "Directa"
    if "(indirecta" in t:
        return "Indirecta"
    return "Sin correspondencia"


def _ref_corta(celda, n=40):
    t = re.sub(r"\s*\([^)]*\)", "", str(celda)).strip()
    t = "Sin correspondencia" if t.startswith("Ninguna") or not t or t == "nan" else t
    return t if len(t) <= n else t[:n - 1].rstrip(" ,;") + "…"


def _matriz(D, hoja, grupos, nota):
    m = pd.read_excel(MAC, sheet_name=hoja).set_index("OP")
    cab_g, cab_c, filas = "", "", []
    for nombre, cols in grupos:
        cab_g += f'<div class="gh" style="grid-column:span {len(cols)}">{e(nombre)}</div>'
        cab_c += "".join(f'<div class="ch">{e(etq)}</div>' for _, etq in cols)
    for o in OP:
        celdas = ""
        for _, cols in grupos:
            for col, _ in cols:
                tipo = _tipo_corresp(m.loc[o, col])
                fondo = CORRESP[tipo]
                tinta = "#fff" if tipo == "Directa" else TINTA
                celdas += f'<div class="cc" style="background:{fondo};color:{tinta}">{e(_ref_corta(m.loc[o, col]))}</div>'
        filas.append(f'<div class="rh"><span style="background:{OP[o]}">{o}</span>{e(op_corto(D, o))}</div>{celdas}')
    n = sum(len(c) for _, c in grupos)
    cuerpo = (f'<div class="mx" style="grid-template-columns:200px repeat({n},1fr)"><div></div>{cab_g}<div></div>'
              f'{cab_c}{"".join(filas)}</div>' + leyenda(list(CORRESP.items())) + f'<div class="nota">{nota}</div>')
    css = f"""
.mx{{display:grid;gap:3px}}
.gh{{background:{AZUL};color:#fff;font-size:11.5px;font-weight:bold;text-align:center;border-radius:4px;padding:4px}}
.ch{{font-size:10.5px;color:{TINTA2};text-align:center;line-height:1.2;padding:3px 2px;display:flex;align-items:end;
     justify-content:center}}
.rh{{font-size:10.5px;line-height:1.2;color:{TINTA2};display:flex;gap:6px;align-items:center}}
.rh span{{color:#fff;font-weight:bold;border-radius:4px;padding:2px 5px;font-size:11px;flex:none}}
.cc{{border-radius:4px;font-size:9.5px;line-height:1.2;padding:5px 4px;text-align:center;display:flex;
     align-items:center;justify-content:center;min-height:46px}}
"""
    return pagina(cuerpo, 1100, css), 1100


def alin_vertical(D):
    per = [("Políticas de Estado [periodo]", "Políticas de Estado"), ("Plan Bicentenario [periodo]", "Plan Bicentenario"),
           ("PEDN al 2050 (2022) [periodo]", "PEDN al 2050 (2022)"),
           ("PGG al 2021 (D.S. 056-2018-PCM) [periodo]", "PGG al 2021"),
           ("PGG 2021–2026 (D.S. 164-2021-PCM) [periodo]", "PGG 2021–2026"),
           ("PGG 2023–2026 (D.S. 042-2023-PCM) [periodo]", "PGG 2023–2026"),
           ("PGG 2025–2026 (D.S. 141-2025-PCM) [periodo]", "PGG 2025–2026"), ("PESEM 2018–2026 [periodo]", "PESEM 2018–2026")]
    tra = [("Políticas de Estado [transición]", "Políticas de Estado"),
           ("PEDN al 2050 actualizado (2023) [transición]", "PEDN actualizado (2023)"),
           ("PGG 2026 (D.S. 043-2026-PCM) [transición]", "PGG 2026"),
           ("PGG 2026–2031 (D.S. 127-2026-PCM) [transición]", "PGG 2026–2031"),
           ("PESEM 2024–2030 [transición]", "PESEM 2024–2030")]
    return _matriz(D, "REV_ALI.2 Por objetivo", [("Periodo 2019–2025", per), ("Transición", tra)],
                   "Cada celda indica el objetivo, eje o lineamiento del instrumento con el que se vincula el OP. "
                   "Directa: el instrumento contiene el mismo resultado; indirecta: contribuye a él.")


def alin_internacional(D):
    cols = [("Ley N.° 32535", "Ley N.° 32535"), ("CEDAW", "CEDAW"),
            ("Observaciones finales del Comité CEDAW (2022)", "Observaciones CEDAW (2022)"),
            ("Convención de Belém do Pará", "Belém do Pará"), ("Plataforma de Acción de Beijing", "Plataforma de Beijing"),
            ("ODS 5 (Agenda 2030)", "ODS 5"), ("Estrategia de Montevideo", "Estrategia de Montevideo")]
    return _matriz(D, "REV_ALI.4 Por objetivo", [("Marco nacional", cols[:1]), ("Compromisos internacionales", cols[1:])],
                   "Cada celda indica el artículo, esfera, meta o medida con que se vincula el OP.")


# ---------------------------------------------------------------------------------------------------------------
# F8. Alineamiento horizontal: políticas afines y servicios de la PNIG
REL = {"Duplicidad": "#ec835a", "Complementariedad": "#1c5cab", "Vacío de atención": "#898781"}


def alin_horizontal(D):
    h = pd.read_excel(MAC, sheet_name="REV_ALI.3 Horizontal", dtype=str)
    pol = h.groupby("Política nacional afín").size().sort_values(ascending=False)
    srv = sorted(h["Servicio o lineamiento de la PNIG"].unique(), key=lambda s: [int(x) for x in re.findall(r"\d+", s)])
    op_de = dict(zip(h["Servicio o lineamiento de la PNIG"], h["OP de la PNIG"]))
    W, fila_i, fila_d = 1000, 30, 21
    H = max(len(pol) * fila_i, len(srv) * fila_d) + 10
    yi = {p: 8 + i * (H - 10) / len(pol) + (H - 10) / len(pol) / 2 for i, p in enumerate(pol.index)}
    yd = {s: 8 + i * (H - 10) / len(srv) + (H - 10) / len(srv) / 2 for i, s in enumerate(srv)}
    xa, xb = 330, 760
    lineas = "".join(
        f'<path d="M{xa},{yi[r["Política nacional afín"]]} C{(xa + xb) / 2},{yi[r["Política nacional afín"]]} '
        f'{(xa + xb) / 2},{yd[r["Servicio o lineamiento de la PNIG"]]} {xb},{yd[r["Servicio o lineamiento de la PNIG"]]}" '
        f'stroke="{REL[r["Tipo de relación"]]}" stroke-width="2" fill="none" stroke-opacity=".75"/>'
        for _, r in h.iterrows())
    izq = "".join(f'<div class="pl" style="top:{yi[p] - 11}px"><span>{e(p)}</span><b>{n}</b></div>' for p, n in pol.items())
    der = "".join(f'<div class="sv" style="top:{yd[s] - 9}px"><span style="background:{OP.get(op_de[s], TENUE)}">'
                  f'{e(op_de[s] if op_de[s] in OP else "Varios")}</span>{e(s)}'
                  f'{"" if op_de[s] in OP else " (" + e(op_de[s]) + ")"}</div>' for s in srv)
    cuerpo = (f'<div style="position:relative;height:{H}px"><svg width="{W}" height="{H}" style="position:absolute;'
              f'left:0;top:0">{lineas}</svg>{izq}{der}'
              f'<div class="hd" style="left:0">Política nacional afín (relaciones)</div>'
              f'<div class="hd" style="left:{xb + 8}px">Servicio o lineamiento de la PNIG</div></div>'
              + leyenda([(f"{t} ({int((h['Tipo de relación'] == t).sum())})", c) for t, c in REL.items()]) +
              f'<div class="nota">{len(h)} relaciones entre {len(pol)} políticas nacionales afines y {len(srv)} servicios '
              'o lineamientos de la PNIG. Duplicidad: la política afín provee el mismo servicio; '
              'complementariedad: aporta al mismo resultado con otro servicio; vacío: la PNIG no tiene servicio '
              'para la población que la política afín atiende.</div>')
    css = f"""
.pl{{position:absolute;left:0;width:{xa - 8}px;font-size:11px;line-height:1.1;color:{TINTA};display:flex;
     justify-content:flex-end;gap:5px}}
.pl span{{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
.pl b{{color:{AZUL};flex:none}}
.sv{{position:absolute;left:{xb + 8}px;font-size:11px;display:flex;gap:5px;align-items:center;white-space:nowrap}}
.sv span{{color:#fff;font-weight:bold;border-radius:3px;padding:0 4px;font-size:9.5px}}
.hd{{position:absolute;top:-16px;font-size:10.5px;font-weight:bold;letter-spacing:.04em;text-transform:uppercase;
     color:{TINTA2}}}
.fig{{padding-top:34px}}
"""
    return pagina(cuerpo, W + 44, css), W + 44


# ---------------------------------------------------------------------------------------------------------------
# F9. Marco conceptual: del enfoque de la PNIG al de igualdad de oportunidades
REL_C = {"Continuidad": "#1c5cab", "Precisión": "#86b6ef", "Divergencia": "#ec835a"}
EFECTO = ["Se mantiene", "Se ajusta", "Se rediseña"]


def conceptual_planos(D):
    return conceptual(D, "planos")


def conceptual_conceptos(D):
    return conceptual(D, "conceptos")


def conceptual(D, parte):
    """Ronda 34 (pedido del usuario): la figura se divide en dos, las relaciones por plano y el plano conceptual."""
    c = pd.read_excel(MAC, sheet_name="REV_ALI.5 Conceptual", dtype=str)
    res = pd.read_excel(MAC, sheet_name="Resumen REV_ALI.5").set_index("Plano")
    barras = ""
    for plano, r in res.iterrows():
        tot = sum(r[t] for t in REL_C)
        seg = "".join(f'<span style="width:{100 * r[t] / tot}%;background:{col};color:{"#fff" if t != "Precisión" else TINTA}">'
                      f'{int(r[t])}</span>' for t, col in REL_C.items() if r[t])
        efe = " · ".join(f"{int(r[x])} {x.lower()}" for x in EFECTO)
        barras += (f'<div class="pr"><div class="pn">{e(plano)} <small>({int(tot)})</small></div><div class="bb">{seg}'
                   f'</div><div class="ef">{efe}</div></div>')
    filas = ""
    for _, r in c[c.Plano == "Conceptual"].iterrows():
        col = REL_C[r["Tipo de relación"]]
        filas += (f'<div class="br"><div class="a">{e(r["Elemento de la PNIG"])}</div>'
                  f'<div class="m"><span style="background:{col};color:{"#fff" if r["Tipo de relación"] != "Precisión" else TINTA}">'
                  f'{e(r["Tipo de relación"])}</span><small>{e(r["Efecto"])}</small></div>'
                  f'<div class="b">{e(r["Categoría del nuevo enfoque"])}</div></div>')
    if parte == "planos":
        cuerpo = (f'<div class="col-t">Relaciones por plano de análisis</div><div style="margin:8px 0 6px">{barras}</div>'
                  + leyenda(list(REL_C.items())) +
                  '<div class="nota">Número de relaciones entre los elementos de la PNIG y las categorías de la Ley N.° '
                  '32535 o del marco conceptual de la DGIGND, por plano (entre paréntesis, el total). Continuidad: '
                  'mismo contenido; precisión: lo acota o redefine; divergencia: lo sustituye. A la derecha, el efecto '
                  'sobre los elementos en la nueva política.</div>')
    else:
        cuerpo = (f'<div class="hdr"><div class="col-t">Concepto de la PNIG</div><div class="col-t" '
                  f'style="text-align:center">Relación y efecto</div><div class="col-t">Categoría de la Ley N.° 32535 o '
                  f'del marco conceptual</div></div>{filas}' + leyenda(list(REL_C.items())) +
                  '<div class="nota">Plano conceptual: cada concepto del documento de la PNIG frente a la categoría de '
                  'la Ley N.° 32535 o del marco conceptual de la DGIGND. El efecto indica qué pasa con el elemento en '
                  'la nueva política.</div>')
    css = f"""
.pr{{display:grid;grid-template-columns:190px 1fr 260px;gap:10px;align-items:center;margin-bottom:6px}}
.pn{{font-size:12px}}
.pn small{{color:{TENUE}}}
.bb{{display:flex;gap:2px;height:22px}}
.bb span{{display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:bold;border-radius:3px}}
.ef{{font-size:11px;color:{TINTA2}}}
.hdr,.br{{display:grid;grid-template-columns:1fr 130px 1fr;gap:10px;align-items:center}}
.hdr{{margin-bottom:4px}}
.br{{border-top:1px solid {GRILLA};padding:5px 0}}
.a,.b{{font-size:11px;line-height:1.25}}
.m{{display:flex;flex-direction:column;align-items:center;gap:2px}}
.m span{{font-size:10.5px;font-weight:bold;border-radius:10px;padding:2px 9px}}
.m small{{font-size:10px;color:{TINTA2}}}
"""
    return pagina(cuerpo, 1000, css), 1000


# ---------------------------------------------------------------------------------------------------------------
# Sesiones y acuerdos de la CONAIG (ronda 34): de la ayuda memoria de la comisión
# (docs/fuentes/pnig/09_conaig_gobernanza/2026_mimp-conaig_ayuda-memoria-sesiones-2021-2026.doc)
SESIONES_CONAIG = [(x["anio"], x["fechas"], x["acuerdos"]) for x in insumo("conaig_sesiones.yaml")]


def conaig_sesiones(D):
    cols = ""
    for anio, fechas, acuerdos in SESIONES_CONAIG:
        fuera = anio > 2025
        chips = "".join(f'<div class="ac{" pr" if "protocolo" in a.lower() else ""}">{e(a)}</div>' for a in acuerdos)
        cols += (f'<div class="co{" fu" if fuera else ""}"><div class="an">{anio}</div><div class="ns"><b>{len(fechas)}</b>'
                 f'<span>{"sesión" if len(fechas) == 1 else "sesiones"}</span></div><div class="fe">{e(", ".join(fechas))}'
                 f'</div>{chips}</div>')
    cuerpo = (f'<div class="g">{cols}</div>'
              + leyenda([("Acuerdo sobre los protocolos pendientes", ESTADO["Implementado parcialmente"])]) +
              '<div class="nota">Sesiones ordinarias por año, con sus fechas y tres acuerdos principales. En 2021 hubo '
              'además una sesión de instalación. 2026 queda fuera del periodo evaluado (una sesión hasta el 12 de junio). '
              'Fuente: ayuda memoria de la CONAIG.</div>')
    css = f"""
.g{{display:grid;grid-template-columns:repeat({len(SESIONES_CONAIG)},1fr);gap:8px;align-items:start}}
.co{{background:#f7f7f4;border-radius:8px;padding:8px}}
.co.fu{{opacity:.6}}
.an{{background:{AZUL_MEDIO};color:#fff;font-weight:bold;font-size:13px;text-align:center;border-radius:5px;padding:4px 0}}
.ns{{display:flex;align-items:baseline;gap:5px;justify-content:center;margin-top:6px}}
.ns b{{font-size:26px;color:{AZUL_MEDIO}}}
.ns span{{font-size:11px;color:{TINTA2}}}
.fe{{font-size:10px;color:{TENUE};text-align:center;margin-bottom:6px}}
.ac{{font-size:10.5px;line-height:1.25;background:#fff;border-radius:5px;padding:5px 6px;margin-top:5px;color:{TINTA}}}
.ac.pr{{background:{ESTADO["Implementado parcialmente"]};font-weight:bold}}
"""
    return pagina(cuerpo, 1000, css), 1000


# ---------------------------------------------------------------------------------------------------------------
# F10. Flujo de búsqueda y selección de la evidencia (REV_EVI)
REV = A / "REV_EVI_matriz_evidencia.xlsx"


def _canal(origen):
    o = str(origen)
    if o.startswith("Búsqueda académica"):
        return "Búsqueda académica"
    if o.startswith("Búsqueda en el 3ie DEP"):
        return "3ie DEP"
    if o.startswith(("Búsqueda en Scopus", "Búsqueda en SciELO", "Búsqueda en LILACS", "Búsqueda en JSTOR")):
        return "Búsqueda manual"
    return "Web y portales"


def flujo_evidencia(D):
    ft = pd.read_excel(REV, sheet_name="Flujo por tipo")
    ft = ft[ft.Tipo.astype(str).str.match(r"T\d\d")]
    fd = pd.read_excel(REV, sheet_name="Flujo 3ie DEP")
    fd = fd[fd.Estrategia.astype(str).str.match(r"DEP-\d+")]
    fm = pd.read_excel(REV, sheet_name="Flujo manual")
    fm = fm[fm.Estrategia.astype(str).str.match(r"MAN-\d+")]
    web = pd.read_excel(REV, sheet_name="Búsquedas web y portales")
    ref = pd.read_excel(REV, sheet_name="Referencias")
    ref["canal"] = ref.Origen.map(_canal)
    nuevas = ref.canal.value_counts()
    rev_col = "Registros revisados (primeros por relevancia)"
    canales = [
        ("Búsqueda académica", "PubMed, ERIC, EconBiz, Campbell, 3ie (Crossref), LA Referencia, Banco Mundial, OpenAlex",
         int(ft.Estrategias.sum()), int(ft["Registros recuperados"].sum()), int(ft[rev_col].sum()),
         int(ft["Seleccionados por título"].sum())),
        ("3ie DEP", "Portal de evidencia del 3ie (API pública)", len(fd), int(fd["Registros recuperados"].sum()),
         int(fd[rev_col].sum()), int(fd["Seleccionados por título"].sum())),
        ("Búsqueda manual", "SciELO, LILACS, Scopus y JSTOR", len(fm), int(fm["Registros recuperados"].sum()),
         int(fm[rev_col].sum()), int(fm["Seleccionados por título"].sum())),
        ("Web y portales", "Organismos, repositorios y evaluaciones peruanas", len(web), None, None, None),
    ]
    cols = []
    for nombre, bases, estr, recup, revis, selec in canales:
        cajas = [("Estrategias o búsquedas", estr), ("Registros recuperados", recup),
                 ("Registros revisados", revis), ("Seleccionados por título", selec),
                 ("Fuentes nuevas extraídas", int(nuevas.get(nombre, 0)))]
        cajas_html = "".join(f'<div class="cj{" fin" if i == 4 else ""}"><b>{fmt(v)}</b><span>{e(t)}</span></div>'
                             f'{"<div class=ar>&#9660;</div>" if i < 4 else ""}'
                             for i, (t, v) in enumerate(cajas) if v is not None)
        cols.append(f'<div class="cn"><div class="cnh">{e(nombre)}</div><div class="cnb">{e(bases)}</div>{cajas_html}</div>')
    inc = int((ref["Incluida (criterio de la nota)"] == "Sí").sum())
    cuerpo = (f'<div class="cns">{"".join(cols)}</div><div class="ar" style="font-size:20px">&#9660;</div>'
              f'<div class="tot"><div><b>{len(ref)}</b><span>fuentes extraídas</span></div><div class="mas">&#8594;</div>'
              f'<div class="ok"><b>{inc}</b><span>incluidas con el criterio de inclusión '
              f'(numeral 2.3)</span></div><div><b>{len(ref) - inc}</b><span>excluidas por ese criterio</span></div></div>'
              '<div class="nota">De cada estrategia se revisaron los primeros registros por relevancia. Cada fuente '
              'se cuenta una sola vez, en la vía por la que se incluyó primero.'
              '</div>')
    css = f"""
.cns{{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}}
.cn{{display:flex;flex-direction:column;align-items:stretch}}
.cnh{{background:{AZUL};color:#fff;font-weight:bold;font-size:13px;border-radius:6px 6px 0 0;padding:6px 10px}}
.cnb{{background:#f2f4f8;font-size:10.5px;color:{TINTA2};padding:5px 10px;border-radius:0 0 6px 6px;min-height:42px;
      margin-bottom:8px;line-height:1.25}}
.cj{{border:1px solid {GRILLA};border-radius:6px;padding:6px 10px;display:flex;gap:8px;align-items:baseline}}
.cj b{{font-size:17px;color:{TINTA}}}
.cj span{{font-size:11px;color:{TINTA2}}}
.cj.fin{{background:#eaf1fb;border-color:#86b6ef}}
.cj.fin b{{color:#1c5cab}}
.ar{{text-align:center;color:{TENUE};font-size:11px;line-height:1.2}}
.tot{{display:flex;gap:14px;justify-content:center;align-items:center;background:#f7f7f4;border-radius:10px;
      padding:12px}}
.tot>div{{display:flex;gap:8px;align-items:center;max-width:330px}}
.tot b{{font-size:30px;color:{AZUL}}}
.tot span{{font-size:11.5px;color:{TINTA2};line-height:1.25}}
.tot .ok b{{color:#0a7d0a}}
.mas{{font-size:22px;color:{TENUE}}}
"""
    return pagina(cuerpo, 1000, css), 1000


def fmt(v):
    return f"{v:,}".replace(",", " ")


# ---------------------------------------------------------------------------------------------------------------
# F11. Solidez y dirección de la evidencia por tipo de intervención
SOLIDEZ = {"Sólida": "#1c5cab", "Moderada": "#5598e7", "Limitada": "#b7d3f6"}


def solidez(D):
    mt = pd.read_excel(REV, sheet_name="Matriz por tipo")
    mt["n"] = mt["Fuentes incluidas"].fillna("").astype(str).map(lambda x: len([y for y in x.split(",") if y.strip()]))
    orden = {"Sólida": 0, "Moderada": 1, "Limitada": 2}
    mt = mt.sort_values(["Solidez", "n"], key=lambda s: s.map(orden) if s.name == "Solidez" else -s)
    mx = mt.n.max()
    filas = ""
    for _, r in mt.iterrows():
        sol = r.Solidez
        filas += (f'<div class="fr"><div class="tp"><b>{e(r.Tipo)}</b> {e(r["Tipo de intervención"])}</div>'
                  f'<div class="sv">{e(r.Servicios)}</div>'
                  f'<div><span class="pl" style="background:{SOLIDEZ[sol]};color:{"#fff" if sol != "Limitada" else TINTA}">'
                  f'{e(sol)}</span></div><div class="dr">{e(str(r["Dirección"]).capitalize())}</div>'
                  f'<div class="bb"><span style="width:{100 * r.n / mx}%"></span><b>{r.n}</b></div></div>')
    cuerpo = (f'<div class="fr hd"><div class="col-t">Tipo de intervención</div><div class="col-t">Servicios</div>'
              f'<div class="col-t">Solidez</div><div class="col-t">Dirección</div><div class="col-t">Fuentes '
              f'incluidas</div></div>{filas}' + leyenda(list(SOLIDEZ.items())) +
              '<div class="nota">Solidez con la regla del numeral 2.4 y del Anexo 4; '
              'dirección: sentido de los resultados de las fuentes incluidas. Ordenado por solidez y número de '
              'fuentes.</div>')
    css = f"""
.fr{{display:grid;grid-template-columns:1fr 150px 82px 82px 170px;gap:10px;align-items:center;padding:3px 0;
     border-bottom:1px solid #f0efec}}
.fr.hd{{border-bottom:1px solid {GRILLA}}}
.tp{{font-size:11.5px;line-height:1.2}}
.tp b{{color:{TENUE};font-weight:normal;margin-right:3px}}
.sv{{font-size:10.5px;color:{TINTA2}}}
.pl{{font-size:10.5px;font-weight:bold;border-radius:10px;padding:2px 8px}}
.dr{{font-size:11px;color:{TINTA2}}}
.bb{{display:flex;align-items:center;gap:5px;font-size:11px}}
.bb span{{display:block;height:9px;background:#9ec5f4;border-radius:0 3px 3px 0}}
"""
    return pagina(cuerpo, 1000, css), 1000


# ---------------------------------------------------------------------------------------------------------------
# barras apiladas al 100 % (fichas, recomendaciones)
def _apiladas(filas, categorias, etiqueta_ancho=300, total_txt=True):
    html_f = ""
    for nombre, valores in filas:
        tot = sum(valores.get(k, 0) for k in categorias)
        seg = ""
        for k, (col, tinta) in categorias.items():
            v = valores.get(k, 0)
            if v:
                seg += (f'<span style="width:{100 * v / tot}%;background:{col};color:{tinta}">'
                        f'{v if 100 * v / tot >= 4 else ""}</span>')
        html_f += (f'<div class="sf" style="grid-template-columns:{etiqueta_ancho}px 1fr 34px"><div class="sn">'
                   f'{nombre}</div><div class="sb">{seg}</div><div class="st">{tot if total_txt else ""}</div></div>')
    return html_f


CSS_APIL = f"""
.sf{{display:grid;gap:10px;align-items:center;margin-bottom:6px}}
.sn{{font-size:11.5px;line-height:1.2}}
.sb{{display:flex;gap:2px;height:22px}}
.sb span{{display:flex;align-items:center;justify-content:center;font-size:10.5px;font-weight:bold;border-radius:3px}}
.st{{font-size:11px;color:{TENUE}}}
"""


# F12. Calidad de las fichas vigentes
def fichas(D):
    r = pd.read_excel(A / "REV_IND_2_calidad_fichas.xlsx", sheet_name="Resumen").set_index("Atributo")
    cat = {"Cumple": (ESTADO["Implementado"], "#fff"), "Cumple parcialmente": (ESTADO["Implementado parcialmente"], TINTA),
           "No cumple": (ESTADO["No implementado"], "#fff"), "No consta": ("#c3c2b7", TINTA), "No aplica": ("#e8e7e1", TINTA)}
    filas = [(e(i), {k: int(r.loc[i, k]) for k in cat if k in r.columns}) for i in r.index]
    cuerpo = (_apiladas(filas, cat) + leyenda([(k, v[0]) for k, v in cat.items()]) +
              f'<div class="nota">{int(r.iloc[0][list(cat)].sum())} fichas de indicadores vigentes (13 de objetivo y '
              f'{int(r.iloc[0][list(cat)].sum()) - 13} de servicio) frente a los atributos de la guía de indicadores '
              'del CEPLAN (2024c), incluidas dos verificaciones de su Anexo A-9.</div>')
    return pagina(cuerpo, 1000, CSS_APIL), 1000


# F16. Avance de las recomendaciones de evaluaciones anteriores
def recomendaciones(D):
    res = D["C"]["rec_resumen"]
    cat = {"Implementada": (ESTADO["Implementado"], "#fff"),
           "Parcialmente implementada": (ESTADO["Implementado parcialmente"], TINTA),
           "No implementada": (ESTADO["No implementado"], "#fff"), "Sin evidencia": ("#c3c2b7", TINTA),
           "Sin objeto por la transición": ("#e8e7e1", TINTA)}
    nombres = {"MIMP 2020": "Evaluación de implementación 2020 (MIMP)", "MIMP 2021": "Evaluación de resultados 2021 "
               "(MIMP)", "MIMP 2023": "Evaluación 2023 (MIMP)", "CEPLAN": "Informe técnico del CEPLAN (servicios del "
               "OP1)", "Total": "<b>Total</b>"}
    filas = [(nombres[ev], {k: int(r[k]) for k in cat}) for ev, r in res.iterrows()]
    cuerpo = _apiladas(filas, cat, 290) + leyenda([(k, v[0]) for k, v in cat.items()])
    return pagina(cuerpo, 1000, CSS_APIL), 1000


# ---------------------------------------------------------------------------------------------------------------
# F13. Enfoques transversales
NIVELES_ENF = ["Ausente", "Mencionado", "Operacionalizado, mínimo", "Operacionalizado, parcial",
               "Operacionalizado, amplio"]
RAMPA_ENF = ["#e8e7e1", "#b7d3f6", "#86b6ef", "#3987e5", "#1c5cab"]


def enfoques(D):
    d6 = pd.read_excel(A / "D6_enfoques_transversales.xlsx", sheet_name="Lista de verificación", dtype=str).fillna("")
    filas = ""
    for _, r in d6.iterrows():
        nivel = r.nivel if r.nivel != "Operacionalizado" else f"Operacionalizado, {r.alcance.lower()}"
        k = NIVELES_ENF.index(nivel)
        celdas = "".join(f'<div class="c{" on" if i == k else ""}" style="{"background:" + RAMPA_ENF[i] if i == k else ""};'
                         f'{"color:#fff" if i == k and i >= 3 else ""}">{"&#9679;" if i == k else ""}</div>'
                         for i in range(len(NIVELES_ENF)))
        prot = f'{e(r.protocolos)} protocolos' if r.protocolos else ""
        filas += f'<div class="fr"><div class="en">{e(r.enfoque)}</div>{celdas}<div class="pr">{prot}</div></div>'
    cab = "".join(f'<div class="col-t" style="text-align:center">{e(n)}</div>' for n in NIVELES_ENF)
    cuerpo = (f'<div class="fr"><div></div>{cab}<div class="col-t">En los protocolos</div></div>{filas}'
              '<div class="nota">Nivel en el diseño de la PNIG (documento, matriz de consistencia y fichas): '
              'operacionalizado cuando el enfoque orienta lineamientos, estándares o indicadores; mencionado cuando '
              'solo se enuncia. Última columna: protocolos de servicio que lo incorporan (D6, «Lista de '
              'verificación»).</div>')
    css = f"""
.fr{{display:grid;grid-template-columns:200px repeat(5,1fr) 130px;gap:6px;align-items:center;margin-bottom:5px}}
.en{{font-size:12.5px}}
.c{{height:30px;border-radius:5px;background:#f7f7f4;display:flex;align-items:center;justify-content:center;
    font-size:14px}}
.pr{{font-size:11.5px;color:{TINTA2}}}
"""
    return pagina(cuerpo, 1000, css), 1000


# ---------------------------------------------------------------------------------------------------------------
# F15. Gasto de la conducción 2019–2025
# tonos más claros de los colores de cada órgano (ronda 34, pedido del usuario)
ORGANOS = {"DGIGND": "#7f99c8", "DPIGND": "#62b2aa", "DPPDM": "#cf9f86"}


def recursos(D):
    C = D["C"]
    dev, anios = C["cond_dev"], sorted(C["cond_dev"])
    mx = max(sum(dev[a][o] for o in ORGANOS) for a in anios)
    H = 260
    cols = ""
    for a in anios:
        tot = sum(dev[a][o] for o in ORGANOS)
        seg = "".join(f'<span style="height:{H * dev[a][o] / mx}px;background:{c}"></span>'
                      for o, c in reversed(list(ORGANOS.items())) if dev[a][o])
        cols += (f'<div class="cl"><div class="tv">{f"{tot / 1e6:.2f}".replace(".", ",")}</div><div class="pila">{seg}</div>'
                 f'<div class="an">{a}</div><div class="ej">{fmt(round(C["dgignd_ejec"][a]))} %</div></div>')
    cuerpo = (f'<div class="gr" style="height:{H + 70}px">{cols}</div>'
              + leyenda(list(ORGANOS.items())) +
              '<div class="nota">Devengado de las metas de los órganos conductores, en millones de soles (encima de '
              'cada barra). Bajo el año: ejecución de la meta de la DGIGND (devengado entre PIM). La DPIGND y la DPPDM '
              'tienen meta propia desde el ROF de 2023. Fuente de los montos: datos abiertos del MEF.</div>')
    css = f"""
.gr{{display:grid;grid-template-columns:repeat({len(anios)},1fr);gap:18px;align-items:end;border-bottom:1px solid {GRILLA};
     padding:0 20px}}
.cl{{display:flex;flex-direction:column;align-items:center;justify-content:end;height:100%}}
.pila{{display:flex;flex-direction:column;gap:2px;width:56px}}
.pila span{{display:block;border-radius:3px 3px 0 0}}
.pila span+span{{border-radius:0}}
.tv{{font-size:12px;font-weight:bold;margin-bottom:4px}}
.an{{font-size:12px;margin-top:6px}}
.ej{{font-size:10.5px;color:{TENUE}}}
"""
    return pagina(cuerpo, 900, css), 900


# ---------------------------------------------------------------------------------------------------------------
# F17. Indicadores de objetivo prioritario en 2025
NIVEL_COL = {"Alto": (ESTADO["Implementado"], "#fff"), "Medio": (ESTADO["Implementado parcialmente"], TINTA),
             "Bajo": (ESTADO["No implementado"], "#fff")}


def _nivel(txt):
    t = str(txt)
    for k in NIVEL_COL:
        if t.startswith(k):
            return k, t[len(k):].strip(" ()") or ""
    return None, "Sin valor" if t.startswith("Sin valor") else ("No interpretable" if t.startswith("No interp") else t)


def indicadores_2025(D):
    C = D["C"]
    tab, est = C["op_tabla"].set_index("Código"), C["op_est25"].set_index("Código")

    def pill(txt):
        k, extra = _nivel(txt)
        if k:
            col, tinta = NIVEL_COL[k]
            return f'<span class="pl" style="background:{col};color:{tinta}">{k}</span><small>{e(extra)}</small>'
        return f'<span class="pl sv">{e(extra or "Sin valor")}</span>'

    filas = ""
    for cod, r in tab.iterrows():
        eq = est.loc[cod] if cod in est.index else None
        oficial = _nivel(r["Semáforo 2025"])[0]
        if oficial or eq is None:
            celda = pill(r["Semáforo 2025"])
        else:   # sin valor oficial: el nivel calculado en esta evaluación, con asterisco y su fuente
            k, _ = _nivel(eq.Nivel)
            col, tinta = NIVEL_COL[k] if k else ("#e8e7e1", TINTA2)
            celda = (f'<span class="pl" style="background:{col};color:{tinta}">{e(k or eq.Nivel)}*</span>'
                     f'<small>{e(eq.Fuente)}</small>')
        filas += (f'<div class="fr"><span class="op" style="background:{OP[r.OP]}">{r.OP}</span><div class="nm"><b>{e(cod)}'
                  f'</b> {e(r.Indicador)}</div><div class="cl">{celda}</div></div>')
    cuerpo = (f'<div class="fr hd"><div></div><div class="col-t">Indicador</div><div class="col-t">Nivel 2025</div>'
              f'</div>{filas}' + leyenda([(k, v[0]) for k, v in NIVEL_COL.items()]) +
              '<div class="nota">Nivel de avance frente a la meta 2025 con la escala del numeral 2.4, según el reporte '
              'de seguimiento 2025. * Calculado en esta evaluación, porque el reporte no tuvo valor 2025: con los '
              'microdatos de la ENDES 2025 o la ENAHO 2025, o recalculado con el método de la ficha; no sustituye al '
              'reporte oficial.</div>')
    css = f"""
.fr{{display:grid;grid-template-columns:44px 1fr 300px;gap:10px;align-items:center;padding:4px 0;
     border-bottom:1px solid #f0efec}}
.op{{color:#fff;font-weight:bold;font-size:11px;border-radius:4px;text-align:center;padding:2px 0}}
.nm{{font-size:11px;line-height:1.22}}
.nm b{{color:{TENUE};font-weight:normal}}
.cl{{display:flex;gap:6px;align-items:center;flex-wrap:wrap}}
.pl{{font-size:10.5px;font-weight:bold;border-radius:10px;padding:2px 9px}}
.pl.sv{{background:#e8e7e1;color:{TINTA2}}}
.cl small{{font-size:10px;color:{TENUE}}}
"""
    return pagina(cuerpo, 1000, css), 1000


# ---------------------------------------------------------------------------------------------------------------
# F18. Matriz de evaluación: preguntas e instrumentos
INSTR = ["REV_DOC", "ENTREV", "REP_ENT", "REV_ALI", "REV_IND", "REV_EVI"]
# cabeceras con el nombre corto (el informe no usa siglas; ADR-009)
CABECERA = {"REV_DOC": "Revisión<br>documental", "ENTREV": "Entrevistas", "REP_ENT": "Reportes de<br>entidades",
            "REV_ALI": "Coherencia y<br>alineamiento", "REV_IND": "Indicadores<br>y recursos",
            "REV_EVI": "Revisión de<br>evidencia"}


def matriz_evaluacion(D):
    cob = pd.read_excel(A / "matriz_triangulacion_preguntas.xlsx", sheet_name="Cobertura", dtype=str).set_index("Pregunta")
    filas = ""
    for p, (crit, texto, previstos) in S.PREGUNTAS.items():
        usados = set(x.strip() for x in str(cob.loc[p, "Instrumentos con evidencia en la matriz"]).split(","))
        celdas = ""
        for i in INSTR:
            uso = i in usados
            celdas += f'<div class="c {"u" if uso else ""}">{"&#9679;" if uso else ""}</div>'
        filas += (f'<div class="fr"><div class="pq"><b>{p}</b> <span>{e(crit)}</span><div>{e(texto)}</div></div>'
                  f'{celdas}</div>')
    cab = "".join(f'<div class="col-t" style="text-align:center;font-size:9.5px;letter-spacing:0;line-height:1.2">'
                  f'{CABECERA[i]}</div>' for i in INSTR)
    cuerpo = (f'<div class="fr"><div class="col-t">Pregunta de evaluación y criterio</div>{cab}</div>{filas}'
              '<div class="ley"><span class="i"><span class="c u sm">&#9679;</span>Con evidencia</span><span '
              'class="i"><span class="c sm"></span>Sin evidencia</span></div>'
              '<div class="nota">Con evidencia: el instrumento sustenta al menos un hallazgo de la pregunta.</div>')
    css = f"""
.fr{{display:grid;grid-template-columns:1fr repeat(6,80px);gap:5px;align-items:center;margin-bottom:5px}}
.pq{{font-size:11px;line-height:1.25}}
.pq b{{color:{AZUL};font-size:12px}}
.pq span{{color:{TINTA2};font-size:10.5px;text-transform:uppercase;letter-spacing:.03em}}
.pq div{{color:{TINTA};margin-top:1px}}
.c{{height:36px;border-radius:5px;background:#f7f7f4;display:flex;align-items:center;justify-content:center;
    font-size:15px}}
.c.u{{background:#86b6ef;color:{AZUL}}}
.c.sm{{width:22px;height:18px;font-size:11px}}
"""
    return pagina(cuerpo, 1000, css), 1000


# ---------------------------------------------------------------------------------------------------------------
# Escalas del numeral 2.4 como diagramas (ronda 33): implementación de los servicios y nivel de avance
def escala_implementacion(D):
    """Árbol de decisión de la escala de implementación (2019–2025, ADR-011; dos de tres componentes, ADR-012), con el
    número de servicios por estado."""
    est = D["grd"]["Estado asignado (preliminar)"].value_counts()
    n = lambda k: int(est.get(k, 0))
    caja = lambda t, d: f'<div class="q"><div class="qt">{t}</div><div class="qd">{d}</div></div>'
    res = lambda k, cls: (f'<div class="r {cls}"><b>{e(k)}</b><span>{n(k)} servicio{"s" * (n(k) != 1)}</span></div>')
    cuerpo = (
        '<div class="g">'
        + caja("1. ¿Hay ejecución reportada en algún año de 2019 a 2025?",
               "Información detallada del servicio en algún informe LIO (una sección que describe sus acciones sin "
               "declarar falta de información) o valor de alguno de sus indicadores. Las acciones preparatorias "
               "(diagnóstico, protocolo en elaboración) no cuentan como ejecución.")
        + '<div class="fl">No &#10140;</div>' + res("No implementado", "no")
        + '<div class="fv">Sí &#8595;</div><div></div><div></div>'
        + caja("2. ¿Reúne al menos dos de los tres componentes (protocolo aprobado, recursos identificados en el MEF, "
               "indicadores con valor reportado) y seguía funcionando en 2024 o 2025?",
               "Los componentes se buscan en todo el periodo; un recurso está identificado si un monto que la "
               "entidad reportó coincide con una meta, actividad o producto suyo en los datos del MEF. «Seguía "
               "funcionando» significa ejecución reportada en 2024 o 2025.")
        + '<div class="fl">Sí &#10140;</div>' + res("Implementado", "si")
        + '<div class="fv">No &#8595;</div><div></div><div></div>'
        + '<div class="q pq"><div class="qd">Ejecución en el periodo con dos o más componentes faltantes, o '
          'interrumpida antes de 2024.</div></div><div class="fl">&#10140;</div>' + res("Implementado parcialmente", "pa")
        + '</div><div class="nota">Sin información suficiente: solo secciones mínimas del informe LIO, sin valores de '
          'indicador (ningún servicio quedó en esta categoría). Se considera ejecutado el servicio cuya provisión a la '
          'población objetivo reporta la entidad.</div>')
    css = f"""
.g{{display:grid;grid-template-columns:1fr 66px 200px;gap:6px 10px;align-items:center}}
.q{{background:#f2f4f8;border-radius:10px;padding:12px 14px}}
.q.pq{{background:#f7f7f4}}
.qt{{font-size:17px;font-weight:bold;color:{AZUL};line-height:1.3}}
.qd{{font-size:14px;color:{TINTA2};line-height:1.35;margin-top:4px}}
.fl{{font-size:15px;color:{TENUE};text-align:center;font-weight:bold}}
.fv{{font-size:14px;color:{TENUE};font-weight:bold;padding-left:18px}}
.r{{border-radius:10px;padding:12px 14px;display:flex;flex-direction:column;gap:2px}}
.r b{{font-size:17px}}
.r span{{font-size:14px}}
.r.si{{background:{ESTADO["Implementado"]};color:#fff}}
.r.pa{{background:{ESTADO["Implementado parcialmente"]};color:{TINTA}}}
.r.no{{background:{ESTADO["No implementado"]};color:#fff}}
"""
    return pagina(cuerpo, 760, css), 760


def escala_avance(D):
    """Regla del nivel de avance de los indicadores (valor/meta × 100), con sus cinco tramos y los casos aparte."""
    tope = 180
    tramos = [(0, 15, "Bajo con posible falla de planeación", ESTADO["No implementado"], True),
              (15, 75, "Bajo", ESTADO["No implementado"], False),
              (75, 95, "Medio", ESTADO["Implementado parcialmente"], False),
              (95, 150, "Alto", ESTADO["Implementado"], False),
              (150, tope, "Alto con posible falla de planeación", ESTADO["Implementado"], True)]
    seg, ley = "", []
    for a, b, nombre, col, falla in tramos:
        ancho = 100 * (b - a) / tope
        fondo = (f"repeating-linear-gradient(45deg,{col} 0 6px,{mezcla(col, 0.55)} 6px 12px)" if falla else col)
        tinta = TINTA if col == ESTADO["Implementado parcialmente"] else "#fff"
        seg += (f'<div class="sg" style="width:{ancho:.2f}%;background:{fondo};color:{tinta}">'
                f'{"" if falla else e(nombre)}</div>')
        if falla:
            rango = "menos de 15" if a == 0 else "150 o más"
            ley.append(f'<span class="i"><span class="sw" style="background:{fondo}"></span>{e(nombre)} ({rango})</span>')
    marcas = "".join(f'<span class="mk" style="left:{100 * v / tope:.2f}%">{v}</span>' for v in (0, 15, 75, 95, 150))
    cuerpo = (f'<div class="fx">Avance = valor obtenido / meta × 100 (sentido ascendente) o meta / valor obtenido × 100 '
              f'(sentido descendente)</div><div class="br">{seg}</div>'
              f'<div class="ej">{marcas}</div><div class="ley">{"".join(ley)}</div>'
              '<div class="ap"><b>Casos aparte</b> (no se clasifican y constituyen un hallazgo de calidad técnica): sin '
              'meta, sin valor, sin sentido definido o no interpretable (meta o valor de 0 o menos).</div>'
              '<div class="nota">La escala no considera la línea de base. Cuando el valor no proviene de la fuente que '
              'fija la ficha, el nivel lleva la advertencia «fuente no comparable».</div>')
    css = f"""
.fx{{font-size:15px;color:{TINTA};background:#f2f4f8;border-radius:8px;padding:9px 12px;margin-bottom:12px}}
.br{{display:flex;height:42px;border-radius:6px;overflow:hidden;margin-top:4px}}
.sg{{display:flex;align-items:center;justify-content:center;font-size:15px;font-weight:bold}}
.ej{{position:relative;height:22px;margin-top:3px;font-size:14px;color:{TINTA2}}}
.mk{{position:absolute;transform:translateX(-50%)}}
.ap{{font-size:14px;color:{TINTA2};background:#f7f7f4;border-radius:8px;padding:9px 12px;margin-top:10px;
     line-height:1.35}}
.ap b{{color:{TINTA}}}
"""
    return pagina(cuerpo, 760, css), 760


# ---------------------------------------------------------------------------------------------------------------
# F19. Valoración por pregunta
VALOR = {"Satisfactorio": (ESTADO["Implementado"], "#fff"), "Parcial": (ESTADO["Implementado parcialmente"], TINTA),
         "Insuficiente": (ESTADO["No implementado"], "#fff")}
# confianza de los hallazgos en verde, ámbar y rojo (ronda 34, decisión del usuario); son los mismos de las marcas de
# las tablas del informe (Anexo 2 y conclusiones)
CONF = {"Sólido": ESTADO["Implementado"], "Con respaldo parcial": ESTADO["Implementado parcialmente"],
        "Indicio": ESTADO["No implementado"]}


def valoracion(D):
    C = D["C"]
    V, H = S.valoraciones(C), S.hallazgos(C)
    bloques = ""
    for titulo, pref in (("Diseño", "D"), ("Implementación", "I")):
        filas = ""
        for p, (crit, texto, _) in S.PREGUNTAS.items():
            if not p.startswith(pref):
                continue
            valor = V[p][0]
            # la valoración va sin color (ronda 34): el color queda para la confianza de los hallazgos
            hs = [h for h in H if h["pregunta"] == p]
            puntos = "".join(f'<span class="dt" style="background:{CONF[h["confianza"]]}" title="{h["id"]}"></span>'
                             for h in hs)
            filas += (f'<div class="fr"><div class="pq"><b>{p}</b> {e(crit)}</div><div><span class="pl">'
                      f'{e(valor)}</span></div><div class="dts">{puntos}'
                      f'<small>{len(hs)} hallazgos</small></div></div>')
        bloques += f'<div class="bl"><div class="bh">{titulo}</div>{filas}</div>'
    cuerpo = (f'<div class="bls">{bloques}</div>' +
              leyenda([(f"Hallazgo {k.lower()}", v) for k, v in CONF.items()]) +
              '<div class="nota">Valoración por pregunta con la escala del numeral 2.4. Cada punto es un hallazgo, '
              'coloreado por su nivel de confianza; el nivel también se lee en el Anexo 2.</div>')
    css = f"""
.bls{{display:grid;grid-template-columns:1fr 1fr;gap:18px}}
.bh{{background:{AZUL};color:#fff;font-weight:bold;font-size:13px;border-radius:6px;padding:6px 10px;margin-bottom:6px}}
.fr{{display:grid;grid-template-columns:1fr 100px 150px;gap:8px;align-items:center;padding:5px 0;
     border-bottom:1px solid #f0efec}}
.pq{{font-size:11.5px;line-height:1.2}}
.pq b{{color:{AZUL}}}
.pl{{font-size:11px;font-weight:bold;border-radius:10px;padding:3px 10px;background:#fff;border:1px solid #c3c2b7;color:{TINTA}}}
.dts{{display:flex;gap:3px;align-items:center;flex-wrap:nowrap}}
.dt{{width:12px;height:12px;border-radius:50%;display:inline-block}}
.dts small{{font-size:10px;color:{TENUE};margin-left:4px}}
"""
    return pagina(cuerpo, 1000, css), 1000


# clave: (función, título, fuente, rastro), en el orden del informe. La fuente va al pie de la figura en el informe,
# con nombres y no siglas; el rastro (libro y hoja) queda en el índice y en el registro de rastro (ADR-009)
FIGURAS = {
    "valoracion": (valoracion, "Valoración de las preguntas de evaluación y confianza de sus hallazgos",
                   "Fuente: matriz de triangulación del equipo evaluador.",
                   "matriz_triangulacion_preguntas.xlsx, hoja «Valoración por pregunta»"),
    "linea_tiempo": (linea_tiempo, "Hitos normativos y de implementación de la PNIG, 2007–2025",
                     "Fuente: normas y documentos del corpus de la evaluación.", "docs/fuentes/registro_fuentes.md"),
    "problema_publico": (problema, "Modelo del problema público de la PNIG",
                         "Fuente: documento de la PNIG (D.S. N.° 008-2019-MIMP) y matriz de consistencia (Anexo 5). "
                         "Elaboración: equipo evaluador (análisis de coherencia interna).",
                         "REV_ALI_coherencia_alineamiento.xlsx, hoja «REV_ALI.1 Modelo causal»"),
    "cadena_valor": (cadena, "Cadena de valor de la PNIG: del problema público a los servicios y sus indicadores",
                     "Fuente: matriz de consistencia de la PNIG (Anexo 5), situación futura deseada (numeral 1.4.5) "
                     "e indicadores vigentes (consolidado de la OMEP). Elaboración: equipo evaluador (análisis de "
                     "coherencia interna y base de indicadores).",
                     "REV_ALI_coherencia_alineamiento.xlsx, hojas «REV_ALI.1 Por servicio» y «REV_ALI.1 Por objetivo»; "
                     "REV_IND_base_indicadores.xlsx; D1_D2_brechas_situacion_futura.xlsx, hoja «Situación futura "
                     "deseada»"),
    "indicadores": (indicadores, "Indicadores de la PNIG: de objetivo prioritario, de servicio y del PEMIG",
                    "Fuente: indicadores de los OP y de servicio vigentes (consolidado de la OMEP) y anexo 5 del "
                    "PEMIG. Elaboración: equipo evaluador (base de indicadores y revisión de la calidad de las "
                    "fichas).",
                    "REV_IND_base_indicadores.xlsx; REV_IND_2_calidad_fichas.xlsx, hoja «Tipo PEMIG»"),
    "actores": (actores, "Actores de la PNIG: rectoría, articulación y entidades responsables de los servicios",
                "Fuente: entidades responsables de cada servicio (revisión documental), D.S. N.° 018-2021-MIMP y "
                "entrevistas al nivel conductor. Elaboración: equipo evaluador.",
                "REV_DOC_estado_implementacion_servicios.xlsx, hoja «Estado por servicio»; ENTREV_codificacion.csv"),
    "alin_vertical": (alin_vertical, "Alineamiento vertical de los objetivos prioritarios en el periodo y en la "
                      "transición", "Fuente: Políticas de Estado, Plan Bicentenario, PEDN, PGG y PESEM. "
                      "Elaboración: equipo evaluador (análisis del alineamiento vertical).",
                      "REV_ALI_coherencia_alineamiento.xlsx, hoja «REV_ALI.2 Por objetivo»"),
    "alin_internacional": (alin_internacional, "Alineamiento de los objetivos prioritarios con la Ley N.° 32535 y los "
                           "compromisos internacionales", "Fuente: Ley N.° 32535, CEDAW, Convención de Belém do Pará, Plataforma de Acción de Beijing, "
                           "Agenda 2030 (ODS 5) y Estrategia de Montevideo. Elaboración: equipo evaluador (análisis de "
                           "la correspondencia normativa e internacional).",
                           "REV_ALI_coherencia_alineamiento.xlsx, hoja «REV_ALI.4 Por objetivo»"),
    "alin_horizontal": (alin_horizontal, "Relaciones de los servicios de la PNIG con las políticas nacionales afines",
                        "Fuente: políticas nacionales afines (base de datos de políticas nacionales del CEPLAN al 15 "
                        "de octubre de 2026). Elaboración: equipo evaluador (análisis del alineamiento horizontal).",
                        "REV_ALI_coherencia_alineamiento.xlsx, hoja «REV_ALI.3 Horizontal»"),
    "conceptual_planos": (conceptual_planos, "Relaciones entre la PNIG y el nuevo marco, por plano de análisis",
                          "Correspondencia conceptual por plano (REV_ALI.5).",
                          "REV_ALI_coherencia_alineamiento.xlsx, hoja «Resumen REV_ALI.5»"),
    "conceptual_conceptos": (conceptual_conceptos, "Conceptos de la PNIG frente a las categorías del nuevo marco",
                             "Correspondencia conceptual, plano conceptual (REV_ALI.5).",
                             "REV_ALI_coherencia_alineamiento.xlsx, hoja «REV_ALI.5 Conceptual»"),
    "flujo_evidencia": (flujo_evidencia, "Búsqueda y selección de la evidencia sobre la plausibilidad de los servicios",
                        "Fuente: búsquedas en bases académicas, en el portal de evidencia del 3ie, en SciELO, LILACS, "
                        "Scopus y JSTOR y en la web. Elaboración: equipo evaluador (revisión de evidencia).",
                        "REV_EVI_matriz_evidencia.xlsx, hojas «Flujo por tipo», «Flujo 3ie DEP», «Flujo manual», "
                        "«Búsquedas web y portales» y «Referencias»"),
    "solidez": (solidez, "Solidez y dirección de la evidencia por tipo de intervención",
                "Fuente: estudios incluidos en la revisión de evidencia. Elaboración: equipo evaluador.",
                "REV_EVI_matriz_evidencia.xlsx, hojas «Matriz por tipo» y «Criterios»"),
    "fichas": (fichas, "Cumplimiento de los atributos de calidad en las fichas de los indicadores vigentes",
               "Fuente: fichas técnicas de los indicadores vigentes (consolidado de la OMEP). Elaboración: equipo "
               "evaluador (revisión de la calidad de las fichas).",
               "REV_IND_2_calidad_fichas.xlsx, hoja «Resumen»"),
    "enfoques": (enfoques, "Incorporación de los enfoques transversales en el diseño de la PNIG",
                 "Fuente: documento de la PNIG y sus anexos, PEMIG (anexo 5) y protocolos aprobados. Elaboración: "
                 "equipo evaluador (lista de verificación de enfoques de la revisión documental).",
                 "D6_enfoques_transversales.xlsx, hoja «Lista de verificación»"),
    "mapa_servicios_op1": (lambda D: mapa_servicios(D, ("OP1",)), "Servicios del OP1 por lineamiento, con su estado de implementación",
                            "Estado de implementación por servicio (revisión documental).",
                            "REV_DOC_estado_implementacion_servicios.xlsx, hoja «Estado por servicio»; "
                            "REV_ALI_coherencia_alineamiento.xlsx, hoja «REV_ALI.1 Por servicio»"),
    "mapa_servicios_op2": (lambda D: mapa_servicios(D, ("OP2",)), "Servicios del OP2 por lineamiento, con su estado de implementación",
                            "Estado de implementación por servicio (revisión documental).",
                            "REV_DOC_estado_implementacion_servicios.xlsx, hoja «Estado por servicio»; "
                            "REV_ALI_coherencia_alineamiento.xlsx, hoja «REV_ALI.1 Por servicio»"),
    "mapa_servicios_op3": (lambda D: mapa_servicios(D, ("OP3",)), "Servicios del OP3 por lineamiento, con su estado de implementación",
                            "Estado de implementación por servicio (revisión documental).",
                            "REV_DOC_estado_implementacion_servicios.xlsx, hoja «Estado por servicio»; "
                            "REV_ALI_coherencia_alineamiento.xlsx, hoja «REV_ALI.1 Por servicio»"),
    "mapa_servicios_op4": (lambda D: mapa_servicios(D, ("OP4",)), "Servicios del OP4 por lineamiento, con su estado de implementación",
                            "Estado de implementación por servicio (revisión documental).",
                            "REV_DOC_estado_implementacion_servicios.xlsx, hoja «Estado por servicio»; "
                            "REV_ALI_coherencia_alineamiento.xlsx, hoja «REV_ALI.1 Por servicio»"),
    "mapa_servicios_op5": (lambda D: mapa_servicios(D, ("OP5",)), "Servicios del OP5 por lineamiento, con su estado de implementación",
                            "Estado de implementación por servicio (revisión documental).",
                            "REV_DOC_estado_implementacion_servicios.xlsx, hoja «Estado por servicio»; "
                            "REV_ALI_coherencia_alineamiento.xlsx, hoja «REV_ALI.1 Por servicio»"),
    "mapa_servicios_op6": (lambda D: mapa_servicios(D, ("OP6",)), "Servicios del OP6 por lineamiento, con su estado de implementación",
                            "Estado de implementación por servicio (revisión documental).",
                            "REV_DOC_estado_implementacion_servicios.xlsx, hoja «Estado por servicio»; "
                            "REV_ALI_coherencia_alineamiento.xlsx, hoja «REV_ALI.1 Por servicio»"),
    "recursos": (recursos, "Gasto devengado de la conducción de la PNIG por órgano, 2019–2025",
                 "Fuente: datos abiertos del gasto del MEF, pliego 039. Elaboración: equipo evaluador (análisis de "
                 "los recursos de la conducción).",
                 "REV_IND_3_verificacion_presupuestal.xlsx, hoja «Conducción por órgano»"),
    "conaig_sesiones": (conaig_sesiones, "Sesiones y acuerdos de la CONAIG, 2021–2026",
                        "Ayuda memoria de la CONAIG.",
                        "docs/fuentes/pnig/09_conaig_gobernanza/2026_mimp-conaig_ayuda-memoria-sesiones-2021-2026.doc"),
    "recomendaciones": (recomendaciones, "Avance de las recomendaciones de las evaluaciones anteriores",
                        "Fuente: informes de las evaluaciones anteriores y documentos al 7 de noviembre de 2026. "
                        "Elaboración: equipo evaluador (seguimiento de las recomendaciones).",
                        "seguimiento_recomendaciones_previas.xlsx, hoja «Resumen por evaluación»"),
    "indicadores_2025": (indicadores_2025, "Nivel de avance de los indicadores de objetivo prioritario en 2025",
                         "Fuente: reporte de seguimiento 2025 y estimaciones del equipo evaluador con la ENDES y la "
                         "ENAHO 2025 y registros electorales. Elaboración: equipo evaluador (base de indicadores y "
                         "recálculo de los indicadores del OP3).",
                         "REV_IND_base_indicadores.xlsx, hoja «Base REV_IND.1»; REV_IND_4_recalculo_OP3.xlsx, hoja "
                         "«Resultado 2025»; D1_D2_brechas_situacion_futura.xlsx, hoja «Manifestaciones ENDES»"),
    "escala_implementacion": (escala_implementacion, "Escala de implementación de los servicios",
                              "Árbol de decisión de la escala de implementación (ADR-011).",
                              "REV_DOC_estado_implementacion_servicios.xlsx, hoja «Criterios»"),
    "escala_avance": (escala_avance, "Nivel de avance de los indicadores frente a su meta",
                      "Regla del nivel de avance de los reportes de cumplimiento de la PNIG.",
                      "REV_IND_base_indicadores.xlsx"),
    "matriz_evaluacion": (matriz_evaluacion, "Preguntas de evaluación e instrumentos que las responden",
                          "Fuente: matriz de triangulación (instrumentos con evidencia por pregunta).",
                          "matriz_triangulacion_preguntas.xlsx, hoja «Cobertura»"),
}

# fuente del pie de cada figura en el informe publicable: documento o base pública con su cita (Referencias del
# informe) y «Elaboración propia» (ADR-010); la tercera columna de FIGURAS queda como descripción técnica
FUENTE_PUBLICA = {
    "valoracion": "Fuente: hallazgos y valoraciones de esta evaluación (Anexos 1 y 2). Elaboración propia.",
    "linea_tiempo": "Fuente: Ley N.° 28983, D.S. N.° 008-2019-MIMP, D.S. N.° 002-2020-MIMP, R.M. N.° 116-2021-MIMP, D.S. "
                    "N.° 018-2021-MIMP y Ley N.° 32535. Elaboración propia.",
    "problema_publico": "Fuente: MIMP (2019), modelo del problema público y matriz de consistencia (Anexo 5 de la PNIG). "
                        "Elaboración propia.",
    "cadena_valor": "Fuente: MIMP (2019), numeral 1.4.5 y matriz de consistencia (Anexo 5 de la PNIG); OGMEPGD (2026). "
                    "Elaboración propia.",
    "indicadores": "Fuente: MIMP (2019); OGMEPGD (2026); anexo 5 del PEMIG (D.S. N.° 002-2020-MIMP). Elaboración propia.",
    "actores": "Fuente: MIMP (2019); D.S. N.° 018-2021-MIMP; protocolos de los servicios (Anexo 10); entrevistas al nivel "
               "conductor. Elaboración propia.",
    "alin_vertical": "Fuente: Políticas de Estado, Plan Bicentenario, PEDN, PGG y PESEM (Anexo 10). Elaboración propia.",
    "alin_internacional": "Fuente: Ley N.° 32535; CEDAW, Convención de Belém do Pará, Plataforma de Acción de Beijing, "
                          "Agenda 2030 y Estrategia de Montevideo (Anexo 10). Elaboración propia.",
    "alin_horizontal": "Fuente: políticas nacionales afines (Anexo 10), identificadas en la base de datos de políticas "
                       "nacionales del CEPLAN al 15 de octubre de 2026. Elaboración propia.",
    "conceptual_planos": "Fuente: MIMP (2019); Ley N.° 32535; MIMP, DGIGND (2026b). Elaboración propia.",
    "conceptual_conceptos": "Fuente: MIMP (2019); Ley N.° 32535; MIMP, DGIGND (2026b). Elaboración propia.",
    "flujo_evidencia": "Fuente: búsquedas en bases académicas, en el portal de evidencia del 3ie, en SciELO, LILACS, "
                       "Scopus y JSTOR y en la web (numeral 2.3). Elaboración propia.",
    "solidez": "Fuente: estudios incluidos en la revisión de evidencia (Anexo 4 y Referencias). Elaboración propia.",
    "fichas": "Fuente: OGMEPGD (2026), fichas técnicas de los indicadores vigentes; CEPLAN (2024c). Elaboración propia.",
    "enfoques": "Fuente: MIMP (2019) y sus anexos; anexo 5 del PEMIG (D.S. N.° 002-2020-MIMP); protocolos de los "
                "servicios (Anexo 10). Elaboración propia.",
    "mapa_servicios_op1": "Fuente: protocolos de los servicios (Anexo 10); MEF (2026); OGMEPGD (2026); MIMP (2020-2026). Elaboración propia.",
    "mapa_servicios_op2": "Fuente: protocolos de los servicios (Anexo 10); MEF (2026); OGMEPGD (2026); MIMP (2020-2026). Elaboración propia.",
    "mapa_servicios_op3": "Fuente: protocolos de los servicios (Anexo 10); MEF (2026); OGMEPGD (2026); MIMP (2020-2026). Elaboración propia.",
    "mapa_servicios_op4": "Fuente: protocolos de los servicios (Anexo 10); MEF (2026); OGMEPGD (2026); MIMP (2020-2026). Elaboración propia.",
    "mapa_servicios_op5": "Fuente: protocolos de los servicios (Anexo 10); MEF (2026); OGMEPGD (2026); MIMP (2020-2026). Elaboración propia.",
    "mapa_servicios_op6": "Fuente: protocolos de los servicios (Anexo 10); MEF (2026); OGMEPGD (2026); MIMP (2020-2026). Elaboración propia.",
    "recursos": "Fuente: MEF (2026), pliego 039, Ministerio de la Mujer y Poblaciones Vulnerables. Elaboración propia.",
    "conaig_sesiones": "Fuente: MIMP y CONAIG (2026), ayuda memoria de las sesiones de la CONAIG. Elaboración propia.",
    "recomendaciones": "Fuente: MIMP (2021, 2022, 2024); CEPLAN (s. f.); documentos al 7 de noviembre de 2026 (Anexo 7). "
                       "Elaboración propia.",
    "indicadores_2025": "Fuente: reporte de seguimiento 2025 de la PNIG (Anexo 5); INEI (2026), microdatos de la ENDES y "
                        "la ENAHO 2025; JNE (2022a); ONPE (2026a). Elaboración propia.",
    "matriz_evaluacion": "Fuente: hallazgos de esta evaluación (Anexo 2). Elaboración propia.",
    "escala_implementacion": "Fuente: elaboración propia, con los resultados de la revisión documental (Anexo 3).",
    "escala_avance": "Fuente: MIMP (2026a), escala de los reportes de cumplimiento de la PNIG. Elaboración propia.",
}
assert set(FUENTE_PUBLICA) == set(FIGURAS), "cada figura necesita su fuente pública"


def generar(claves=None):
    SALIDA.mkdir(parents=True, exist_ok=True)
    D = datos()
    perfil = tempfile.mkdtemp(prefix="chrome-figuras-")
    filas = []
    try:
        for clave, (fn, titulo, fuente, rastro) in FIGURAS.items():
            if claves and clave not in claves:
                continue
            doc, ancho = fn(D)
            ruta_html, ruta_png = SALIDA / f"{clave}.html", SALIDA / f"{clave}.png"
            ruta_html.write_text(doc, encoding="utf-8")
            alto = renderizar(ruta_html, ruta_png, ancho, perfil)
            filas.append({"clave": clave, "archivo": ruta_png.as_posix(), "titulo": titulo,
                          "fuente": FUENTE_PUBLICA[clave], "descripcion_tecnica": fuente, "rastro": rastro,
                          "ancho_px": ancho, "alto_px": alto})
            print(f"OK {clave} ({ancho}×{alto})", flush=True)
    finally:
        shutil.rmtree(perfil, ignore_errors=True)
    indice = SALIDA / "indice_figuras_informe.csv"
    previo = pd.read_csv(indice) if indice.exists() and claves else pd.DataFrame(columns=list(filas[0]))
    nuevo = pd.concat([previo[~previo.clave.isin([f["clave"] for f in filas])], pd.DataFrame(filas)])
    orden = {c: i for i, c in enumerate(FIGURAS)}
    nuevo.sort_values("clave", key=lambda s: s.map(orden)).to_csv(indice, index=False, encoding="utf-8")


if __name__ == "__main__":
    generar(sys.argv[1:] or None)
