"""Revisión de evidencia para el informe de cierre de la PNIG: marco por tipo de intervención (numeral 3.4) y detalle
de los estudios (Anexo 4) (ADR-010).

Todo sale de la matriz de evidencia (04_analisis/REV_EVI_matriz_evidencia.xlsx): la tipología (intervención,
población, resultado esperado, causa o efecto), los estudios incluidos con su diseño, ámbito, calidad, resultado
principal y dirección, y la clasificación por tipo. El resultado de cada estudio es el resumen que registró el equipo
evaluador (en la mayoría, tomado del resumen del artículo, como consta en la hoja «Referencias»).

motivo() reproduce la regla de solidez de rev_evi_evidencia.solidez() y comprueba que da la misma clasificación que la
hoja «Matriz por tipo»; así el Anexo 4 explica por qué cada tipo tiene su nivel (por ejemplo, por qué un tipo con
muchas revisiones puede quedar en «limitada»: sus revisiones son de contextos poco comparables).

Lo importa informe_cierre.py.
"""
import re

import pandas as pd

from docx_mimp import lista_es
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

REV = f"{EVAL}/04_analisis/REV_EVI_matriz_evidencia.xlsx"
PRIMARIOS = ("Experimental", "Cuasi experimental")
DISENO = {"Revisión sistemática": "revisión sistemática", "Experimental": "evaluación experimental",
          "Cuasi experimental": "evaluación cuasiexperimental", "Revisión narrativa": "revisión narrativa",
          "Descriptivo": "estudio descriptivo"}
DISENO_PL = {"Revisión sistemática": ("revisión sistemática", "revisiones sistemáticas"),
             "Experimental": ("evaluación experimental", "evaluaciones experimentales"),
             "Cuasi experimental": ("evaluación cuasiexperimental", "evaluaciones cuasiexperimentales")}


def datos():
    tip = pd.read_excel(REV, sheet_name="Tipología").set_index("tipo")
    mt = pd.read_excel(REV, sheet_name="Matriz por tipo").set_index("Tipo")
    me = pd.read_excel(REV, sheet_name="Matriz de evidencia")
    ref = pd.read_excel(REV, sheet_name="Referencias").set_index("Id")
    me["Id"] = me.Fuente.str.extract(r"^(E\d+):")[0]
    me["comparable"] = ~me["Ámbito"].astype(str).str.startswith(("Ingreso alto", "No informado"))
    me["incluida"] = me["Criterio de inclusión"] == "Incluida"
    return tip, mt, me, ref


def cita(fuente):
    """«E01: Sviatschi, M. M. y Trako, I. (2024)» -> «Sviatschi y Trako (2024)»; tres o más autores, «et al.» (APA 7)."""
    m = re.match(r"^(?:E\d+:\s*)?(.+?)\s*\((\d{4}[a-z]?|s\. f\.)\)", str(fuente))
    if not m:
        return re.sub(r"^E\d+:\s*", "", str(fuente))
    autores, anio = m.groups()
    # apellidos con inicial final («Labrador R.») incluidos
    apellidos = re.findall(r"([A-ZÁÉÍÓÚÑ][\w\-'’ .]*?), (?:[A-ZÁÉÍÓÚÑ]\.[\s\-]?)+", autores) or [autores.strip()]
    if "et al" in autores or len(apellidos) >= 3:
        return f"{apellidos[0]} et al. ({anio})"
    if len(apellidos) == 2:
        return f"{apellidos[0]} y {apellidos[1]} ({anio})"
    return f"{apellidos[0]} ({anio})"


def _minuscula(t):
    t = str(t).strip()
    if len(t) > 1 and t[0].isupper() and not t[1].isupper() and not t.split(" ")[0].isupper():
        return t[0].lower() + t[1:]
    return t


def _ambito(a):
    a = str(a)
    if a.startswith("No informado"):
        return "ámbito no informado en el resumen"
    if a == "Global":
        return "alcance global"
    return a[0].lower() + a[1:] if a.startswith(("Ingreso", "Varios")) else a


def motivo(g, solidez_registrada):
    """Por qué el tipo tiene su nivel de solidez: conteo de revisiones favorables y de evaluaciones primarias en
    contextos comparables (regla del numeral 2.4). Comprueba que el nivel coincide con el de la matriz."""
    g = g[g.incluida]
    comp = g[g.comparable]
    rs = comp[comp["Diseño del estudio"] == "Revisión sistemática"]
    ev = comp[comp["Diseño del estudio"].isin(PRIMARIOS)]
    rs_fav = rs[(rs["Dirección"] == "Favorable") & ~rs["Calidad según la fuente"].astype(str).str.match(r"(Muy b|B)aja")]
    no_comp = len(g) - len(comp)
    if len(rs_fav) or (len(ev) >= 2 and (ev["Dirección"] == "Favorable").all()):
        nivel = "Sólida"
        razon = (f"{len(rs_fav)} {'revisión sistemática favorable' if len(rs_fav) == 1 else 'revisiones sistemáticas favorables'}"
                 f" en contextos comparables, sin calidad baja" if len(rs_fav) else
                 f"{len(ev)} evaluaciones con diseño experimental o cuasiexperimental en contextos comparables, todas "
                 f"favorables")
    elif len(ev) or len(rs):
        nivel = "Moderada"
        partes = []
        if len(rs):
            partes.append(f"{len(rs)} {'revisión sistemática' if len(rs) == 1 else 'revisiones sistemáticas'} en "
                          f"contextos comparables, con resultados mixtos o de calidad baja")
        if len(ev):
            partes.append(f"{len(ev)} {'evaluación primaria' if len(ev) == 1 else 'evaluaciones primarias'} en "
                          f"contextos comparables" + (", no todas favorables" if len(ev) >= 2 else ""))
        razon = lista_es(partes)
    elif len(g):
        nivel = "Limitada"
        razon = ("solo estudios descriptivos, revisiones narrativas o estudios de contextos poco comparables"
                 + (f" ({no_comp} de {len(g)} fuentes son de países de ingreso alto o no informan su ámbito)"
                    if no_comp else ""))
    else:
        nivel, razon = "Sin evidencia identificada", "no se identificaron estudios"
    assert nivel == solidez_registrada, f"la regla da {nivel} y la matriz {solidez_registrada}"
    return f"{nivel}: {razon}."


def resumen_tipo(g):
    """Composición de la evidencia incluida de un tipo: revisiones, evaluaciones, contextos y dirección."""
    g = g[g.incluida]
    n = len(g)
    dis = g["Diseño del estudio"].value_counts()
    partes = [f"{dis[k]} {DISENO_PL[k][0] if dis[k] == 1 else DISENO_PL[k][1]}" for k in DISENO_PL if dis.get(k, 0)]
    otros = n - sum(dis.get(k, 0) for k in DISENO_PL)
    if otros:
        partes.append(f"{otros} {'estudio descriptivo o revisión narrativa' if otros == 1 else 'estudios descriptivos o revisiones narrativas'}")
    d = g["Dirección"].value_counts()
    peru = int(g["Ámbito"].astype(str).str.contains("Perú").sum())
    no_comp = int((~g.comparable).sum())
    fav, mix, nul = d.get('Favorable', 0), d.get('Mixto', 0), d.get('Nulo', 0) + d.get('Desfavorable', 0)
    sin_dir = n - fav - mix - nul   # estudios descriptivos o revisiones sin resultado de efecto (observación 35)
    texto = (f"La revisión incluyó {n} {'fuente' if n == 1 else 'fuentes'} sobre este tipo ({lista_es(partes)}): "
             f"{fav} con {'resultado favorable' if fav == 1 else 'resultados favorables'}, {mix} "
             f"{'mixta' if mix == 1 else 'mixtas'}, {nul} {'nula' if nul == 1 else 'nulas'}"
             + (f" y {sin_dir} sin resultado de efecto" if sin_dir else "") + ".")
    if no_comp:
        texto += (f" De ellas, {no_comp} {'proviene' if no_comp == 1 else 'provienen'} de países de ingreso alto o "
                  f"no {'informa' if no_comp == 1 else 'informan'} su ámbito en el resumen, por lo que su "
                  f"aplicabilidad al Perú es menor.")
    texto += ((" Una de ellas es del Perú." if peru == 1 else f" {peru} son del Perú.") if peru else
              " Ninguna es del Perú, de modo que su aplicabilidad depende de que el contexto sea comparable.")
    return texto


def estudios_clave(g, maximo=4):
    """Estudios que más pesan en la clasificación: primero las revisiones sistemáticas, luego las evaluaciones
    experimentales y cuasiexperimentales; las del Perú se citan siempre (hasta dos)."""
    g = g[g.incluida].drop_duplicates("Id").copy()
    g["k"] = g["Diseño del estudio"].map({"Revisión sistemática": 0, "Experimental": 1, "Cuasi experimental": 2})
    g = g.assign(k=g["k"].fillna(3)).sort_values("k", kind="stable")
    peru = g[g["Ámbito"].astype(str).str.contains("Perú")].head(2)
    otros = g[~g.Id.isin(peru.Id)].head(max(2, maximo - len(peru)))
    return pd.concat([otros, peru])


def frase_estudio(r):
    return (f"{cita(r.Fuente)}, {DISENO.get(r['Diseño del estudio'], r['Diseño del estudio'].lower())} "
            f"({_ambito(r['Ámbito'])}): {_minuscula(r['Resultado principal']).rstrip('.')}.")


def _cita_parentetica(fuente):
    """«Rivas et al. (2015)» -> «Rivas et al., 2015» (cita entre paréntesis, APA 7)."""
    return re.sub(r"\s*\((\d{4}[a-z]?|s\. f\.)\)$", r", \1", cita(fuente))


def _frase_ambito(a, peru_en_conector=False):
    a0 = _ambito(a)
    if a0 == "alcance global":
        return "de alcance global"
    if a0.startswith("ámbito no informado"):
        return "cuyo resumen no informa el ámbito"
    if a0 == "Perú":
        return "" if peru_en_conector else "en el Perú"
    if a0.startswith("ingreso bajo y medio"):
        return f"en contextos de {a0}"
    if a0.startswith("ingreso"):
        return f"en un contexto de {a0}"
    return f"en {a0}"


def relato_evidencia(clave):
    """Relato de la evidencia de un tipo (ronda 33, pedido del usuario: lectura continua, sin etiquetas): empieza por
    el estudio de mayor peso y enlaza los demás según la dirección de su resultado; los del Perú cierran el relato."""
    frases, peru_previos, vistos = [], 0, {}
    for k, (_, r) in enumerate(clave.iterrows()):
        diseno = r["Diseño del estudio"]
        nombre = DISENO.get(diseno, str(diseno).lower())
        art = "un" if nombre.startswith("estudio") else "una"
        peru = str(r["Ámbito"]).startswith("Perú")
        amb = _frase_ambito(r["Ámbito"], peru_en_conector=peru)
        autor = cita(r.Fuente)
        plural = " et al." in autor or " y " in autor
        # cita narrativa al inicio («Rivas et al. (2015), una revisión sistemática de alcance global»)
        np_ = f"{autor}, {art} {nombre}{' ' + amb if amb else ''}"
        res = _minuscula(r["Resultado principal"]).rstrip(".").replace(": ", "; ", 1)
        direccion = str(r["Dirección"])
        if k == 0:
            inicio = ("La evidencia más sólida proviene de" if diseno in ("Revisión sistemática", "Experimental",
                                                                          "Cuasi experimental")
                      else "La evidencia disponible proviene de")
            frases.append(f"{inicio} {np_}: {res}.")
        elif peru:
            frases.append(f"{'En el Perú' if not peru_previos else 'También en el Perú'}, {np_}, "
                          f"{'encontraron' if plural else 'encontró'} lo siguiente: {res}.")
        elif direccion == "Favorable":
            inicio = ("En la misma línea apunta", "Coincide con ello", "También lo respalda")[vistos.get("fav", 0) % 3]
            vistos["fav"] = vistos.get("fav", 0) + 1
            frases.append(f"{inicio} {np_}: {res}.")
        elif direccion == "Mixto":
            inicio = ("Con resultados mixtos", "También con resultados mixtos")[min(vistos.get("mix", 0), 1)]
            vistos["mix"] = vistos.get("mix", 0) + 1
            frases.append(f"{inicio}, {np_}, {'informan' if plural else 'informa'} lo siguiente: {res}.")
        elif direccion in ("Nulo", "Desfavorable"):
            frases.append(f"En cambio, {np_}, {'matizan' if plural else 'matiza'} esa conclusión: {res}.")
        else:
            frases.append(f"Además, {np_}, {'aportan' if plural else 'aporta'} contexto: {res}.")
        peru_previos += peru
    return " ".join(frases)


# precisiones de comparabilidad entre estudios de un mismo tipo (revisión institucional del 13-11-2026, observación 34)
NOTA_COMPARABILIDAD = {
    "T01": "Los dos estudios peruanos no son directamente comparables: Sviatschi y Trako (2024) miden la violencia "
           "doméstica autorreportada, las muertes y las hospitalizaciones, y Calderon-Anyosa et al. (2025) miden el "
           "reporte de violencia de pareja en la ENDES de 2004 a 2016; un aumento del reporte puede reflejar mayor "
           "disposición a declarar o a denunciar, y no necesariamente más violencia.",
}


def parrafos_tipo(t, tip, mt, me, nombre_serv, entidad_serv):
    """Párrafos del numeral 3.4 para un tipo: qué es y qué servicios lo implementaron, mecanismo, qué dice la
    literatura, consistencia y aplicabilidad, y conclusión para los servicios."""
    x = tip.loc[t]
    m = mt.loc[t]
    g = me[me.Tipo == t]
    servicios = [s.strip() for s in str(x.servicios).split(";")]
    nodo = x.nodo_modelo.split(": ", 1)
    nodo_txt = (f"la causa «{nodo[1]}»" if nodo[0] == "Causa" else f"el efecto «{nodo[1]}»")
    p1 = (f"La PNIG implementó este tipo con {'el servicio' if len(servicios) == 1 else 'los servicios'} "
          + lista_es([f"{s} ({nombre_serv[s]})" for s in servicios]) +
          f". La intervención consiste en {_minuscula(x.intervencion)}, dirigida a {_minuscula(x.poblacion)}. Su "
          f"mecanismo esperado es que produzca {_minuscula(x.resultado_esperado)} y actúe así sobre {nodo_txt} del "
          f"modelo del problema público.")
    clave = estudios_clave(g)
    p2 = relato_evidencia(clave)
    if t in NOTA_COMPARABILIDAD:
        p2 += " " + NOTA_COMPARABILIDAD[t]
    resto = g[g.incluida].Id.nunique() - len(clave)
    if resto > 0:
        p2 += (f" Las otras {resto} fuentes que respaldan este tipo están en el Anexo 4." if resto > 1 else
               " La otra fuente que respalda este tipo está en el Anexo 4.")
    sol = str(m.Solidez).lower()
    pert = "directa" if m["Pertinencia del resultado"] == "Directa" else "parcial (la evidencia mide otro resultado o " \
                                                                            "cubre solo una parte de la población)"
    p3 = (resumen_tipo(g) + f" Con ese respaldo, la lógica causal de {'este servicio' if len(servicios) == 1 else 'estos servicios'} "
          f"tiene evidencia {sol}, de dirección {str(m['Dirección']).lower()} y de pertinencia {pert}.")
    return p1, p2, p3
