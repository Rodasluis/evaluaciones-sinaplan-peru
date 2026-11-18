"""Información reportada por las entidades (REP_ENT): lo que las entidades responsables informan sobre sus servicios,
codificado con el libro común, y, si la evaluación lo usa, el formulario que se les envía por oficio.

Entradas (02_insumos/):
    rep_ent_segmentos.csv              id_segmento, entidad, servicio, fuente, pagina, texto: fragmentos de los
                                       reportes de las entidades (informes de seguimiento, memorias, respuestas)
    rep_ent_segmentos_formulario.csv   (opcional) los mismos campos, escritos por consolidar() desde las respuestas
    rep_ent_codificacion.csv           id_segmento, codigos, sentido; opcional critico (Sí o No)
    rep_ent_segunda_codificacion.csv   (opcional) id_segmento, codigos: la del codificador-ciego
    libro_codigos.csv                  el común
Formulario (modalidad.formulario_entidades): formulario() arma un Excel por entidad responsable en
00_encargo/formularios/; las respuestas se dejan en 03_datos/raw/formularios/ (no se versiona) y consolidar() las lleva a
02_insumos/. El formulario no pide nombres ni datos de contacto de personas, solo el cargo del órgano que responde.
Salida: 04_analisis/REP_ENT_informacion_entidades.xlsx.
"""
import pandas as pd

from ..fuentes import slug
from . import base, codificados, rev_doc, rev_ind

NOMBRE = "Información reportada por las entidades"
LIBRO = "REP_ENT_informacion_entidades.xlsx"
PUBLICO = "Informacion_de_las_entidades.xlsx"
SEGMENTOS = ["id_segmento", "entidad", "servicio", "fuente", "pagina", "texto"]
CAMPOS_ABIERTOS = {"dificultades": "Dificultades para implementar el servicio",
                   "factores_favorables": "Factores que favorecieron la implementación",
                   "propuestas": "Propuestas de ajuste para la nueva política"}


def entradas(cfg):
    return [("insumos", "libro_codigos.csv", codificados.LIBRO_CODIGOS + ["incluye", "excluye"]),
            ("insumos", "rep_ent_segmentos.csv", SEGMENTOS),
            ("insumos", "rep_ent_codificacion.csv", codificados.CODIFICACION + ["critico"])]


def leer_segmentos(cfg, p, servicios):
    partes = []
    for nombre, opcional in (("rep_ent_segmentos.csv", False), ("rep_ent_segmentos_formulario.csv", True)):
        seg = base.leer(cfg, nombre, SEGMENTOS, p, opcional=opcional)
        if seg is not None:
            base.obligatorio(seg, "texto", nombre, p)
            base.obligatorio(seg, "fuente", nombre, p)
            if servicios is not None:
                base.existentes(seg[seg.servicio != ""], "servicio", servicios.codigo, nombre, p)
            partes.append(seg)
    if not partes:
        return None
    seg = pd.concat(partes, ignore_index=True)
    base.unicos(seg, "id_segmento", "segmentos de las entidades", p)
    return seg


def validar(cfg):
    p = base.Problemas("REP_ENT")
    s = rev_doc.leer_servicios(cfg, p)
    lib = codificados.leer_libro(cfg, p)
    seg = leer_segmentos(cfg, p, s)
    cod = codificados.leer_codificacion(cfg, "insumos", "rep_ent_codificacion.csv", seg, lib, p)
    if cod is not None:
        if "critico" not in cod.columns:
            cod["critico"] = ""
        base.en_escala(cod, "critico", base.SI_NO, "rep_ent_codificacion.csv", p, vacio=True)
    seg2 = codificados.leer_codificacion(cfg, "insumos", "rep_ent_segunda_codificacion.csv", seg, lib, p, segunda=True)
    p.revisar()
    return s, lib, seg, cod, seg2


def correr(cfg):
    s, lib, seg, cod, seg2 = validar(cfg)
    larga, frec = codificados.frecuencias(seg, cod, lib, "entidad")
    frec = frec.rename(columns={"Grupos": "Entidades que lo reportan"})
    nombres = dict(zip(lib.codigo, lib.nombre))
    por_servicio = []
    for c, g in larga[larga.servicio != ""].groupby("servicio"):
        top = g.codigo.value_counts().head(3)
        por_servicio.append({"Código": c, "Servicio": dict(zip(s.codigo, s.servicio)).get(c, ""),
                             "Segmentos": g.id_segmento.nunique(),
                             "Códigos más frecuentes": "; ".join(f"{nombres[k]} ({v})" for k, v in top.items()),
                             "Hallazgos críticos": int(g[g.critico == "Sí"].id_segmento.nunique()),
                             "Fuentes": "; ".join(dict.fromkeys(f"{f} p. {pg}" if pg else f
                                                               for f, pg in zip(g.fuente, g.pagina)))})
    registro = seg.merge(cod, on="id_segmento", how="left")
    registro["Nombres de los códigos"] = registro.codigos.map(lambda x: "; ".join(nombres.get(c, c)
                                                                                 for c in base.lista(x or "")))
    hojas = {"Registro codificado": registro, "Frecuencias por código": frec,
             "Por servicio": pd.DataFrame(por_servicio)}
    salida = {"segmentos": len(seg), "entidades": int(seg.entidad.nunique())}
    if seg2 is not None:
        resumen, detalle = codificados.concordancia(cod, seg2, lib)
        hojas |= {"Concordancia": resumen, "Contraste de codificaciones": detalle}
        salida["acuerdo_exacto"] = resumen.iloc[0]["Acuerdo exacto (%)"]
    crit = [("Objeto", "Lo que las entidades responsables reportan sobre la ejecución de sus servicios, sus dificultades "
                       "y sus propuestas, con su fuente y página.", "CEPLAN (Guía de seguimiento y evaluación: "
                                                                   "implementación)")] + codificados.CRITERIOS
    destino = base.ruta(cfg, "analisis", LIBRO)
    base.escribir_libro(destino, hojas, crit)
    return {"libro": destino, **salida}


# ------------------------------------------------------------------ formulario a las entidades
def _entidades(s):
    return sorted({e for x in s.responsable for e in base.lista(x)})


def formulario(cfg):
    """Un Excel por entidad responsable, con sus servicios e indicadores, para enviarlo por oficio."""
    p = base.Problemas("REP_ENT")
    s = rev_doc.leer_servicios(cfg, p)
    ind = rev_ind.leer_indicadores(cfg, p)
    if s is not None and (s.responsable == "").all():
        p("servicios.csv: el formulario necesita la columna responsable")
    p.revisar()
    periodo = base.anios(cfg)
    salidas = []
    for e in _entidades(s):
        suyos = s[s.responsable.map(lambda x: e in base.lista(x))]
        serv = pd.DataFrame({"codigo": suyos.codigo, "servicio": suyos.servicio,
                             **{f"ejecucion_{a}": "" for a in periodo}, "protocolo_aprobado": "", "norma_protocolo": "",
                             **{c: "" for c in CAMPOS_ABIERTOS}})
        inds = ind[(ind.nivel == "Servicio") & ind.servicio.isin(suyos.codigo)]
        filas_i = [{"codigo": i.codigo, "indicador": i.nombre, "anio": a, "meta": "", "valor": "", "fuente": "",
                    "observacion": ""} for i in inds.itertuples() for a in periodo]
        instrucciones = pd.DataFrame({"Instrucción": [
            f"Formulario de la {cfg['titulo']}.",
            "Hoja Servicios: para cada año, escriba Sí si el servicio se ejecutó, No si no se ejecutó o Sin información.",
            "Indique si el servicio tiene protocolo aprobado (Sí o No) y la norma que lo aprueba.",
            "En las columnas " + ", ".join(CAMPOS_ABIERTOS) + " describa " + "; ".join(
                v.lower() for v in CAMPOS_ABIERTOS.values()) + ".",
            "Hoja Indicadores: meta y valor de cada año, con la fuente del dato.",
            "No consigne nombres ni datos de contacto de personas; indique solo el cargo del órgano que responde."]})
        destino = cfg.abs("encargo", "formularios", f"Formulario_{slug(e, 40)}.xlsx")
        destino.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(destino, engine="openpyxl") as w:
            instrucciones.to_excel(w, sheet_name="Instrucciones", index=False)
            serv.to_excel(w, sheet_name="Servicios", index=False)
            pd.DataFrame(filas_i, columns=["codigo", "indicador", "anio", "meta", "valor", "fuente",
                                           "observacion"]).to_excel(w, sheet_name="Indicadores", index=False)
            pd.DataFrame({"entidad": [e]}).to_excel(w, sheet_name="Entidad", index=False)
        salidas.append(destino)
    return salidas


def consolidar(cfg):
    """Lleva las respuestas de 03_datos/raw/formularios/*.xlsx a rep_ent_segmentos_formulario.csv y
    rev_ind_valores_formulario.csv (02_insumos/). Los valores de indicadores se contrastan con rev_ind_valores.csv; no
    lo reemplazan."""
    carpeta = cfg.abs("datos_raw", "formularios")
    segmentos, valores = [], []
    for archivo in sorted(carpeta.glob("*.xlsx")):
        entidad = pd.read_excel(archivo, sheet_name="Entidad", dtype=str, engine="openpyxl").entidad.iloc[0]
        serv = pd.read_excel(archivo, sheet_name="Servicios", dtype=str, engine="openpyxl").fillna("")
        fuente = f"Formulario de {entidad}"
        for r in serv.itertuples():
            for campo in CAMPOS_ABIERTOS:
                texto = str(getattr(r, campo)).strip()
                if texto:
                    segmentos.append({"id_segmento": f"F-{slug(entidad, 20)}-{r.codigo}-{campo}", "entidad": entidad,
                                      "servicio": r.codigo, "fuente": fuente, "pagina": "", "texto": texto})
        ind = pd.read_excel(archivo, sheet_name="Indicadores", dtype=str, engine="openpyxl").fillna("")
        for r in ind.itertuples():
            if r.meta or r.valor:
                valores.append({"codigo": r.codigo, "anio": r.anio, "meta": r.meta, "valor": r.valor,
                                "fuente": r.fuente or fuente, "entidad": entidad, "observacion": r.observacion})
    pd.DataFrame(segmentos, columns=SEGMENTOS).to_csv(base.ruta(cfg, "insumos", "rep_ent_segmentos_formulario.csv"),
                                                      index=False, encoding="utf-8", lineterminator="\n")
    pd.DataFrame(valores, columns=["codigo", "anio", "meta", "valor", "fuente", "entidad", "observacion"]).to_csv(
        base.ruta(cfg, "insumos", "rev_ind_valores_formulario.csv"), index=False, encoding="utf-8", lineterminator="\n")
    return len(segmentos), len(valores)
