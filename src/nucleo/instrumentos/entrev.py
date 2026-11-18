"""Entrevistas (ENTREV): guías por rol del informante, registro codificado y concordancia de la doble codificación.

Solo se usa si la evaluación hace entrevistas (config.yaml, modalidad.entrevistas).

Entradas:
    05_entrevistas/entrev_preguntas.csv       id, rol, bloque, pregunta_evaluacion, texto; opcionales repreguntas y
                                              contrasta (hallazgo preliminar o evidencia que la pregunta contrasta)
    02_insumos/libro_codigos.csv              codigo, nombre, definicion; opcionales incluye y excluye
    05_entrevistas/entrev_segmentos.csv       id_segmento, rol, pregunta_id, texto (de las transcripciones
                                              anonimizadas, sin nombres)
    05_entrevistas/entrev_codificacion.csv    id_segmento, codigos, sentido (Facilitador, Obstáculo o Neutro);
                                              opcionales hallazgo y contraste
    05_entrevistas/entrev_segunda_codificacion.csv   (opcional) id_segmento, codigos: la del codificador-ciego
Salidas: 05_entrevistas/ENTREV_guia_<rol>.docx (guias()) y 04_analisis/ENTREV_registro_codificado.xlsx (correr()).
Las grabaciones, las notas en bruto y la relación entre el rol y la persona van fuera del repositorio (CLAUDE.md,
regla 6).
"""
from ..fuentes import slug
from . import base, codificados

NOMBRE = "Entrevistas"
LIBRO = "ENTREV_registro_codificado.xlsx"
PUBLICO = "Entrevistas_registro_codificado.xlsx"
PREGUNTAS = ["id", "rol", "bloque", "pregunta_evaluacion", "texto"]
SEGMENTOS = ["id_segmento", "rol", "pregunta_id", "texto"]
CONSENTIMIENTO = ("Antes de empezar se explica a la persona entrevistada el objetivo de la evaluación y que su "
                  "participación es voluntaria. Lo que diga se registra por su cargo, no por su nombre. La grabación, "
                  "si la autoriza, es de uso interno del equipo evaluador y se guarda fuera del repositorio, conforme "
                  "a la Ley N.° 29733.")


def entradas(cfg):
    return [("insumos", "libro_codigos.csv", codificados.LIBRO_CODIGOS + ["incluye", "excluye"]),
            ("entrevistas", "entrev_preguntas.csv", PREGUNTAS + ["repreguntas", "contrasta"]),
            ("entrevistas", "entrev_segmentos.csv", SEGMENTOS),
            ("entrevistas", "entrev_codificacion.csv", codificados.CODIFICACION + ["hallazgo", "contraste"])]


def leer_preguntas(cfg, p):
    pre = base.leer(cfg, "entrev_preguntas.csv", PREGUNTAS, p, carpeta="entrevistas")
    if pre is not None:
        n = "entrev_preguntas.csv"
        base.unicos(pre, "id", n, p)
        for c in PREGUNTAS:
            base.obligatorio(pre, c, n, p)
        base.existentes(pre, "pregunta_evaluacion", cfg["preguntas"], n, p)
        for i, k in enumerate(pre.pregunta_evaluacion, start=2):
            q = cfg["preguntas"].get(k)
            if q and "ENTREV" not in q["instrumentos"]:
                p(f"{n}, fila {i}: la pregunta {k} no tiene las entrevistas entre sus instrumentos en config.yaml")
        for c in ("repreguntas", "contrasta"):
            if c not in pre.columns:
                pre[c] = ""
    return pre


def guias(cfg):
    """Una guía en Word por rol de informante, con el consentimiento, los bloques y lo que cada pregunta contrasta."""
    from ..documento import Documento
    p = base.Problemas("ENTREV")
    pre = leer_preguntas(cfg, p)
    p.revisar()
    salidas = []
    periodo = cfg["periodo"]
    for rol, g in pre.groupby("rol", sort=False):
        d = Documento(pie=f"Guía de entrevista · {cfg['politica']['sigla']} {periodo['inicio']}–{periodo['fin']}",
                      titulo_doc=f"Guía de entrevista: {rol}")
        d.titulo("Guía de entrevista semiestructurada", cfg["titulo"], f"Informante: {rol}")
        d.h1("Presentación y consentimiento")
        d.p(CONSENTIMIENTO)
        for bloque, gb in g.groupby("bloque", sort=False):
            d.h1(bloque)
            for r in gb.itertuples():
                d.p((f"{r.id}. ", {"b": True}), r.texto, (f" (pregunta {r.pregunta_evaluacion})", {"i": True}))
                for rp in base.lista(r.repreguntas):
                    d.vineta(rp)
                if r.contrasta:
                    d.p(("Contrasta: ", {"i": True}), r.contrasta)
        d.h1("Cierre")
        d.p("Se agradece la participación y se pregunta si hay documentos que la persona quiera ofrecer para la "
            "evaluación; si los ofrece, se registran como fuentes.")
        destino = base.ruta(cfg, "entrevistas", f"ENTREV_guia_{slug(rol, 40)}.docx")
        destino.parent.mkdir(parents=True, exist_ok=True)
        d.guardar(destino)
        salidas.append(destino)
    return salidas


def validar(cfg):
    p = base.Problemas("ENTREV")
    pre = leer_preguntas(cfg, p)
    lib = codificados.leer_libro(cfg, p)
    seg = base.leer(cfg, "entrev_segmentos.csv", SEGMENTOS, p, carpeta="entrevistas")
    if seg is not None:
        base.unicos(seg, "id_segmento", "entrev_segmentos.csv", p)
        if pre is not None:
            base.existentes(seg[seg.pregunta_id != ""], "pregunta_id", pre.id, "entrev_segmentos.csv", p)
    cod = codificados.leer_codificacion(cfg, "entrevistas", "entrev_codificacion.csv", seg, lib, p)
    seg2 = codificados.leer_codificacion(cfg, "entrevistas", "entrev_segunda_codificacion.csv", seg, lib, p,
                                         segunda=True)
    p.revisar()
    return pre, lib, seg, cod, seg2


def correr(cfg):
    pre, lib, seg, cod, seg2 = validar(cfg)
    larga, frec = codificados.frecuencias(seg, cod, lib, "rol")
    frec = frec.rename(columns={"Grupos": "Roles que lo mencionan"})
    larga = larga.merge(pre[["id", "pregunta_evaluacion"]], left_on="pregunta_id", right_on="id", how="left")
    por_pregunta = (larga.groupby(["pregunta_evaluacion", "codigo"]).agg(Segmentos=("id_segmento", "nunique"),
                                                                           Roles=("rol", "nunique")).reset_index()
                    .rename(columns={"pregunta_evaluacion": "Pregunta de evaluación", "codigo": "Código"}))
    registro = seg.merge(cod, on="id_segmento", how="left")
    nombres = dict(zip(lib.codigo, lib.nombre))
    registro["Nombres de los códigos"] = registro.codigos.map(lambda x: "; ".join(nombres.get(c, c)
                                                                                 for c in base.lista(x or "")))
    hojas = {"Registro codificado": registro, "Frecuencias por código": frec, "Por pregunta": por_pregunta,
             "Preguntas formuladas": pre}
    salida = {"segmentos": len(seg), "roles": int(seg.rol.nunique())}
    if seg2 is not None:
        resumen, detalle = codificados.concordancia(cod, seg2, lib)
        hojas |= {"Concordancia": resumen, "Contraste de codificaciones": detalle}
        salida["acuerdo_exacto"] = resumen.iloc[0]["Acuerdo exacto (%)"]
    destino = base.ruta(cfg, "analisis", LIBRO)
    base.escribir_libro(destino, hojas, [("Objeto", "Percepción de los órganos que conducen la política sobre su "
                                                     "diseño e implementación, contrastada con los hallazgos "
                                                     "documentales.", "Equipo evaluador"),
                                         ("Consentimiento", CONSENTIMIENTO, "Ley N.° 29733")] + codificados.CRITERIOS)
    return {"libro": destino, **salida}


def para_codificador_ciego(cfg):
    """Ruta de los segmentos sin códigos que se dan al agente codificador-ciego (nunca la primera codificación)."""
    return base.ruta(cfg, "entrevistas", "entrev_segmentos.csv")

