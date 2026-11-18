"""Coherencia y alineamiento (REV_ALI): coherencia interna de la cadena de resultados y alineamiento de la política con
el planeamiento, otras políticas, los compromisos internacionales y su marco conceptual.

Entradas (02_insumos/):
    servicios.csv, indicadores.csv  las de la revisión documental y de indicadores
    cadena.csv          codigo (servicio), causa_efecto (causa o efecto del modelo del problema que atiende),
                        tipo_vinculo (Causa, Efecto o vacío); opcionales estandar y fuente
    lineamientos.csv    (opcional) lineamiento, op: para detectar lineamientos sin servicio
    alineamiento.csv    tipo (Vertical, Horizontal, Internacional, Normativa o Conceptual), op, instrumento, referencia,
                        documento (clave de documentos.csv), cita (literal), correspondencia, sustento; opcionales
                        momento y pagina
    documentos.csv      clave, titulo, ruta (relativa a la raíz del repositorio)
Coherencia por servicio: completa si tiene indicador vigente (y estándar, si se registra) y está vinculado a una causa o
efecto del modelo; con vacíos en otro caso. Un lineamiento sin servicio está roto.
Citas: cada cita se ubica en su documento (citas.ubicar). Si el documento está en el repositorio y la cita no aparece,
la corrida se detiene; si el documento no está (por ejemplo, en la versión pública), la cita queda «sin verificar».
Salida: 04_analisis/REV_ALI_coherencia_alineamiento.xlsx.
"""
import pandas as pd

from . import base, citas, rev_doc, rev_evi, rev_ind

NOMBRE = "Coherencia y alineamiento"
LIBRO = "REV_ALI_coherencia_alineamiento.xlsx"
PUBLICO = "Coherencia_y_alineamiento.xlsx"
CORRESPONDENCIA = ("Directa", "Indirecta", "Sin correspondencia")
ESCALAS = {"Vertical": CORRESPONDENCIA, "Internacional": CORRESPONDENCIA, "Normativa": CORRESPONDENCIA,
           "Horizontal": ("Complementariedad", "Duplicidad", "Vacío de atención"),
           "Conceptual": ("Continuidad", "Precisión", "Divergencia")}
CADENA = ["codigo", "causa_efecto", "tipo_vinculo"]
ALINEAMIENTO = ["tipo", "op", "instrumento", "referencia", "documento", "cita", "correspondencia", "sustento"]
DOCUMENTOS = ["clave", "titulo", "ruta"]


def entradas(cfg):
    return [("insumos", "cadena.csv", CADENA + ["estandar", "fuente"]),
            ("insumos", "lineamientos.csv", ["lineamiento", "op"]),
            ("insumos", "alineamiento.csv", ["tipo", "momento", "op", "instrumento", "referencia", "documento", "pagina",
                                             "cita", "correspondencia", "sustento"]),
            ("insumos", "documentos.csv", DOCUMENTOS + ["enlace"])]


def ubicar_citas(cfg, ali, docs, p):
    """Página o párrafo donde aparece cada cita, y su verificación."""
    rutas = dict(zip(docs.clave, docs.ruta)) if docs is not None else {}
    encontradas, estado = [], []
    for i, r in enumerate(ali.itertuples(), start=2):
        if not r.cita:
            encontradas.append("")
            estado.append("Sin cita")
            continue
        archivo = cfg.repo / rutas.get(r.documento, "")
        if not rutas.get(r.documento) or not archivo.is_file():
            encontradas.append("")
            estado.append("Sin verificar: documento no disponible")
            continue
        paginas = citas.ubicar(r.cita, archivo)
        encontradas.append(", ".join(map(str, paginas)))
        if not paginas:
            estado.append("No ubicada")
            p(f"alineamiento.csv, fila {i}: la cita no aparece en {rutas[r.documento]}: «{r.cita[:80]}»")
        elif getattr(r, "pagina", "") and r.pagina not in [str(x) for x in paginas]:
            estado.append(f"Ubicada en otra página (declarada {r.pagina})")
        else:
            estado.append("Ubicada")
    return encontradas, estado


def validar(cfg, ubicar=True):
    p = base.Problemas("REV_ALI")
    s = rev_doc.leer_servicios(cfg, p)
    ind = rev_ind.leer_indicadores(cfg, p)
    cad = base.leer(cfg, "cadena.csv", CADENA, p)
    lin = base.leer(cfg, "lineamientos.csv", ["lineamiento", "op"], p, opcional=True)
    ali = base.leer(cfg, "alineamiento.csv", ALINEAMIENTO, p)
    docs = base.leer(cfg, "documentos.csv", DOCUMENTOS, p, opcional=True)
    if cad is not None and s is not None:
        base.unicos(cad, "codigo", "cadena.csv", p)
        base.existentes(cad, "codigo", s.codigo, "cadena.csv", p)
        faltan = sorted(set(s.codigo) - set(cad.codigo))
        if faltan:
            p(f"cadena.csv: servicios sin registro: {', '.join(faltan)}")
        base.en_escala(cad, "tipo_vinculo", ("Causa", "Efecto"), "cadena.csv", p, vacio=True)
    if ali is not None:
        n = "alineamiento.csv"
        base.en_escala(ali, "tipo", list(ESCALAS), n, p)
        for i, r in enumerate(ali.itertuples(), start=2):
            if r.tipo in ESCALAS and r.correspondencia not in ESCALAS[r.tipo]:
                p(f"{n}, fila {i}: correspondencia «{r.correspondencia}» no admitida para {r.tipo} "
                  f"({', '.join(ESCALAS[r.tipo])})")
            if not r.sustento:
                p(f"{n}, fila {i}: falta sustento")
            if r.tipo != "Conceptual" and not r.op:
                p(f"{n}, fila {i}: falta op")
            if r.correspondencia not in ("Sin correspondencia", "Vacío de atención") and not r.cita:
                p(f"{n}, fila {i}: una correspondencia «{r.correspondencia}» requiere la cita literal")
        if s is not None:
            base.existentes(ali[ali.op != ""], "op", s.op, n, p)
        if docs is not None:
            base.existentes(ali[ali.documento != ""], "documento", docs.clave, n, p)
        elif (ali.documento != "").any():
            p(f"{n}: cita documentos, pero falta documentos.csv")
        for c in ("momento", "pagina"):
            if c not in ali.columns:
                ali[c] = ""
        if ubicar and not p:
            ali["Ubicación en el documento"], ali["Verificación de la cita"] = ubicar_citas(cfg, ali, docs, p)
    p.revisar()
    return s, ind, cad, lin, ali


def libro_previo(cfg, modulo, hoja):
    archivo = base.ruta(cfg, "analisis", modulo.LIBRO)
    return (pd.read_excel(archivo, sheet_name=hoja, dtype=str, engine="openpyxl").fillna("") if archivo.exists()
            else None)


def coherencia(s, ind, cad, lin, estado=None, plaus=None):
    vig = ind[(ind.nivel == "Servicio") & (ind.vigente == "Sí")]
    por_servicio = vig.groupby("servicio").codigo.apply(", ".join).to_dict()
    cad = cad.set_index("codigo")
    con_estandar = "estandar" in cad.columns
    filas = []
    for sv in s.itertuples():
        c = cad.loc[sv.codigo]
        vacios = []
        if not por_servicio.get(sv.codigo):
            vacios.append("Sin indicador vigente")
        if con_estandar and not c.estandar:
            vacios.append("Sin estándar de cumplimiento")
        if not c.causa_efecto:
            vacios.append("Sin causa ni efecto del modelo")
        fila = {"Código": sv.codigo, "Objetivo prioritario": sv.op, "Lineamiento": sv.lineamiento,
                "Servicio": sv.servicio, "Causa o efecto que atiende": c.causa_efecto, "Tipo de vínculo": c.tipo_vinculo,
                "Indicadores vigentes": por_servicio.get(sv.codigo, "")}
        if con_estandar:
            fila["Estándar de cumplimiento"] = c.estandar
        fila |= {"Vacíos detectados": "; ".join(vacios), "Valoración": "Con vacíos" if vacios else "Completa"}
        if estado is not None:
            fila["Estado de implementación"] = estado.get(sv.codigo, "")
        if plaus is not None:
            fila["Plausibilidad causal"] = plaus.get(sv.codigo, "")
        filas.append(fila)
    srv = pd.DataFrame(filas)
    lineas = []
    universo = (lin[["lineamiento", "op"]].values.tolist() if lin is not None
                else s[["lineamiento", "op"]].drop_duplicates().values.tolist())
    for li, op in universo:
        g = srv[srv.Lineamiento == li]
        val = "Rota" if g.empty else ("Completa" if (g.Valoración == "Completa").all() else "Con vacíos")
        lineas.append({"Lineamiento": li, "Objetivo prioritario": op, "Servicios": len(g),
                       "Servicios completos": int((g.Valoración == "Completa").sum()), "Valoración": val})
    lin_df = pd.DataFrame(lineas)
    obj = (lin_df.groupby("Objetivo prioritario", sort=False).Valoración.value_counts().unstack(fill_value=0)
           .reindex(columns=["Completa", "Con vacíos", "Rota"], fill_value=0)
           .rename(columns={"Completa": "Lineamientos completos", "Con vacíos": "Lineamientos con vacíos",
                            "Rota": "Lineamientos rotos"}).reset_index())
    obj.insert(1, "Servicios", obj["Objetivo prioritario"].map(srv["Objetivo prioritario"].value_counts()).fillna(0)
               .astype(int))
    return srv, lin_df, obj


def resumen_alineamiento(ali):
    filas = []
    for (tipo, momento, op), g in ali.groupby(["tipo", "momento", "op"], sort=False):
        escala = ESCALAS[tipo]
        n = g.correspondencia.value_counts()
        filas.append({"Tipo": tipo, "Momento": momento, "Objetivo prioritario": op,
                      **{k: int(n.get(k, 0)) for k in escala},
                      "Valoración": ", ".join(f"{n.get(k, 0)} {k.lower()}" for k in escala),
                      "Referencias": "\n".join(f"{r.instrumento}: {r.referencia} ({r.correspondencia.lower()})"
                                               for r in g.itertuples())})
    return pd.DataFrame(filas)


def correr(cfg):
    s, ind, cad, lin, ali = validar(cfg)
    est = libro_previo(cfg, rev_doc, "Estado por servicio")
    pla = libro_previo(cfg, rev_evi, "Plausibilidad por servicio")
    srv, lin_df, obj = coherencia(s, ind, cad, lin,
                                  None if est is None else dict(zip(est["Código"], est["Estado asignado"])),
                                  None if pla is None else dict(zip(pla["Código"], pla["Plausibilidad"])))
    registro = ali.rename(columns={"tipo": "Tipo", "momento": "Momento", "op": "Objetivo prioritario",
                                   "instrumento": "Instrumento", "referencia": "Referencia", "documento": "Documento",
                                   "pagina": "Página declarada", "cita": "Cita", "correspondencia": "Correspondencia",
                                   "sustento": "Sustento"})
    crit = [
        ("Coherencia interna", "Cada lineamiento tiene al menos un servicio; cada servicio tiene indicador vigente (y "
                               "estándar, si se registra) y está vinculado a una causa o efecto del modelo del problema.",
         "CEPLAN (Guía de políticas nacionales: cadena de resultados)"),
        ("Valoración", "Servicio: completa si cumple las condiciones; con vacíos en otro caso. Lineamiento: rota sin "
                       "servicio, completa si todos sus servicios lo están.", "Equipo evaluador"),
        ("Alineamiento", "Correspondencia por objetivo e instrumento: " + "; ".join(
            f"{t.lower()}: {', '.join(e)}" for t, e in ESCALAS.items()) + ".", "Equipo evaluador"),
        ("Citas", "Literales, ubicadas en su documento; una cita que no aparece detiene la corrida.",
         "Repositorio (CLAUDE.md, regla 1)"),
    ]
    destino = base.ruta(cfg, "analisis", LIBRO)
    base.escribir_libro(destino, {"Coherencia por servicio": srv, "Coherencia por lineamiento": lin_df,
                                  "Coherencia por objetivo": obj, "Alineamiento por objetivo": resumen_alineamiento(ali),
                                  "Registro de alineamiento": registro}, crit)
    verif = registro["Verificación de la cita"].value_counts().to_dict()
    return {"libro": destino, "servicios_completos": int((srv.Valoración == "Completa").sum()), "servicios": len(srv),
            "citas_ubicadas": int(sum(v for k, v in verif.items() if k.startswith("Ubicada"))),
            "citas_sin_verificar": int(verif.get("Sin verificar: documento no disponible", 0))}
