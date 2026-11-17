"""Síntesis de la evaluación de cierre de la PNIG: cifras y hallazgos por pregunta de evaluación.

Módulo sin salidas propias. Lee los libros de análisis ya generados (REV_DOC, REV_ALI, REV_IND, REP_ENT y REV_EVI), calcula
las cifras que se citan en la matriz de triangulación, en las guías ENTREV y en el borrador del informe, y define los
hallazgos preliminares con sus evidencias. Así, una cifra se calcula una sola vez y los tres productos la citan igual.

Planos (nota metodológica, numeral 7 y Tabla 9):
  - CEPLAN: estructura del informe (Guía de SyE, Tabla A7.1) y matriz de recomendaciones (Tabla 4).
  - Nota metodológica: preguntas D1 a D7 e I1 a I5, escalas de valoración y regla de confianza del hallazgo.
  - Equipo evaluador: enunciado de cada hallazgo, nivel de confianza y valoración por pregunta. El usuario los validó
    el 2026-11-06 (estado VALIDADO y VALIDADA); D3-H2 e I4-H5, actualizados después, los validó el 2026-11-07.

Regla de confianza (nota, numeral 7): sólido, cuando al menos dos fuentes o instrumentos independientes coinciden; con
respaldo parcial, cuando coinciden con discrepancias explicables; indicio, cuando se sustenta en una sola fuente. Cada
evidencia declara su «base» (la fuente primaria de la que proviene); las bases se descomponen en documentos (PARTES) y
dos evidencias son independientes si sus documentos difieren. La entrevista a la OGMEPGD no cuenta como fuente
independiente, porque la OMEP, que conduce la evaluación, es parte de ese órgano (ADR-010); con esa regla ningún
hallazgo cambia de nivel. hallazgos() verifica que la confianza declarada sea compatible con el conteo.

Cada hallazgo lleva, además del enunciado, un «detalle» (1 a 3 oraciones con su cifra clave, para el Anexo 2 del
informe), una «salvedad» en lenguaje llano que el informe pone en el cuerpo, después del enunciado, y una «nota»
técnica para el Anexo 2 y la matriz de triangulación (ADR-010).

Uso:
    import sintesis_hallazgos as S
    C = S.cifras()            # diccionario de cifras; C.origen[clave] dice de dónde sale cada una
    H = S.hallazgos(C)        # lista de hallazgos con evidencias
    V = S.valoraciones(C)     # valoración validada por pregunta
Lo importan sintesis_triangulacion.py, entrev_guias_entrevista.py e informe_cierre.py (se corren desde la raíz).
"""
import re
import warnings

import pandas as pd

from docx_mimp import fmt_num, lista_es, monto, pct
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
import nucleo.hallazgos as _NH  # noqa: E402  (modelo de hallazgo y reglas de confianza del núcleo)
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

# estado de hallazgos y valoraciones: el usuario validó todo el informe el 2026-11-06 y, el 2026-11-07, los dos
# hallazgos que cambiaron después (evaluación 2020 en la 4.7, tipo de indicadores del PEMIG)
VALIDADO = "Validado por el equipo evaluador"
VALIDADO_0711 = "Validado por el equipo evaluador el 2026-11-07"
ACTUALIZADO = "Validado; actualizado el 2026-11-07, por revisar"  # ya no se usa; lo conserva la matriz de triangulación
VALIDADA = "Validada por el equipo evaluador"
# fecha de cada entrevista al nivel conductor (entrev_guias_entrevista.APLICADA; informe_cierre.py lo comprueba)
FECHA_ENTREV = {"OGPPM": "4 de noviembre de 2026", "OGMEPGD": "5 de noviembre de 2026",
                "DGIGND": "5 de noviembre de 2026"}
import svy

warnings.filterwarnings("ignore", category=UserWarning)
A = f"{EVAL}/04_analisis"
GRD = f"{A}/REV_DOC_estado_implementacion_servicios.xlsx"
MAC = f"{A}/REV_ALI_coherencia_alineamiento.xlsx"
MIC = f"{A}/REV_DOC_matriz_incompatibilidades.xlsx"   # contraste de D7 con la matriz de la DPIGND (ADR-016)
PSI = f"{A}/REV_IND_base_indicadores.xlsx"
PSI2 = f"{A}/REV_IND_2_calidad_fichas.xlsx"
PSI3V = f"{A}/REV_IND_3_verificacion_presupuestal.xlsx"
PSI3R = f"{A}/REV_IND_3_recursos_conduccion.xlsx"
PSI4 = f"{A}/REV_IND_4_recalculo_OP3.xlsx"
FIE = f"{A}/REP_ENT_informacion_entidades.xlsx"
FIEDC = f"{A}/REP_ENT_doble_codificacion.xlsx"
REV = f"{A}/REV_EVI_matriz_evidencia.xlsx"
LIBRO = f"{A}/libro_codigos_pnig.csv"
D1D2 = f"{A}/D1_D2_brechas_situacion_futura.xlsx"
DEP = f"{A}/D1_D2_brechas_departamentales.xlsx"
D6 = f"{A}/D6_enfoques_transversales.xlsx"
REC = f"{A}/seguimiento_recomendaciones_previas.xlsx"
GES = f"{EVAL}/05_entrevistas"
GES_COD = f"{GES}/ENTREV_codificacion.csv"
GES_DED = f"{GES}/ENTREV_deducciones.csv"
GES_REG = f"{GES}/ENTREV_registro_codificacion.xlsx"

# preguntas de evaluación: se leen de la configuración de la evaluación (evaluaciones/<id>/config.yaml, Fase 2)
PREGUNTAS = {p: (q["criterio"], " ".join(str(q["texto"]).split()), list(q["instrumentos"]))
             for p, q in _cargar_config()["preguntas"].items()}
CONFIANZA = ("Sólido", "Con respaldo parcial", "Indicio", "No determinado")  # el último, solo con análisis pendiente
VALORACION = ("Satisfactorio", "Parcial", "Insuficiente")
# Semáforo de la Tabla 9 de la nota (RC 2025, Anexo 01, p. 16): los mismos umbrales de rev_ind_indicadores.PNIG, que se
# comprueban contra la hoja «Criterios» de REV_IND para no depender de pdfplumber
NIVELES = [(150, "Alto con posible falla de planeación"), (95, "Alto"), (75, "Medio"), (15, "Bajo"),
           (float("-inf"), "Bajo con posible falla de planeación")]
UMBRALES_TXT = "Alto con posible falla de planeación ≥ 150; alto [95–150>; medio [75–95>; bajo [15–75>"
NOMBRE_EST = {  # indicadores de OP sin valor oficial 2025 que el equipo estimó o recalculó (I2-H3)
    "IOP.01.01": "la violencia física o sexual de pareja", "IOP.02.01": "el uso de métodos anticonceptivos",
    "IOP.02.02": "la demanda insatisfecha de planificación familiar", "IOP.02.03": "el embarazo adolescente",
    "IOP.03.01": "el índice de empoderamiento político", "IOP.03.02": "las funcionarias y directivas del servicio civil",
}
INDICES = ("IOP.03.01",)  # se expresan como índice (tres decimales), no en porcentaje
MOTIVO_SIN_VALOR = {  # indicadores de OP sin valor oficial 2025 que el equipo no estimó (verificado al 2026-11-06)
    "IOP.06.01": "IOP.06.01 no tiene fuente en 2025, porque la ENARES no se aplicó ese año",
}


def fmt_est(cod, valor):
    """Valor de un indicador de OP con su unidad: «0,583» para los índices, «40,3 %» para los porcentajes."""
    return fmt_num(valor, 3) if cod in INDICES else f"{fmt_num(valor, 1)} %"


def nivel_pnig(avance):
    return next(etq for umbral, etq in NIVELES if avance >= umbral)


def frase_est25(C):
    """«en nivel alto el embarazo adolescente (7,4 %, meta 10,3) y …; y en nivel bajo, …» (estimaciones ENDES 2025)."""
    e = C["op_est25"]
    segs = []
    for _, niv in NIVELES:
        sub = e[e.Nivel == niv]
        if len(sub):
            segs.append(f"en nivel {niv.lower()}{',' if segs else ''} " + lista_es(
                [f"{NOMBRE_EST[r.Código]} ({fmt_est(r.Código, r.Estimación)}, meta "
                 f"{fmt_num(r['Meta 2025'], 3 if r.Código in INDICES else 1)})" for _, r in sub.iterrows()]))
    return lista_pc(segs)


def queda_sin_valor(C):
    """«queda 1 indicador sin valor 2025, del OP6» o «quedan 3 indicadores sin valor 2025, de los OP3 y OP6»."""
    n, ops = len(C["op_sinvalor_resto"]), C["op_sin_valor_op"]
    de = f"del {ops[0]}" if len(ops) == 1 else f"de los {lista_es(ops)}"
    return f"queda 1 indicador sin valor 2025, {de}" if n == 1 else f"quedan {n} indicadores sin valor 2025, {de}"


def lista_pc(items):
    """Enumeración con punto y coma, para elementos que ya llevan comas."""
    return items[0] if len(items) == 1 else "; ".join(items[:-1]) + "; y " + items[-1]


class Cifras(dict):
    """Diccionario de cifras con su origen (libro y hoja, o cálculo)."""

    def __init__(self):
        super().__init__()
        self.origen = {}

    def put(self, clave, valor, origen):
        self[clave] = valor
        self.origen[clave] = origen


def _leer(libro, hoja):
    return pd.read_excel(libro, sheet_name=hoja)


def _hoja(libro, hoja):
    return f"{libro.split('/')[-1]}, hoja «{hoja}»"


def _conteo_valoracion(texto):
    m = re.match(r"(\d+) directa, (\d+) indirecta, (\d+) sin correspondencia", str(texto))
    return tuple(int(x) for x in m.groups()) if m else (0, 0, 0)


def _num(x):
    return pd.to_numeric(x, errors="coerce")


def cifras():
    C = Cifras()
    # REV_DOC: estado de implementación por servicio
    g = _leer(GRD, "Estado por servicio")
    o = _hoja(GRD, "Estado por servicio")
    est = g["Estado asignado (preliminar)"]
    C.put("serv_total", len(g), o)
    C.put("serv_impl", int((est == "Implementado").sum()), o)
    C.put("serv_parc", int((est == "Implementado parcialmente").sum()), o)
    C.put("serv_noimpl", int((est == "No implementado").sum()), o)
    C.put("serv_noimpl_cod", g.loc[est == "No implementado", "Código"].astype(str).tolist(), o)
    C.put("serv_reclas", int((g["Estado asignado (preliminar)"] != g["Clasificación por criterios"]).sum()), o)
    C.put("serv_prot_si", int((g["Protocolo aprobado"] == "Sí").sum()), o)
    C.put("serv_prot_no", int((g["Protocolo aprobado"] == "No").sum()), o)
    C.put("serv_asig_si", int((g["Asignación presupuestal identificable"] == "Sí").sum()), o)
    C.put("serv_con_ind", int((_num(g["Indicadores del servicio"]) > 0).sum()), o)
    C.put("serv_falta_tres", int((g["Componentes faltantes"] == "protocolo, asignación identificable, reporte de "
                                                               "indicador").sum()), o)
    por_op = g.groupby("OP")["Estado asignado (preliminar)"].value_counts().unstack(fill_value=0)
    C.put("estado_por_op", por_op, o)
    C.put("serv_impl_cod", g.loc[est == "Implementado", "Código"].astype(str).tolist(), o)
    C.put("grd", g, o)

    # REV_ALI.1: coherencia interna
    lin = _leer(MAC, "REV_ALI.1 Por lineamiento")
    o = _hoja(MAC, "REV_ALI.1 Por lineamiento")
    C.put("lin_total", len(lin), o)
    sin = lin[_num(lin["N.° de servicios"]).fillna(0) == 0]
    C.put("lin_sin_serv", [str(x).split(" ")[0].rstrip(".") for x in sin["Lineamiento"]], o)
    ps = _leer(MAC, "REV_ALI.1 Por servicio")
    o = _hoja(MAC, "REV_ALI.1 Por servicio")
    C.put("serv_sin_causa", ps.loc[ps["Causa o efecto del modelo (Anexo 5)"].isna(), "Código"].astype(str).tolist(), o)
    C.put("serv_sin_ind_vig", int(ps["Vacíos detectados"].astype(str).str.contains("Sin indicador").sum()), o)
    pob = _leer(MAC, "REV_ALI.1 Por objetivo")
    o = _hoja(MAC, "REV_ALI.1 Por objetivo")
    dic = " ".join(pob["Dictamen de las fichas del OP (REV_IND.2)"].astype(str))
    C.put("op_fichas_obs", dic.count("Observada"), o)
    C.put("op_fichas_total", len(re.findall(r"IOP\.\d\d\.\d\d", dic)), o)

    # REV_ALI.2: alineamiento vertical
    v = _leer(MAC, "REV_ALI.2 Por objetivo")
    o = _hoja(MAC, "REV_ALI.2 Por objetivo")
    per = [_conteo_valoracion(x) for x in v["Valoración (periodo)"]]
    tra = [_conteo_valoracion(x) for x in v["Valoración (transición)"]]
    C.put("vert_periodo", tuple(sum(x[i] for x in per) for i in range(3)), o)
    C.put("vert_transicion", tuple(sum(x[i] for x in tra) for i in range(3)), o)
    col43 = [c for c in v.columns if "043-2026" in c][0]
    col127 = [c for c in v.columns if "127-2026" in c][0]
    C.put("pgg043_sin", int(v[col43].astype(str).str.contains("sin correspondencia").sum()), o)
    C.put("pgg127_sin_op", v.loc[v[col127].astype(str).str.contains("sin correspondencia"), "OP"].tolist(), o)
    C.put("pgg127_ind_op", v.loc[v[col127].astype(str).str.contains("indirecta"), "OP"].tolist(), o)
    C.put("vert_por_op", v[["OP", "Valoración (periodo)", "Valoración (transición)"]], o)
    directas = {r.OP: _conteo_valoracion(r["Valoración (periodo)"])[0] for _, r in v.iterrows()}
    C.put("vert_fuerte", [k for k, x in directas.items() if x == max(directas.values())], o)
    C.put("vert_debil", [k for k, x in directas.items() if x == min(directas.values())], o)
    colpesem = [c for c in v.columns if "PESEM 2024" in c][0]
    C.put("pesem24_directa", int(v[colpesem].astype(str).str.contains("directa").sum()), o)

    # REV_ALI.3: alineamiento horizontal
    h = _leer(MAC, "REV_ALI.3 Horizontal")
    o = _hoja(MAC, "REV_ALI.3 Horizontal")
    tr = h["Tipo de relación"].value_counts()
    C.put("hor_rel", len(h), o)
    C.put("hor_pol", int(h["Política nacional afín"].nunique()), o)
    C.put("hor_dup", int(tr.get("Duplicidad", 0)), o)
    C.put("hor_comp", int(tr.get("Complementariedad", 0)), o)
    C.put("hor_vac", int(tr.get("Vacío de atención", 0)), o)
    cem = h[(h["Servicio o lineamiento de la PNIG"].astype(str) == "1.1.3") & (h["Tipo de relación"] == "Duplicidad")]
    C.put("hor_cem", len(cem), o)
    C.put("hor_cem_pol", sorted(cem["Política nacional afín"].unique().tolist()), o)

    # REV_ALI.4: normativa e internacional
    i4 = _leer(MAC, "REV_ALI.4 Por objetivo")
    o = _hoja(MAC, "REV_ALI.4 Por objetivo")
    inter = [_conteo_valoracion(x) for x in i4["Valoración (internacional)"]]
    C.put("intl", tuple(sum(x[i] for x in inter) for i in range(3)), o)
    C.put("ley_dir_op", i4.loc[i4["Valoración (Ley N.° 32535)"] == "Directa", "OP"].tolist(), o)
    C.put("ley_ind_op", i4.loc[i4["Valoración (Ley N.° 32535)"] == "Indirecta", "OP"].tolist(), o)

    # REV_ALI.5: correspondencia conceptual
    r5 = _leer(MAC, "Resumen REV_ALI.5").set_index("Plano")
    C.put("mac5", r5, _hoja(MAC, "Resumen REV_ALI.5"))
    m5 = _leer(MAC, "REV_ALI.5 Conceptual")
    o = _hoja(MAC, "REV_ALI.5 Conceptual")
    C.put("mac5_serv_rediseña", m5.loc[(m5.Plano == "Servicios") & (m5.Efecto == "Se rediseña"), "Código"]
          .astype(str).tolist(), o)
    C.put("mac5_rediseña", m5.loc[m5.Efecto == "Se rediseña", ["Plano", "Código", "Elemento de la PNIG"]], o)
    C.put("mac5_n", len(m5), o)
    C.put("mac5_entrevista", _leer(MAC, "Para la entrevista DGIGND"), _hoja(MAC, "Para la entrevista DGIGND"))
    # contraste de la clasificación de los servicios con la matriz de incompatibilidades de la DPIGND (ADR-016)
    C.put("mic_conc", _leer(MIC, "Concordancia").set_index("Medida")["Valor"], _hoja(MIC, "Concordancia"))
    ms = _leer(MIC, "Servicios").assign(Código=lambda x: x["Código"].map(str))
    C.put("mic_serv", ms, _hoja(MIC, "Servicios"))
    C.put("mic_norma", _leer(MIC, "Resumen por norma"), _hoja(MIC, "Resumen por norma"))
    C.put("mic_cruce", _leer(MIC, "Tabla cruzada"), _hoja(MIC, "Tabla cruzada"))

    # REV_IND: indicadores
    b = _leer(PSI, "Base REV_IND.1")
    o = _hoja(PSI, "Base REV_IND.1")
    op = b[b.Nivel == "OP"].copy()
    C.put("op_ind", len(op), o)
    C.put("op_base", op, o)
    C.put("op_sinvalor_2025_cod", op.loc[op["Semáforo 2025"].astype(str).str.startswith("Sin valor"), "Código"].tolist(),
          o)
    C.put("serv_ind", int((b.Nivel == "Servicio").sum()), o)
    for anio in (2024, 2025):
        sem = op[f"Semáforo {anio}"].astype(str)
        C.put(f"op_sinvalor_{anio}", int(sem.str.startswith("Sin valor").sum()), o)
        C.put(f"op_alto_{anio}", int(sem.str.startswith("Alto").sum()), o)
        C.put(f"op_medio_{anio}", int(sem.str.startswith("Medio").sum()), o)
        C.put(f"op_bajo_{anio}", int(sem.str.startswith("Bajo").sum()), o)
        C.put(f"op_noint_{anio}", int(sem.str.startswith("No interpretable").sum()), o)
        C.put(f"op_fnc_{anio}", int(sem.str.contains("fuente no comparable").sum()), o)
        C.put(f"op_bajofalla_{anio}", int(sem.str.startswith("Bajo con posible falla").sum()), o)
    filas = []
    for _, r in op.iterrows():
        lb = _num(pd.Series([r["Valor LB"]])).iloc[0]
        ult, anio_ult = None, None
        for anio in range(2025, 2018, -1):
            val = _num(pd.Series([r[f"Valor {anio}"]])).iloc[0]
            if pd.notna(val):
                ult, anio_ult = val, anio
                break
        if pd.isna(lb) or ult is None:
            mejora = "Sin línea de base numérica" if pd.isna(lb) else "Sin valor"
        elif str(r[f"Semáforo {anio_ult}"]).startswith("No interpretable"):
            mejora = "No interpretable"
        elif r["Sentido aplicado"] == "Descendente":
            mejora = "Mejora" if ult < lb else "No mejora"
        else:
            mejora = "Mejora" if ult > lb else "No mejora"
        filas.append({"Código": r["Código"], "OP": r["OP"].replace("OP.0", "OP"), "Indicador": r["Indicador"],
                      "Sentido": r["Sentido aplicado"], "Año LB": r["Año LB"], "Valor LB": r["Valor LB"],
                      "Último valor": ult, "Año del último valor": anio_ult, "Meta 2025": r["Meta 2025"],
                      "Semáforo 2024": r["Semáforo 2024"], "Semáforo 2025": r["Semáforo 2025"],
                      "Respecto de la LB": mejora})
    av25 = _num(op.set_index("Código")["Avance 2025 (%)"]).dropna()
    assert av25.idxmin() == "IOP.05.01", "el menor avance 2025 ya no es el de IOP.05.01; revisar el texto de I2-H2"
    C.put("iop0501_avance", float(av25["IOP.05.01"]), o)
    opt = pd.DataFrame(filas)
    C.put("op_tabla", opt, o + "; cálculo: último valor disponible frente a la línea de base, según el sentido aplicado")
    C.put("op_mejora", int((opt["Respecto de la LB"] == "Mejora").sum()), C.origen["op_tabla"])
    C.put("op_nomejora", opt.loc[opt["Respecto de la LB"] == "No mejora", "Código"].tolist(), C.origen["op_tabla"])
    C.put("op_sinlb", opt.loc[opt["Respecto de la LB"].str.startswith("Sin"), "Código"].tolist(),
          C.origen["op_tabla"])
    C.put("op_noint_lb", opt.loc[opt["Respecto de la LB"] == "No interpretable", "Código"].tolist(),
          C.origen["op_tabla"])
    res = _leer(PSI, "Resumen PNIG")
    o = _hoja(PSI, "Resumen PNIG")
    C.put("resumen_semaforo", res, o)
    sv = res[res.Nivel == "Servicio"].set_index("Semáforo PNIG")[2025]
    C.put("serv25", sv.to_dict(), o)
    C.put("serv25_total", int(sv.sum()), o)
    C.put("serv25_falla", int(sv[[i for i in sv.index if "falla" in i]].sum()), o)
    base = _leer(PSI, "Base REV_IND.1")
    C.put("serv25_con_valor", int(base.loc[base.Nivel == "Servicio", "Valor 2025"].notna().sum()),
          _hoja(PSI, "Base REV_IND.1"))

    # REV_IND.2: calidad de las fichas
    r2 = _leer(PSI2, "Resumen").set_index("Atributo")
    o = _hoja(PSI2, "Resumen")
    C.put("fichas_total", int(r2.iloc[0].sum()), o)
    C.put("ficha_lb_nc", int(r2.loc["Línea de base con año y valor", "No cumple"]), o)
    C.put("ficha_lb_parc", int(r2.loc["Línea de base con año y valor", "Cumple parcialmente"]), o)
    fila_des = [i for i in r2.index if i.startswith("Desagregaciones")][0]
    C.put("ficha_desag_nc", int(r2.loc[fila_des, "No consta"]), o)
    C.put("ficha_meta_nc", int(r2.loc["Meta anual hasta el fin del horizonte", "No cumple"]), o)
    fila_sen = [i for i in r2.index if i.startswith("Sentido esperado")][0]
    C.put("ficha_sentido_obs", int(r2.loc[fila_sen, "No cumple"] + r2.loc[fila_sen, "Cumple parcialmente"]), o)
    cj = _leer(PSI2, "Conjunto")
    txt = cj.loc[cj.Elemento == "Servicios sin indicador", "Observación"].iloc[0]
    C.put("serv_sin_ind_omep", int(re.match(r"(\d+)", txt).group(1)), _hoja(PSI2, "Conjunto"))
    C.put("serv_falta_cob", int(cj.Observación.astype(str).str.contains("le falta un indicador").sum()),
          _hoja(PSI2, "Conjunto"))
    tp = _leer(PSI2, "Tipo PEMIG")
    o = _hoja(PSI2, "Tipo PEMIG")
    por_tipo = tp.tipo.value_counts()
    C.put("pemig_ind", len(tp), o)
    C.put("pemig_serv", int(tp.servicio.astype(str).nunique()), o)
    C.put("pemig_cob", int(por_tipo.get("Cobertura", 0)), o)
    C.put("pemig_cal", int(por_tipo.get("Calidad", 0)), o)
    C.put("pemig_part", int((tp.atributo == "Participación de las mujeres entre los beneficiarios").sum()), o)
    C.put("pemig_serv_cal", int(tp.loc[tp.tipo == "Calidad", "servicio"].astype(str).nunique()), o)

    # REV_IND.3: recursos de la conducción y verificación presupuestal
    co = _leer(PSI3V, "Conducción por órgano").set_index("Órgano")
    o = _hoja(PSI3V, "Conducción por órgano")
    org = {k: [i for i in co.index if i.startswith(k)][0] for k in ("DGIGND", "DPIGND", "DPPDM", "DGTEG")}
    dev = {a: {k: float(_num(pd.Series([co.loc[v, f"{a} Devengado"]])).fillna(0).iloc[0]) for k, v in org.items()}
           for a in range(2019, 2026)}
    pim = {a: {k: float(_num(pd.Series([co.loc[v, f"{a} PIM"]])).fillna(0).iloc[0]) for k, v in org.items()}
           for a in range(2019, 2026)}
    C.put("cond_dev", dev, o)
    C.put("cond_pim", pim, o)
    cond = {a: dev[a]["DGIGND"] + dev[a]["DPIGND"] + dev[a]["DPPDM"] for a in dev}
    C.put("cond_suma", cond, o + "; cálculo: DGIGND + DPIGND + DPPDM (las dos últimas desde el ROF 2023)")
    C.put("cond_var_22_23", (cond[2023] / cond[2022] - 1) * 100, C.origen["cond_suma"])
    C.put("dgignd_ejec", {a: dev[a]["DGIGND"] / pim[a]["DGIGND"] * 100 for a in dev}, o)
    vd = _leer(PSI3R, "Vinculada directa").set_index("anio")
    o = _hoja(PSI3R, "Vinculada directa")
    C.put("act5001072", vd, o)
    rv = _leer(PSI3V, "Resumen")
    o = _hoja(PSI3V, "Resumen")
    coinc = int(rv["Coincide"].sum() + rv["Coincide con el año anterior"].sum())
    cerc = int(sum(rv[c].sum() for c in rv.columns if c.startswith("Cercano")))
    sinc = int(rv["Sin coincidencia"].sum())
    C.put("mef_coinc", coinc, o)
    C.put("mef_cerc", cerc, o)
    C.put("mef_sinc", sinc, o)
    C.put("mef_total", coinc + cerc + sinc, o)
    mo = _leer(PSI3V, "Sin coincidencia por motivo").set_index("Motivo probable (sin coincidencia)")
    mot = mo.sum(axis=1).astype(int)
    clave = {"Actividad operativa": "ao", "Monto de una norma": "norma", "No es gasto del pliego": "pliego",
             "Sin explicación": "sinexp", "Total de varias": "varias"}
    C.put("mef_motivos", {v: int(mot[[i for i in mot.index if i.startswith(k)]].sum()) for k, v in clave.items()},
          _hoja(PSI3V, "Sin coincidencia por motivo"))
    assert sum(C["mef_motivos"].values()) == sinc, "los motivos no suman los montos sin coincidencia"

    # D1 y D2: brechas, manifestaciones (ENDES) y situación futura deseada
    ev = _leer(D1D2, "Evolución de las brechas")
    o = _hoja(D1D2, "Evolución de las brechas")
    C.put("brechas", ev, o)
    # una brecha seguida con dos encuestas sucesivas se cuenta una vez; ambas fuentes deben coincidir en la
    # clasificación (ninguna desde que el empleo informal de la ENAHO 2024 y 2025 se reconstruye)
    cont = ev[ev.Clave != ev["Brecha del diagnóstico"]]
    for _, r in cont.iterrows():
        orig = ev.loc[ev.Clave == r["Brecha del diagnóstico"], "Clasificación"].iloc[0]
        assert orig == r["Clasificación"], f"{r.Clave}: la continuación no coincide con la serie original"
    ev1 = ev[ev.Clave == ev["Brecha del diagnóstico"]]
    C.put("brechas_n", len(ev1), o)
    corto = {"tasa_actividad": "tasa de actividad", "tasa_desempleo": "desempleo",
             "sin_ingresos": "población sin ingresos propios",
             "informal": "empleo informal", "credito_solicito": "crédito agropecuario",
             "tfnr": "trabajo familiar no remunerado", "trabajo_hogar": "trabajo del hogar",
             "afiliado_pension": "afiliación a pensiones", "ingreso_laboral": "ingreso laboral",
             "ingreso_publico": "ingreso en el sector público", "analfabetismo": "analfabetismo",
             "internet": "uso de internet", "nini_15_29": "jóvenes que no estudian ni están ocupados",
             "horas_no_remunerado": "trabajo no remunerado entre 2010 y 2024",
             "tolerancia_10": "tolerancia social a la violencia entre 2019 y 2024"}
    for clase, clave in (("Se reduce", "reduce"), ("Sin cambio significativo", "sincambio"), ("Se amplía", "amplia")):
        C.put(f"brechas_{clave}", [corto[k] for k in ev1.loc[ev1["Clasificación"] == clase, "Clave"]], o)
    C.put("brechas_persisten", int((ev1["Persiste en el año final"] == "Sí").sum()), o)
    serie = _leer(D1D2, "Brechas por año")
    C.put("brechas_serie", serie, _hoja(D1D2, "Brechas por año"))
    man = _leer(D1D2, "Manifestaciones ENDES")
    o = _hoja(D1D2, "Manifestaciones ENDES")
    C.put("endes", man, o)
    val = man[man.oficial.notna()]
    C.put("endes_contrastes", len(val), o)
    C.put("endes_fuera", int((val.diferencia_oficial.abs() > 1.959964 * val.ee).sum()), o)
    # Indicadores de OP sin valor oficial 2025: estimación del equipo con la ENDES 2025 frente a la meta (I2-H3)
    crit = _leer(PSI, "Criterios")
    assert crit.Regla.astype(str).str.contains(UMBRALES_TXT, regex=False).any(), "cambiaron los umbrales de la Tabla 9"
    base = C["op_base"].set_index("Código")

    def nivel_de(cod, valor):
        meta, sentido = float(base.loc[cod, "Meta 2025"]), base.loc[cod, "Sentido aplicado"]
        av = 100 * (valor / meta if sentido == "Ascendente" else meta / valor)
        return av, nivel_pnig(av)

    filas = []
    for _, r in man[(man.anio == 2025) & man.codigo_op.notna()].iterrows():
        filas.append({"Código": r.codigo_op, "Estimación": r.estimacion, "IC inferior": r.ic_inf,
                      "IC superior": r.ic_sup, "Fuente": "ENDES 2025", "Tipo": "Estimado"})
    r4 = _leer(PSI4, "Resultado 2025")
    for _, r in r4.iterrows():  # OP3: ENAHO 2025 (IOP.03.02) y registros electorales (IOP.03.01), REV_IND.4
        assert abs(r["Meta 2025"] - float(base.loc[r.Código, "Meta 2025"])) < 1e-6, f"{r.Código}: meta distinta"
        assert r.Tipo in ("Estimado", "Recalculado"), f"{r.Código}: el método de REV_IND.4 no cumplió su regla"
        filas.append({"Código": r.Código, "Estimación": r.Estimación, "IC inferior": r["IC inferior"],
                      "IC superior": r["IC superior"],
                      "Fuente": "ENAHO 2025" if r.Código == "IOP.03.02" else "Registros electorales", "Tipo": r.Tipo})
    for f in filas:
        av, niv = nivel_de(f["Código"], f["Estimación"])
        niveles_ic = {nivel_de(f["Código"], x)[1] for x in (f["IC inferior"], f["IC superior"])} | {niv}
        f |= {"OP": base.loc[f["Código"], "OP"].replace("OP.0", "OP"),
              "Meta 2025": float(base.loc[f["Código"], "Meta 2025"]), "Avance (%)": av, "Nivel": niv,
              "Niveles en el IC": [n for _, n in NIVELES if n in niveles_ic]}
    est25 = pd.DataFrame(filas).sort_values("Código").reset_index(drop=True)
    assert set(est25.Código) == set(NOMBRE_EST), "cambiaron los indicadores de OP estimados o recalculados"
    assert set(est25.Código) <= set(C["op_sinvalor_2025_cod"]), "un indicador estimado ya tiene valor oficial 2025"
    C.put("op_est25", est25, f"{o}; {_hoja(PSI4, 'Resultado 2025')}; avance y nivel con la regla de la Tabla 9 y la "
                             f"meta 2025 de REV_IND")
    C.put("op3_serie", _leer(PSI4, "IOP.03.02 serie"), _hoja(PSI4, "IOP.03.02 serie"))
    C.put("op3_indice", _leer(PSI4, "IOP.03.01 componentes"), _hoja(PSI4, "IOP.03.01 componentes"))
    resto = [c for c in C["op_sinvalor_2025_cod"] if c not in set(est25.Código)]
    assert set(resto) == set(MOTIVO_SIN_VALOR), f"revisar MOTIVO_SIN_VALOR: {resto}"
    C.put("op_sinvalor_resto", resto, C.origen["op_sinvalor_2025_cod"])
    con_valor = set(base.index[_num(base["Valor 2025"]).notna()]) | set(est25.Código)
    sin_op = sorted(set(base.OP) - set(base.loc[sorted(con_valor), "OP"]))
    C.put("op_sin_valor_op", [x.replace("OP.0", "OP") for x in sin_op], C.origen["op_est25"])
    sfd = _leer(D1D2, "Situación futura deseada")
    C.put("sfd", sfd, _hoja(D1D2, "Situación futura deseada"))
    pp = _leer(D1D2, "Participación política")
    C.put("participacion", pp, _hoja(D1D2, "Participación política"))
    v = _leer(D1D2, "Validación")
    C.put("d1_validacion", v, _hoja(D1D2, "Validación"))
    # brechas por departamento (d1_d2_departamental.py)
    C.put("dep_resumen", _leer(DEP, "Resumen por indicador"), _hoja(DEP, "Resumen por indicador"))
    C.put("dep_cambio", pd.read_excel(DEP, sheet_name="Cambio por ámbito", dtype={"ambito": str}),
          _hoja(DEP, "Cambio por ámbito"))
    vd = _leer(DEP, "Validación con el INEI")
    C.put("dep_validacion", (int((vd["Dentro del IC 95 %"] == "Sí").sum()), len(vd)), _hoja(DEP, "Validación con el INEI"))
    C.put("informal_validacion", pd.read_csv(f"{A}/d1_brechas/validacion_informal_enaho.csv"),
          f"{A}/d1_brechas/validacion_informal_enaho.csv")
    C.put("dep_participacion", pd.read_excel(DEP, sheet_name="Participación política", dtype={"ambito": str}),
          _hoja(DEP, "Participación política"))
    # D6: enfoques transversales
    lv = _leer(D6, "Lista de verificación")
    C.put("enfoques", lv, _hoja(D6, "Lista de verificación"))
    C.put("enfoques_res", _leer(D6, "Resumen de la evidencia"), _hoja(D6, "Resumen de la evidencia"))

    # REP_ENT: información reportada por las entidades
    cod = _leer(FIE, "Codificación")
    o = _hoja(FIE, "Codificación")
    C.put("seg_total", len(cod), o)
    C.put("seg_form", int(cod.fuente.astype(str).str.startswith("Formulario").sum()), o)
    C.put("seg_lio", int(cod.fuente.astype(str).str.startswith("Informe LIO").sum()), o)
    C.put("seg_conaig", int(cod.fuente.astype(str).str.contains("CONAIG").sum()), o)
    libro = pd.read_csv(LIBRO).set_index("codigo")
    ex = cod.assign(c=cod.codigos.astype(str).str.split(";")).explode("c")
    ex["c"] = ex.c.str.strip()
    dif = ex[ex.sentido.isin(["Dificultad", "Mixto"]) & ex.c.str.startswith("I1.")]
    top = dif.c.value_counts().head(6)
    C.put("fie_dif_top", [(k, libro.loc[k, "nombre"], int(n)) for k, n in top.items()], o +
          "; cálculo: segmentos con sentido «Dificultad» o «Mixto» por código I1")
    C.put("fie_cod_total", ex.c.value_counts().to_dict(), o)
    C.put("fie_cod_dif", ex[ex.sentido.isin(["Dificultad", "Mixto"])].c.value_counts().to_dict(), o)
    C.put("fie_seg", cod.set_index("id"), o)
    hc = _leer(FIE, "Hallazgos críticos")
    C.put("hall_crit", len(hc), _hoja(FIE, "Hallazgos críticos"))

    # Sección 3.7: recomendaciones de evaluaciones anteriores (seguimiento_recomendaciones.py)
    rec = _leer(REC, "Matriz")
    C.put("rec", rec, _hoja(REC, "Matriz"))
    C.put("rec_resumen", _leer(REC, "Resumen por evaluación").set_index("Evaluación"),
          _hoja(REC, "Resumen por evaluación"))
    dp = _leer(FIE, "Discrepancias del panel")
    C.put("panel_disc", len(dp), _hoja(FIE, "Discrepancias del panel"))
    dc = _leer(FIEDC, "Resumen").set_index("Indicador")["Valor"]
    o = _hoja(FIEDC, "Resumen")
    C.put("dc_n", int(dc["Segmentos doblemente codificados"]), o)
    C.put("dc_kappa_cod", float(dc["Kappa medio por código, ponderado por frecuencia"]), o)
    C.put("dc_atrib", float(dc["Acuerdo en la atribución (%)"]), o)
    C.put("libro", libro, LIBRO.split("/")[-1])

    # REV_EVI: revisión de evidencia
    mt = _leer(REV, "Matriz por tipo")
    o = _hoja(REV, "Matriz por tipo")
    C.put("rev_tipos", len(mt), o)
    sol = mt.Solidez.value_counts()
    for k in ("Sólida", "Moderada", "Limitada"):
        C.put(f"rev_{k.lower()}", int(sol.get(k, 0)), o)
    lim = mt[mt.Solidez == "Limitada"]
    C.put("rev_lim", [(r.Tipo, r["Tipo de intervención"], r.Servicios) for _, r in lim.iterrows()], o)
    C.put("rev_parcial", mt.loc[mt["Pertinencia del resultado"] == "Parcial", "Tipo de intervención"].tolist(), o)
    C.put("rev_mixta", int((mt.Dirección == "mixta").sum()), o)
    C.put("rev_tabla", mt[["Tipo", "Tipo de intervención", "Servicios", "Solidez", "Dirección",
                           "Pertinencia del resultado"]], o)
    pl = _leer(REV, "Plausibilidad por servicio")
    o = _hoja(REV, "Plausibilidad por servicio")
    sp = pl.Solidez.value_counts()
    C.put("rev_serv", {k: int(sp.get(k, 0)) for k in ("Sólida", "Moderada", "Limitada")}, o)
    ref = _leer(REV, "Referencias")
    C.put("rev_fuentes", len(ref), _hoja(REV, "Referencias"))
    # ENTREV: segmentos codificados de las entrevistas al nivel conductor y deducciones del equipo evaluador
    gs = pd.read_csv(GES_COD, dtype=str).fillna("").set_index("id_segmento")
    C.put("entrev", gs, GES_COD.split("/")[-1])
    C.put("entrev_ded", pd.read_csv(GES_DED, dtype=str).fillna("").set_index("id"), GES_DED.split("/")[-1])
    C.put("entrev_n", gs.groupby("informante").size().to_dict(), GES_COD.split("/")[-1])
    C.put("entrev_criticos", int((gs.hallazgo_critico == "Candidato").sum()), GES_COD.split("/")[-1])
    # lo que el usuario validó y los vacíos salen del registro que arma entrev_guias_entrevista.py (corrida anterior)
    try:
        reg = pd.read_excel(GES_REG, sheet_name="Registro", dtype=str).fillna("")
        vac = pd.read_excel(GES_REG, sheet_name="Vacíos", dtype=str).fillna("")
    except FileNotFoundError:
        reg = pd.DataFrame(columns=["informante", "validacion"])
        vac = pd.DataFrame(columns=["tipo", "referencia"])
    C.put("entrev_validados", int((reg.validacion == "Conforme").sum()), _hoja(GES_REG, "Registro"))
    C.put("entrev_val_inf", reg[reg.validacion == "Conforme"].groupby("informante").size().to_dict(),
          _hoja(GES_REG, "Registro"))
    C.put("entrev_sin_contraste", vac.loc[vac.tipo == "Hallazgo sin contraste", "referencia"].tolist(),
          _hoja(GES_REG, "Vacíos"))
    return C


def corto(texto):
    """Abrevia las rutas de los libros: «REV_ALI_coherencia_alineamiento.xlsx, hoja «X»» -> «REV_ALI, hoja «X»»."""
    return re.sub(r"\b(REV_DOC|ENTREV|REP_ENT|REV_ALI|REV_IND|REV_EVI)(?:_(\d))?_[\w\-]+\.xlsx",
                  lambda m: m.group(1) + (f".{m.group(2)}" if m.group(2) else ""), texto)


# Nombres de los instrumentos y de sus análisis para el texto del informe (ADR-009, actualización del 2026-11-07): el
# cuerpo usa el nombre; la sigla queda en el glosario del numeral 2.3 y en el rastro (Anexo 1 y registro de tablas y
# figuras)
NOMBRES = {
    "REV_DOC": "revisión documental", "ENTREV": "entrevistas al nivel conductor",
    "REP_ENT": "información reportada por las entidades", "REV_ALI": "análisis de coherencia y alineamiento",
    "REV_ALI.1": "análisis de coherencia interna", "REV_ALI.2": "análisis del alineamiento vertical",
    "REV_ALI.3": "análisis del alineamiento horizontal",
    "REV_ALI.4": "análisis de la correspondencia normativa e internacional",
    "REV_ALI.5": "análisis de la correspondencia conceptual", "REV_IND": "revisión de indicadores y recursos",
    "REV_IND.1": "base de indicadores", "REV_IND.2": "revisión de la calidad de las fichas",
    "REV_IND.3": "análisis de los recursos de la conducción", "REV_IND.4": "recálculo de los indicadores del OP3",
    "REV_EVI": "revisión de evidencia"}
CORTOS = {"REV_DOC": "Revisión documental", "ENTREV": "Entrevistas", "REP_ENT": "Reportes de las entidades",
          "REV_ALI": "Coherencia y alineamiento", "REV_IND": "Indicadores y recursos", "REV_EVI": "Revisión de evidencia"}
# primeras palabras de las fuentes que van en minúscula dentro de la frase; las demás son nombres propios (ENDES, PEMIG)
_COMUNES = {"Documento", "Matriz", "Informes", "Reportes", "Fichas", "Formulario", "Protocolos", "Instrumentos", "Políticas",
            "Normas", "Actas", "Datos", "Microdatos", "Literatura", "Resultados", "Consolidado", "Entrevista"}


# documentos que componen cada «base» de evidencia: sirven para contar fuentes independientes y para nombrarlas sin
# repetir (ADR-010). Las bases que no figuran son un solo documento con su propio nombre.
PARTES = {
    "Encuesta Nacional de Hogares (ENAHO)": ["ENAHO"],
    "ENAHO; INEI, Brechas de Género 2025; JNE, Perfil Electoral N.° 4":
        ["ENAHO", "INEI, Brechas de Género 2025", "JNE, Perfil Electoral N.° 4"],
    "Encuesta Demográfica y de Salud Familiar (ENDES)": ["ENDES"],
    "Encuesta Nacional de Uso del Tiempo (ENUT)": ["ENUT"],
    "Encuesta Nacional sobre Relaciones Sociales (ENARES)": ["ENARES"],
    "Encuesta Permanente de Empleo Nacional (EPEN)": ["EPEN"],
    "Encuesta Nacional Agropecuaria (ENA)": ["ENA"],
    "Microdatos del INEI por departamento": ["microdatos del INEI"],
    "Documento de la PNIG (Anexo 5)": ["documento de la PNIG"],
    "Documento de la PNIG": ["documento de la PNIG"],
    "Documento de la PNIG y marco conceptual": ["documento de la PNIG", "marco conceptual de la DGIGND"],
    "PEMIG (anexo 5)": ["PEMIG"],
    "Formulario de seguimiento e informes LIO": ["formulario de seguimiento", "informes LIO"],
    "MIMP (LIO), JNE y ONPE": ["informes LIO", "JNE", "ONPE"],
    "Jurado Nacional de Elecciones (JNE) y ONPE": ["JNE", "ONPE"],
    "Oficina Nacional de Procesos Electorales (ONPE)": ["ONPE"],
    "Resultados de 2026 (JNE y ONPE)": ["JNE", "ONPE"],
    "Reportes de seguimiento (RS); INEI; JNE": ["reportes de seguimiento", "INEI", "JNE"],
    "Reportes de seguimiento (RS)": ["reportes de seguimiento"],
    "Reportes de cumplimiento (RC)": ["reportes de cumplimiento"],
    "Literatura especializada": ["literatura especializada"],
}
# origen primario de cada fuente (revisión institucional del 13-11-2026, observaciones 20 y 60): dos documentos que
# reproducen el dato de un mismo productor cuentan como un solo origen. Las fuentes que no figuran son su propio origen.
_AUTORREPORTE = "información reportada por las entidades"
# los reportes de seguimiento y de cumplimiento consolidan los valores que informan las entidades responsables
_REPORTE_OFICIAL = _AUTORREPORTE
ORIGEN = {
    "informes LIO": _AUTORREPORTE, "formulario de seguimiento": _AUTORREPORTE,
    "informes de gobernanza de la CONAIG": _AUTORREPORTE,
    "reportes de seguimiento": _REPORTE_OFICIAL, "reportes de cumplimiento": _REPORTE_OFICIAL,
    "consolidado de indicadores de la OMEP": _REPORTE_OFICIAL,
    "datos abiertos del MEF": "SIAF (MEF)", "Información presupuestal de la DGIGND": "SIAF (MEF)",
    "JNE": "organismos electorales", "ONPE": "organismos electorales", "JNE, Perfil Electoral N.° 4": "organismos electorales",
    "instrumentos de planeamiento nacional": "instrumentos de planeamiento nacional",
    "normas de aprobación de la PGG": "instrumentos de planeamiento nacional",
    "INEI, Cuenta Satélite de la Economía Informal 2022–2024": "ENAHO",
    "fichas técnicas": "fichas técnicas", "actas de validación de fichas": "fichas técnicas",
    "protocolos aprobados": "protocolos", "protocolos de los servicios": "protocolos",
    # informes técnicos del 9-11-2026 y su matriz: un solo productor (ADR-016)
    "Informes técnicos de la DGIGND": "informes técnicos de la DGIGND (2026)",
    "Matriz de incompatibilidades de la DGIGND": "informes técnicos de la DGIGND (2026)",
}
# entrevista que no cuenta como fuente independiente: la OMEP, que conduce la evaluación, es parte de la OGMEPGD
NO_INDEPENDIENTE = "Entrevista al nivel conductor (OGMEPGD)"
# términos internos que no pueden aparecer en el informe publicable (ADR-010)
INTERNOS = re.compile(r"\b(?:REV_DOC|ENTREV|REP_ENT|REV_ALI|REV_IND|REV_EVI)\b|\bENTREV-|\bDED-\d|\bADR-\d|"
                      r"\b(?:LIO|FORM|CONAIG)-20\d\d-|\bEV20\d\d-\d|\bCEPLAN-\d\d\b|\bP\d\.\d+\b|\bPS-[A-Z]|"
                      r"decisión del usuario|nota metodológica|svy\.|SHA-256|repositorio de la evaluación|"
                      r"carpeta administrativa|libros? de análisis|matriz de triangulación|\.xlsx\b")


def _partes(base):
    if base in PARTES:
        return PARTES[base]
    return [base[0].lower() + base[1:] if base.split(" ")[0] in _COMUNES else base]


def fuentes_legibles(h):
    """Fuentes primarias de un hallazgo en lenguaje llano y sin repetir, desde la «base» de cada evidencia; las
    entrevistas se agrupan en «entrevistas a la OGPPM, la OGMEPGD y la DGIGND»."""
    fuentes, organos = [], []
    for e in h["evidencias"]:
        m = re.match(r"Entrevista al nivel conductor \((\w+)\)", e["base"])
        if m:
            organos += [m.group(1)] if m.group(1) not in organos else []
            continue
        for b in _partes(e["base"]):
            fuentes += [b] if b not in fuentes else []
    if organos:
        fuentes.append(("entrevistas a " if len(organos) > 1 else "entrevista a ") +
                       lista_es([f"la {o}" for o in organos]))
    return fuentes


# nombres de los libros y de sus hojas con que el informe los cita y el anexo técnico los publica (decisión del usuario,
# 2026-11-07, ADR-010); el repositorio conserva los nombres con sigla (ADR-009)
LIBROS_PUBLICOS = {
    "D1_D2_brechas_situacion_futura.xlsx": "Brechas_situacion_futura.xlsx",
    "D1_D2_brechas_departamentales.xlsx": "Brechas_departamentales.xlsx",
    "D6_enfoques_transversales.xlsx": "Enfoques_transversales.xlsx",
    "REV_ALI_coherencia_alineamiento.xlsx": "Coherencia_alineamiento.xlsx",
    "REV_DOC_estado_implementacion_servicios.xlsx": "Estado_implementacion_servicios.xlsx",
    "REV_EVI_matriz_evidencia.xlsx": "Matriz_evidencia.xlsx",
    "REV_IND_base_indicadores.xlsx": "Base_indicadores.xlsx",
    "REV_IND_2_calidad_fichas.xlsx": "Calidad_fichas.xlsx",
    "REV_IND_3_recursos_conduccion.xlsx": "Recursos_conduccion.xlsx",
    "REV_IND_3_verificacion_presupuestal.xlsx": "Verificacion_presupuestal.xlsx",
    "REV_IND_3_presupuesto_dgignd.xlsx": "Presupuesto_DGIGND.xlsx",
    "REV_IND_4_recalculo_OP3.xlsx": "Recalculo_indicadores_OP3.xlsx",
    "REP_ENT_informacion_entidades.xlsx": "Informacion_entidades.xlsx",
    "REP_ENT_doble_codificacion.xlsx": "Doble_codificacion.xlsx",
    "seguimiento_recomendaciones_previas.xlsx": "Seguimiento_recomendaciones.xlsx",
    "matriz_triangulacion_preguntas.xlsx": "Matriz_triangulacion.xlsx",
    "REV_DOC_matriz_incompatibilidades.xlsx": "Contraste_matriz_incompatibilidades.xlsx",
    "ENTREV_registro_codificacion.xlsx": "Registro_entrevistas.xlsx",
    "validacion_informal_enaho.csv": "Validacion_empleo_informal_ENAHO.csv",
}
# nombres publicados de las hojas con siglas (ronda 36): 31 caracteres o menos, el límite de Excel, para que las citas
# del informe coincidan con las hojas de los libros del anexo técnico
HOJAS_PUBLICAS = {
    "REV_ALI.1 Por servicio": "Coherencia, por servicio",
    "REV_ALI.1 Por lineamiento": "Coherencia, por lineamiento",
    "REV_ALI.1 Por objetivo": "Coherencia, por objetivo",
    "REV_ALI.1 Modelo causal": "Coherencia, modelo causal",
    "REV_ALI.2 Por objetivo": "Vertical, por objetivo",
    "REV_ALI.2 Registro": "Vertical, registro",
    "REV_ALI.2 Declarado 2019": "Vertical, declarado 2019",
    "REV_ALI.2 Indicadores PEDN": "Vertical, indicadores PEDN",
    "REV_ALI.3 Horizontal": "Horizontal",
    "REV_ALI.3 Por servicio PNIG": "Horizontal, por servicio",
    "REV_ALI.3 Universo": "Horizontal, universo",
    "REV_ALI.4 Por objetivo": "Ley e internacional, por OP",
    "REV_ALI.4 Registro": "Ley e internacional, registro",
    "REV_ALI.5 Conceptual": "Correspondencia conceptual",
    "Resumen REV_ALI.5": "Resumen conceptual",
    "Factores REP_ENT": "Factores reportados",
    "Base REV_IND.1": "Base de indicadores",
}


def hoja_publica(hoja):
    """Nombre publicado de una hoja (nucleo.hallazgos.hoja_publica con la tabla del piloto)."""
    return _NH.hoja_publica(hoja, HOJAS_PUBLICAS)


def libro_publico(libro):
    return _NH.libro_publico(libro, LIBROS_PUBLICOS)


# documentos del corpus que las evidencias citan por su ruta
_RUTAS = {
    "docs/fuentes/inei_metodologia/2025_inei_produccion-empleo-informal-cuenta-satelite-2022-2024.pdf":
        "INEI, Producción y Empleo Informal en el Perú: Cuenta Satélite de la Economía Informal 2022–2024 (2025)",
    "docs/fuentes/pnig/09_conaig_gobernanza/2026_mimp-conaig_ayuda-memoria-sesiones-2021-2026.doc":
        "Ayuda memoria de las sesiones de la CONAIG, 2021–2026 (MIMP y CONAIG, 2026)",
    "docs/fuentes/pnig/07_transicion_ley_32535/2026_mimp-ogaj_informe-d000926-2026_vigencia-normas-ley-28983.pdf":
        "Informe N.° D000926-2026-MIMP-OGAJ (MIMP, 2026e)",
    "docs/fuentes/pnig/07_transicion_ley_32535/2026_mimp-dgignd_informe-d000041-2026_sustento-vigencia-normas-ley-28983.pdf":
        "Informe N.° D000041-2026-MIMP-DGIGND (MIMP, 2026g)",
    "docs/fuentes/pnig/07_transicion_ley_32535/2026_mimp-dteg_informe-tecnico-d000040-2026_sustento-vigencia-normas-ley-28983.pdf":
        "Informe Técnico N.° D000040-2026-MIMP-DTEG (MIMP, 2026i)",
    "docs/fuentes/pnig/07_transicion_ley_32535/2026_mimp-dpignd_matriz-incompatibilidades-ley-32535.xlsx":
        "Matriz técnica de incompatibilidades de la DPIGND, anexa al Informe N.° D000041-2026-MIMP-DGIGND "
        "(MIMP, 2026g)",
}


def _respaldo_pieza(t):
    t = t.strip()
    if not t or t.startswith(("cálculo", "avance y nivel")):
        return None
    if t in _RUTAS:
        return _RUTAS[t]
    m = re.match(r"(?:evaluaciones/[\w/]+/)?([\w\-]+\.(?:xlsx|csv))(?:, hojas? (.+))?$", t)
    if m:
        hojas = re.findall(r"«([^»]+)»", m.group(2) or "")
        libro = libro_publico(m.group(1))
        if not hojas:
            return libro
        return f"{libro}, {'hoja' if len(hojas) == 1 else 'hojas'} " + lista_es([f"«{hoja_publica(x)}»" for x in hojas])
    t = re.sub(r"\s*\[[^\]]+\]", "", t)
    t = re.sub(r"^Informe LIO (\d{4})", r"Informe de cumplimiento de la Ley N.° 28983, \1", t)
    t = re.sub(r"^Formulario (\d{4})", r"Formulario de seguimiento \1", t)
    assert "docs/" not in t and "evaluaciones/" not in t, f"ruta sin nombre legible en el respaldo: {t}"
    return t


def _agrupar(piezas):
    """«Libro, hoja «A»» y «Libro, hoja «B»» -> «Libro, hojas «A» y «B»»; «Informe, p. 171» y «Informe, p. 172» ->
    «Informe, pp. 171 y 172»."""
    orden, locs = [], {}
    for p in piezas:
        m = re.match(r"^(.+?), (?:hojas? (.+)|pp?\. (.+))$", p)
        doc, loc, tipo = (m.group(1), m.group(2) or m.group(3), "hoja" if m.group(2) else "p") if m else (p, None, None)
        clave = (doc, tipo)
        if clave not in locs:
            orden.append(clave)
            locs[clave] = []
        if loc:
            for x in (re.findall(r"«[^»]+»", loc) if tipo == "hoja" else re.split(r", | y ", loc)):
                locs[clave] += [x] if x not in locs[clave] else []
    salida = []
    for doc, tipo in orden:
        ls = locs[(doc, tipo)]
        if not ls:
            salida.append(doc)
        elif tipo == "hoja":
            salida.append(f"{doc}, {'hoja' if len(ls) == 1 else 'hojas'} {lista_es(ls)}")
        else:
            salida.append(f"{doc}, {'p.' if len(ls) == 1 and ' a ' not in ls[0] else 'pp.'} {lista_es(ls)}")
    return salida


def respaldo(h):
    """Respaldo legible de un hallazgo (Anexo 2): documento con su página, libro con su hoja y entrevista con su fecha;
    sin rutas ni identificadores de segmento."""
    piezas, organos, deduce = [], [], False
    for e in h["evidencias"]:
        if e["instrumento"] == "ENTREV":
            inf = re.match(r"Entrevista al nivel conductor \((\w+)\)", e["base"]).group(1)
            organos += [inf] if inf not in organos else []
            deduce = deduce or "deducción" in e["fuente"]
            continue
        for t in str(e["fuente"]).split(";"):
            p = _respaldo_pieza(t)
            if p and p not in piezas:
                piezas.append(p)
    piezas = _agrupar(piezas)
    piezas += [f"Entrevista a la {o}, {FECHA_ENTREV[o]}" for o in organos]
    if deduce:
        piezas.append("deducción a partir de las entrevistas (Anexo 8)")
    return "; ".join(piezas)


def rastro(h):
    """Rastro técnico de un hallazgo: libro y hoja, página o segmento de cada evidencia (Anexo 1 del informe)."""
    return "; ".join(dict.fromkeys(corto(e["fuente"]) for e in h["evidencias"]))


def pl(n, singular, plural):
    """Concordancia de número, sin el número: pl(1, 'no mejora', 'no mejoran') -> 'no mejora'. El número se escribe
    aparte en la frase (f"{n} {pl(n, ...)}"), para que no se repita."""
    return singular if n == 1 else plural


def seg(C, ident):
    """Cita de un segmento de REP_ENT: «Informe LIO 2025, p. 171» o «Formulario 2025»."""
    r = C["fie_seg"].loc[ident]
    pag = _num(pd.Series([r["pagina"]])).iloc[0]
    return f"{r['fuente']}" + (f", p. {int(pag)}" if pd.notna(pag) else "") + f" [{ident}]"


def texto_seg(C, ident):
    return str(C["fie_seg"].loc[ident, "texto"]).strip()


def _ev(plano, instrumento, texto, fuente, base, n_fuentes=1, origen=None):
    """n_fuentes > 1 cuando la evidencia agrega varias fuentes primarias independientes (p. ej., la matriz REV_EVI);
    origen: origen primario cuando no se deduce de la base."""
    return _NH.evidencia(plano, instrumento, texto, fuente, base, n_fuentes, origen)


def ges(C, texto, *ids, ded=None):
    """Evidencias del nivel conductor (ENTREV): una por órgano, porque cada entrevista es una fuente independiente.
    Los segmentos deben existir en la codificación; ded cita la deducción del equipo evaluador que se incorpora."""
    g = C["entrev"]
    por = {}
    for i in ids:
        assert i in g.index, f"segmento ENTREV inexistente: {i}"
        por.setdefault(g.loc[i, "informante"], []).append(i)
    assert ded is None or ded in C["entrev_ded"].index, f"deducción inexistente: {ded}"
    return [_ev("Nivel conductor", "ENTREV", texto, ", ".join(ss) + (f"; deducción {ded}" if ded else ""),
                f"Entrevista al nivel conductor ({inf})") for inf, ss in por.items()]


# Datos que los hallazgos toman de segmentos de REP_ENT: cada frase debe constar en el texto literal del segmento.
CONSTA = {
    "LIO-2024-6.4.1-07": ["3 sesiones ordinarias", "3 titulares y 11 representantes alternas/os"],
    "LIO-2025-6.3.3-11": ["tres sesiones ordinarias", "57,1 %", "42,9 %", "90,5 %"],
    "LIO-2025-6.3.3-14": ["12 medidas urgentes", "4 cumplieron con la meta hasta septiembre de 2025",
                       "6 medidas aún se encuentran en proceso", "2 medidas con limitaciones"],
    "CONAIG-2024-040": ["16 servicios operativos", "solo 2 servicios"],
    "CONAIG-2023-049": ["ninguna de las entidades realizó o proyectó realizar actividades específicas"],
    "LIO-2025-6.3.3-21": ["38,7 % (19)", "30,6 % (15)", "OP 2 no reportaron articulación nacional"],
    "LIO-2021-5.5.1-03": ["459 personas", "solo 139 cuentan con formación en género", "30%"],
    "LIO-2023-5.3.1-03": ["dificultades para el cumplimiento de la meta", "rotación de personal"],
    "LIO-2025-5.3.1-04": ["no logró alcanzar la meta", "alta rotación de personal"],
    "LIO-2025-5.1.1-01": ["de 21 a 10 entidades", "rotación de personal"],
    "LIO-2025-5.1.1-03": ["Proyecto de Ley N°8731", "asistencias virtuales", "presupuesto adicional"],
    "LIO-2025-6.3.3-02": ["a la espera de las acciones", "Ley N°32535"],
    "CONAIG-2024-035": ["limitaciones presupuestales", "formulario virtual de seguimiento"],
    "FORM-2025-IOP.06.01-contexto": ["ENARES con periodicidad bienal"],
    "FORM-2025-IOP.05.02-falla_abajo": ["planilla electrónica", "ENAHO"],
    "LIO-2023-5.5.1-05": ["no se realiza desde 2020"],
    "FORM-2025-IOP.04.01-causa_dificultad": ["Inexistencia de convenios o protocolos de articulación",
                                          "Duplicidad de funciones"],
    "FORM-2025-ISS.04.02.02.01-falla_abajo": ["evaluar el retiro del servicio de la PNIG"],
    "CONAIG-2024-042": ["no le resulta aplicable"],
    "CONAIG-2023-022": ["sin diferencias basadas en sexo"],
    "LIO-2019-4.4.1-03": ["forma indirecta"],
    "FORM-2025-ISS.04.03.05.01-efectos_no_previstos": ["sobrecarga laboral", "durante la noche"],
    "CONAIG-2025-030": ["horarios poco compatibles", "dificultades de traslado"],
    "CONAIG-2025-024": ["limitaciones presupuestales para contratar personal"],
    "LIO-2022-6.4.1-01": ["etapa de diseño"],
    "LIO-2025-6.4.1-01": ["aún no cuenta con protocolo"],
}


def verificar_citas(C):
    faltan = [(i, f) for i, frases in CONSTA.items() for f in frases if f not in texto_seg(C, i)]
    assert not faltan, f"frases que no constan en el texto literal del segmento: {faltan}"


def hallazgos(C):
    """Hallazgos preliminares del equipo evaluador. Plano: equipo evaluador (propuesta por validar)."""
    n = C["serv_total"]
    vp, vt = C["vert_periodo"], C["vert_transicion"]
    m5 = C["mac5"]
    dif = C["fie_dif_top"]
    dev = C["cond_dev"]
    cs = C["cond_suma"]
    vd = C["act5001072"]
    s25 = C["serv25"]
    lim = C["rev_lim"]
    H = []

    def add(ident, preg, aspecto, enunciado, evid, confianza, estado=VALIDADO, nota="", detalle=None, salvedad=""):
        """detalle: resumen de 1 a 3 oraciones con la cifra clave (Anexo 2); si falta, el enunciado, que entonces debe
        ser breve. salvedad: matiz en lenguaje llano que va en el cuerpo. nota: precisión técnica para el Anexo 2."""
        H.append({"id": ident, "pregunta": preg, "aspecto": aspecto, "enunciado": enunciado, "evidencias": evid,
                  "confianza": confianza, "estado": estado, "nota": nota, "salvedad": salvedad,
                  "detalle": detalle or enunciado})

    # D1
    br = C["brechas"].set_index("Clave")
    sr = C["brechas_serie"]

    def ult(clave, col):
        g = sr[sr.indicador == clave]
        return float(g.loc[g.anio.idxmax(), col])

    def prim(clave, col):
        g = sr[sr.indicador == clave]
        return float(g.loc[g.anio.idxmin(), col])
    add("D1-H1", "D1", "Evolución de las brechas entre mujeres y hombres",
        f"Las brechas entre mujeres y hombres que describe el diagnóstico siguen presentes al cierre: las "
        f"{C['brechas_n']} brechas medidas persisten en el último año disponible. De las {C['brechas_n'] - 1} que se "
        f"comparan desde 2019 con el mismo instrumento, {len(C['brechas_reduce'])} se reducen "
        f"({lista_es(C['brechas_reduce'])}) y {len(C['brechas_sincambio'])} no cambian de forma significativa "
        f"({lista_es(C['brechas_sincambio'])}); el trabajo no remunerado, que la ENUT solo permite comparar entre 2010 y "
        f"2024 y con un cambio de instrumento, se amplía en una comparación indicativa. En 2025, la tasa de actividad es de "
        f"{fmt_num(ult('tasa_actividad', 'Mujeres'), 1)} % en las mujeres y {fmt_num(ult('tasa_actividad', 'Hombres'), 1)} "
        f"% en los hombres, y la población sin ingresos propios, de {fmt_num(ult('sin_ingresos', 'Mujeres'), 1)} % y "
        f"{fmt_num(ult('sin_ingresos', 'Hombres'), 1)} %; en 2024, las mujeres dedican "
        f"{fmt_num(ult('horas_no_remunerado', 'Mujeres'), 1)} horas semanales al trabajo no remunerado y los hombres "
        f"{fmt_num(ult('horas_no_remunerado', 'Hombres'), 1)}. Entre los productores agropecuarios, las mujeres "
        f"conducen el {fmt_num(ult('productora_mujer', 'Mujeres'), 1)} % de las unidades en 2025, solicitan menos "
        f"crédito ({fmt_num(ult('credito_solicito', 'Mujeres'), 1)} % frente a "
        f"{fmt_num(ult('credito_solicito', 'Hombres'), 1)} %) y trabajan parcelas más pequeñas (mediana de "
        f"{fmt_num(ult('superficie_mediana_ha', 'Mujeres'), 1)} ha frente a "
        f"{fmt_num(ult('superficie_mediana_ha', 'Hombres'), 1)} ha; la razón de medianas va de "
        f"{fmt_num(prim('superficie_mediana_ha', 'Razón M/H'), 2)} en 2023 a "
        f"{fmt_num(ult('superficie_mediana_ha', 'Razón M/H'), 2)} en 2025). En cambio, son propietarias de al menos una "
        f"parcela con más frecuencia que los hombres ({fmt_num(ult('propietario', 'Mujeres'), 1)} % frente a "
        f"{fmt_num(ult('propietario', 'Hombres'), 1)} %): la brecha de tierra aparece en la superficie que trabajan, "
        f"no en la tenencia de al menos una parcela; los datos no informan sobre titulación, copropiedad ni control de "
        f"la tierra.",
        [_ev("Cuantitativa", "REV_DOC", "Brechas ENAHO 2016–2025 con diseño muestral complejo; las definiciones "
             "reproducen el diagnóstico y la serie oficial del INEI.", C.origen["brechas"],
             "Encuesta Nacional de Hogares (ENAHO)"),
         _ev("Cuantitativa", "REV_DOC", "Trabajo no remunerado, ENUT 2010 y 2024.", C.origen["brechas"],
             "Encuesta Nacional de Uso del Tiempo (ENUT)"),
         _ev("Cuantitativa", "REV_DOC", "Tolerancia social por sexo, ENARES 2019 y 2024.", C.origen["brechas"],
             "Encuesta Nacional sobre Relaciones Sociales (ENARES)"),
         _ev("Cuantitativa", "REV_DOC", "Empleo informal por sexo, EPEN 2023–2025; reproduce las cifras del INEI.",
             C.origen["brechas"], "Encuesta Permanente de Empleo Nacional (EPEN)"),
         _ev("Cuantitativa", "REV_DOC", "Productoras agropecuarias, crédito, propiedad y superficie, ENA 2019–2025; "
             "reproduce las cifras del INEI de 2024.", C.origen["brechas"], "Encuesta Nacional Agropecuaria (ENA)"),
         _ev("Documental", "REV_DOC", "Brechas citadas en el diagnóstico de la PNIG.", "Documento de la PNIG, pp. 17 a "
             "26", "Documento de la PNIG")],
        "Sólido",
        detalle=f"Las {C['brechas_n']} brechas del diagnóstico persisten al cierre: {len(C['brechas_reduce'])} se "
                f"reducen desde 2019 con el mismo instrumento, {len(C['brechas_sincambio'])} no cambian de forma significativa y "
                f"el trabajo no remunerado se amplía ({lista_es(C['brechas_amplia'])}, en una "
                f"comparación indicativa). En 2025, la tasa de actividad es de "
                f"{fmt_num(ult('tasa_actividad', 'Mujeres'), 1)} % en las mujeres y "
                f"{fmt_num(ult('tasa_actividad', 'Hombres'), 1)} % en los hombres.",
        salvedad="La comparación del trabajo no remunerado entre 2010 y 2024 es indicativa, porque la ENUT cambió de "
                 "instrumento entre ambas rondas.")
    add("D1-H4", "D1", "Correspondencia con el nuevo marco normativo y conceptual",
        f"El enunciado del problema público de la PNIG (discriminación estructural contra las mujeres) diverge del "
        f"marco conceptual de igualdad de oportunidades y se clasifica como «se rediseña»; en el plano del problema "
        f"público, {m5.loc['Problema público', 'Divergencia']} de "
        f"{int(m5.loc['Problema público', ['Continuidad', 'Precisión', 'Divergencia']].sum())} elementos divergen y "
        f"{m5.loc['Problema público', 'Precisión']} requieren precisión. La OGAJ concluye que el problema público "
        f"de la PNIG, explicado desde patrones socioculturales y relaciones de género (numeral 1.4.2), y su finalidad, "
        f"la igualdad de género (numeral 1.4.3), son incompatibles con el artículo 6.1 de la Ley N.° 32535 "
        f"(MIMP, 2026e). Llega a esa conclusión sobre la base de los informes técnicos de la DGIGND, que sostienen que "
        f"la cadena causal de la Política deriva de ese problema público, y de su matriz de incompatibilidades, que "
        f"marca como incompatibles las {int(C['mic_norma'].iloc[0]['Disposiciones revisadas'])} disposiciones "
        f"revisadas del decreto que aprobó la PNIG y de su Anexo (MIMP, 2026g, numeral 2.11).",
        [_ev("Documental", "REV_ALI", "Correspondencia conceptual del problema público (elemento P01 y plano del "
             "problema).", C.origen["mac5"], "Documento de la PNIG y marco conceptual"),
         _ev("Documental", "REV_DOC", "Incompatibilidad del objeto de la PNIG con la Ley N.° 32535 (numerales 2.6.1 y "
             "3.2).", "docs/fuentes/pnig/07_transicion_ley_32535/2026_mimp-ogaj_informe-d000926-2026_vigencia-normas-ley-28983.pdf", "Informe de la OGAJ"),
         _ev("Documental", "REV_DOC", "Incompatibilidad del problema público y de la cadena causal de la PNIG "
             "(numeral 2.11) y disposiciones de la PNIG marcadas como incompatibles en la matriz.", "docs/fuentes/pnig/07_transicion_ley_32535/2026_mimp-dgignd_informe-d000041-2026_sustento-vigencia-normas-ley-28983.pdf",
             "Informes técnicos de la DGIGND")],
        "Sólido", nota="No se contrastó en las entrevistas: la pregunta sobre el problema público de la nueva "
        "política no se formuló.",
        detalle="El problema público de la PNIG diverge del marco de igualdad de oportunidades; la OGAJ considera "
                "incompatibles su modelo del problema y su finalidad con el artículo 6.1 de la Ley N.° 32535.")
    add("D1-H2", "D1", "Manifestaciones medidas por los indicadores de los OP",
        f"Las manifestaciones del problema que miden los indicadores de los OP mejoran sin desaparecer: "
        f"{C['op_mejora']} de {C['op_ind']} indicadores mejoran frente a su línea de base, "
        f"{len(C['op_nomejora'])} {pl(len(C['op_nomejora']), 'no mejora', 'no mejoran')} ({lista_es(C['op_nomejora'])}), "
        f"{len(C['op_sinlb'])} {pl(len(C['op_sinlb']), 'no tiene', 'no tienen')} línea de base numérica ({lista_es(C['op_sinlb'])}) y "
        f"{len(C['op_noint_lb'])} {pl(len(C['op_noint_lb']), 'tiene', 'tienen')} un último valor no interpretable "
        f"({lista_es(C['op_noint_lb'])}).",
        [_ev("Cuantitativa", "REV_IND", "Último valor disponible de cada indicador de OP frente a su línea de base.",
             C.origen["op_tabla"], "Reportes de seguimiento (RS)"),
         _ev("Cuantitativa", "REV_DOC", "Serie ENDES 2019–2025 de los indicadores de los OP1 y OP2.",
             C.origen["endes"], "Encuesta Demográfica y de Salud Familiar (ENDES)", origen=_REPORTE_OFICIAL)],
        "Indicio")
    inf = br.loc["informal"]
    vi = C["informal_validacion"]
    add("D1-H3", "D1", "Cambios en las fuentes de seguimiento",
        f"Dos fuentes del diagnóstico cambiaron durante la vigencia de la Política. Desde 2024 la ENAHO no publica la "
        f"variable de empleo informal, porque el INEI pasó esa medición a la EPEN; la regla del INEI para la ENAHO "
        f"(Cuenta Satélite de la Economía Informal 2022–2024, pp. 87–90) reproduce la variable de 2019 a 2023 con una "
        f"concordancia individual de {fmt_num(vi.concordancia.min(), 2)} % o más y tasas por sexo a "
        f"{fmt_num(vi.diferencia.abs().max(), 2)} puntos o menos, de modo que la brecha se sigue con la misma encuesta: "
        f"pasa de {fmt_num(inf['Brecha inicial'], 1)} puntos en 2019 a {fmt_num(inf['Brecha final'], 1)} en 2025, "
        f"{'sin cambio significativo' if inf['Clasificación'] == 'Sin cambio significativo' else inf['Clasificación'].lower()}"
        f". La ENARES cambió las afirmaciones de tolerancia entre 2015, 2019 y 2024, de modo que el índice de IOP.06.01 "
        f"no puede recalcularse con los microdatos.",
        [_ev("Cuantitativa", "REV_DOC", "Variable ocupinf ausente en los módulos 500 de 2024 y 2025; regla del INEI "
             "validada contra ocupinf en 2019–2023.", C.origen["informal_validacion"],
             "Encuesta Nacional de Hogares (ENAHO)"),
         _ev("Documental", "REV_DOC", "Regla del empleo informal con la ENAHO (anexo A2.2, pp. 87–90).",
             "docs/fuentes/inei_metodologia/2025_inei_produccion-empleo-informal-cuenta-satelite-2022-2024.pdf",
             "INEI, Cuenta Satélite de la Economía Informal 2022–2024"),
         _ev("Cuantitativa", "REV_DOC", "Cuestionarios CRS.02 de 2015, 2019 y 2024; validación sin coincidencia.",
             C.origen["d1_validacion"], "Encuesta Nacional sobre Relaciones Sociales (ENARES)")],
        "Sólido",
        detalle=f"Dos fuentes del diagnóstico cambiaron: la ENAHO dejó de publicar el empleo informal en 2024 (se "
                f"reconstruyó con la regla del INEI; la brecha pasa de {fmt_num(inf['Brecha inicial'], 1)} puntos en "
                f"2019 a {fmt_num(inf['Brecha final'], 1)} en 2025) y la ENARES cambió las preguntas de tolerancia, por lo "
                f"que el índice del OP6 no puede recalcularse.",
        salvedad="La EPEN 2023–2025 sirve de contraste: mide niveles de empleo informal unos 3 puntos menores que la "
                 "ENAHO y una brecha similar.")

    dr = C["dep_resumen"].set_index("Clave")
    dc = C["dep_cambio"]

    def desv(clave):
        return int(dr.loc[clave, "Brecha significativa en desventaja de las mujeres"])

    def deps(clave, clase):
        return dc.loc[(dc.indicador == clave) & (dc.clase == clase), "departamento"].tolist()

    def referenciales(clave):
        return int(dr.loc[clave, "Ámbitos con estimación (último año)"] - dr.loc[clave, "Confiables (último año)"])

    def n_amb(clave):
        return int(dr.loc[clave, "Ámbitos con estimación (último año)"])

    dpp = C["dep_participacion"]
    alc = dpp[dpp.cargo.str.startswith("Alcaldías")]
    reg = dpp[dpp.cargo == "Regidurías"]
    cong = dpp[dpp.cargo.str.startswith("Congreso")]
    ok, tot = C["dep_validacion"]
    mejora_v = deps("violencia_fisica_o_sexual_12m", "Mejora")
    cambio_act = len(deps("tasa_actividad", "Se amplía")) + len(deps("tasa_actividad", "Se reduce"))
    add("D1-H5", "D1", "Brechas por departamento",
        f"Las brechas se presentan en casi todo el territorio y su reducción se concentra en pocos departamentos. En 2025, la "
        f"brecha en desventaja de las mujeres es significativa en {desv('tasa_actividad')} de "
        f"{n_amb('tasa_actividad')} departamentos (incluidos el Callao, Lima Metropolitana y Región Lima) en la tasa de "
        f"actividad, en {desv('ingreso_laboral')} en el ingreso laboral y en {desv('sin_ingresos')} en la población "
        f"sin ingresos propios. Entre 2019 y 2025, la brecha de población sin ingresos propios se redujo de forma "
        f"significativa en {len(deps('sin_ingresos', 'Se reduce'))} departamentos "
        f"({lista_es(deps('sin_ingresos', 'Se reduce'))}) y la de uso de internet en "
        f"{len(deps('internet', 'Se reduce'))}; en los demás no cambió de forma significativa. La brecha en la tasa de "
        f"actividad cambió solo en {cambio_act} {pl(cambio_act, 'departamento', 'departamentos')}, un conteo compatible con el azar (1,3 "
        f"esperados en 26 pruebas al 5 %). Con la ENDES, la violencia de pareja, el embarazo adolescente y la demanda "
        f"insatisfecha de planificación familiar no tienen estimaciones confiables por departamento en 2025 (coeficiente "
        f"de variación mayor a 15 % en {referenciales('violencia_fisica_o_sexual_12m')}, "
        f"{referenciales('embarazo_adolescente')} y {referenciales('demanda_insatisfecha')} de 26 departamentos); con esa "
        f"salvedad, la violencia de pareja disminuye de forma significativa entre 2019 y 2025 en {len(mejora_v)} departamentos. "
        f"En participación política, los departamentos sin ninguna alcaldesa son {int((alc['2019/2022'] == 0).sum())} en "
        f"2019–2022, {int((alc['2023/2026'] == 0).sum())} en 2023–2026 y {int((alc['2027/2030'] == 0).sum())} en la "
        f"composición elegida para 2027–2030; las regidoras superan el {fmt_num(reg['2023/2026'].min(), 0)} % en todos "
        f"los departamentos desde 2023, y la Cámara de Diputados 2026–2031 tiene al menos una diputada en cada "
        f"departamento, frente "
        f"a {int((cong['2021/2026'] == 0).sum())} departamentos sin congresista mujer en 2021–2026.",
        [_ev("Cuantitativa", "REV_DOC", f"Brechas por departamento con la ENAHO, la EPEN, la ENA, la ENDES y la ENARES 2024; "
             f"{ok} de {tot} cifras del INEI de 2024 por departamento dentro del IC 95 %.",
             C.origen["dep_resumen"], "Microdatos del INEI por departamento"),
         _ev("Documental", "REV_DOC", "Autoridades electas por departamento, cuadros 4.2 a 4.4 del anexo del INEI.",
             C.origen["dep_participacion"], "Jurado Nacional de Elecciones (JNE) y ONPE"),
         _ev("Documental", "REV_DOC", "Diputados y senadores 2026–2031 por circunscripción (Resoluciones N.° "
             "1411-2026-JNE y 1412-2026-JNE), alcaldías y gobiernos regionales 2027–2030 (resultados preliminares de "
             "la ONPE), con sexo por nombre.", C.origen["dep_participacion"], "Resultados de 2026 (JNE y ONPE)")],
        "Sólido",
        detalle=f"En 2025 la brecha en desventaja de las mujeres es significativa en {desv('tasa_actividad')} de "
                f"{n_amb('tasa_actividad')} departamentos en la tasa de actividad y en {desv('ingreso_laboral')} en el "
                f"ingreso laboral. Desde 2019, la brecha de población sin ingresos propios solo se redujo en "
                f"{len(deps('sin_ingresos', 'Se reduce'))} departamentos.",
        nota="Las autoridades de 2027–2030 provienen de resultados preliminares de la ONPE, pendientes de proclamación "
             "del JNE.")

    # D2
    add("D2-H2", "D2", "Trayectoria de los indicadores frente a sus metas",
        f"Con el reporte oficial, la situación futura deseada no puede valorarse para la mayoría de los OP con datos "
        f"de 2025: {C['op_sinvalor_2025']} de {C['op_ind']} indicadores de OP no tienen valor 2025 y "
        f"{C['op_noint_2025']} no permite calcular el avance. Con las estimaciones y los recálculos de esta evaluación "
        f"(I2-H3), {len(C['op_est25'])} de esos indicadores se comparan con su meta 2025 y "
        f"{queda_sin_valor(C)}. Con los datos de 2024, {C['op_alto_2024']} indicadores están en "
        f"nivel alto, {C['op_medio_2024']} en medio, {C['op_bajo_2024']} en bajo y {C['op_sinvalor_2024']} sin valor. "
        f"Las metas de la ficha llegan a 2025; la trayectoria hasta el horizonte de 2030 no se ha proyectado en las fichas oficiales, aunque esta "
        f"evaluación presenta escenarios referenciales (numeral 5.4).",
        [_ev("Cuantitativa", "REV_IND", "Semaforización anual de los indicadores de OP.",
             C.origen["resumen_semaforo"], "Reportes de seguimiento (RS)"),
         _ev("Cuantitativa", "REV_IND", "Contraste de valores y metas con el consolidado de la OMEP y el formulario.",
             _hoja(PSI, "Contraste de fuentes"), "Consolidado de indicadores de la OMEP")],
        "Indicio",
        detalle=f"Con el reporte oficial, {C['op_sinvalor_2025']} de {C['op_ind']} indicadores de OP no tienen valor "
                f"2025, de modo que la situación futura no puede valorarse para la mayoría de los OP; con las "
                f"estimaciones de esta evaluación, {queda_sin_valor(C)}.")
    add("D2-H3", "D2", "Calidad de las metas",
        f"Las metas al fin del horizonte tienen problemas de estimación: de {C['fichas_total']} fichas, "
        f"{C['ficha_meta_nc']} no fijan meta anual hasta el fin del horizonte y {C['ficha_lb_nc']} no tienen línea de "
        f"base con año y valor; en 2025, "
        f"{C['serv25_falla']} de {C['serv25_total']} indicadores de servicio muestran una posible falla de planeación "
        f"(avance de 150 % o más, o menor de 15 %). Las entidades atribuyen las desviaciones a metas mal estimadas "
        f"({C['fie_cod_total'].get('D5.META', 0)} segmentos codificados).",
        [_ev("Documental", "REV_IND", "Atributos de las fichas técnicas.", C.origen["ficha_lb_nc"], "Fichas técnicas"),
         _ev("Cuantitativa", "REV_IND", "Semaforización 2025 de los indicadores de servicio.",
             C.origen["serv25"], "Reportes de seguimiento (RS)"),
         _ev("Entidades", "REP_ENT", "Código D5.META en el formulario y los informes LIO.",
             C.origen["fie_cod_total"], "Formulario de seguimiento e informes LIO")],
        "Sólido",
        detalle=f"De {C['fichas_total']} fichas, {C['ficha_meta_nc']} no fijan meta anual hasta el fin del horizonte y "
                f"{C['ficha_lb_nc']} no tienen línea de base con año y valor; en 2025, {C['serv25_falla']} de "
                f"{C['serv25_total']} indicadores de servicio muestran una posible falla de planeación.")
    sfd = C["sfd"].set_index("Código")
    # SFD4 enuncia dos metas (congresistas y alcaldesas) en un solo enunciado
    corr = C["sfd"].drop_duplicates("Enunciado (cita literal)")["Correspondencia con el indicador del OP"].value_counts()
    pp = C["participacion"]
    alc = pp[pp.cargo.str.startswith("Alcaldías") & pp.porcentaje_mujeres.notna()].porcentaje_mujeres
    ad = C["dep_participacion"]
    ad = ad[ad.cargo.str.startswith("Alcaldías")].set_index("departamento")
    sobre_meta = ad[ad["2027/2030"] >= 20]["2027/2030"]
    ppn = C["participacion"]
    gob = ppn[ppn.cargo == "Gobiernos regionales: gobernadoras"].set_index("periodo").mujeres.astype(int)
    camara = ppn[(ppn.cargo == "Cámara de Diputados") & (ppn.periodo == "2026/2031")].iloc[0]
    senado = ppn[(ppn.cargo == "Senado") & (ppn.periodo == "2026/2031")].iloc[0]
    add("D2-H1", "D2", "Correspondencia y alcanzabilidad de la situación futura deseada",
        f"La situación futura deseada cubre las tres causas y los cuatro efectos del modelo, pero se mide con "
        f"indicadores que no siempre son los de los OP: de {corr.sum()} enunciados, {corr.get('Misma', 0)} usan el mismo "
        f"indicador del OP, {corr.get('Aproximada', 0)} uno aproximado, {corr.get('Distinta', 0)} uno distinto y "
        f"{corr.get('Sin indicador', 0)} ninguno. Al ritmo observado, solo el embarazo adolescente alcanzaría su meta "
        f"al 2030 ({fmt_num(sfd.loc['SFD7', 'Último valor'], 1)} % en 2025 frente a 7,2 %); la tolerancia social "
        f"(proyección de {fmt_num(sfd.loc['SFD1', 'Valor 2030'], 1)} % frente a 36,8 %) y la violencia "
        f"de pareja ({fmt_num(sfd.loc['SFD3', 'Valor 2030'], 1)} % frente a 2,4 %) no la alcanzarían. La meta de "
        f"congresistas se alcanzó en el periodo 2021–2026 ({fmt_num(sfd.loc['SFD4a', 'Último valor'], 1)} %), pero el "
        f"Congreso elegido en 2026, que estará en funciones en 2030, tiene "
        f"{fmt_num(sfd.loc['SFD4a', 'Valor 2030'], 1)} % de mujeres (70 de 190, según la ONPE): "
        f"{fmt_num(camara.porcentaje_mujeres, 1)} % en la Cámara de Diputados y "
        f"{fmt_num(senado.porcentaje_mujeres, 1)} % en el Senado. La de alcaldesas (20 %) no se alcanza: en las "
        f"elecciones de 2026 se eligió {fmt_num(sfd.loc['SFD4b', 'Valor 2030'], 1)} % de alcaldesas, menos que el "
        f"{fmt_num(sfd.loc['SFD4b', 'Último valor'], 1)} % de 2022, y solo "
        f"{lista_es([f'{k} ({fmt_num(v, 1)} %)' for k, v in sobre_meta.items()])} superan el 20 %, con pocos "
        f"municipios. En los gobiernos regionales, cargo sin meta en la situación futura deseada, las gobernadoras "
        f"fueron {gob['2015/2018']} de 25 en 2015–2018, {gob['2019/2022']} en 2019–2022 y {gob['2023/2026']} en "
        f"2023–2026, y serán {gob['2027/2030']} en 2027–2030: la del Cusco ganó en primera vuelta y las 15 "
        f"circunscripciones que van a la segunda elección del 6 de diciembre enfrentan a dos fórmulas encabezadas por "
        f"hombres. El nivel conductor coincide en que la situación futura era poco alcanzable: la OGPPM considera que la "
        f"PNIG abordaba dimensiones demasiado grandes para resolverlas en diez años y la DGIGND, que pretendió abarcar "
        f"todos los derechos de las mujeres; ambas proponen acotar el problema y los servicios a lo que es posible "
        f"lograr con los recursos disponibles.",
        [_ev("Documental", "REV_DOC", "Enunciados de la situación futura deseada, con cita literal verificada.",
             "Documento de la PNIG, numeral 1.4.5, p. 20", "Documento de la PNIG"),
         _ev("Cuantitativa", "REV_IND", "Valores oficiales de IOP.06.01 e IOP.05.01.", C.origen["op_tabla"],
             "Reportes de seguimiento (RS)"),
         _ev("Cuantitativa", "REV_DOC", "Valores 2025 de violencia de pareja y embarazo adolescente.",
             C.origen["sfd"], "Encuesta Demográfica y de Salud Familiar (ENDES)"),
         _ev("Documental", "REV_DOC", "Congresistas, alcaldesas y regidoras por periodo, 1956–2026.",
             C.origen["participacion"], "Jurado Nacional de Elecciones (JNE) y ONPE"),
         _ev("Documental", "REV_DOC", "Congreso 2026–2031 por sexo y cámara (Analítica Electoral n.° 3, Tabla 3, "
             "p. 22).", C.origen["participacion"], "Oficina Nacional de Procesos Electorales (ONPE)"),
         _ev("Documental", "REV_DOC", "Gobernadoras 2015–2026 (informes de cumplimiento de la LIO 2017, p. 33; 2018, "
             "p. 26; 2022, p. 103) y 2027–2030 (dos fórmulas más votadas por circunscripción, ONPE).",
             C.origen["participacion"], "MIMP (LIO), JNE y ONPE")]
        + ges(C, "Alcance de la PNIG frente a su horizonte y a sus recursos.", "ENTREV-OGPPM-026", "ENTREV-OGPPM-027",
              "ENTREV-DGIGND-036", "ENTREV-DGIGND-053"),
        "Sólido",
        detalle=f"De {corr.sum()} enunciados de la situación futura, solo {corr.get('Misma', 0)} se miden con el mismo "
                f"indicador del OP. Al ritmo observado, solo el embarazo adolescente alcanzaría su meta al 2030; la de "
                f"congresistas se alcanzó en 2021, pero el Congreso elegido en 2026 tiene "
                f"{fmt_num(sfd.loc['SFD4a', 'Valor 2030'], 1)} % de mujeres.",
        salvedad="La proyección al 2030 es lineal, con el ritmo observado, y no es un pronóstico.",
        nota="Las alcaldías y los gobiernos regionales 2027–2030 provienen de resultados preliminares de la ONPE, "
             "pendientes de proclamación del JNE.")

    # D3
    add("D3-H1", "D3", "Lineamientos sin servicio y servicios sin causa",
        f"La cadena causal está casi completa: {C['lin_total'] - len(C['lin_sin_serv'])} de {C['lin_total']} "
        f"lineamientos tienen al menos un servicio (el lineamiento {lista_es(C['lin_sin_serv'])} no tiene ninguno) y "
        f"{n - len(C['serv_sin_causa'])} de {n} servicios tienen una causa o un efecto del modelo consignado en la "
        f"matriz de consistencia (faltan {lista_es(C['serv_sin_causa'])}). La DGIGND matiza la solidez de esa "
        f"cadena: la prisa de la formulación no dejó tiempo para analizarla, y la evaluación de diseño del CEPLAN "
        f"recomendó revisar la teoría de cambio de los servicios.",
        [_ev("Documental", "REV_ALI", "Lineamientos y servicios frente al Anexo 5 de la PNIG.",
             C.origen["lin_total"], "Documento de la PNIG (Anexo 5)"),
         _ev("Documental", "REV_ALI", "Servicios del PEMIG por lineamiento.", C.origen["serv_sin_causa"], "PEMIG")]
        + ges(C, "Análisis de la cadena causal durante la formulación.", "ENTREV-DGIGND-050"),
        "Sólido", nota="La lectura de la DGIGND matiza el hallazgo, sin contradecir el conteo.",
        detalle=f"{C['lin_total'] - len(C['lin_sin_serv'])} de {C['lin_total']} lineamientos tienen servicio (falta el "
                f"{lista_es(C['lin_sin_serv'])}) y {n - len(C['serv_sin_causa'])} de {n} servicios tienen una causa o un "
                f"efecto del modelo asignado en la matriz de consistencia.")
    assert C["serv_sin_ind_vig"] == C["serv_sin_ind_omep"], "REV_ALI.1 y el consolidado difieren en servicios sin indicador"
    ind, cob, cal = C["pemig_ind"], C["pemig_cob"], C["pemig_cal"]
    amplio = cob + C["pemig_part"]
    # la cifra de la DGIGND («algo más del 60 %» solo de cobertura) queda entre la lectura estricta y la amplia
    assert 100 * cob / ind < 61 < 100 * amplio / ind
    add("D3-H2", "D3", "Servicios sin indicador",
        f"Más de la mitad de los servicios no puede seguirse con indicadores aprobados: {C['serv_sin_ind_vig']} de "
        f"{n} servicios no tienen protocolo aprobado ni, por tanto, indicador aprobado, aunque el PEMIG había "
        f"asignado {ind} indicadores a sus {C['pemig_serv']} servicios; y {C['serv_falta_cob']} servicios con "
        f"indicador carecen del de cobertura o del de calidad. Los indicadores del PEMIG miden sobre todo cobertura: "
        f"{cob} de {ind} ({pct(100 * cob / ind, 0)}) miden la provisión del servicio, {amplio} "
        f"({pct(100 * amplio / ind, 0)}) si se suman los que miden la participación de las mujeres entre los "
        f"beneficiarios, y solo {cal} ({pct(100 * cal / ind, 0)}) un atributo de calidad, en "
        f"{C['pemig_serv_cal']} servicios; la DGIGND estimó que algo más del 60 % solo mide cobertura, cifra que "
        f"queda dentro de ese rango. Según el nivel conductor, el PEMIG no tuvo seguimiento propio: la OGPPM lo "
        f"describe como «congelado» desde su aprobación, un informe de la OMEP concluyó que sus indicadores no eran "
        f"aptos para el seguimiento y, según la DGIGND, los dos indicadores que se crearon con el Plan nunca se "
        f"midieron.",
        [_ev("Documental", "REV_ALI", "Servicios sin indicador aprobado.", C.origen["serv_sin_ind_vig"],
             "Consolidado de indicadores de la OMEP"),
         _ev("Documental", "REV_IND", "Observaciones de conjunto sobre las fichas.", C.origen["serv_sin_ind_omep"],
             "Consolidado de indicadores de la OMEP"),
         _ev("Documental", "REV_IND", "Tipo de los indicadores de servicio del PEMIG (criterio del Recuadro 2 de la "
             "guía de indicadores del CEPLAN).", C.origen["pemig_ind"], "PEMIG (anexo 5)")]
        + ges(C, "Seguimiento e indicadores del PEMIG.", "ENTREV-OGPPM-010", "ENTREV-OGMEPGD-012", "ENTREV-OGMEPGD-013",
              "ENTREV-DGIGND-021", "ENTREV-DGIGND-023"),
        "Con respaldo parcial", estado=VALIDADO_0711,
        detalle=f"{C['serv_sin_ind_vig']} de {n} servicios no tienen indicador aprobado, porque no tienen protocolo "
                f"aprobado. De los {ind} indicadores que el PEMIG asignó a los servicios, {cob} ({pct(100 * cob / ind, 0)}) "
                f"miden cobertura y solo {cal} ({pct(100 * cal / ind, 0)}) un atributo de calidad.",
        salvedad="El informe de la OMEP sobre los indicadores del PEMIG que menciona el nivel conductor se solicitó, "
                 "pero no se obtuvo.",
        nota="Los indicadores del PEMIG se clasificaron por tipo con las definiciones de la guía de indicadores del "
             "CEPLAN (CEPLAN, 2024c, pp. 61 y 62).")
    add("D3-H3", "D3", "Solidez de la evidencia por tipo de servicio",
        f"La evidencia respalda la lógica causal de la mayoría de los servicios: de {C['rev_tipos']} tipos de "
        f"intervención, {C['rev_sólida']} tienen evidencia sólida, {C['rev_moderada']} moderada y "
        f"{C['rev_limitada']} limitada ({lista_es([t[1].lower() for t in lim])}), que comprenden "
        f"{C['rev_serv']['Limitada']} servicios. La dirección de los efectos es mixta en {C['rev_mixta']} tipos y la "
        f"evidencia mide un resultado solo en parte pertinente en {len(C['rev_parcial'])}. La DGIGND advierte que, "
        f"al decidir las alternativas de solución, se descartaron las más costosas y se adoptaron sustitutos mínimos.",
        [_ev("Literatura", "REV_EVI", f"Matriz por tipo de intervención ({C['rev_fuentes']} fuentes: revisiones "
             "sistemáticas y evaluaciones), con la regla de solidez del Anexo 6.", C.origen["rev_tipos"],
             "Literatura especializada", n_fuentes=C["rev_fuentes"]),
         _ev("Documental", "REV_ALI", "Causa o efecto que atiende cada servicio.", C.origen["serv_sin_causa"],
             "Documento de la PNIG (Anexo 5)")]
        + ges(C, "Selección de alternativas de solución por su costo.", "ENTREV-DGIGND-052"),
        "Sólido",
        detalle=f"De {C['rev_tipos']} tipos de intervención, {C['rev_sólida']} tienen evidencia sólida, "
                f"{C['rev_moderada']} moderada y {C['rev_limitada']} limitada; la dirección de los efectos es mixta en "
                f"{C['rev_mixta']} tipos.")

    # D4
    add("D4-H1", "D4", "Alineamiento vertical en el periodo",
        f"Durante el periodo, los OP se alinean con los instrumentos de planeamiento nacional: de "
        f"{sum(vp)} relaciones valoradas, {vp[0]} son directas, {vp[1]} indirectas y {vp[2]} sin correspondencia. "
        f"El alineamiento directo es mayor en {lista_es(C['vert_fuerte'])} y menor en {lista_es(C['vert_debil'])}.",
        [_ev("Documental", "REV_ALI", "Relaciones por OP con Políticas de Estado, Plan Bicentenario, PEDN, PGG y "
             "PESEM.", C.origen["vert_periodo"], "Instrumentos de planeamiento nacional"),
         _ev("Documental", "REV_ALI", "Alineamiento declarado por la PNIG en 2019.",
             _hoja(MAC, "REV_ALI.2 Declarado 2019"), "Documento de la PNIG")],
        "Sólido")
    add("D4-H2", "D4", "Alineamiento con los instrumentos de la transición",
        f"El alineamiento se debilita con los instrumentos de la transición: de {sum(vt)} relaciones, {vt[0]} son "
        f"directas, {vt[1]} indirectas y {vt[2]} sin correspondencia. La PGG 2026 (D.S. N.° 043-2026-PCM) no "
        f"corresponde a ningún OP ({C['pgg043_sin']} de 6) y la PGG 2026–2031 (D.S. N.° 127-2026-PCM) no "
        f"corresponde a {lista_es(C['pgg127_sin_op'])}.",
        [_ev("Documental", "REV_ALI", "Relaciones por OP con PGG 2026, PGG 2026–2031, PEDN actualizado y PESEM "
             "2024–2030.", C.origen["vert_transicion"], "Instrumentos de planeamiento nacional"),
         _ev("Documental", "REV_ALI", "Registro de citas por instrumento.", _hoja(MAC, "REV_ALI.2 Registro"),
             "Normas de aprobación de la PGG"),
         _ev("Documental", "REV_ALI", "Objetivos prioritarios de la PNIG que se contrastan.",
             _hoja(MAC, "REV_ALI.2 Por objetivo"), "Documento de la PNIG")],
        "Sólido")
    add("D4-H3", "D4", "Duplicidades y complementariedades con otras políticas",
        f"La articulación horizontal muestra duplicidades: en {C['hor_rel']} relaciones con {C['hor_pol']} políticas "
        f"nacionales afines hay {C['hor_dup']} duplicidades, {C['hor_comp']} complementariedades y {C['hor_vac']} "
        f"vacíos de atención. El servicio de los Centros Emergencia Mujer (1.1.3) figura en {C['hor_cem']} políticas "
        f"además de la PNIG.",
        [_ev("Documental", "REV_ALI", "Relaciones con políticas nacionales afines, con cita y página.",
             C.origen["hor_rel"], "Políticas nacionales afines"),
         _ev("Entidades", "REP_ENT", "El MIMP reporta duplicidad de funciones y falta de convenios de articulación.",
             seg(C, "FORM-2025-IOP.04.01-causa_dificultad"), "Formulario de seguimiento")],
        "Con respaldo parcial", nota="El formulario se refiere a la duplicidad de funciones entre entidades, no de "
        "servicios entre políticas: coincide en el sentido, no en el objeto.")
    add("D4-H4", "D4", "Correspondencia con compromisos internacionales y con la Ley N.° 32535",
        f"La correspondencia con los compromisos internacionales es alta: de {sum(C['intl'])} relaciones, "
        f"{C['intl'][0]} son directas, {C['intl'][1]} indirectas y {C['intl'][2]} sin correspondencia. Frente a la "
        f"Ley N.° 32535, la relación es directa para {lista_es(C['ley_dir_op'])} e indirecta para "
        f"{lista_es(C['ley_ind_op'])}.",
        [_ev("Documental", "REV_ALI", "CEDAW, observaciones finales 2022, Belém do Pará, Beijing, ODS 5 y Estrategia "
             "de Montevideo.", C.origen["intl"], "Instrumentos internacionales"),
         _ev("Documental", "REV_ALI", "Artículos de la Ley N.° 32535.", C.origen["ley_dir_op"], "Ley N.° 32535")],
        "Sólido")

    # D5
    add("D5-H1", "D5", "Atributos de las fichas técnicas",
        f"Las fichas técnicas no cumplen el estándar del SINAPLAN en atributos críticos: de {C['fichas_total']} "
        f"fichas, {C['ficha_lb_nc']} no tienen línea de base con año y valor (y {C['ficha_lb_parc']} la tienen "
        f"incompleta), en ninguna consta la desagregación por territorio y grupo ({C['ficha_desag_nc']}), "
        f"{C['ficha_sentido_obs']} tienen observado el sentido esperado, y las {C['op_fichas_total']} fichas de OP "
        f"quedan observadas.",
        [_ev("Documental", "REV_IND", "Matriz de atributos frente a la Guía de indicadores del CEPLAN.",
             C.origen["fichas_total"], "Fichas técnicas"),
         _ev("Documental", "REV_ALI", "Dictamen de las fichas por OP.", C.origen["op_fichas_obs"],
             "Documento de la PNIG (Anexo 5)")]
        + ges(C, "Formulación de los indicadores de los OP.", "ENTREV-OGMEPGD-019", "ENTREV-OGPPM-014"),
        "Sólido",
        salvedad="La OGMEPGD considera bien formulados la mayoría de los indicadores de los OP. La diferencia se "
                 "explica porque la revisión valora los atributos de la ficha (línea de base, desagregación y sentido) y "
                 "no el enunciado del indicador. La OGPPM agrega que los indicadores se formularon como descriptivos y "
                 "se aprobaron sin fichas técnicas.")
    add("D5-H2", "D5", "Modificaciones de fichas y trazabilidad del reporte",
        f"Las fichas se modificaron mediante actas de validación de 2022 y 2023, varias no ratificadas, y el reporte "
        f"tiene discrepancias: el panel del formulario registra {C['panel_disc']} discrepancias con las demás "
        f"fuentes de seguimiento. Además, el aplicativo de seguimiento que el CEPLAN puso en marcha en 2026 solo "
        f"considera los indicadores aprobados por decreto supremo, los de la PNIG y del PEMIG, y no los aprobados en "
        f"los protocolos; según la OGPPM, el MIMP no logró que se incorporaran.",
        [_ev("Documental", "REV_IND", "Observación de conjunto «Modificación de fichas».", _hoja(PSI2, "Conjunto"),
             "Actas de validación de fichas"),
         _ev("Cuantitativa", "REP_ENT", "Discrepancias del panel por indicador.", C.origen["panel_disc"],
             "Formulario de seguimiento")]
        + ges(C, "Indicadores que considera el aplicativo del CEPLAN.", "ENTREV-OGPPM-015"),
        "Con respaldo parcial", nota="Cada fuente sustenta una parte del enunciado: las actas, las modificaciones, y "
        "el panel del formulario, las discrepancias.",
        detalle=f"Las fichas se modificaron con actas de validación de 2022 y 2023, varias no ratificadas, y el panel del "
                f"formulario registra {C['panel_disc']} discrepancias con las demás fuentes de seguimiento.")
    add("D5-H3", "D5", "Factores que explican las deficiencias",
        f"Las entidades explican las fallas de los indicadores sobre todo por la fuente o el método de cálculo "
        f"({C['fie_cod_total'].get('D5.FUENTE', 0)} segmentos), la meta mal estimada "
        f"({C['fie_cod_total'].get('D5.META', 0)}) y el diseño de la ficha ({C['fie_cod_total'].get('D5.FICHA', 0)}). "
        f"Ejemplos: SERVIR calcula la brecha salarial con la planilla electrónica y no con la ENAHO que fija la ficha, "
        f"y la ENARES, fuente del indicador del OP6, no se aplicó con la periodicidad prevista. El nivel conductor "
        f"agrega una causa de origen: la premura con que se aprobaron la Política y el PEMIG no dejó tiempo para "
        f"negociar con las entidades indicadores de calidad, y se aceptaron los que ellas propusieron (OGPPM y DGIGND).",
        [_ev("Entidades", "REP_ENT", "Códigos D5 del libro de códigos.", C.origen["fie_cod_total"],
             "Formulario de seguimiento e informes LIO"),
         _ev("Entidades", "REP_ENT", "Fuente distinta de la ficha (SERVIR).", seg(C, "FORM-2025-IOP.05.02-falla_abajo"),
             "Formulario de seguimiento"),
         _ev("Entidades", "REP_ENT", "ENARES sin periodicidad.", seg(C, "LIO-2023-5.5.1-05"), "Informes LIO")]
        + ges(C, "Premura de la formulación y negociación de los indicadores.", "ENTREV-OGPPM-012", "ENTREV-DGIGND-020"),
        "Sólido",
        detalle=f"Las entidades explican las fallas de los indicadores por la fuente o el método de cálculo "
                f"({C['fie_cod_total'].get('D5.FUENTE', 0)} segmentos), la meta mal estimada "
                f"({C['fie_cod_total'].get('D5.META', 0)}) y el diseño de la ficha "
                f"({C['fie_cod_total'].get('D5.FICHA', 0)}); el nivel conductor agrega la premura de la formulación.")

    assert C["endes_fuera"] == 2, "revisar el texto de D5-H4: cambió el número de contrastes fuera del intervalo"
    add("D5-H4", "D5", "Discrepancias en valores oficiales de 2020",
        f"Dos valores oficiales de 2020 no coinciden con los que se obtienen de los microdatos de la ENDES 2020, "
        f"levantada en buena parte por teléfono: la demanda insatisfecha de planificación familiar (oficial 6,1 %, "
        f"microdatos 4,4 %) y la violencia física o sexual de pareja (oficial 8,8 %, microdatos 7,9 %). Son los únicos "
        f"de {C['endes_contrastes']} contrastes con la ENDES fuera del intervalo de confianza.",
        [_ev("Cuantitativa", "REV_DOC", "Estimación con microdatos de la ENDES 2020.", C.origen["endes"],
             "Encuesta Demográfica y de Salud Familiar (ENDES)"),
         _ev("Cuantitativa", "REV_IND", "Valor oficial del reporte de seguimiento.", _hoja(PSI, "Serie anual"),
             "Reportes de seguimiento (RS)")],
        "Con respaldo parcial", salvedad="Los reportes no explican la causa de esas diferencias.",
        detalle="Dos valores oficiales de 2020 no coinciden con los microdatos de la ENDES 2020: la demanda "
                "insatisfecha de planificación familiar (6,1 % frente a 4,4 %) y la violencia física o sexual de pareja "
                "(8,8 % frente a 7,9 %).")
    s3 = C["op3_serie"].set_index("Año").sort_index()
    cambios = [svy.cambio(s3.loc[a, "Estimación (%)"], s3.loc[a, "EE"], s3.loc[a + 1, "Estimación (%)"],
                          s3.loc[a + 1, "EE"])[2] for a in s3.index[:-1]]
    k, nc = sum(cambios), len(cambios)
    per = s3.loc[2019:2025]
    i3 = C["op3_indice"].set_index("Año")
    assert i3.loc[2021, "Valor oficial"] == i3.loc[2022, "Valor oficial"], "la serie oficial ya no repite 2021 en 2022"
    signif = (f"ninguno de los {nc} cambios entre años consecutivos de {s3.index.min()} a {s3.index.max()} es "
              f"significativo al 95 %" if k == 0 else
              f"{k} de los {nc} cambios entre años consecutivos de {s3.index.min()} a {s3.index.max()} "
              f"{'es significativo' if k == 1 else 'son significativos'} al 95 %")
    add("D5-H5", "D5", "Sensibilidad y precisión de los indicadores del OP3",
        f"Los dos indicadores del OP3 no permiten seguir el avance anual. El índice de empoderamiento político "
        f"(IOP.03.01) solo cambia con las elecciones: la serie oficial repite su valor en 2021 y 2022, y el valor "
        f"recalculado es el mismo de 2023 a 2025. El porcentaje de funcionarias y directivas (IOP.03.02) se estima con "
        f"{int(per.n.min())} a {int(per.n.max())} personas de la ENAHO por año entre 2019 y 2025: en 2025 su intervalo "
        f"de confianza al 95 % va de {fmt_num(s3.loc[2025, 'IC inferior'], 1)} % a "
        f"{fmt_num(s3.loc[2025, 'IC superior'], 1)} %, y {signif}. SERVIR, que lo publica, cambió además de método "
        f"en 2023.",
        [_ev("Cuantitativa", "REV_IND", "Serie del IOP.03.02 con la ENAHO, error estándar con estrato y conglomerado "
             "(REV_IND.4).", C.origen["op3_serie"], "Encuesta Nacional de Hogares (ENAHO)"),
         _ev("Cuantitativa", "REV_IND", "Componentes del IOP.03.01 por año y serie oficial (REV_IND.4).",
             C.origen["op3_indice"], "Reportes de seguimiento (RS); INEI; JNE")],
        "Sólido",
        detalle=f"El índice de empoderamiento político (IOP.03.01) solo cambia con las elecciones y el porcentaje de "
                f"funcionarias y directivas (IOP.03.02) se estima con {int(per.n.min())} a {int(per.n.max())} personas "
                f"de la ENAHO por año: {signif}.")

    # D6
    ef = C["enfoques"].set_index("enfoque")
    niv = C["enfoques"].nivel.value_counts()
    oper = ef[ef.nivel == "Operacionalizado"]
    add("D6-H1", "D6", "Nivel de incorporación de los siete enfoques",
        f"De los siete enfoques transversales, {niv.get('Operacionalizado', 0)} se operacionalizan en el diseño, con "
        f"alcance desigual: el de género de forma amplia (lineamientos, estándar de 38 de 52 servicios e indicador de "
        f"OP), el intercultural de forma parcial (estándar de pertinencia cultural en 24 servicios) y los de "
        f"interseccionalidad y ciclo de vida de forma mínima. El de derechos humanos solo se menciona y los de "
        f"{lista_es([e.lower() for e in ef[ef.nivel == 'Ausente'].index])} están ausentes.",
        [_ev("Documental", "REV_DOC", "Lista de verificación con cita y ubicación verificadas.", C.origen["enfoques"],
             "Documento de la PNIG"),
         _ev("Documental", "REV_DOC", "Estándares e indicadores de los 52 servicios.", C.origen["enfoques_res"],
             "Documento de la PNIG (Anexo 5)"),
         _ev("Documental", "REV_DOC", "Fichas de indicadores de los OP y de servicio.", C.origen["enfoques_res"],
             "Fichas técnicas")],
        "Sólido", salvedad="La gestión por resultados está ausente como enfoque declarado, aunque la Política sigue la "
        "cadena de resultados que exige el CEPLAN.",
        detalle=f"De los siete enfoques transversales, {niv.get('Operacionalizado', 0)} se operacionalizan (solo el de "
                f"género con amplitud), el de derechos humanos solo se menciona y los de "
                f"{lista_es([e.lower() for e in ef[ef.nivel == 'Ausente'].index])} están ausentes.")
    add("D6-H2", "D6", "Enfoques en los protocolos y en los indicadores",
        f"Los protocolos aprobados incorporan los enfoques más que el diseño: {ef.loc['Género', 'protocolos']} "
        f"protocolos incluyen el de género, {ef.loc['Interseccionalidad', 'protocolos']} el de interseccionalidad y "
        f"{ef.loc['Interculturalidad', 'protocolos']} el intercultural. En cambio, ninguna de las "
        f"{C['fichas_total']} fichas de indicador prevé desagregaciones que permitan seguir los enfoques.",
        [_ev("Documental", "REV_DOC", "Menciones en los 28 protocolos aprobados.", C.origen["enfoques_res"],
             "Protocolos aprobados"),
         _ev("Documental", "REV_IND", "Atributo de desagregación de las fichas.", C.origen["ficha_desag_nc"],
             "Fichas técnicas")],
        "Sólido")

    # D7
    tot = lambda plano: int(m5.loc[plano, ["Continuidad", "Precisión", "Divergencia"]].sum())
    add("D7-H1", "D7", "Correspondencia conceptual en cuatro planos",
        f"El nuevo marco mantiene la mayor parte de la estructura y de los servicios, pero diverge en el plano "
        f"conceptual: en ese plano {m5.loc['Conceptual', 'Divergencia']} de {tot('Conceptual')} elementos divergen, "
        f"frente a {m5.loc['Estructura de la Política', 'Divergencia']} de {tot('Estructura de la Política')} en la "
        f"estructura de la Política y {m5.loc['Servicios', 'Divergencia']} de {tot('Servicios')} en los servicios. "
        f"{m5.loc['Servicios', 'Se mantiene']} servicios se mantienen, {m5.loc['Servicios', 'Se ajusta']} se "
        f"ajustan y {m5.loc['Servicios', 'Se rediseña']} se rediseñan ({lista_es(C['mac5_serv_rediseña'])}). En "
        f"la entrevista (5 de noviembre de 2026), la DGIGND no preveía cambios de fondo: consideraba que la definición "
        f"del enfoque era todavía general y que su operacionalización dependía de los lineamientos de "
        f"transversalización en elaboración. Sus informes técnicos del 9 de noviembre concluyen en cambio que la "
        f"incompatibilidad alcanza al objeto de la Política (su problema público, su finalidad y la cadena causal de la "
        f"que derivan objetivos, lineamientos y servicios) (MIMP, 2026g, 2026h, 2026i), y sobre esa base la OGAJ "
        f"considera que la PNIG perdió vigencia el 25 de diciembre de 2025 (MIMP, 2026e). La matriz de "
        f"incompatibilidades que acompaña esos informes coincide con la clasificación de esta evaluación en "
        f"{int(C['mic_conc']['Coinciden'])} de los {int(C['mic_conc']['Servicios'])} servicios (acuerdo de "
        f"{fmt_num(C['mic_conc']['Acuerdo simple (%)'], 1)} % y kappa de {fmt_num(C['mic_conc']['Kappa de Cohen'], 2)}): "
        f"las dos rediseñan los servicios de educación sexual integral y las "
        f"{int(C['mic_conc']['Discrepancias entre «Se mantiene» y «Se ajusta»'])} diferencias son entre mantener y "
        f"ajustar un servicio.",
        [_ev("Documental", "REV_ALI", "Clasificación por plano con cita de la PNIG y del nuevo marco.",
             C.origen["mac5"], "Documento de la PNIG y marco conceptual")]
        + ges(C, "Lectura de la DGIGND sobre el nuevo marco conceptual.", "ENTREV-DGIGND-041", "ENTREV-DGIGND-042",
              "ENTREV-DGIGND-044")
        + [_ev("Documental", "REV_DOC", "Alcance de la incompatibilidad según los informes técnicos de la DGIGND, la "
               "DPIGND y la DTEG del 9 de noviembre de 2026 (conclusiones 3.6 a 3.10).", "docs/fuentes/pnig/07_transicion_ley_32535/2026_mimp-dgignd_informe-d000041-2026_sustento-vigencia-normas-ley-28983.pdf",
               "Informes técnicos de la DGIGND"),
           _ev("Documental", "REV_DOC", "Contraste de la clasificación de los servicios con la matriz de "
               "incompatibilidades.", "REV_DOC_matriz_incompatibilidades.xlsx, hojas «Concordancia» y «Servicios»",
               "Matriz de incompatibilidades de la DGIGND"),
           _ev("Documental", "REV_DOC", "Pérdida de vigencia de la PNIG (numerales 2.3, 2.7.2 y 3.2).", "docs/fuentes/pnig/07_transicion_ley_32535/2026_mimp-ogaj_informe-d000926-2026_vigencia-normas-ley-28983.pdf",
               "Informe de la OGAJ")],
        "Con respaldo parcial", nota="Las fuentes coinciden en la divergencia conceptual; la lectura de la DGIGND en la "
        "entrevista, previa a sus informes técnicos y a la opinión de la OGAJ, no preveía cambios de fondo. La "
        "entrevista fue grupal y se atribuye al órgano. La matriz se tradujo a las categorías de la evaluación con una "
        "regla fija (numeral 3.2).",
        detalle=f"El nuevo marco mantiene la mayoría de los servicios, pero diverge en el plano conceptual "
                f"({m5.loc['Conceptual', 'Divergencia']} de {tot('Conceptual')} elementos); la matriz de la DPIGND "
                f"coincide en {int(C['mic_conc']['Coinciden'])} de {int(C['mic_conc']['Servicios'])} servicios. Según la "
                f"DGIGND y la OGAJ, la incompatibilidad alcanza al objeto de la Política, que perdió vigencia el 25 de "
                f"diciembre de 2025.")
    add("D7-H2", "D7", "Elementos que se rediseñan",
        "Las divergencias se concentran en el enfoque de género como categoría de análisis, el problema público de "
        "discriminación estructural, la diversidad de las mujeres, la autonomía física y la relación con el enfoque "
        "de familia. La DGIGND coincide en los cambios que derivan de la Ley N.° 32535 (la ley ya no habla de "
        "educación sexual integral, pone énfasis en la familia y la salud integral debe abordarse desde las "
        "competencias del MIMP), pero ve continuidad en el marco contra la discriminación, que remite a la CEDAW, a la "
        "Convención de Belém do Pará y a la Ley N.° 30364. Su matriz de incompatibilidades precisa esos cambios "
        "servicio por servicio: declara incompatible la educación sexual integral (servicios 2.2.1 y 2.2.2) y pide "
        "prestar otros sin elementos del enfoque de género, como las «nuevas masculinidades» del servicio 1.2.2 o la "
        "formación técnica y superior «con enfoque de género» del 4.3.3 (MIMP, 2026g).",
        [_ev("Documental", "REV_ALI", "Elementos con divergencia para contrastar con la DGIGND.",
             C.origen["mac5_entrevista"], "Documento de la PNIG y marco conceptual")]
        + ges(C, "Cambios y continuidades del nuevo marco según la DGIGND.", "ENTREV-DGIGND-040", "ENTREV-DGIGND-043")
        + [_ev("Documental", "REV_DOC", "Contenidos incompatibles y compatibles por grupo de servicios (hoja de la "
               "PNIG, numeral 3 del Anexo).", "docs/fuentes/pnig/07_transicion_ley_32535/2026_mimp-dpignd_matriz-incompatibilidades-ley-32535.xlsx", "Matriz de incompatibilidades de la DGIGND"),
           _ev("Documental", "REV_DOC", "Contenidos incompatibles y compatibles identificados en la matriz de la DGIGND "
               "(numerales 2.3.2 y 2.8.2).", "docs/fuentes/pnig/07_transicion_ley_32535/2026_mimp-ogaj_informe-d000926-2026_vigencia-normas-ley-28983.pdf", "Informe de la OGAJ")],
        "Con respaldo parcial",
        detalle="Las divergencias con el nuevo marco se concentran en el enfoque de género como categoría, el problema "
                "público, la diversidad de las mujeres, la autonomía física y la relación con el enfoque de familia; la "
                "DGIGND ve continuidad en el marco contra la discriminación.")

    # I1
    g_op2 = C["grd"][C["grd"]["Código"].astype(str).str.startswith("2")]
    op2_n, op2_prot = len(g_op2), int((g_op2["Protocolo aprobado"] == "Sí").sum())
    assert op2_prot == 0, "revisar I1-H1: el OP2 ya tiene protocolo aprobado"
    add("I1-H1", "I1", "Estado de implementación de los servicios",
        f"La implementación fue {'mayoritariamente parcial' if C['serv_parc'] > n / 2 else 'parcial en buena parte de los servicios'}: "
        f"de {n} servicios, {C['serv_impl']} están implementados, "
        f"{C['serv_parc']} implementados parcialmente y {C['serv_noimpl']} "
        f"{pl(C['serv_noimpl'], 'no implementado', 'no implementados')} ({lista_es(C['serv_noimpl_cod'])}). Solo {C['serv_prot_si']} servicios tienen protocolo aprobado al 31 de "
        f"diciembre de 2025 y {C['serv_falta_tres']} carecen a la vez de protocolo, de asignación identificable y de "
        f"reporte de indicador. El nivel conductor da cifras concordantes: unos 25 de los 52 servicios con protocolo "
        f"(OGMEPGD) y 28 protocolos aprobados entre 2021 y 2025, según lo informado al CEPLAN en julio de 2026 "
        f"(DGIGND), que coinciden "
        f"con los que registra la revisión documental. En el OP2, cuyos servicios están sobre todo a cargo del Minsa, "
        f"no se aprobó ningún protocolo ({op2_prot} de {op2_n} servicios).",
        [_ev("Documental", "REV_DOC", "Protocolos aprobados y evidencia de ejecución por servicio.",
             C.origen["serv_total"], "Protocolos aprobados"),
         _ev("Entidades", "REV_DOC", "Ejecución reportada en los informes LIO 2019–2025.", C.origen["serv_total"],
             "Informes LIO"),
         _ev("Cuantitativa", "REV_IND", "Montos verificados en los datos abiertos del MEF.", C.origen["mef_coinc"],
             "Datos abiertos del MEF")]
        + ges(C, "Protocolos aprobados según el nivel conductor.", "ENTREV-OGMEPGD-014", "ENTREV-DGIGND-013", "ENTREV-DGIGND-046"),
        "Sólido",
        detalle=f"De {n} servicios, {C['serv_impl']} están implementados, {C['serv_parc']} implementados parcialmente y "
                f"{C['serv_noimpl']} {pl(C['serv_noimpl'], 'no implementado', 'no implementados')}; solo "
                f"{C['serv_prot_si']} tienen protocolo aprobado al 31 de "
                f"diciembre de 2025, y ninguno de los del OP2.")
    add("I1-H2", "I1", "Asignación presupuestal identificable",
        f"La asignación presupuestal es identificable para {C['serv_asig_si']} de {n} servicios. De "
        f"{C['mef_total']} montos que las entidades reportaron, {C['mef_coinc']} coinciden con una meta, actividad o "
        f"producto del MEF, {C['mef_cerc']} son cercanos y {C['mef_sinc']} no coinciden. De estos últimos, "
        f"{C['mef_motivos']['sinexp']} no tienen explicación en el texto del reporte, {C['mef_motivos']['ao']} "
        f"corresponden a actividades operativas del POI, {C['mef_motivos']['pliego']} no son gasto del pliego "
        f"(créditos, subvenciones o cofinanciamientos), {C['mef_motivos']['norma']} son montos de una norma o "
        f"transferencia y {C['mef_motivos']['varias']} suman varias entidades o programas. Según el nivel "
        f"conductor, las entidades no registran en el POI actividades propias de cada servicio sino actividades "
        f"operativas genéricas que, en lo interno, los alimentan (OGPPM), y no sabían cuánto gastaban en cada servicio "
        f"(DGIGND).",
        [_ev("Cuantitativa", "REV_IND", "Verificación de cifras reportadas.", C.origen["mef_coinc"],
             "Datos abiertos del MEF"),
         _ev("Entidades", "REV_DOC", "Montos reportados en los informes LIO y en el formulario 2023.",
             C.origen["serv_asig_si"], "Informes LIO")]
        + ges(C, "Registro del gasto de los servicios.", "ENTREV-OGPPM-003", "ENTREV-DGIGND-019"),
        "Sólido",
        detalle=f"Los recursos presupuestales se identificaron para {C['serv_asig_si']} de {n} servicios. De "
                f"{C['mef_total']} montos reportados por las entidades, {C['mef_coinc']} coinciden con una meta, "
                f"actividad o producto de los datos abiertos del MEF y {C['mef_sinc']} no coinciden.")
    add("I1-H3", "I1", "Factores de la implementación parcial",
        "Los factores de dificultad más frecuentes en lo que reportaron las entidades son "
        + lista_es([f"{nombre.lower()} ({k}, {v} segmentos)" for k, nombre, v in dif[:5]]) +
        ". Los problemas de personal y de presupuesto aparecen en las tres fuentes: el formulario, los informes LIO y "
        "los informes de gobernanza de la CONAIG. El nivel conductor coincide y precisa por qué se demoraron los "
        "protocolos: los que estaban listos a nivel técnico se retrasaban en la alta dirección de las entidades o "
        "volvían a revisarse con los cambios de gestión; pesaron la baja prioridad del tema y la falta de personal en "
        "sectores como el Minsa; normas sectoriales y la nueva ley frenaron los del Minedu; la implementación de la Política se vio afectada por la "
        "pandemia desde 2020, y algunas entidades no aprobaban protocolos que no podían financiar (DGIGND). Las frecuencias indican cuántas veces se mencionó cada factor, no su peso ni su efecto.",
        [_ev("Entidades", "REP_ENT", "Segmentos con sentido de dificultad, por código I1.", C.origen["fie_dif_top"],
             "Formulario de seguimiento e informes LIO"),
         _ev("Entidades", "REP_ENT", "Limitaciones de presupuesto y personal.", seg(C, "CONAIG-2025-024"),
             "Informes de gobernanza de la CONAIG")]
        + ges(C, "Factores en la aprobación de los protocolos.", "ENTREV-DGIGND-012", "ENTREV-DGIGND-014", "ENTREV-DGIGND-015",
              "ENTREV-DGIGND-016", "ENTREV-DGIGND-034", "ENTREV-DGIGND-035"),
        "Sólido",
        detalle="Los factores de dificultad que más reportaron las entidades son "
                + lista_es([f"{nombre.lower()} ({v} segmentos)" for _, nombre, v in dif[:3]]) +
                "; el nivel conductor agrega la demora de los protocolos en la alta dirección de las entidades.",
        nota=f"Los {C['dc_n']} segmentos críticos se codificaron dos veces de forma independiente (kappa medio por "
             f"código de {fmt_num(C['dc_kappa_cod'], 2)}).")

    # I2
    add("I2-H1", "I2", "Avance anual de los indicadores de OP",
        f"En el reporte oficial, el avance de los OP no puede valorarse en 2025 para más de la mitad de los "
        f"indicadores: {C['op_sinvalor_2025']} de {C['op_ind']} no tienen valor. De los que tienen valor, "
        f"{C['op_alto_2025']} están en nivel alto ({C['op_fnc_2025']} con fuente no comparable), {C['op_medio_2025']} "
        f"en medio, {C['op_bajo_2025']} en bajo ({C['op_bajofalla_2025']} con posible falla de planeación) y "
        f"{C['op_noint_2025']} no permite calcular el avance. El vacío es del reporte y no siempre de la información: para "
        f"{len(C['op_est25'])} de esos {C['op_sinvalor_2025']} indicadores la fuente existe y esta evaluación "
        f"estimó o recalculó el valor (I2-H3).",
        [_ev("Cuantitativa", "REV_IND", "Semaforización 2019–2025.", C.origen["resumen_semaforo"],
             "Reportes de seguimiento (RS)"),
         _ev("Cuantitativa", "REV_IND", "Contraste con el consolidado de la OMEP.", _hoja(PSI, "Contraste de fuentes"),
             "Consolidado de indicadores de la OMEP")],
        "Indicio",
        detalle=f"En el reporte oficial, {C['op_sinvalor_2025']} de {C['op_ind']} indicadores de OP no tienen valor 2025; "
                f"de los que lo tienen, {C['op_alto_2025']} están en nivel alto, {C['op_medio_2025']} en medio, "
                f"{C['op_bajo_2025']} en bajo y {C['op_noint_2025']} no permite calcular el avance.")
    add("I2-H2", "I2", "Tendencia frente a la línea de base",
        f"Frente a su línea de base, {C['op_mejora']} de {C['op_ind']} indicadores de OP mejoran; "
        f"{'no mejora' if len(C['op_nomejora']) == 1 else 'no mejoran'} {lista_es(C['op_nomejora'])}; "
        f"{lista_es(C['op_sinlb'])} {'no tiene' if len(C['op_sinlb']) == 1 else 'no tienen'} línea de base "
        f"numérica, y el último valor de {lista_es(C['op_noint_lb'])} no permite calcular el avance. El indicador de entidades "
        f"públicas que incorporan el enfoque de género (IOP.05.01) es el de menor avance en 2025 "
        f"({pct(C['iop0501_avance'])} de su meta).",
        [_ev("Cuantitativa", "REV_IND", "Último valor frente a la línea de base.", C.origen["op_tabla"],
             "Reportes de seguimiento (RS)"),
         _ev("Cuantitativa", "REV_IND", "Series de los reportes de cumplimiento.", _hoja(PSI, "Serie anual"),
             "Reportes de cumplimiento (RC)")],
        "Indicio",
        salvedad="Los años de línea de base (2015 a 2018) son anteriores a la Política. Donde el reporte de "
                 "seguimiento 2025 y los reportes de cumplimiento difieren, se usa el primero.")
    # todas las categorías de 2025, en el orden de la escala, para que la enumeración sume el total (ADR-010)
    orden25 = ["Alto", "Alto con posible falla de planeación", "Medio", "Bajo", "Bajo con posible falla de planeación",
               "Sin sentido definido", "Sin valor (No disponible)", "Sin meta"]
    assert set(k for k, v in s25.items() if v) <= set(orden25), f"categorías 2025 sin texto: {s25}"
    cat25 = [(s25.get(k, 0), k) for k in orden25 if s25.get(k, 0)]
    assert sum(v for v, _ in cat25) == C["serv25_total"]
    frase25 = lista_es([f"{v} {('están en nivel ' if i == 0 else '') + k.lower().replace(' (no disponible)', '')}"
                        for i, (v, k) in enumerate(cat25)])
    add("I2-H4", "I2", "Indicadores de servicio",
        f"Los indicadores de servicio pasaron de no tener meta a tenerla: en 2025, de {C['serv25_total']} "
        f"indicadores, {frase25}. En total, {C['serv25_con_valor']} tienen valor 2025, la misma cifra que informa "
        f"el MIMP (2026b, p. 11).",
        [_ev("Cuantitativa", "REV_IND", "Semaforización de servicios 2019–2025.", C.origen["serv25"],
             "Reportes de seguimiento (RS)"),
         _ev("Entidades", "REP_ENT", "Panel del formulario de seguimiento 2023–2025.",
             _hoja(FIE, "Panel por indicador"), "Formulario de seguimiento")],
        "Indicio",
        detalle=f"En 2025, de {C['serv25_total']} indicadores de servicio, {s25.get('Alto', 0)} están en nivel alto y "
                f"{C['serv25_falla']} muestran una posible falla de planeación; {C['serv25_con_valor']} tienen valor 2025.")

    e = C["op_est25"]
    endes = e[e.Fuente == "ENDES 2025"]
    resto = C["op_sinvalor_resto"]
    var_ic = [f"{NOMBRE_EST[r.Código]} (entre {lista_es([n.lower() for n in r['Niveles en el IC']])})"
              for _, r in e.iterrows() if len(r["Niveles en el IC"]) > 1]
    sin_op = C["op_sin_valor_op"]
    cierra = f"el {sin_op[0]} cierra" if len(sin_op) == 1 else f"los {lista_es(sin_op)} cierran"
    sigue = "El otro sigue" if len(resto) == 1 else f"Los otros {len(resto)} siguen"
    nota_ic = (f"El nivel se calcula con la estimación puntual; con el intervalo de confianza al 95 % podría variar "
               f"para {lista_es(var_ic)}. " if var_ic else "El nivel se calcula con la estimación puntual. ")
    s3 = C["op3_serie"].set_index("Año")
    rep = s3[s3.Reproduce == "Sí"]
    s25 = s3.loc[2025]
    assert e.Fuente.isin(["ENDES 2025", "ENAHO 2025", "Registros electorales"]).all()
    add("I2-H3", "I2", "Indicadores sin valor oficial 2025: estimaciones de esta evaluación y vacíos que persisten",
        f"Esta evaluación estimó o recalculó {len(e)} de los {C['op_sinvalor_2025']} indicadores de OP sin valor "
        f"oficial 2025, con métodos que reproducen los valores oficiales o publicados de años anteriores: "
        f"{len(endes)} de los {lista_es(sorted(set(endes.OP)))} con los microdatos de la ENDES 2025, IOP.03.02 con la "
        f"ENAHO 2025 e IOP.03.01 con registros electorales. Frente a su meta 2025, quedan {frase_est25(C)}. {sigue} "
        f"sin valor: {lista_pc([MOTIVO_SIN_VALOR[c] for c in resto])}. Así, {cierra} sin ningún indicador con valor "
        f"2025.",
        [_ev("Cuantitativa", "REV_DOC", "Estimación con diseño muestral complejo.", C.origen["endes"],
             "Encuesta Demográfica y de Salud Familiar (ENDES)"),
         _ev("Cuantitativa", "REV_IND", "IOP.03.02 con la ENAHO 2025, con el método que reproduce la serie de SERVIR, e "
             "IOP.03.01 con registros electorales, con el método que reproduce la serie oficial (REV_IND.4).",
             f"{C.origen['op3_serie']}; {C.origen['op3_indice']}",
             "ENAHO; INEI, Brechas de Género 2025; JNE, Perfil Electoral N.° 4"),
         _ev("Cuantitativa", "REV_IND", "Nivel frente a la meta 2025 con la regla de la Tabla 9.", C.origen["op_est25"],
             "Reportes de seguimiento (RS)")],
        "Sólido",
        detalle=f"Esta evaluación estimó o recalculó {len(e)} de los {C['op_sinvalor_2025']} indicadores de OP sin "
                f"valor oficial 2025, con métodos que reproducen los valores oficiales de años anteriores; "
                f"{queda_sin_valor(C)}.",
        salvedad="Las estimaciones y los recálculos son de esta evaluación: no son valores reportados por las "
                 "entidades ni sustituyen al reporte oficial.",
        nota=nota_ic.strip())
    # salvedades largas: van como «Nota:» bajo la tabla de indicadores de los OP (numeral 5.2 del informe)
    H[-1]["nota_tabla"] = (
        f"IOP.03.02 reproduce a ±{fmt_num(rep.Diferencia.abs().max(), 1)} puntos la serie que SERVIR publicó con la "
        f"ENAHO de {rep.index.min()} a {rep.index.max()}; en 2025 se apoya en {int(s25.n)} personas de la muestra y su "
        f"intervalo de confianza va de {fmt_num(s25['IC inferior'], 1)} % a {fmt_num(s25['IC superior'], 1)} %. SERVIR "
        f"publicó 2023 y 2024 con una metodología actualizada (INEI, s. f.) que los microdatos públicos no reproducen: "
        f"con el método de la serie esos años dan {fmt_num(s3.loc[2023, 'Estimación (%)'], 1)} y "
        f"{fmt_num(s3.loc[2024, 'Estimación (%)'], 1)}, frente a {fmt_num(s3.loc[2023, 'SERVIR (%)'], 1)} y "
        f"{fmt_num(s3.loc[2024, 'SERVIR (%)'], 1)}, que son los valores de los reportes de seguimiento; además, su edición 2026 presenta en "
        f"el texto como de 2025 el 44 % que su Gráfico 12 (p. 22) atribuye a la ENAHO 2024. IOP.03.01 reproduce a cuatro "
        f"decimales los valores oficiales de 2019 a 2022 y el 0,58 de 2023; toma la composición del Congreso elegido en "
        f"2021 y de las ERM 2022 hasta las elecciones de 2026, por lo que su valor de 2024 y 2025 es el de 2023. El "
        f"componente de candidatas de las ERM 2022 usa el porcentaje que publica el JNE (43,0 %), que no da conteos.")

    # I3
    add("I3-H1", "I3", "Recursos de la conducción por órgano",
        f"El devengado de la meta de la DGIGND bajó de {monto(dev[2022]['DGIGND'])} en 2022 a "
        f"{monto(dev[2023]['DGIGND'])} en 2023, año en que aparecen metas propias de la DPIGND y de la DPPDM, "
        f"órganos del ROF 2023. Sumadas las tres metas, la conducción devengó {monto(cs[2022])} en 2022 y "
        f"{monto(cs[2023])} en 2023 (una reducción de {pct(-C['cond_var_22_23'])}), {monto(cs[2024])} en 2024 y "
        f"{monto(cs[2025])} en 2025. Según la OGPPM, el presupuesto se asigna por meta a cada órgano y la meta de la "
        f"dirección general financia también otros compromisos, como la evaluación de la LIO y los mecanismos, por lo "
        f"que la parte destinada a la conducción de la Política solo puede estimarse.",
        [_ev("Cuantitativa", "REV_IND", "Devengado y PIM por meta de los órganos conductores.",
             C.origen["cond_dev"], "Datos abiertos del MEF"),
         _ev("Documental", "REV_IND", "Actividades operativas de la conducción en el POI 2019–2025.",
             _hoja(PSI3R, "POI"), "POI del MIMP"),
         _ev("Documental", "REV_IND", "Ejecución de la actividad 5001072 por finalidad remitida por la DGIGND: "
             "coincide al sol con el devengado de cada órgano.", _hoja("REV_IND_3_presupuesto_dgignd.xlsx", "Conciliación MEF"), "Información presupuestal de la DGIGND")]
        + ges(C, "Asignación por meta y gasto de la conducción.", "ENTREV-OGPPM-019"),
        "Con respaldo parcial", VALIDADO,
        "El POI no indica la actividad presupuestal de cada actividad operativa; el traslado entre órganos se "
        "infiere de los datos abiertos del MEF.",
        detalle=f"Sumadas las metas de la DGIGND, la DPIGND y la DPPDM, el devengado de la conducción pasó de "
                f"{monto(cs[2022])} en 2022 a {monto(cs[2024])} en 2024 y a {monto(cs[2025])} en 2025; la parte "
                f"destinada a la Política solo puede estimarse.")
    add("I3-H2", "I3", "Ejecución y actividad presupuestal vinculada",
        f"La ejecución de la meta de la DGIGND fue alta en todo el periodo (entre "
        f"{pct(min(C['dgignd_ejec'].values()))} y {pct(max(C['dgignd_ejec'].values()))} del PIM). La actividad "
        f"5001072 «Promoción de la mujer», que agrupa la promoción de la mujer de la Administración Central y cuya "
        f"denominación corresponde a la finalidad de la PNIG, pasó de un PIM de {monto(vd.loc[2019, 'pim'])} en 2019 "
        f"a {monto(vd.loc[2025, 'pim'])} en 2025, con ejecución entre {pct(vd['ejecucion_%'].min())} y "
        f"{pct(vd['ejecucion_%'].max())}.",
        [_ev("Cuantitativa", "REV_IND", "Ejecución por meta y por actividad.", C.origen["act5001072"],
             "Datos abiertos del MEF"),
         _ev("Documental", "REV_IND", "PIM y ejecución de la actividad 5001072 por finalidad, 2019–2026, remitidos "
             "por la DGIGND.", _hoja("REV_IND_3_presupuesto_dgignd.xlsx", "Conciliación MEF"), "Información presupuestal de la DGIGND")],
        "Indicio", nota="La información presupuestal remitida por la DGIGND coincide al sol con los datos abiertos del "
        "MEF en cada órgano y año de 2019 a 2025; ambas se originan en el registro presupuestal del Estado. No se "
        "contrastó en las entrevistas: las preguntas sobre la ejecución y la actividad 5001072 no se formularon.",
        detalle=f"La ejecución de la meta de la DGIGND fue alta en todo el periodo (entre "
                f"{pct(min(C['dgignd_ejec'].values()))} y {pct(max(C['dgignd_ejec'].values()))} del PIM); la actividad "
                f"5001072 «Promoción de la mujer» pasó de un PIM de {monto(vd.loc[2019, 'pim'])} en 2019 a "
                f"{monto(vd.loc[2025, 'pim'])} en 2025.")
    add("I3-H3", "I3", "Limitaciones presupuestales de la conducción",
        "Las limitaciones presupuestales afectaron la asistencia técnica presencial y el seguimiento: la DTEG priorizó "
        "asistencias virtuales y gestionó presupuesto adicional, el seguimiento de lactarios pasó a un formulario "
        "virtual y la ENARES no pudo aplicarse cada dos años por falta de presupuesto. La DGIGND, única fuente sobre "
        "este punto, sostiene además que la conducción no contó con los recursos necesarios: la llevó con una oficina "
        "pequeña y con otras prioridades, sin recursos para medir los indicadores del PEMIG ni para llegar al nivel "
        "descentralizado, y advierte que, si el MIMP no destina recursos suficientes a la Política, menos lo harán las "
        "otras entidades.",
        [_ev("Entidades", "REP_ENT", "Asistencia técnica de la DTEG.", seg(C, "LIO-2025-5.1.1-03"), "Informes LIO"),
         _ev("Entidades", "REP_ENT", "Seguimiento de lactarios.", seg(C, "CONAIG-2024-035"),
             "Informes de gobernanza de la CONAIG"),
         _ev("Entidades", "REP_ENT", "ENARES.", seg(C, "FORM-2025-IOP.06.01-contexto"), "Formulario de seguimiento")]
        + ges(C, "Recursos de la conducción.", "ENTREV-DGIGND-023", "ENTREV-DGIGND-026", "ENTREV-DGIGND-034", "ENTREV-DGIGND-037",
              ded="DED-02"),
        "Sólido",
        detalle="Las limitaciones presupuestales afectaron la asistencia técnica presencial y el seguimiento (informes "
                "LIO, CONAIG y formulario). Que la conducción no contó con los recursos necesarios es la lectura de la "
                "DGIGND, única fuente sobre ese punto.",
        nota="La confianza corresponde a lo que respaldan las cuatro fuentes. La pregunta sobre limitaciones "
             "presupuestales no se formuló; su contenido se dedujo de otras respuestas de la DGIGND "
             "(Anexo 8).")

    # I4
    add("I4-H2", "I4", "Funcionamiento de la CONAIG",
        "La CONAIG sesionó con regularidad formal (tres sesiones ordinarias en 2024 y en 2025, el mínimo de su "
        "reglamento) pero con representación mayoritaria de alternos y asistencia variable en 2025 (12, 9 y 19 de sus "
        "21 integrantes: 57,1 %, 42,9 % y 90,5 %, según el informe LIO 2025). De las 12 medidas urgentes a su cargo, 4 cumplieron la meta hasta septiembre de 2025, 6 estaban en "
        "proceso y 2 con limitaciones. La DGIGND confirma la regularidad formal (dos sesiones en 2021, año de su "
        "instalación, y tres por año hasta 2025) y la baja asistencia de los titulares: hubo sesiones sin quórum, que es "
        "la mitad más uno de titulares o alternos, y con representantes de los representantes, sin capacidad de "
        "decisión, según la DGIGND. Aun así, la CONAIG facilitó que algunas decisiones se tradujeran en protocolos "
        "aprobados, como el de SERVIR. Fuera del periodo evaluado, la OGAJ considera que la CONAIG agotó su objeto el "
        "25 de diciembre de 2025, con la pérdida de vigencia del D.S. N.° 018-2021-MIMP (MIMP, 2026e), y la DGIGND "
        "informa que la PCM ya la registró como comisión extinguida en su reporte del tercer trimestre de 2026 "
        "(MIMP, 2026g, numeral 2.12); su tercera sesión de 2025 (29 de diciembre) y la de 2026 son posteriores a esa "
        "fecha.",
        [_ev("Entidades", "REP_ENT", "Sesiones 2024.", seg(C, "LIO-2024-6.4.1-07"), "Informes LIO"),
         _ev("Entidades", "REP_ENT", "Sesiones 2025.", seg(C, "LIO-2025-6.3.3-11"), "Informes LIO"),
         _ev("Entidades", "REP_ENT", "Medidas urgentes.", seg(C, "LIO-2025-6.3.3-14"), "Informes LIO"),
         _ev("Documental", "REV_DOC", "Sesiones y acuerdos por año, 2021–2026.",
             "docs/fuentes/pnig/09_conaig_gobernanza/2026_mimp-conaig_ayuda-memoria-sesiones-2021-2026.doc",
             "Ayuda memoria de la CONAIG")]
        + ges(C, "Sesiones, quórum y efecto de la CONAIG.", "ENTREV-DGIGND-007", "ENTREV-DGIGND-008", "ENTREV-DGIGND-009",
              "ENTREV-DGIGND-010", "ENTREV-DGIGND-011", "ENTREV-DGIGND-017")
        + [_ev("Documental", "REV_DOC", "Agotamiento del objeto de la CONAIG (numerales 2.6.2 y 2.8.3).", "docs/fuentes/pnig/07_transicion_ley_32535/2026_mimp-ogaj_informe-d000926-2026_vigencia-normas-ley-28983.pdf", "Informe de la OGAJ"),
           _ev("Documental", "REV_DOC", "Derogación tácita del D.S. N.° 018-2021-MIMP y registro de la CONAIG como "
               "extinguida por la PCM (numeral 2.12).", "docs/fuentes/pnig/07_transicion_ley_32535/2026_mimp-dgignd_informe-d000041-2026_sustento-vigencia-normas-ley-28983.pdf", "Informes técnicos de la DGIGND")],
        "Con respaldo parcial",
        detalle="La CONAIG sesionó con la regularidad mínima de su reglamento (tres sesiones por año), pero con "
                "representación mayoritaria de alternos, asistencia variable y sesiones sin quórum; de sus 12 medidas "
                "urgentes, 4 cumplieron la meta hasta septiembre de 2025.",
        nota="Las evidencias documentales provienen de los informes LIO 2024 y 2025 y de la ayuda memoria de la "
             "CONAIG, que confirma el número de sesiones por año; las actas de la CONAIG no se revisaron. La DGIGND "
             "matiza el efecto de la CONAIG sobre los protocolos.")
    add("I4-H3", "I4", "Seguimiento de las recomendaciones de la CONAIG",
        "Las recomendaciones de gobernanza de la CONAIG tuvieron baja adopción: solo 2 de 16 servicios incorporaron "
        "acciones de seguimiento en los planes de sus mecanismos para la igualdad de género, ninguna entidad realizó "
        "capacitaciones específicas para los equipos de los servicios, y algunas entidades declararon que las "
        "recomendaciones no les eran aplicables porque prestan sus servicios sin distinción por sexo.",
        [_ev("Entidades", "REP_ENT", "Planes de los MIG.", seg(C, "CONAIG-2024-040"), "Informes de gobernanza de la CONAIG"),
         _ev("Entidades", "REP_ENT", "Capacitación de los equipos.", seg(C, "CONAIG-2023-049"),
             "Informes de gobernanza de la CONAIG"),
         _ev("Entidades", "REP_ENT", "Respuesta de COFOPRI.", seg(C, "FORM-2025-ISS.04.02.02.01-falla_abajo"),
             "Formulario de seguimiento")],
        "Indicio", nota="No se contrastó en las entrevistas: las preguntas sobre los medios de la "
        "CONAIG para el cumplimiento de sus recomendaciones no se formularon.")
    add("I4-H4", "I4", "Articulación efectiva frente a la formal",
        f"La articulación reportada es desigual: según el informe LIO 2025, 19 servicios (38,7 %) reportaron "
        f"coordinación con entidades del nivel nacional y 15 (30,6 %) con gobiernos regionales y locales, sobre todo en "
        f"el OP4, y los servicios del OP2 no reportaron articulación nacional; el informe no precisa el total sobre el "
        f"que calcula esos porcentajes. A ello se suman {C['hor_dup']} duplicidades con otras políticas "
        f"nacionales. El nivel conductor coincide en que la articulación efectiva fue limitada: para la DGIGND, la "
        f"CONAIG fue el único mecanismo de articulación; los tres órganos coinciden en que el PEMIG no cumplió ese papel "
        f"y solo sirvió de sustento para exigir los protocolos; la OGPPM califica la articulación territorial de «muy "
        f"declarativa» y la DGIGND reconoce que la Política no llegó al nivel descentralizado.",
        [_ev("Entidades", "REP_ENT", "Coordinación por OP.", seg(C, "LIO-2025-6.3.3-21"), "Informes LIO"),
         _ev("Documental", "REV_ALI", "Duplicidades horizontales.", C.origen["hor_dup"], "Políticas nacionales afines")]
        + ges(C, "Mecanismos de articulación según el nivel conductor.", "ENTREV-DGIGND-024", "ENTREV-DGIGND-025",
              "ENTREV-DGIGND-026", "ENTREV-OGPPM-011", "ENTREV-OGPPM-029", "ENTREV-OGPPM-030", "ENTREV-OGMEPGD-012"),
        "Sólido",
        detalle=f"Según el informe LIO 2025, 19 servicios (38,7 %) reportaron coordinación con el nivel nacional y 15 "
                f"(30,6 %) con gobiernos regionales y locales; hay {C['hor_dup']} duplicidades con otras políticas, y el "
                f"nivel conductor considera la articulación efectiva limitada.")
    rec = C["rec"]
    est = rec["Estado al corte (p3.4)"].value_counts()
    asp = pd.crosstab(rec["Aspecto de mejora (Tabla 5)"], rec["Estado al corte (p3.4)"])
    n_ev = rec["Evaluación"].value_counts()
    impl = asp.loc["Implementación"]
    dis = asp.loc["Diseño"]
    # protocolos de los servicios del OP1 después del informe del CEPLAN: solo sus recomendaciones por servicio
    rc = rec[rec["Evaluación"] == "CEPLAN"]
    medios = rc["Medio de verificación"].map(str)
    sin_prot = sorted({s for s, m in zip(rc["Servicio"].map(str), medios)
                       if "sin protocolo aprobado" in m and re.fullmatch(r"\d\.\d\.\d", s)})
    prot_post = sorted({s for s, e, m in zip(rc["Servicio"].map(str), rc["Estado al corte (p3.4)"], medios)
                        if e == "Implementada" and "protocolo aprobado por" in m and re.fullmatch(r"\d\.\d\.\d", s)})
    r20 = rec[rec["Evaluación"] == "MIMP 2020"].set_index("N.°")["Estado al corte (p3.4)"]
    # lo que el enunciado nombra de la evaluación 2020: ENUT 2024 (EV2020-17) y autodiagnóstico de la DTEG (EV2020-21)
    # implementados; servicio 4.3.2 sin protocolo (EV2020-15) y lineamiento 3.2 sin servicio (EV2020-22)
    assert set(r20[r20 == "Implementada"].index) == {"EV2020-17", "EV2020-21"}
    assert r20["EV2020-15"] == r20["EV2020-22"] == "No implementada"
    e20 = r20.value_counts()
    add("I4-H5", "I4", "Seguimiento de las recomendaciones de evaluaciones anteriores",
        f"De las {len(rec)} recomendaciones de las cuatro evaluaciones anteriores ({n_ev['MIMP 2020']} de la "
        f"evaluación de implementación 2020, {n_ev['MIMP 2021']} de la evaluación "
        f"de 2021, {n_ev['MIMP 2023']} de la de 2023 y {n_ev['CEPLAN']} del informe técnico del CEPLAN), al corte "
        f"{est.get('Implementada', 0)} están implementadas, {est.get('Parcialmente implementada', 0)} parcialmente, "
        f"{est.get('No implementada', 0)} no implementadas y {est.get('Sin evidencia', 0)} sin evidencia; "
        f"{est.get('Sin objeto por la transición', 0)} quedó sin objeto por la Ley N.° 32535. Avanzaron las de "
        f"implementación, que dependían de las entidades responsables "
        f"({impl.get('Implementada', 0) + impl.get('Parcialmente implementada', 0)} de {impl.sum()} implementadas o "
        f"parcialmente implementadas), y no las de diseño ({dis.get('No implementada', 0)} de {dis.sum()} no "
        f"implementadas): después del informe del CEPLAN solo se aprobó el protocolo del servicio "
        f"{lista_es(prot_post)}, los servicios {lista_es(sin_prot)} siguen sin protocolo aprobado (hay propuestas "
        f"para el 1.1.4 y el 1.2.4) y no se modificaron los "
        f"indicadores de los demás servicios del OP1, porque el MIMP optó por hacerlo mediante la actualización de la "
        f"Política. De las {len(r20)} recomendaciones de la evaluación 2020, {e20.get('Implementada', 0)} se "
        f"implementaron (la ENUT 2024 y la herramienta de autodiagnóstico de la DTEG), "
        f"{e20.get('Parcialmente implementada', 0)} en parte y {e20.get('Sin evidencia', 0)} no tienen evidencia; el "
        f"servicio 4.3.2 siguió sin protocolo y el lineamiento 3.2 sin servicio. "
        f"De las dos recomendaciones de la evaluación de 2023, con plazo al primer trimestre de 2025, se "
        f"conformó el equipo de trabajo, pero no consta el plan de trabajo ni la revisión de evidencia del modelo "
        f"conceptual; la matriz de compromisos se elaboró una sola vez, en 2023. En las entrevistas, la única medida "
        f"que el nivel conductor refiere ante las recomendaciones de la evaluación de 2023 y del CEPLAN son las "
        f"coordinaciones para actualizar la PNIG, que no llegó a concretarse; las acciones de 2026 responden ya a la Ley "
        f"N.° 32535. La OMEP emitió informes de mejora del diseño de la Política que no tuvieron respuesta de la oficina "
        f"de planeamiento.",
        [_ev("Documental", "REV_DOC", "Recomendaciones de las evaluaciones de 2020, 2021 y 2023 y del informe técnico "
             "del CEPLAN, con su estado y su medio de verificación.", C.origen["rec"],
             "Informes de evaluación anteriores"),
         _ev("Documental", "REV_DOC", "Equipo de trabajo, ruta de actualización de los indicadores y matriz de "
             "compromisos.", "Reporte de cumplimiento de la PNIG 2024, p. 13; Informe N.° D000112-2026-MIMP-OP, "
             "p. 3; Informe N.° D000039-2026-MIMP-DPIGND, p. 11", "Reportes e informes del MIMP"),
         _ev("Entidades", "REP_ENT", "Medidas de 2025 de las entidades responsables de los servicios del OP1.",
             "Informe LIO 2025, pp. 38 a 78; formulario de seguimiento 2025", "Informes LIO"),
         _ev("Documental", "REV_DOC", "Protocolos e indicadores de los servicios del OP1.", C.origen["grd"],
             "Protocolos de los servicios")]
        + ges(C, "Medidas ante las recomendaciones de evaluaciones anteriores.", "ENTREV-OGMEPGD-011", "ENTREV-OGMEPGD-016",
              "ENTREV-OGPPM-017", "ENTREV-DGIGND-050", ded="DED-01"),
        "Con respaldo parcial", estado=VALIDADO_0711,
        detalle=f"De {len(rec)} recomendaciones de las cuatro evaluaciones anteriores, {est.get('Implementada', 0)} "
                f"están implementadas, {est.get('Parcialmente implementada', 0)} parcialmente, "
                f"{est.get('No implementada', 0)} no implementadas y {est.get('Sin evidencia', 0)} sin evidencia. "
                f"Avanzaron las de implementación y no las de diseño ({dis.get('No implementada', 0)} de {dis.sum()} no "
                f"implementadas).",
        salvedad="Las entrevistas no aportaron medios de verificación de otras medidas, por lo que las recomendaciones "
                 "sin evidencia se mantienen así. La evaluación de implementación 2020 (MIMP, 2021) se identificó "
                 "después de las entrevistas, por lo que sus recomendaciones se valoran solo con la revisión documental.",
        nota="La pregunta sobre las medidas ante las recomendaciones se formuló solo en parte, a la OGMEPGD; su "
             "contenido se dedujo de las respuestas (Anexo 8).")
    add("I4-H1", "I4", "Rectoría y conducción según el nivel conductor",
        "La lectura de la rectoría difiere entre los órganos del nivel conductor, según su función. La DGIGND destaca "
        "como fortaleza la asistencia técnica para elaborar los protocolos y aclara que la conducción operativa de la "
        "Política corresponde a la Dirección General, mientras que la DPIGND tiene a su cargo la formulación, el diseño "
        "y la asistencia técnica. La OGMEPGD valora como fluida la coordinación con la dirección conductora, con la que "
        "separó el reporte de la PNIG del de la LIO. La OGPPM la ve «bastante débil»: la dirección conductora se adecua "
        "a lo que proponen los sectores en lugar de fijar el cambio que busca. La OGMEPGD y la OGPPM coinciden en que la "
        "oficina de planeamiento actuó sobre todo como puente con el CEPLAN.",
        ges(C, "Valoración de la rectoría y de la conducción.", "ENTREV-DGIGND-004", "ENTREV-DGIGND-029", "ENTREV-OGMEPGD-007",
            "ENTREV-OGMEPGD-008", "ENTREV-OGMEPGD-010", "ENTREV-OGPPM-007", "ENTREV-OGPPM-008", "ENTREV-OGPPM-020"),
        "Con respaldo parcial", nota="Hallazgo surgido de las entrevistas al nivel conductor. Las lecturas difieren por la "
        "función de cada órgano: conducción técnica, seguimiento y planeamiento.",
        detalle="La lectura de la rectoría difiere según la función de cada órgano: la DGIGND destaca la asistencia "
                "técnica para los protocolos, la OGMEPGD valora la coordinación como fluida y la OGPPM ve la rectoría "
                "«bastante débil».")

    # I5
    add("I5-H1", "I5", "Capacidades institucionales y rotación",
        "La rotación de personal es la brecha de capacidad más persistente: en 2021, solo el 30 % de los integrantes "
        "de los mecanismos regionales para la igualdad de género tenía formación en género; en 2023 y 2025 la DTEG "
        "atribuyó a la rotación sus dificultades para cumplir la meta de transversalización, y en 2025 la cobertura "
        "de su asistencia técnica nacional se redujo de 21 a 10 entidades. La DGIGND coincide: los cambios "
        "permanentes de personal en las entidades frenaron la elaboración de los protocolos.",
        [_ev("Entidades", "REP_ENT", "Mecanismos regionales.", seg(C, "LIO-2021-5.5.1-03"), "Informes LIO"),
         _ev("Entidades", "REP_ENT", "Transversalización 2023.", seg(C, "LIO-2023-5.3.1-03"), "Informes LIO"),
         _ev("Entidades", "REP_ENT", "Transversalización 2025.", seg(C, "LIO-2025-5.3.1-04"), "Informes LIO"),
         _ev("Entidades", "REP_ENT", "Cobertura de la asistencia técnica 2025.", seg(C, "LIO-2025-5.1.1-01"),
             "Informes LIO"),
         _ev("Entidades", "REP_ENT", "Escasez de profesionales con enfoque de género.", seg(C, "CONAIG-2025-024"),
             "Informes de gobernanza de la CONAIG")]
        + ges(C, "Cambios de personal en las entidades.", "ENTREV-DGIGND-015", "ENTREV-DGIGND-016"),
        "Sólido",
        detalle="La rotación de personal es la brecha de capacidad más persistente: en 2021 solo el 30 % de los "
                "integrantes de los mecanismos regionales tenía formación en género y en 2025 la asistencia técnica "
                "nacional de la DTEG bajó de 21 a 10 entidades.")
    add("I5-H2", "I5", "Riesgos ante la transición",
        "La transición normativa ya afecta la continuidad: un protocolo de servicio quedó a la espera de las acciones "
        "derivadas de la Ley N.° 32535 y la DTEG reportó la incertidumbre del proyecto de nueva ley como su principal "
        "dificultad en 2025. Las entrevistas confirman que la transición ya condiciona la gestión: la nueva ley y "
        "normas sectoriales frenaron protocolos casi consensuados en el Minedu (DGIGND), las acciones de 2026 hacia la "
        "actualización responden a la Ley N.° 32535 (OGMEPGD) y la OGPPM sostuvo unas cuatro reuniones con el CEPLAN "
        "para acelerar la formulación ante el plazo legal del 31 de diciembre. La DGIGND y la OGAJ coinciden en que la "
        "pérdida de vigencia de la PNIG no deja un vacío: las obligaciones de incorporar el enfoque, prestar los "
        "servicios y reportar su cumplimiento nacen de la Ley N.° 32535 (artículos 5, 6.2, 8, 11 y 12) y, en "
        "violencia, de la Ley N.° 30364, y los protocolos conservan su vigencia. En cambio, con la PNIG perdieron "
        "vigencia los decretos que crearon los mecanismos para la igualdad, los lineamientos de transversalización, "
        "la capacitación obligatoria y el Oficial para la Igualdad, por lo que esos mecanismos y Oficiales quedan sin "
        "mandato reglamentario (MIMP, 2026g, 2026i; MIMP, 2026e). La adecuación de los instrumentos de cada entidad "
        "vence el 25 de diciembre de 2026, y la derogación tácita de los decretos debe comunicarse al Ministerio de "
        "Justicia y Derechos Humanos y al Congreso a más tardar el 31 de diciembre de 2026, conforme a la Ley N.° 31772 "
        "(MIMP, 2026g, conclusión 3.12).",
        [_ev("Entidades", "REP_ENT", "Protocolo del servicio 6.3.3.", seg(C, "LIO-2025-6.3.3-02"), "Informes LIO"),
         _ev("Entidades", "REP_ENT", "Incertidumbre normativa.", seg(C, "LIO-2025-5.1.1-03"), "Informes LIO")]
        + ges(C, "Efectos de la transición.", "ENTREV-DGIGND-015", "ENTREV-OGMEPGD-016", "ENTREV-OGPPM-022")
        + [_ev("Documental", "REV_DOC", "Continuidad sin vacío normativo y plazo de adecuación (numerales 2.7.4 a "
               "2.7.6, 2.8.2 y 2.8.3).", "docs/fuentes/pnig/07_transicion_ley_32535/2026_mimp-ogaj_informe-d000926-2026_vigencia-normas-ley-28983.pdf", "Informe de la OGAJ"),
           _ev("Documental", "REV_DOC", "Pérdida de vigencia de los decretos de los mecanismos, los lineamientos, la "
               "capacitación y el Oficial para la Igualdad; continuidad sin vacío y comunicación por la Ley N.° 31772 "
               "(conclusiones 3.7 a 3.12).", "docs/fuentes/pnig/07_transicion_ley_32535/2026_mimp-dgignd_informe-d000041-2026_sustento-vigencia-normas-ley-28983.pdf", "Informes técnicos de la DGIGND"),
           _ev("Documental", "REV_DOC", "Normas de transversalización sin vigencia (conclusiones del informe técnico).",
               "docs/fuentes/pnig/07_transicion_ley_32535/2026_mimp-dteg_informe-tecnico-d000040-2026_sustento-vigencia-normas-ley-28983.pdf", "Informes técnicos de la DGIGND")],
        "Sólido", nota="Las dos evidencias de las entidades provienen del informe LIO 2025.",
        detalle="La transición normativa ya afecta la continuidad: un protocolo quedó a la espera de las acciones "
                "derivadas de la Ley N.° 32535, normas sectoriales frenaron protocolos del Minedu y la formulación de la "
                "nueva política corre contra el plazo legal; con la PNIG perdieron vigencia los mecanismos para la "
                "igualdad y el Oficial para la Igualdad, sin vacío normativo.")
    add("I5-H3", "I5", "Propuestas de ajuste de las entidades",
        f"Las entidades plantearon ajustes que la nueva política debe resolver: "
        f"{C['fie_cod_total'].get('I5.AJSERV', 0)} segmentos proponen ajustar un servicio y "
        f"{C['fie_cod_total'].get('I5.AJIND', 0)} un indicador. COFOPRI pidió evaluar el retiro del servicio 4.2.2 y "
        f"el Programa Juntos señaló que su contribución al servicio 4.4.1 es indirecta.",
        [_ev("Entidades", "REP_ENT", "COFOPRI.", seg(C, "FORM-2025-ISS.04.02.02.01-falla_abajo"),
             "Formulario de seguimiento"),
         _ev("Entidades", "REP_ENT", "COFOPRI ante la CONAIG.", seg(C, "CONAIG-2024-042"),
             "Informes de gobernanza de la CONAIG"),
         _ev("Entidades", "REP_ENT", "Juntos.", seg(C, "LIO-2019-4.4.1-03"), "Informes LIO")],
        "Indicio")
    add("I5-H4", "I5", "Efectos no previstos",
        "La provisión del servicio 4.3.5 generó un efecto no previsto: el horario nocturno de las capacitaciones "
        "recargó la jornada de las artesanas rurales; la consulta de la CONAIG en Piura identificó las mismas barreras "
        "de horario, traslado y cuidado.",
        [_ev("Entidades", "REP_ENT", "MINCETUR.", seg(C, "FORM-2025-ISS.04.03.05.01-efectos_no_previstos"),
             "Formulario de seguimiento"),
         _ev("Entidades", "REP_ENT", "Consulta en Piura.", seg(C, "CONAIG-2025-030"), "Informes de gobernanza de la CONAIG")],
        "Indicio")
    verificar_citas(C)
    _verificar(H)
    return H


def fuentes_independientes(h):
    """Documentos distintos (PARTES), sin contar la entrevista a la OGMEPGD (NO_INDEPENDIENTE); una evidencia que agrega
    varias fuentes primarias, como la revisión de evidencia, cuenta por cada una (nucleo.hallazgos)."""
    return _NH.fuentes_independientes(h, NO_INDEPENDIENTE, _partes, lambda f: ORIGEN.get(f, f))


def _verificar(H):
    """Reglas de consistencia de los hallazgos (nucleo.hallazgos.verificar) con los términos internos del piloto."""
    _NH.verificar(H, PREGUNTAS, fuentes_independientes, INTERNOS)


def valoraciones(C):
    """Valoración por pregunta (escala de la nota, Tabla 9). Plano: equipo evaluador, validada por el usuario."""
    return {
        "D1": ("Parcial", "Las brechas del diagnóstico persisten al cierre (varias se redujeron y la del trabajo no "
               "remunerado se amplió, en una comparación indicativa), por lo que el diagnóstico conserva vigencia empírica; pero, "
               "según la OGAJ, su modelo del problema y su finalidad son incompatibles con la Ley N.° 32535, y dos fuentes "
               "del diagnóstico cambiaron: el empleo informal de la ENAHO se reconstruyó y la ENARES ya no permite "
               "recalcular el índice del OP6.", VALIDADA),
        "D2": ("Parcial", "La situación futura cubre las causas y los efectos del modelo, pero solo dos de sus siete "
               "enunciados se miden con el indicador del OP; de los evaluables, solo la meta de embarazo adolescente "
               "se alcanzaría al 2030, y las metas anuales presentan fallas de estimación.", VALIDADA),
        "D3": ("Parcial", f"La cadena causal está casi completa y la evidencia respalda la mayoría de los tipos de "
               f"servicio, pero más de la mitad de los servicios no tiene indicador aprobado y {C['rev_limitada']} tipos "
               f"tienen evidencia limitada.", VALIDADA),
        "D4": ("Parcial", "Alineamiento alto en el periodo y con los compromisos internacionales, que se debilita con "
               "los instrumentos de la transición; duplicidades con otras políticas. La correspondencia con la Ley N.° 32535 "
               "es la del contenido de los objetivos y servicios; según la OGAJ, el objeto de la Política es "
               "incompatible con la ley.", VALIDADA),
        "D5": ("Insuficiente", f"Las fichas incumplen atributos críticos del estándar (línea de base y "
               f"desagregación), {C['serv_sin_ind_vig']} servicios no tienen indicador aprobado y las "
               f"modificaciones no siguieron el procedimiento formal.", VALIDADA),
        "D6": ("Parcial", "Cuatro de los siete enfoques se operacionalizan, solo el de género con amplitud; el de "
               "derechos humanos se menciona y los de territorialidad y gestión por resultados están ausentes. Ningún "
               "indicador prevé desagregaciones para seguirlos.", VALIDADA),
        "D7": ("Parcial", "El nuevo marco diverge del de la PNIG en el plano conceptual y, según la OGAJ, esa "
               "divergencia alcanza al objeto de la Política, que perdió vigencia el 25 de diciembre de 2025; la "
               "estructura y la mayoría de los servicios pueden continuar con sustento directo en la Ley N.° 32535, como "
               "confirma la matriz de incompatibilidades de la DPIGND, pero la operacionalización del nuevo enfoque "
               "depende de lineamientos aún en elaboración.", VALIDADA),
        "I1": ("Parcial", f"La mayoría de los servicios se implementó parcialmente, con menos de la mitad de "
               f"protocolos aprobados; {C['serv_noimpl']} {pl(C['serv_noimpl'], 'servicio no se implementó', 'servicios no se implementaron')}.", VALIDADA),
        "I2": ("Parcial", f"Los indicadores con valor muestran avances mixtos. {C['op_sinvalor_2025']} de "
               f"{C['op_ind']} no tienen valor oficial 2025; con las estimaciones y los recálculos de esta evaluación para "
               f"{len(C['op_est25'])} de ellos, {queda_sin_valor(C)}.",
               VALIDADA),
        "I3": ("Parcial", "La ejecución fue alta, pero según la DGIGND la conducción no contó con los recursos "
               "necesarios, y su gasto no puede separarse de los demás compromisos de la meta de la dirección "
               "general.", VALIDADA),
        "I4": ("Parcial", "La CONAIG funcionó con regularidad formal, pero con baja asistencia de los titulares y, "
               "según la DGIGND, sesiones sin quórum; para la DGIGND fue el único mecanismo de articulación efectivo, "
               "el PEMIG no cumplió ese papel y la articulación territorial fue declarativa. Las lecturas de la rectoría "
               "difieren entre órganos y las recomendaciones de diseño de las evaluaciones anteriores no se "
               "concretaron. Según la OGAJ, la CONAIG agotó su objeto el 25 de diciembre de 2025, y la PCM ya la registró "
               "como extinguida.", VALIDADA),
        "I5": ("Parcial", "Las capacidades se ven afectadas por la rotación de personal y por una conducción con "
               "pocos recursos, y la transición ya frenó protocolos y condiciona la formulación de la nueva política. "
               "Según la DGIGND y la OGAJ no hay vacío normativo, pero los mecanismos para la igualdad y los Oficiales "
               "para la Igualdad de las entidades quedan sin mandato.", VALIDADA),
    }
