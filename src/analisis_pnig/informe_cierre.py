"""Informe de evaluación de cierre de la PNIG (2019–2025), versión autocontenida (ADR-010).

El informe se entiende sin la nota metodológica, sin el repositorio y sin el equipo evaluador: la metodología completa
está en la sección 2; cada tabla y figura se introduce y se interpreta; cada sección abre con su pregunta de evaluación
y cierra con la respuesta; la línea de respaldo de cada hallazgo va en una nota al pie, y el detalle técnico, en los
anexos. No lleva siglas de instrumentos ni identificadores internos: el rastro técnico (libro, hoja y segmento de cada
evidencia, tabla y figura) queda en 06_informe/rastro_hallazgos.csv y rastro_tablas_figuras.csv, para el anexo técnico,
y la correspondencia entre los nombres de los libros que cita el informe y los del repositorio, en
correspondencia_libros.csv.

Estructura (ronda 33): portada (módulo src/nucleo/portada, con el diseño elegido en config/portada.yaml), índices,
abreviaturas; resumen ejecutivo (abstract extenso, solo texto, dos páginas como máximo);
1. la Política; 2. metodología; 3. diseño; 4. implementación; 5. resultados; 6. conclusiones; 7. recomendaciones;
referencias; Anexos 1 a 11. Contiene lo que exige la Tabla A7.1 de la Guía de seguimiento y evaluación. Cada figura y
cada tabla del cuerpo lleva su entrada, su lectura y su análisis; el informe no presenta al equipo evaluador como
actor: explica los criterios.
Toda cifra y todo hallazgo vienen de sintesis_hallazgos.py; los diagramas, de figuras_informe.py; el marco de la
evidencia, de informe_evidencia.py, y las referencias, el glosario y los documentos de respaldo, de
informe_referencias.py.

Los índices son campos de Word (TOC y SEQ). El script genera el documento dos veces (títulos y leyendas, y luego los
índices completos) y lo termina con Word (actualizar_campos_word.ps1), que actualiza los campos con su paginación, guarda
el archivo y exporta un PDF para la revisión; si Word no responde, quedan en caché las páginas leídas del PDF de
LibreOffice.

Salida: evaluaciones/pnig_2026/06_informe/2026_mimp-ogmepgd_informe-evaluacion-cierre-pnig.docx. Una vez que el
usuario edite el Word a mano, los cambios se hacen sobre ese archivo y no regenerándolo.

Uso (desde la raíz del repositorio; antes, figuras_informe.py):
    uv run --no-project --with pandas --with openpyxl --with python-docx --with pillow --with pyyaml python -X utf8         src/analisis_pnig/informe_cierre.py
"""
import os
import re
import subprocess
import sys
import tempfile

import pandas as pd

sys.path.insert(0, os.path.dirname(__file__))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))   # src/, para el módulo portada
import figuras_informe as FIGI  # noqa: E402
import informe_evidencia as EV  # noqa: E402
import informe_referencias as R  # noqa: E402
import sintesis_hallazgos as S  # noqa: E402
from d2_autoridades_2026 import NOTA_2026  # noqa: E402
from docx_mimp import Documento, fmt_num, lista_es, millones, monto  # noqa: E402
from entrev_guias_entrevista import APLICACION, APLICADA, ORIENTADORAS, SALIDA as GES  # noqa: E402
from nucleo.portada import agregar_portada  # noqa: E402
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
from nucleo.config import insumo  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

SALIDA = f"{EVAL}/06_informe/2026_mimp-ogmepgd_informe-evaluacion-cierre-pnig.docx"
RASTRO = f"{EVAL}/06_informe/rastro_tablas_figuras.csv"
RASTRO_HALLAZGOS = f"{EVAL}/06_informe/rastro_hallazgos.csv"
CORRESPONDENCIA = f"{EVAL}/06_informe/correspondencia_libros.csv"
FIGURAS = f"{EVAL}/04_analisis/d1_brechas/figuras/indice_figuras.csv"
FIG_INF = f"{EVAL}/06_informe/figuras/indice_figuras_informe.csv"
MAC = f"{EVAL}/04_analisis/REV_ALI_coherencia_alineamiento.xlsx"
PSI3V = f"{EVAL}/04_analisis/REV_IND_3_verificacion_presupuestal.xlsx"
PSI = f"{EVAL}/04_analisis/REV_IND_base_indicadores.xlsx"
INVENTARIO = "docs/fuentes/inventario_fuentes.csv"
REGISTRO = "docs/fuentes/registro_fuentes.md"
ENLACES = "docs/fuentes/enlaces_publicos_2026-11-06.tsv"
DEP_LIBRO = "D1_D2_brechas_departamentales.xlsx"
LIO_COLECCION = "https://www.gob.pe/institucion/mimp/colecciones/16-informes-lio"
EVAL_2020 = "https://cdn.www.gob.pe/uploads/document/file/1958363/Evaluacion-implementacion-PNIG-2020.pdf"
from nucleo.entorno import herramienta  # noqa: E402
SOFFICE = herramienta("soffice")  # config/entorno.yaml o variable SOFFICE
PS_WORD = "src/nucleo/actualizar_campos_word.ps1"
PDF_WORD = os.path.join(tempfile.gettempdir(), "informe_cierre_pnig_word.pdf")
PORTADA = _cargar_config()["productos"]["portada"]   # config/portada.yaml en el piloto (config.yaml, productos)
MAPAS_CUERPO = ("mapa_sin_ingresos", "mapa_tasa_actividad", "mapa_informal", "mapa_violencia_fisica_o_sexual_12m")
MAPAS_D2 = ("mapa_participacion_alcaldias", "mapa_participacion_gobiernos_regionales")
FECHA = "7 de noviembre de 2026"
VERSION = "Versión final"
PNIG = "MIMP, 2019"
CONF = {"Sólido": "sólido", "Con respaldo parcial": "con respaldo parcial", "Indicio": "indicio",
        "No determinado": "no determinado"}
# relleno de la celda (claro) y color del marcador (fuerte) de cada nivel de confianza; los mismos de la Figura 1
CONF_FILL = {"Sólido": "E2EFDA", "Con respaldo parcial": "FFF2CC", "Indicio": "FADBD8"}
# verde, ámbar y rojo (ronda 34, decisión del usuario): los mismos de los puntos de la figura de valoración
CONF_COLOR = {"Sólido": "0CA30C", "Con respaldo parcial": "FAB219", "Indicio": "D03B3B"}
VALOR_TXT = {"Satisfactorio": "satisfactoria", "Parcial": "parcial", "Insuficiente": "insuficiente"}
# fuente de cada figura de brechas según la encuesta con que empieza su nota (d1_d2_departamental.py)
ENCUESTA_REF = {"ENAHO": "INEI (2026), microdatos de la ENAHO", "ENDES": "INEI (2026), microdatos de la ENDES",
                "ENA": "INEI (2026), microdatos de la ENA", "EPEN": "INEI (2026), microdatos de la EPEN",
                "ENARES": "INEI (2026), microdatos de la ENARES", "ENUT": "INEI (2026), microdatos de la ENUT"}
LIMPIEZA = [(" y las funciones del repositorio «Indicadores ENDES»", ""), ("repositorio «Indicadores ENDES»", ""),
            ("informes de cumplimiento de la LIO (MIMP, 2018, p. 26, y 2022, p. 103; ",
             "informes LIO 2018, p. 26, y 2022, p. 103 (")]
TITULOS, LEYENDAS = [], []


# ---------------------------------------------------------------------------------------------------------------
# Primera mención: servicios, indicadores y siglas con su nombre (ADR-010)
class Nombrador:
    """En la primera mención del cuerpo, agrega el nombre de cada servicio («1.1.3 (Atención integral…)») y de cada
    indicador, y escribe el nombre completo de cada sigla («Oficina General … (OGPPM)»). Lo ya nombrado no se repite."""

    def __init__(self, servicios, indicadores, siglas):
        self.serv, self.ind, self.siglas = servicios, indicadores, siglas
        self.vistos = set()
        self.activo = False
        cods = sorted(self.serv, key=len, reverse=True)
        self.re_serv = re.compile(r"(?<![\d.\w])(" + "|".join(map(re.escape, cods)) + r")(?![\d])")
        self.re_ind = re.compile(r"\b(IOP\.\d\d\.\d\d|ISS(?:\.\d\d){3,4})\b")
        orden = sorted(self.siglas, key=len, reverse=True)
        self.re_sig = re.compile(r"(?<![\w\-.])(" + "|".join(map(re.escape, orden)) + r")(?![\w\-])")

    def _insertar(self, texto, regex, nombres, clave):
        salida, pos = [], 0
        for m in regex.finditer(texto):
            cod = m.group(1)
            if (clave, cod) in self.vistos or cod not in nombres:
                continue
            antes = texto[max(0, m.start() - 9):m.start()]
            if clave == "serv" and re.search(r"(numeral|Anexo|anexo|p\.|pp\.)\s*$", antes):
                continue
            self.vistos.add((clave, cod))
            # ya nombrado: el nombre (sus primeras palabras) aparece en las 160 letras anteriores o el código va
            # seguido de su nombre entre paréntesis
            inicio_nombre = " ".join(nombres[cod].lower().split()[:4])
            ya = inicio_nombre in texto[max(0, m.start() - 160):m.start()].lower() or texto[m.end():m.end() + 2] == " ("
            entre_par = texto[m.start() - 1:m.start()] == "(" and texto[m.end():m.end() + 1] == ")"
            salida.append((texto[pos:m.end()], {}))
            if not ya:
                salida += ([(", ", {}), (nombres[cod], {"i": True})] if entre_par else
                           [(" (", {}), (nombres[cod], {"i": True}), (")", {})])
            pos = m.end()
        salida.append((texto[pos:], {}))
        return salida

    def _siglas(self, texto):
        salida, pos = [], 0
        for m in self.re_sig.finditer(texto):
            s = m.group(1)
            if ("sig", s) in self.vistos:
                continue
            # autor de una cita («(CEPLAN, 2024a)»): no se expande; la mención siguiente lo hará
            if texto[m.start() - 1:m.start()] == "(" and re.match(r", (?:\d{4}|s\. f\.)", texto[m.end():]):
                continue
            self.vistos.add(("sig", s))
            nombre = self.siglas[s]
            if nombre.lower() in texto[:m.start()].lower():
                continue
            if texto[m.end():m.end() + len(s) + 2] == f" ({s}":
                # «la OCDE (OCDE, 2019)»: el nombre basta, la sigla ya va en la cita
                salida.append((texto[pos:m.start()] + nombre, {}))
                pos = m.end()
                continue
            if texto[m.start() - 1:m.start()] == "(" and texto[m.end():m.end() + 1] == ")":
                salida.append((texto[pos:m.start()] + f"{nombre}, {s}", {}))
            else:
                salida.append((texto[pos:m.start()] + f"{nombre} ({s})", {}))
            pos = m.end()
        salida.append((texto[pos:], {}))
        return salida

    def piezas(self, piezas):
        if not self.activo:
            return piezas
        salida = []
        for pz in piezas:
            texto, fmt = (pz, {}) if isinstance(pz, str) else pz
            if fmt.get("url") or fmt.get("i"):
                salida.append((texto, fmt))
                continue
            for t1, f1 in self._siglas(texto):
                for t2, f2 in self._insertar(t1, self.re_serv, self.serv, "serv"):
                    for t3, f3 in (self._insertar(t2, self.re_ind, self.ind, "ind") if not f2 else [(t2, {})]):
                        if t3:
                            salida.append((t3, {**fmt, **f1, **f2, **f3}))
        return salida


class Informe(Documento):
    """Documento del informe: aplica la primera mención en párrafos, viñetas y celdas del cuerpo."""

    def __init__(self, nombrador, **kw):
        super().__init__(**kw)
        self.N = nombrador
        self._sin_nombrar = False

    def _par(self, piezas, size=None, **kw):
        if not self._sin_nombrar:
            piezas = self.N.piezas(list(piezas))
        return super()._par(piezas, size=size, **kw)

    def _sin(self, fn, *a, **kw):
        previo, self._sin_nombrar = self._sin_nombrar, True
        try:
            return fn(*a, **kw)
        finally:
            self._sin_nombrar = previo

    def h1(self, texto):
        return self._sin(super().h1, texto)

    def h2(self, texto):
        return self._sin(super().h2, texto)

    def h3(self, texto):
        return self._sin(super().h3, texto)

    def fuente(self, texto):
        return self._sin(super().fuente, texto)

    def leyenda(self, texto):
        return self._sin(super().leyenda, texto)

    def figura(self, *a, **kw):
        return self._sin(super().figura, *a, **kw)

    def cita(self, texto, fuente):
        return self._sin(super().cita, texto, fuente)

    def titulo_indice(self, texto):
        return self._sin(super().titulo_indice, texto)

    def tabla(self, encabezado, filas, anchos, nombrar=True, **kw):
        if nombrar and self.N.activo:
            nuevas = []
            for fila in filas:
                nf = []
                for c in fila:
                    if isinstance(c, dict):
                        nf.append({**c, "piezas": self.N.piezas(c["piezas"] if isinstance(c["piezas"], list)
                                                                else [c["piezas"]])})
                    else:
                        nf.append(self.N.piezas(c if isinstance(c, list) else [str(c)]))
                nuevas.append(nf)
            filas = nuevas
        return self._sin(super().tabla, encabezado, filas, anchos, **kw)


# ---------------------------------------------------------------------------------------------------------------
def num(x):
    """Valor de indicador: tres decimales para índices, uno para porcentajes."""
    v = pd.to_numeric(pd.Series([x]), errors="coerce").iloc[0]
    if pd.isna(v):
        return re.sub(r"(\d)\.(\d)", r"\1,\2", str(x)) if str(x) not in ("nan", "None") else "s. d."
    return fmt_num(v, 3 if abs(v) < 2 else 1)


def pie_hallazgo(h):
    """Nota al pie de un hallazgo, de 50 palabras como máximo: identificador, confianza y respaldo."""
    fuentes = S.fuentes_legibles(h)
    base = f"Hallazgo {h['id']} (confianza: {CONF[h['confianza']]}). Respaldo: "
    cola = ". Véase el Anexo 2."
    for k in range(len(fuentes), 0, -1):
        lista = fuentes[:k] if k == len(fuentes) else fuentes[:k] + ["otras fuentes"]
        sep = "; " if any("," in f for f in lista if not f.startswith("entrevista")) else ", "
        txt = base + (sep.join(lista[:-1]) + " y " + lista[-1] if len(lista) > 1 else lista[0]) + cola
        if len(txt.split()) <= 50:
            return txt
    return base + "véase el Anexo 2."


def hallazgo(d, H, ident):
    h = H[ident]
    par = d.p(h["enunciado"], after=40 if h["salvedad"] else None, recuadro="hallazgo")
    d.nota_pie(par, pie_hallazgo(h))
    if h["salvedad"]:
        d.p((h["salvedad"], {"i": True}), size=20)


def abre(d, p):
    crit, texto, _ = S.PREGUNTAS[p]
    d.p((f"Pregunta {p} ({crit.lower()}). ", {"b": True, "color": "1F3864"}), (texto, {"i": True}),
        recuadro="pregunta", after=200)


def cierra(d, p, H, V):
    valor, sust, _ = V[p]
    hs = [h for h in H.values() if h["pregunta"] == p]
    conf = lista_es([f"{h['id']}, {CONF[h['confianza']]}" for h in hs])
    d.p((f"Respuesta a la pregunta {p}. ", {"b": True, "color": "1F3864"}),
        f"Valoración: {VALOR_TXT[valor]}. {sust} Confianza de los hallazgos que la sostienen: {conf}.",
        recuadro="respuesta", after=240)


def marcador(conf):
    """Celda con el nivel de confianza: relleno claro, marcador de color y el texto del nivel."""
    return {"piezas": [("■ ", {"color": CONF_COLOR[conf]}), (CONF[conf], {})], "fill": CONF_FILL[conf]}


def leyenda_confianza():
    return [("Confianza: ", {"b": True})] + [x for k in CONF_FILL for x in
                                             (("■ ", {"color": CONF_COLOR[k]}), (f"{CONF[k]}   ", {}))]


# ---------------------------------------------------------------------------------------------------------------
def preparar(C):
    """Conteos y tablas que usan las secciones descriptivas, la revisión de evidencia y los anexos."""
    g = C["grd"]
    entidades = {("MIMP" if x.strip().startswith("MIMP") else x.strip())
                 for v in g["Entidad u órgano responsable"] for x in str(v).split(";")}
    normas = [x for v in g["Norma del protocolo"].dropna() for x in str(v).split("|") if x.strip()]
    rev_x = f"{EVAL}/04_analisis/REV_EVI_matriz_evidencia.xlsx"
    ft = pd.read_excel(rev_x, sheet_name="Flujo por tipo")
    ft = ft[ft.Tipo.astype(str).str.match(r"T\d\d")]
    fd = pd.read_excel(rev_x, sheet_name="Flujo 3ie DEP")
    fm = pd.read_excel(rev_x, sheet_name="Flujo manual")
    ea = pd.read_excel(rev_x, sheet_name="Estrategias académicas")
    bases = ea.base.value_counts()
    rev = {"academica": int(ft.Estrategias.sum()), "dep": int(fd.Estrategia.astype(str).str.match(r"DEP-\d+").sum()),
           "manual": int(fm.Estrategia.astype(str).str.match(r"MAN-\d+").sum()),
           "web": len(pd.read_excel(rev_x, sheet_name="Búsquedas web y portales")),
           "bases": lista_es([f"{b} ({n})" for b, n in bases.items()])}
    assert rev["academica"] == int(bases.sum()), "estrategias académicas: el flujo y el registro difieren"
    tip, mt, me, ref = EV.datos()
    resumen_res = int(ref["Verificación"].astype(str).str.contains("resultado tomado del resumen").sum())
    obj = pd.read_excel(MAC, sheet_name="REV_ALI.1 Por objetivo").set_index("OP")
    # modelo del problema: servicios por causa o efecto según la matriz de consistencia (Anexo 5 de la PNIG)
    mod = pd.read_excel(MAC, sheet_name="REV_ALI.1 Modelo causal")
    ps = pd.read_excel(MAC, sheet_name="REV_ALI.1 Por servicio", dtype={"Código": str})
    vinculo = ps["Tipo de vínculo"].value_counts().to_dict()
    # indicadores del PEMIG por tipo y vigentes de servicio por OP (mismos datos que la figura de indicadores)
    tp = pd.read_excel(f"{EVAL}/04_analisis/REV_IND_2_calidad_fichas.xlsx", sheet_name="Tipo PEMIG")
    srv = pd.read_csv("src/formulario_fie03/datos/servicios.csv", dtype=str).fillna("")
    vig_op = srv.assign(n=srv.indicadores.str.count(r"ISS\.")).groupby("op").n.sum()
    # alineamiento con la Ley N.° 32535 y los compromisos internacionales (mismas columnas que la figura)
    m4 = pd.read_excel(MAC, sheet_name="REV_ALI.4 Por objetivo").set_index("OP")
    cols4 = ["Ley N.° 32535", "CEDAW", "Observaciones finales del Comité CEDAW (2022)", "Convención de Belém do Pará",
             "Plataforma de Acción de Beijing", "ODS 5 (Agenda 2030)", "Estrategia de Montevideo"]
    tipos4 = pd.Series([FIGI._tipo_corresp(m4.loc[o, c]) for o in m4.index for c in cols4])
    # montos reportados que coinciden con una estructura del MEF (mismo filtro que rev_doc_estado_servicios.py; desde la
    # ronda 33, todo el periodo 2019–2025, ADR-011)
    cv = pd.read_excel(PSI3V, sheet_name="Cifras verificadas", dtype={"Servicio": str})
    cv = cv[cv.Servicio.notna() & cv["Año"].between(2019, 2025)]
    cv = cv[cv.Resultado.isin({"Coincide", "Cercano (≤ 0,5 % del mismo agregado)", "Coincide con el año anterior"})
            & (cv.Alcance == "Entidad responsable")
            & cv.Nivel.fillna("").str.startswith(("Meta", "Metas del mismo nombre", "Actividad",
                                                  "Unidad ejecutora y actividad", "Producto"))
            & ~cv["Observación"].fillna("").str.contains("casual")]
    base = pd.read_excel(PSI, sheet_name="Base REV_IND.1")
    return {"n_entidades": len(entidades), "n_protocolos": len(normas), "n_corpus": len(pd.read_csv(INVENTARIO)),
            "rev": rev, "obj": obj, "tip": tip, "mt": mt, "me": me, "ref": ref, "resumen_res": resumen_res,
            "montos": cv, "base_ind": base, "mod": mod, "vinculo": vinculo, "tp": tp, "vig_op": vig_op,
            "int4": tipos4.value_counts().to_dict(), "int4_n": len(tipos4)}


def tabla_departamental(d, C):
    """Resumen por indicador mapeado: departamentos confiables, con brecha en desventaja de las mujeres y con cambio."""
    r = C["dep_resumen"]
    r = r[r.Mapa == "Sí"]
    filas = []
    for _, x in r.iterrows():
        nivel = x.Tipo.startswith("Nivel")
        reduce = x.get("Cambio: mejora") if nivel else x.get("Cambio: se reduce")
        amplia = x.get("Cambio: empeora") if nivel else x.get("Cambio: se amplía")
        desv = x.get("Brecha significativa en desventaja de las mujeres")
        n = lambda v: str(int(v)) if pd.notna(v) else "n. a."
        nombre = x.Indicador if f"({x.Encuesta})" in x.Indicador else f"{x.Indicador} ({x.Encuesta})"
        filas.append([nombre, x["Años comparados"],
                      f"{int(x['Confiables (último año)'])} de {int(x['Ámbitos con estimación (último año)'])}",
                      "n. a." if nivel else n(desv), n(reduce), n(x.get("Cambio: sin cambio significativo")), n(amplia)])
    d.tabla(["Indicador (encuesta)", "Años", "Confiables", "Desventaja de las mujeres", "Se reduce o mejora",
             "Sin cambio", "Se amplía o empeora"], filas, [2926, 1000, 1100, 1100, 1000, 900, 1000], size=14,
            leyenda="Brechas por departamento: confiabilidad, brecha en 2025 y cambio desde 2019",
            rastro=f"{DEP_LIBRO}, hoja «Resumen por indicador»",
            fuente=f"Fuente: INEI (2026), microdatos de la {lista_es(sorted(r.Encuesta.unique()))}. Elaboración propia.",
            nota="Confiables: departamentos en los que las estimaciones de mujeres y de hombres tienen un coeficiente "
                 "de variación de 15 % o menos; la significancia de la brecha se prueba con su propio error estándar. "
                 "Desventaja de las mujeres: departamentos con brecha significativa al 95 % en el último año. Cambio: "
                 "2019 frente a 2025, o los años más cercanos con datos; el análisis es exploratorio. Propietarios de "
                 "alguna parcela y propietarios con título (ENA): comparación indicativa, porque la ENA cambió el capítulo "
                 "de parcelas en 2023. En los indicadores de la ENDES, que solo encuesta a mujeres, el cambio es mejora o "
                 "empeora. n. a.: no aplica.", nombrar=False)
    return r


# nombre completo de cada indicador mapeado (ronda 35, pedido del usuario), tomado de las definiciones de los scripts
# d1_brechas_*.py; el resumen departamental usa nombres cortos
NOMBRE_INDICADOR = insumo("nombres_indicadores.yaml")


def frase_mapa(x):
    """Lectura de un mapa con el resumen departamental, con el nombre completo del indicador como sujeto:
    confiabilidad, desventaja de las mujeres y cambio entre los años comparados (rondas 33 y 35)."""
    nombre = NOMBRE_INDICADOR[x.name if x.name in NOMBRE_INDICADOR else x["Clave"]]
    nivel = str(x.Tipo).startswith("Nivel")
    tot, conf = int(x["Ámbitos con estimación (último año)"]), int(x["Confiables (último año)"])
    v = lambda k: int(x[k]) if pd.notna(x.get(k)) else 0
    comp = str(x["Años comparados"])
    if "–" not in comp:       # una sola medición departamental: no hay comparación en el tiempo
        return (f"La estimación de la brecha en «{nombre}» es confiable en {conf} de los {tot} departamentos. Solo "
                f"hay medición departamental de {comp}, de modo que no se compara en el tiempo; en ese año, el valor es "
                f"significativamente mayor en los hombres en {v('Brecha significativa, mayor en hombres')} "
                f"departamentos y en las mujeres en {v('Brecha significativa, mayor en mujeres')}, y la diferencia no "
                f"es significativa en {v('Brecha no significativa')}.")
    a0, a1 = comp.split("–")
    azar = (f" (por azar se esperarían {fmt_num(x['Esperados por azar (5 %)'], 1)} cambios significativos; el "
            f"análisis es exploratorio y la falta de significancia no prueba que no haya cambio)")
    if nivel:
        return (f"La estimación de «{nombre}» es confiable en {conf} de los {tot} departamentos; entre {a0} y {a1}, el "
                f"indicador mejoró de forma significativa en {v('Cambio: mejora')}, se mantuvo sin cambio "
                f"significativo en {v('Cambio: sin cambio significativo')} y empeoró en {v('Cambio: empeora')}"
                f"{azar}.")
    return (f"La estimación de la brecha en «{nombre}» es confiable en {conf} de los {tot} departamentos. En {a1}, la "
            f"brecha es significativa en desventaja de las mujeres en "
            f"{v('Brecha significativa en desventaja de las mujeres')} "
            f"{S.pl(v('Brecha significativa en desventaja de las mujeres'), 'departamento', 'departamentos')}; entre {a0} y {a1} se redujo en "
            f"{v('Cambio: se reduce')}, se mantuvo sin cambio significativo en {v('Cambio: sin cambio significativo')} "
            f"y se amplió en {v('Cambio: se amplía')}{azar}.")


def figura_brecha(d, F, clave, ancho=16.0, alt=None):
    """Figura de brechas: la nota del índice empieza por la encuesta, que pasa a la fuente; el resto queda como nota."""
    r = F.loc[clave]
    nota = str(r.nota)
    for a, b in LIMPIEZA:
        nota = nota.replace(a, b)
    m = re.match(r"^(ENAHO|ENDES|ENA|EPEN|ENARES|ENUT)\.\s*", nota)
    anios = " y ".join(re.findall(r"(?:19|20)\d\d", r.titulo)[-2:]) or ""
    if m:
        fuente = f"Fuente: {ENCUESTA_REF[m.group(1)]}{' ' + anios if anios else ''}. Elaboración propia."
        nota = nota[m.end():]
    elif clave == "series_endes":
        fuente = "Fuente: INEI (2026), microdatos de la ENDES 2019 a 2025. Elaboración propia."
    elif clave == "series_otras":
        fuente = "Fuente: INEI (2026), microdatos de la ENA, la EPEN, la ENARES y la ENUT. Elaboración propia."
    elif clave == "series_enaho":
        fuente = "Fuente: INEI (2026), microdatos de la ENAHO 2016 a 2025; INEI (2025c). Elaboración propia."
        nota = re.sub(r"^ENAHO;\s*", "", nota)
    else:
        fuente = ("Fuente: JNE y ONPE, compilados por INEI (2025a); JNE (2026a, 2026b); ONPE (2026a, 2026b); "
                  "MIMP (2020-2026). Elaboración propia.")
    d.figura(r.archivo, r.titulo, fuente=fuente, ancho_cm=ancho, rastro=f"{DEP_LIBRO}; {r.archivo.split('/')[-1]}",
             alt=alt, nota=(nota[0].upper() + nota[1:]) if nota.strip() else None)


def fig(d, FI, clave, ancho=16.0, alt=None, nota=None):
    r = FI.loc[clave]
    d.figura(r.archivo, r.titulo, fuente=r.fuente, ancho_cm=ancho, rastro=r.rastro, alt=alt, nota=nota)


def _listas(C):
    """Limitaciones, lecciones y temas vigentes, que se citan en más de una sección."""
    limitaciones = (
        "La OMEP, que condujo esta evaluación, también participa en el seguimiento de la Política: aplica el "
        "formulario de seguimiento a las entidades y consolida los indicadores que la evaluación analiza. Además, forma "
        "parte de la OGMEPGD, uno de los tres órganos entrevistados. Para acotar ese conflicto de interés se tomaron tres "
        "medidas: la entrevista a la OGMEPGD no se cuenta como fuente independiente al clasificar la confianza de los "
        "hallazgos (con esa regla ningún hallazgo cambió de nivel, porque todos los que la citan tienen al menos otras "
        "dos fuentes independientes); el consolidado de indicadores y el formulario se usan como registro de lo que "
        "reportaron las entidades y se contrastan con los reportes de seguimiento y de cumplimiento, las encuestas del "
        "INEI y los datos del MEF, y los hallazgos y las valoraciones se someten a la validación de la DGIGND, del "
        "Despacho Viceministerial de la Mujer y del CEPLAN.",
        "La información de las entidades proviene de lo que reportaron al MIMP entre 2019 y 2025, sin verificación en "
        "el punto de prestación ni consulta a la población usuaria o a los gobiernos regionales y locales.",
        "La asignación presupuestal de los servicios solo se verifica cuando el monto reportado coincide con una meta, "
        "actividad o producto de los datos abiertos del MEF; esos datos no identifican el gasto por política.",
        "La calidad de los hallazgos sobre avance depende de la calidad del reporte; los vacíos se consignan como "
        "hallazgo y no se estiman, salvo los indicadores de OP que pudieron estimarse con el método que reproduce sus "
        "valores oficiales (numeral 5.2).",
        "La revisión de evidencia es rápida y no sistemática: una sola persona hizo la selección de documentos, se revisaron los "
        "primeros registros por relevancia de cada búsqueda y los resultados se tomaron del resumen de los estudios. "
        "Su aplicabilidad al Perú se valora caso por caso.",
        f"Las entrevistas al nivel conductor no cubrieron todas las preguntas de la guía. Los hallazgos "
        f"{lista_es(C['entrev_sin_contraste'])}, cuyo contraste estaba previsto, no se contrastaron con el nivel "
        f"conductor. Las medidas ante las recomendaciones de evaluaciones anteriores y las limitaciones presupuestales "
        f"de la conducción se dedujeron de otras respuestas. La entrevista de la DGIGND fue grupal y se atribuye al "
        f"órgano.",
        "La evaluación de implementación 2020 (MIMP, 2021) se identificó después de las entrevistas: sus "
        "recomendaciones (numeral 4.7) se valoran solo con la revisión documental.",
        "Algunas fuentes no son comparables en el tiempo: la ENUT cambió de instrumento entre 2010 y 2024 y la ENARES "
        "cambió las afirmaciones de tolerancia, de modo que su índice oficial no se reproduce con los microdatos. La "
        "ENAHO dejó de publicar la variable de empleo informal en 2024; se reconstruyó con la regla del INEI "
        "(INEI, 2025c), que reproduce la variable oficial de 2019 a 2023.",
        "Los errores estándar de la ENA son aproximados, porque el INEI no publica su estrato geográfico. La "
        "participación política se toma de la serie del JNE y la ONPE compilada por el INEI hasta 2022, de los informes "
        "LIO para las gobernadoras y, para 2026, de las resoluciones del JNE (Congreso) y de los resultados "
        "preliminares de la ONPE (alcaldías y gobiernos regionales), con el sexo identificado por el nombre.",
        "Por departamento, varias estimaciones son referenciales (coeficiente de variación mayor a 15 %): la ENDES no "
        "permite estimar con confiabilidad la violencia de pareja, el embarazo adolescente ni la demanda insatisfecha "
        "de planificación familiar en 2025; la EPEN no separa el Callao de Lima Metropolitana, y la ENUT 2024 y la "
        "ENARES 2019 no tienen inferencia por departamento.",
        "Los valores 2025 de los indicadores de los OP1 y OP2 son cálculos de esta evaluación con los microdatos de "
        "la ENDES; reproducen los valores oficiales de 2019 y de 2021 a 2024 (en 2020 no, porque la ENDES se levantó "
        "en buena parte por teléfono; numeral 5.2), pero no son valores reportados por las entidades.",
        "Los segmentos de las entrevistas los validó el mismo especialista que las condujo: es un control de "
        "consistencia, no una doble codificación independiente.",
    )
    lecciones = (
        "Una política que agrega servicios de muchas entidades necesita, desde su aprobación, protocolos e indicadores "
        "por servicio; sin ellos, más de la mitad de los servicios no pudo seguirse.",
        "Las metas fijadas sin línea de base ni fuente asegurada produjeron fallas de planeación y años sin dato.",
        "La rectoría formal (sesiones y recomendaciones de la CONAIG) no bastó para que las entidades adoptaran las "
        "recomendaciones de gobernanza ni aprobaran sus protocolos: con baja asistencia de los titulares, los avances "
        "dependieron de coordinaciones de alto nivel caso por caso.",
        "Un plan de implementación sin seguimiento propio no ordena la implementación: el PEMIG sirvió para exigir "
        "protocolos, pero no para articular a las entidades ni para medir.",
        "Los problemas de diseño deben resolverse antes de exigir el seguimiento: sin servicios con protocolo e "
        "indicadores medibles, el seguimiento no puede valorar el aporte de los servicios a la Política.",
        "Las capacidades en género se pierden con la rotación de personal si no se institucionalizan en perfiles, "
        "formación y asistencia técnica continua.",
    )
    temas = ("retomar en el diagnóstico las brechas que persisten, en especial la carga de trabajo no remunerado, "
             "la participación laboral y los ingresos de las mujeres;",
             "dar a cada servicio de la nueva política causa, protocolo e indicadores de cobertura y calidad con "
             "línea de base, desagregación y fuente financiada;",
             "resolver las duplicidades con otras políticas nacionales y explicitar el alineamiento con la PGG "
             "2026–2031;",
             "definir, antes de aprobar la nueva política, el destino de los protocolos en trámite y de los "
             "servicios no implementados;",
             "conservar la documentación de seguimiento de la PNIG para la formulación y el seguimiento de la nueva "
             "política.",)
    return limitaciones, lecciones, temas


def recomendaciones(C, nombre_serv):
    """Matriz de recomendaciones (CEPLAN, Tabla 4). Plano: equipo evaluador, validada por el usuario (2026-11-06)."""
    impl = "; ".join(f"{c} ({nombre_serv[c]})" for c in C["serv_impl_cod"])
    noimpl = "; ".join(f"{c} ({nombre_serv[c]})" for c in C["serv_noimpl_cod"])
    noimpl_txt = (f"del servicio no implementado ({noimpl})" if len(C["serv_noimpl_cod"]) == 1 else
                  f"de los servicios no implementados ({noimpl})")
    return [
        ["Diseño: temas vigentes", "Retomar en el diagnóstico de la nueva política las brechas que persisten: la "
         "carga de trabajo no remunerado (la única que se amplió, en una comparación indicativa), la menor "
         "participación laboral y los menores ingresos de las mujeres, la violencia de pareja, la demanda insatisfecha "
         "de planificación familiar y la tolerancia social a la violencia, que es mayor entre los hombres.", "DGIGND",
         "Antes de la aprobación de la nueva política", "D1, D2, I2"],
        ["Diseño: enfoques", "Operacionalizar en la nueva política los enfoques que la PNIG solo mencionó o no incorporó "
         "(derechos humanos, territorialidad, gestión por resultados) y extender los demás a lineamientos, estándares e "
         "indicadores con desagregación.", "DGIGND", "Antes de la aprobación de la nueva política", "D6"],
        ["Diseño: coherencia", "Asegurar que cada lineamiento de la nueva política tenga al menos un servicio y que cada "
         "servicio tenga causa asociada en el modelo, protocolo e indicadores de cobertura y de calidad.",
         "DGIGND, con la OGPPM", "Antes de la aprobación de la nueva política", "D3, I1"],
        ["Diseño: evidencia", "Priorizar los servicios de tipos con evidencia sólida y, para los de evidencia limitada "
         "(hostigamiento sexual laboral, tratamiento de agresores sexuales, estadísticas de género), prever una "
         "evaluación con medición desde su diseño, sin condicionar su continuidad cuando atienden obligaciones "
         "institucionales o necesidades de la población; en las estadísticas y los estudios, valorar también su aporte "
         "al sistema de información.", "DGIGND, con la OGMEPGD",
         "Antes de la aprobación de la nueva política", "D3"],
        ["Diseño: priorización", "Acotar el problema público y los servicios de la nueva política a lo que es posible "
         "lograr en su horizonte y con los recursos disponibles, sin incluir servicios cuyo alcance no basta para el "
         "objetivo que atienden.", "DGIGND, con la OGPPM", "Antes de la aprobación de la nueva política", "D2, D3"],
        ["Diseño: articulación horizontal", "Ordenar los servicios que la PNIG comparte con otras políticas "
         "nacionales con el mismo proveedor, como los Centros Emergencia Mujer o los lactarios: comprobar si hay "
         "duplicidad de población, prestaciones y recursos, y definir reglas de registro y de conteo que eviten el doble "
         "conteo sin perder la contribución de cada política, en coordinación con el CEPLAN y los rectores de esas "
         "políticas.", "DGIGND, con la OGPPM",
         "Antes de la aprobación de la nueva política", "D4, I4"],
        ["Diseño: alineamiento", f"Explicitar el alineamiento de la nueva política con la PGG 2026–2031 y el PESEM "
         f"2024–2030, en particular en los ámbitos de {lista_es(C['pgg127_sin_op'])}.", "DGIGND, con la OGPPM",
         "Antes de la aprobación de la nueva política", "D4"],
        ["Seguimiento y evaluación: indicadores", "Elaborar las fichas de la nueva política con línea de base, "
         "desagregación por territorio y grupo, sentido coherente y fuente con financiamiento asegurado y comparable en "
         "el tiempo (ENARES con afirmaciones estables, empleo informal con la EPEN desde su propia línea de base), "
         "medir la situación futura deseada con los mismos indicadores de los OP, y modificar los indicadores solo con "
         "el procedimiento formal de la Guía de Políticas Nacionales.",
         "DGIGND, OGPPM y OGMEPGD", "Antes de la aprobación de la nueva política y en cada modificación de fichas", "D5, I2"],
        # ronda 36 (pedido del usuario): los servicios implementados van en una nota al pie, no en la celda
        ["Implementación: componentes a preservar", [
            ("Preservar en la nueva política, con sus protocolos y previa verificación de su necesidad y su "
             "desempeño, los servicios implementados",
             {"pie": f"Servicios implementados al cierre ({len(C['serv_impl_cod'])}): {impl}."}),
            ". Según la OGAJ, sus contenidos compatibles con la Ley N.° 32535 siguen ejecutándose con sustento directo "
            "en ella (MIMP, 2026e). Incorporar lo principal de cada protocolo en la ficha de servicio, para que el "
            "servicio no cambie con cada "
            "gestión, y ajustar los que el nuevo marco conceptual obliga a redefinir, como los vinculados a la educación "
            "sexual y a la salud integral."], "DGIGND", "Antes de la aprobación de la nueva política", "I1, D7"],
        ["Implementación: transición", f"Definir, antes de aprobar la nueva política, el destino de los protocolos en "
         f"trámite y {noimpl_txt}, teniendo en cuenta que los protocolos aprobados conservan su vigencia y se leen "
         f"conforme a la Ley N.° 32535 (MIMP, 2026e), y comunicarlo a las entidades responsables.", "DGIGND",
         "Antes de la aprobación de la nueva política", "I1, I5"],
        ["Implementación: transición normativa", "Usar la matriz de incompatibilidades de la DPIGND como insumo del "
         "diseño de la nueva política y de la continuidad de los servicios; informar al CEPLAN el criterio de la OGAJ "
         "sobre la pérdida de vigencia de la PNIG, como recomienda esa oficina (MIMP, 2026e); comunicar a las entidades "
         "qué contenidos siguen ejecutándose y cuáles perdieron vigencia, y comunicar la derogación tácita de los "
         "decretos al Ministerio de Justicia y Derechos Humanos y al Congreso, conforme a la Ley N.° 31772 "
         "(MIMP, 2026g).", "DGIGND, con la OGPPM y la OGAJ",
         "Informe al CEPLAN y comunicación a las entidades: de inmediato; Ley N.° 31772: hasta el 31 de diciembre de "
         "2026; matriz: antes de la aprobación de la nueva política", "D7, I5"],
        ["Implementación: asignación de recursos", "Identificar en la estructura presupuestal (meta, actividad o "
         "producto) el gasto de cada servicio y de la conducción, de modo que su asignación pueda verificarse con datos "
         "públicos, y dimensionar el equipo de la conducción a la asistencia técnica y al seguimiento que exige la "
         "nueva política.", "OGPPM, con la DGIGND y las entidades responsables", "Programación multianual siguiente",
         "I1, I3"],
        ["Implementación: coordinación", "Prever en la nueva política el mecanismo de articulación multisectorial "
         "que asuma las funciones de la CONAIG, cuyo objeto se agotó con la PNIG (MIMP, 2026e), con titulares o "
         "alternos con capacidad de decisión y un seguimiento de sus acuerdos, y un mecanismo de articulación con los "
         "gobiernos regionales y locales.",
         "DGIGND", "Antes de la aprobación de la nueva política", "I4"],
        ["Implementación: capacidades", "Incorporar en la asistencia técnica de la nueva política medidas para retener "
         "capacidades ante la rotación de personal, como perfiles de puesto con enfoque, formación certificada y "
         "acompañamiento continuo, y orientar a las entidades sobre la continuidad o la adecuación de sus mecanismos "
         "para la igualdad y de sus Oficiales para la Igualdad, que desde el 25 de diciembre de 2025 no tienen mandato "
         "reglamentario (MIMP, 2026e).", "DGIGND", "Antes del 25 de diciembre de 2026 (plazo de adecuación)", "I5"],
        ["Seguimiento y evaluación: compromisos", "Elaborar con las entidades responsables la matriz de compromisos "
         "de esta evaluación (Tabla 6 de la Guía de seguimiento y evaluación) y registrar su cumplimiento; trasladar a "
         "la formulación de la nueva política las recomendaciones de diseño pendientes de las evaluaciones anteriores "
         "(protocolos con estándares de calidad e indicadores con línea de base).",
         "OGMEPGD, con la DGIGND", "Dentro de los 30 días de aprobado el informe de evaluación", "I4"],
        ["Seguimiento y evaluación: gestión del conocimiento", "Conservar la documentación de la PNIG (series 2019–2025, "
         "formularios, actas de validación, base de evidencia y matrices de esta evaluación) en un archivo accesible "
         "para la formulación y el seguimiento de la nueva política.", "OGMEPGD",
         "Antes de la aprobación de la nueva política", "Transversal"],
    ]


# ---------------------------------------------------------------------------------------------------------------
def portada(d):
    """Portada: módulo src/nucleo/portada con su único diseño, el centrado (G), y los datos de config/portada.yaml. Es la
    primera sección del documento, A4, sin encabezado, pie ni número de página."""
    info = agregar_portada(d.d, PORTADA)
    print(f"Portada: diseño {info['opcion']} ({info['diseno']}), letra {info['fuente']}, contraste del blanco sobre "
          f"el rojo {info['contraste_blanco_primario']}:1")
    for aviso in info["avisos"]:
        print("  aviso de la portada:", aviso)
    return info


# fuentes que vienen con Windows u Office: Word las incrusta junto con Poppins (Calibri y Calibri Light, del tema de la
# plantilla, suman unos 3 MB), pero cualquier equipo con Office ya las tiene
FUENTES_DE_OFFICE = ("Calibri", "Calibri Light", "Cambria", "Arial", "Times New Roman", "Symbol", "Wingdings")


def quitar_fuentes_de_office(docx_path):
    """Quita del .docx las fuentes de Office incrustadas por Word y deja las demás (Poppins de la portada; ronda 27).
    Devuelve los nombres de las fuentes que quedan incrustadas."""
    import zipfile
    with zipfile.ZipFile(docx_path) as z:
        partes = [(i, z.read(i.filename)) for i in z.infolist()]
    datos = {i.filename: b for i, b in partes}
    if "word/_rels/fontTable.xml.rels" not in datos:
        return []
    tabla = datos["word/fontTable.xml"].decode("utf-8")
    rels = datos["word/_rels/fontTable.xml.rels"].decode("utf-8")
    quitar, quedan = set(), []
    for m in re.finditer(r'<w:font w:name="([^"]+)">(.*?)</w:font>', tabla, re.S):
        ids = re.findall(r'<w:embed\w+ r:id="([^"]+)"[^>]*/>', m.group(2))
        if not ids:
            continue
        if m.group(1) in FUENTES_DE_OFFICE:
            quitar.update(ids)
        else:
            quedan.append(m.group(1))
    if not quitar:
        return quedan
    destinos = set()
    for rid in quitar:
        rel = re.search(r'<Relationship [^>]*Id="%s"[^>]*/>' % rid, rels).group(0)
        destinos.add("word/" + re.search(r'Target="([^"]+)"', rel).group(1))
        rels = rels.replace(rel, "")
    tabla = re.sub(r'<w:embed\w+ r:id="(%s)"[^>]*/>' % "|".join(sorted(quitar)), "", tabla)
    datos["word/fontTable.xml"] = tabla.encode("utf-8")
    datos["word/_rels/fontTable.xml.rels"] = rels.encode("utf-8")
    tmp = docx_path + ".tmp"
    with zipfile.ZipFile(tmp, "w") as z:
        for i, _ in partes:
            if i.filename not in destinos:
                z.writestr(i, datos[i.filename], compress_type=i.compress_type)
    os.replace(tmp, docx_path)
    return quedan


def actualizar_con_word(docx_path, pdf_path):
    """Actualiza los campos con Word (paginación real), guarda el .docx y exporta el PDF; devuelve (ok, salida)."""
    try:
        r = subprocess.run(["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", PS_WORD,
                            "-Docx", os.path.abspath(docx_path), "-Pdf", os.path.abspath(pdf_path)],
                           capture_output=True, text=True, timeout=1200)
    except (OSError, subprocess.TimeoutExpired) as e:
        return False, str(e)
    return r.returncode == 0 and os.path.exists(pdf_path), (r.stdout + r.stderr).strip()


# ---------------------------------------------------------------------------------------------------------------
def construir(paginas=None):
    """Arma el informe. paginas: {(«t», título) o (rótulo, número): página}, de la vuelta anterior, para los índices."""
    paginas = paginas or {}
    C = S.cifras()
    H = {h["id"]: h for h in S.hallazgos(C)}
    F = pd.read_csv(FIGURAS).set_index("clave")
    E = pd.read_csv(ENLACES, sep="\t").set_index("documento")
    V = S.valoraciones(C)
    FI = pd.read_csv(FIG_INF).set_index("clave")
    P = preparar(C)
    refx = P["ref"]
    _, LECCIONES, TEMAS = _listas(C)
    for org, f in S.FECHA_ENTREV.items():
        iso = APLICADA[org]["fecha"]
        assert f.startswith(str(int(iso[8:10]))) and "noviembre" in f and iso[5:7] == "11", f"fecha de {org}: {iso}"
    g = C["grd"].copy()
    g["Código"] = g["Código"].astype(str)
    nombre_serv = {r["Código"]: re.sub(r"\s+", " ", str(r["Servicio"])).strip().rstrip(".") for _, r in g.iterrows()}
    base = P["base_ind"]
    nombre_ind = {}
    for _, r in base.iterrows():
        t = re.sub(r"\s+", " ", str(r["Indicador"])).strip().rstrip(".")
        nombre_ind[r["Código"]] = t[0].lower() + t[1:]
    excluir = ("D.S.", "R.M.", "IC", "CV", "OP", "IOP", "ISS", "RS", "RC", "SFD", "LIO", "ODS", "PIA", "PIM", "ROF",
               "POI", "PEI", "RPCD", "ERM")
    # en la primera mención, el nombre sin la precisión entre paréntesis de la lista de abreviaturas («(MIMP)»)
    siglas = {k: re.sub(r"\s*\([^)]*\)$", "", v) for k, v in R.ABREVIATURAS.items() if k not in excluir}
    N = Nombrador(nombre_serv, nombre_ind, siglas)
    por_op = C["estado_por_op"]
    anios = [2019, 2020, 2021, 2022, 2023, 2024, 2025]
    n = C["serv_total"]
    cs, dev = C["cond_suma"], C["cond_dev"]
    est25 = C["op_est25"].set_index("Código")
    sfd = C["sfd"]
    d = Informe(N, pie=f"Informe de evaluación de cierre · PNIG 2019–2025 · OMEP–OGMEPGD · {VERSION}",
                titulo_doc=f"Informe de evaluación de cierre de la PNIG (2019–2025), {VERSION.lower()}", informe=True)

    # ------------------------------------------------------------------ portada (sección sin encabezado ni pie)
    portada(d)

    # ------------------------------------------------------------------ índices (numeración romana; sin ficha técnica
    # desde la ronda 33, a pedido del usuario)
    def entradas(tipo):
        if tipo == "titulos":
            return [(t, paginas.get(("t", t), ""), nivel) for t, nivel in TITULOS]
        return [(f"{tipo} {k}. {t}", paginas.get((tipo, k), ""), 1) for (r_, k, t) in LEYENDAS if r_ == tipo]

    d.titulo_indice("Índice")
    d.indice(r'TOC \o "1-2" \h \z \u', entradas("titulos"))
    d.salto()
    d.titulo_indice("Índice de tablas")
    d.indice(r'TOC \h \z \c "Tabla"', entradas("Tabla"), nivel_de="Tabla")
    d.salto()
    d.titulo_indice("Índice de figuras")
    d.indice(r'TOC \h \z \c "Figura"', entradas("Figura"), nivel_de="Figura")
    d.h1("Abreviaturas y siglas")
    d.tabla(["Sigla", "Significado"], [[k, v] for k, v in sorted(R.ABREVIATURAS.items(), key=lambda x: x[0].lower())],
            [1800, 7226], size=16, nombrar=False)
    d.cerrar_seccion(numeracion="lowerRoman")

    # ------------------------------------------------------------------ resumen ejecutivo
    # Ronda 33 (pedido del usuario): resumen como un abstract extenso, solo texto y de dos páginas como máximo, con la
    # estructura del informe (objeto, qué se evalúa, por qué, cómo y resultados). Sin figura, sin viñetas, sin
    # presentación ni guía de lectura.
    N.activo = True
    cuenta = est25.Nivel.value_counts()
    niveles_est = lista_es([f"{cuenta[niv]} en {'nivel ' * (i == 0)}{niv.lower()}" for i, niv in
                            enumerate(x for _, x in S.NIVELES if x in cuenta)])
    vv = pd.Series({p: V[p][0] for p in S.PREGUNTAS}).value_counts()
    confs = pd.Series([h["confianza"] for h in H.values()]).value_counts()
    insuf = [p for p in S.PREGUNTAS if V[p][0] == "Insuficiente"]
    noimpl = C["serv_noimpl"]
    d.h1("Resumen ejecutivo")
    d.p(f"La Política Nacional de Igualdad de Género (PNIG), aprobada en abril de 2019 con el D.S. N.° 008-2019-MIMP, "
        f"aborda como problema público la discriminación estructural contra las mujeres. Su modelo identifica tres "
        f"causas (patrones socioculturales que privilegian lo masculino sobre lo femenino, la asignación desigual de "
        f"roles productivos y reproductivos, y normas y cultura institucional que refuerzan la subordinación de las "
        f"mujeres) y cuatro efectos sobre sus derechos a una vida libre de violencia, a la salud sexual y "
        f"reproductiva, a participar en la toma de decisiones y a los derechos económicos y sociales. Para "
        f"enfrentarlos, la Política se organiza en seis objetivos prioritarios, {C['lin_total']} lineamientos y {n} "
        f"servicios a cargo de {P['n_entidades']} entidades de los tres niveles de gobierno, que el Plan Estratégico "
        f"Multisectorial de Igualdad de Género (PEMIG) ordenó desde 2020. La Ley N.° 32535, publicada en diciembre de "
        f"2025, dispuso sustituirla por una nueva política con el enfoque de igualdad de oportunidades entre mujeres y "
        f"hombres, y la Oficina General de Asesoría Jurídica del MIMP concluyó en noviembre de 2026 que la PNIG perdió "
        f"vigencia el 25 de diciembre de 2025, por incompatibilidad de su objeto con esa ley (MIMP, 2026e).")
    d.p("Este informe presenta la evaluación de cierre de la PNIG, de su diseño y de su implementación, desde su "
        "aprobación hasta el 31 de diciembre de 2025, último año de su vigencia. Se hace porque, conforme a la Guía de "
        "seguimiento y evaluación del CEPLAN, una política que se reemplaza debe evaluarse al cierre: para valorar lo que "
        "logró en todo su periodo y para que sus recomendaciones retroalimenten la formulación de la política que la "
        "reemplaza. La evidencia se recibió hasta noviembre de 2026; lo ocurrido en 2026 se presenta como contexto de "
        "la transición, separado del periodo evaluado. La "
        "evaluación responde doce preguntas, siete de diseño y cinco de implementación, ordenadas en seis criterios: "
        "pertinencia, coherencia, eficacia, eficiencia (con alcance acotado a la conducción), implementación y "
        "sostenibilidad.")
    d.p(f"El enfoque es mixto y no pidió información nueva a las entidades. Combina seis instrumentos: una revisión "
        f"documental de {P['n_corpus']} documentos, que determina el estado de implementación de cada servicio en todo "
        f"el periodo; entrevistas a profesionales de los tres órganos del MIMP que conducen la Política o intervienen "
        f"en su seguimiento; el análisis de lo que las entidades reportaron entre 2019 y 2025 en los reportes de "
        f"seguimiento y de cumplimiento, en los informes anuales de la Ley de Igualdad de Oportunidades y en el "
        f"formulario de seguimiento, con {fmt_num(C['seg_total'])} segmentos codificados y una segunda codificación "
        f"ciega de los que podían cambiar una valoración; el análisis de la coherencia interna y del alineamiento con "
        f"el planeamiento nacional, otras políticas y los compromisos internacionales; la revisión de los indicadores y "
        f"de los recursos, con los microdatos de las encuestas del INEI y los datos abiertos del gasto del MEF, y una "
        f"revisión rápida de {C['rev_fuentes']} estudios sobre la eficacia de los tipos de servicio. Cada hallazgo "
        f"contrasta fuentes, métodos y actores, y su nivel de confianza depende de cuántas fuentes independientes lo "
        f"respaldan.")
    d.p(f"En el diseño, la cadena causal es coherente: {C['lin_total'] - len(C['lin_sin_serv'])} de "
        f"{C['lin_total']} lineamientos tienen servicio y la evidencia respalda de forma sólida {C['rev_sólida']} de "
        f"los {C['rev_tipos']} tipos de servicio. El diagnóstico conserva vigencia: las {C['brechas_n']} brechas entre "
        f"mujeres y hombres que describe persisten; {len(C['brechas_reduce'])} se redujeron desde 2019, "
        f"{len(C['brechas_sincambio'])} no cambiaron de forma significativa y "
        f"{len(C['brechas_amplia'])} {S.pl(len(C['brechas_amplia']), 'se amplió', 'se ampliaron')} "
        f"({lista_es(C['brechas_amplia'])}, en una "
        f"comparación indicativa), y por departamento su reducción se concentra en pocos lugares. La debilidad "
        f"principal está en la medición: {C['serv_sin_ind_vig']} de los {n} servicios no tienen indicador aprobado, "
        f"{C['ficha_lb_nc']} de {C['fichas_total']} fichas de indicador carecen de línea de base y ninguna prevé "
        f"desagregación. De los siete enfoques transversales, solo el de género se operacionaliza con amplitud. El "
        f"alineamiento con las Políticas de Estado, el PEDN y los compromisos internacionales es alto en el periodo, "
        f"pero se debilita con la PGG 2026–2031, que no corresponde a {lista_es(C['pgg127_sin_op'])}, y la Política "
        f"comparte {C['hor_dup']} duplicidades con otras políticas nacionales.")
    d.p(f"En la implementación, de los {n} servicios, {C['serv_impl']} están implementados, {C['serv_parc']} lo están "
        f"parcialmente y {noimpl} {S.pl(noimpl, 'no se implementó', 'no se implementaron')}; solo "
        f"{C['serv_prot_si']} cuentan con protocolo aprobado, que es el componente que más separa a los servicios "
        f"implementados de los parciales. Las entidades atribuyen las dificultades sobre todo al presupuesto, al "
        f"contexto externo y al personal. La conducción ejecutó casi todo lo que se le asignó, pero su devengado bajó "
        f"de {monto(cs[2022])} en 2022 a {monto(cs[2024])} en 2024 antes de recuperarse en 2025, y la "
        f"DGIGND considera que no contó con los recursos necesarios. Para la DGIGND, la CONAIG fue el único mecanismo "
        f"de articulación efectivo, aunque con baja asistencia de los titulares y baja adopción de sus recomendaciones; el PEMIG no "
        f"tuvo seguimiento propio. En los resultados, {C['op_mejora']} de los {C['op_ind']} indicadores de los "
        f"objetivos prioritarios mejoran frente a su línea de base; el reporte oficial no tiene valor 2025 para "
        f"{C['op_sinvalor_2025']} de ellos, y esta evaluación estimó o recalculó {len(est25)} con la ENDES, la ENAHO y "
        f"registros electorales ({niveles_est}, frente a su meta). Al ritmo observado, de las metas de la situación "
        f"futura deseada que pueden evaluarse, solo la de embarazo adolescente se alcanzaría al 2030.")
    d.p(f"En conjunto, ninguna de las doce preguntas alcanza una valoración satisfactoria: {vv.get('Parcial', 0)} son "
        f"parciales y {vv.get('Insuficiente', 0)} {S.pl(vv.get('Insuficiente', 0), 'insuficiente', 'insuficientes')} "
        f"({lista_es(insuf)}, calidad de los indicadores), con {confs.get('Sólido', 0)} de los {len(H)} hallazgos de "
        f"confianza sólida. El desempeño de la PNIG en 2019–2025 es parcial: su diseño es en general coherente y "
        f"respaldado por evidencia, pero su sistema de medición no permite valorar el cumplimiento de la mayoría de "
        f"los objetivos, la implementación de los servicios quedó incompleta y las brechas que la motivaron persisten. "
        f"Como la Ley N.° 32535 dispone su sustitución, las recomendaciones se dirigen a la nueva política: "
        f"{TEMAS[0].rstrip(';')}; {TEMAS[1].rstrip(';')}; {TEMAS[2].rstrip(';')}; {TEMAS[3].rstrip(';')}, y "
        f"{TEMAS[4].rstrip('.')}.")

    # ------------------------------------------------------------------ 1. la Política
    seccion_1(d, C, P, FI, sfd)
    # ------------------------------------------------------------------ 2. metodología
    metodologia(d, C, H, V, P, FI)
    # ------------------------------------------------------------------ 3. diseño
    seccion_3(d, C, H, V, P, F, FI, nombre_serv)
    # ------------------------------------------------------------------ 4. implementación
    seccion_4(d, C, H, V, P, FI, g, por_op, anios, n, cs, dev)
    # ------------------------------------------------------------------ 5. resultados
    seccion_5(d, C, H, V, P, F, FI, est25, sfd, anios)

    # ------------------------------------------------------------------ 6. conclusiones
    d.h1("6. Conclusiones")
    d.p("Las conclusiones responden, una por una, las doce preguntas de evaluación. La tabla siguiente reúne, para cada "
        "pregunta, su texto, la conclusión que se desprende de sus hallazgos, su valoración con la escala del numeral "
        "2.4 y la confianza de cada hallazgo que la sostiene, con los mismos colores de la figura de valoración del "
        "numeral 2.5 y del Anexo 2.")
    filas = []
    for p, (crit, texto, _) in S.PREGUNTAS.items():
        valor, sust, _ = V[p]
        hs = [h for h in H.values() if h["pregunta"] == p]
        conf = []
        for k, h in enumerate(hs):
            conf += [("\n" if k else "", {}), ("■ ", {"color": CONF_COLOR[h["confianza"]]}),
                     (f"{h['id']}: {CONF[h['confianza']]}", {})]
        filas.append([p, [(texto, {"i": True}), (f"\n{sust}", {})], VALOR_TXT[valor].capitalize(),
                      [c for c in conf if c[0]]])
    d.tabla(["Pregunta", "Pregunta y conclusión", "Valoración", "Confianza de los hallazgos"], filas,
            [900, 5126, 1300, 1700], size=16, leyenda="Conclusiones por pregunta de evaluación",
            fuente="Fuente: hallazgos de esta evaluación (Anexo 2). Elaboración propia.", nombrar=False)
    d.p(*leyenda_confianza(), size=18)
    vv6 = pd.Series({p: V[p][0] for p in S.PREGUNTAS}).value_counts()
    confs6 = pd.Series([h["confianza"] for h in H.values()]).value_counts()
    d.p(f"Leída de arriba abajo, la tabla cuenta una misma historia en el diseño y en la implementación: ninguna "
        f"pregunta llega a una valoración satisfactoria, {vv6.get('Parcial', 0)} quedan en parcial y solo la de la "
        f"calidad de los indicadores cae a insuficiente. Que los problemas aparezcan en ambas etapas indica que no "
        f"dependen de un solo eslabón de la Política: un diseño con vacíos de medición se tradujo en una "
        f"implementación difícil de seguir. Las valoraciones descansan en hallazgos firmes: {confs6.get('Sólido', 0)} "
        f"de los {len(H)} son sólidos, {confs6.get('Con respaldo parcial', 0)} tienen respaldo parcial y "
        f"{confs6.get('Indicio', 0)} son indicios, de modo que la mayoría se apoya en dos o más fuentes "
        f"independientes. Esa lectura de conjunto es la base del juicio que sigue.")
    d.h2("Juicio sobre el desempeño de la Política")
    d.p("El juicio es de conjunto y lo formula el equipo evaluador sin ponderación numérica: ninguna de las doce "
        "preguntas alcanza la valoración satisfactoria, la de calidad de los indicadores es insuficiente y ninguna de "
        "las tres dimensiones que sostienen la Política (su diseño, su sistema de medición y la implementación de sus "
        "servicios) muestra un desempeño satisfactorio.")
    d.p("Con la evidencia disponible, el desempeño de la PNIG en 2019–2025 es parcial: su diseño es en general coherente "
        "y está respaldado por evidencia, pero su sistema de medición no permite valorar el cumplimiento de la mayoría "
        "de los objetivos al cierre, la implementación de los servicios fue incompleta y las brechas que motivaron la "
        "Política persisten. Las entrevistas al nivel conductor confirman este juicio; la DGIGND agrega que la "
        "conducción no contó con los recursos necesarios, y los tres órganos coinciden en que el PEMIG no funcionó como "
        "instrumento de articulación ni de seguimiento.")
    d.h2("Lecciones aprendidas")
    for t in LECCIONES:
        d.vineta(t)
    d.h2("Temas vigentes y componentes que conviene preservar")
    for t in TEMAS:
        d.vineta(t[0].upper() + t[1:].rstrip(";.") + ".")
    d.h2("Enseñanzas del proceso de evaluación")
    for t in (
        "La información que las entidades ya reportaron (formulario de seguimiento e informes LIO) permitió evaluar sin "
        "solicitarles información nueva, pero sus vacíos limitan la verificación.",
        "Los datos abiertos del MEF permiten verificar los montos reportados cuando coinciden con una estructura "
        "presupuestal; la mayoría de los que no coinciden no explican su origen.",
        f"La doble codificación de los segmentos críticos ({C['dc_n']}) mostró un acuerdo aceptable en los códigos "
        f"(kappa medio por código de {fmt_num(C['dc_kappa_cod'], 2)}) y obligó a precisar el libro de códigos.",
        "Las entrevistas por videollamada se transcribieron con herramientas automáticas que confunden siglas, por lo "
        "que los pasajes dudosos se verificaron en el audio; la entrevista grupal de la DGIGND impidió distinguir "
        "voces, y varias preguntas de la guía no llegaron a formularse.",
    ):
        d.vineta(t)

    # ------------------------------------------------------------------ 7. recomendaciones
    d.h1("7. Recomendaciones")
    d.p("Por tratarse del cierre de una política que será sustituida, las recomendaciones se dirigen a la formulación "
        "de la nueva política y a la transición. Siguen el formato de la Tabla 4 de la Guía de seguimiento y "
        "evaluación y los aspectos de mejora de su Tabla 5 (CEPLAN, 2024a).")
    RR = recomendaciones(C, nombre_serv)
    d.p("La matriz siguiente presenta cada recomendación con el aspecto de mejora al que responde, la unidad de "
        "organización que debe llevarla a cabo, el plazo y las preguntas de evaluación de las que se deriva.")
    d.tabla(["N.°", "Aspecto de mejora", "Recomendación", "Unidad de organización responsable", "Plazo", "Deriva de"],
            [[str(i + 1), r[0].replace(": ", ":\n")] + r[1:] for i, r in enumerate(RR)],
            [450, 1450, 3576, 1400, 1150, 1000], size=16, leyenda="Matriz de recomendaciones",
            fuente="Fuente: hallazgos y conclusiones de esta evaluación. Elaboración propia.",
            nota="Deriva de: preguntas de evaluación de las que surge la recomendación. Las recomendaciones incorporan "
                 "las entrevistas al nivel conductor, el criterio de la OGAJ (MIMP, 2026e), los informes técnicos de "
                 "la DGIGND sobre la transición (MIMP, 2026g, 2026h, 2026i) y la revisión institucional del informe del "
                 "13 de noviembre de 2026.", nombrar=False)
    asp = pd.Series([r[0].split(":")[0] for r in RR]).value_counts()
    resp = pd.Series([r[2] for r in RR])
    d.p(f"La matriz se concentra en el momento en que todavía se puede corregir el rumbo: "
        f"{int(asp.get('Diseño', 0))} de las {len(RR)} recomendaciones son de diseño y deben aplicarse durante la "
        f"formulación de la nueva política, {int(asp.get('Implementación', 0))} son de implementación y "
        f"{int(asp.get('Seguimiento y evaluación', 0))} de seguimiento y evaluación. La DGIGND, como órgano que conduce "
        f"la Política, participa en {int(resp.str.contains('DGIGND').sum())} de ellas, sola o con la OGPPM y la "
        f"OGMEPGD. Varias tienen un plazo anterior a la aprobación de la nueva política, porque se refieren a "
        f"decisiones que no pueden esperar: el destino de los protocolos en trámite y de los servicios no "
        f"implementados, la representación en la CONAIG y la conservación de la documentación de seguimiento.")

    # ------------------------------------------------------------------ referencias y anexos
    d.h1("Referencias")
    referencias(d, E, refx)
    anexos(d, C, H, P, F, E, nombre_serv)
    d.numerar_cuerpo_desde_1()
    return d, H, RR


def seccion_1(d, C, P, FI, sfd):
    """Sección 1 (ronda 33): cada figura y tabla con su entrada, su lectura, su mensaje y su implicación."""
    n = C["serv_total"]
    d.h1("1. La Política Nacional de Igualdad de Género")
    d.h2("1.1 Antecedentes y marco normativo")
    d.p("La Ley N.° 28983, Ley de Igualdad de Oportunidades entre Mujeres y Hombres, publicada el 16 de marzo de 2007, "
        "fue el marco legal de la Política durante el periodo evaluado. La PNIG se aprobó con el D.S. N.° "
        "008-2019-MIMP, publicado el 4 de abril de 2019, con una situación futura deseada al 2030. Su implementación se "
        "ordenó con el Plan Estratégico Multisectorial de Igualdad de Género (PEMIG), aprobado por el D.S. N.° "
        "002-2020-MIMP (7 de marzo de 2020), que dispuso que las entidades formularan o adecuaran los protocolos de "
        "sus servicios con la asistencia técnica del MIMP, según un cronograma (MIMP, 2026b, p. 8). Los lineamientos "
        "para elaborar esos protocolos se aprobaron con la R.M. N.° 116-2021-MIMP (21 de abril de 2021), y la "
        "Comisión Nacional de Alto Nivel para la Igualdad de Género (CONAIG), comisión multisectorial permanente, se creó con el D.S. N.° 018-2021-MIMP (22 de julio de "
        "2021) "
        "para el seguimiento de la implementación. La Ley N.° 32535, publicada el 24 de diciembre de 2025 y vigente "
        "desde el día siguiente, derogó la Ley N.° 28983 (Única Disposición Complementaria Derogatoria), ordenó al MIMP "
        "sustituir la PNIG por la Política Nacional de Igualdad de Oportunidades entre Mujeres y Hombres (Segunda "
        "Disposición Complementaria Final) y dio a las entidades un año para adecuar sus políticas, planes y programas "
        "(Cuarta Disposición Complementaria Final).")
    d.p("La figura siguiente ordena en el tiempo esos hitos normativos y de implementación, desde la ley que dio "
        "origen a la Política hasta la que dispone su sustitución.")
    fig(d, FI, "linea_tiempo", 14.5)
    sin_prot = n - C["serv_prot_si"]
    d.p("Los hitos se concentran al inicio del periodo. En los dos años y medio que siguieron a la aprobación de la "
        "Política se dictaron todos los instrumentos que debían ponerla en marcha: el PEMIG, los lineamientos para los "
        "protocolos y la creación de la CONAIG (en verde en la figura; en azul, los hitos del marco normativo). "
        f"Después no se dictó otro instrumento general de conducción o articulación durante más de cuatro años (los "
        f"protocolos de los servicios, que sí se aprobaron, se analizan en el numeral 4.1), hasta que la Ley N.° 32535 "
        f"cerró el periodo "
        f"con la disposición de sustituir la Política, cuando {sin_prot} de sus {n} servicios aún no tenían protocolo "
        "aprobado. Por ello, la evaluación toma dos referencias: la PNIG tal como se aprobó y el marco que regirá la "
        "política que la sustituye.")
    d.p("Tras la Ley N.° 32535, en 2026, el CEPLAN brindó asistencia técnica sobre esa disposición (CEPLAN, 2026); la "
        "RPCD N.° 042-2026/CEPLAN/PCD (20 de junio de 2026) fijó la lista de políticas con fichas técnicas por modificar "
        "y los lineamientos para hacerlo, y en julio el MIMP comunicó al CEPLAN que la PNIG seguiría un proceso de "
        "actualización (MIMP, 2026b, 2026d). Ese criterio cambió en noviembre. En setiembre, la Oficina General de "
        "Asesoría Jurídica (OGAJ) había opinado que la PNIG seguía vigente hasta su sustitución, con su marco conceptual "
        "desplazado por la ley, y había encargado a los órganos técnicos identificar, disposición por disposición, los "
        "contenidos incompatibles (MIMP, 2026f). Con esa identificación, la DGIGND, la DPIGND y la DTEG concluyeron el "
        "9 de noviembre que la incompatibilidad alcanza al objeto de la Política: su matriz técnica marca como "
        f"incompatibles las {int(C['mic_norma']['Disposiciones revisadas'].sum())} disposiciones revisadas de los "
        "decretos de la PNIG, el PEMIG y la CONAIG (MIMP, 2026g, 2026h, 2026i). Sobre esa base, la OGAJ revisó su "
        "criterio y concluyó que la PNIG, el PEMIG, la CONAIG y otros cuatro decretos supremos y "
        "dos resoluciones ministeriales dictados en el marco de la Ley N.° 28983 perdieron vigencia el 25 de diciembre "
        "de 2025, porque su objeto es incompatible con la Ley N.° 32535; que los contenidos compatibles siguen "
        "ejecutándose con sustento directo en esa ley hasta que se aprueben los nuevos instrumentos, y que el MIMP debe "
        "informar ese criterio al CEPLAN (MIMP, 2026e). Es el criterio jurídico del órgano competente del MIMP, que esta "
        "evaluación adopta como referencia normativa. Así, el periodo evaluado (2019–2025) coincide con la vigencia de "
        "la PNIG. La DGIGND elaboró el marco conceptual del enfoque de igualdad de "
        "oportunidades, cuya versión final es del 16 de setiembre de 2026 (MIMP, DGIGND, 2026b) y que se analiza en el "
        "numeral 3.2.")

    d.h2("1.2 Problema público y población")
    d.p("La PNIG aborda como problema público la discriminación estructural contra las mujeres, que genera "
        "desigualdad y limita el ejercicio de sus derechos y oportunidades de desarrollo, y que se reproduce y se "
        f"transmite entre generaciones ({PNIG}, p. 6). El documento no contiene una sección de población objetivo; "
        "su diagnóstico se refiere a las mujeres, que constituyen el 50,8 % de la población (p. 6).")
    d.p("El modelo causal identifica tres causas (patrones socioculturales discriminatorios que privilegian lo "
        "masculino sobre lo femenino; asignación desigual de roles productivos y reproductivos; y normas y cultura "
        "institucional que refuerzan la subordinación de las mujeres) y cuatro efectos sobre los derechos a una vida "
        "libre de violencia, a la salud sexual y reproductiva, a la participación en la toma de decisiones y a los "
        "derechos económicos y sociales.")
    vin = P["vinculo"]
    sin_causa = C["serv_sin_causa"]
    total_nodos = int(P["mod"].Servicios.sum())
    d.p(f"La vinculación de cada servicio con el modelo causal está establecida en la matriz de consistencia de la "
        f"Política (Anexo 5 del D.S. N.° 008-2019-MIMP), cuya columna «Causa/Componentes» asocia cada servicio con una "
        f"causa o un efecto del modelo, junto con la alternativa de solución, el objetivo prioritario y el lineamiento "
        f"al que pertenece. Esa matriz vincula {n - len(sin_causa)} de los {n} "
        f"servicios: {vin.get('Causa', 0)} solo con una causa, {vin.get('Efecto', 0)} solo con un efecto y "
        f"{vin.get('Efecto y causa', 0)} con un efecto y una causa a la vez; en "
        f"{S.pl(len(sin_causa), 'el servicio', 'los servicios')} {lista_es(sin_causa)} no se consigna ninguna. La "
        f"figura siguiente presenta, para cada causa y cada efecto, el número de servicios asociados y los objetivos "
        f"prioritarios a los que pertenecen.")
    fig(d, FI, "problema_publico")
    mod = P["mod"].sort_values("Servicios", ascending=False)
    top, ult = mod.iloc[0], mod.iloc[-1]
    causas = mod[mod.Tipo == "Causa"]
    d.p(f"La distribución de los servicios en el modelo es desigual. La causa "
        f"«{str(top['Causa o efecto del modelo']).lower()}» concentra {int(top.Servicios)} servicios de casi todos los "
        f"objetivos prioritarios, mientras que el efecto sobre la participación en la toma de decisiones tiene "
        f"{int(ult.Servicios)} y la causa institucional, "
        f"{int(causas[causas['Causa o efecto del modelo'].str.startswith('Normas')].Servicios.iloc[0])}, todos del OP5. "
        f"Como {vin.get('Efecto y causa', 0)} servicios se asocian a la vez con un efecto y con una causa, la suma de "
        f"los servicios por causa y efecto ({total_nodos}) supera el número de servicios. Esta distribución es "
        f"relevante para la evaluación, porque la evidencia y la implementación se analizan según la causa o el efecto "
        f"que atiende cada servicio: donde hay pocos servicios, la Política depende de muy pocas intervenciones para "
        f"incidir en un eslabón completo de su modelo.")

    d.h2("1.3 Situación futura deseada")
    d.p("La situación futura deseada al 2030 se enuncia así:")
    d.cita("Al 2030 se reducirá la discriminación estructural que afecta a las mujeres", f"{PNIG}, numeral 1.4.5, p. 20")
    n_enun = sfd["Enunciado (cita literal)"].nunique()
    d.p(f"El mismo numeral la concreta en {n_enun} enunciados, cada uno vinculado a una causa o un efecto del modelo. "
        f"El cuarto enunciado (SFD4) fija dos metas, de congresistas y de alcaldesas, por lo que la tabla siguiente "
        f"tiene {len(sfd)} filas. Para saber si la Política podía medir su propia situación futura, cada enunciado se "
        f"compara con el indicador del objetivo prioritario al que corresponde. La correspondencia es «misma» cuando el "
        f"indicador tiene la misma definición (población, numerador, denominador, unidad y fuente) y la misma línea de base que el enunciado; «aproximada», cuando comparte "
        f"la línea de base pero cambia su denominación o su alcance, por lo que no se trata como equivalente; «distinta», cuando mide otra unidad o un índice "
        f"compuesto, o no reproduce la línea de base del enunciado, y «sin indicador», cuando ningún indicador de "
        f"objetivo lo mide.")
    d.tabla(["Código", "Enunciado (cita literal, p. 20)", "OP", "Indicador del OP"],
            [[r["Código"], [(r["Enunciado (cita literal)"], {"i": True})], r.OP,
              (r['Indicador del OP'] if r['Indicador del OP'] == r['Correspondencia con el indicador del OP'] else
               f"{r['Indicador del OP']} ({r['Correspondencia con el indicador del OP'].lower()})")]
             for _, r in sfd.iterrows()],
            [800, 5126, 700, 2400], size=15, leyenda="Enunciados de la situación futura deseada",
            fuente=f"Fuente: {PNIG}, numeral 1.4.5. Elaboración propia.",
            rastro="D1_D2_brechas_situacion_futura.xlsx, hoja «Situación futura deseada»", nombrar=False)
    corr = sfd.drop_duplicates("Enunciado (cita literal)")["Correspondencia con el indicador del OP"].value_counts()
    d.p(f"La situación futura cubre todos los objetivos prioritarios, pero solo {corr.get('Misma', 0)} de sus {n_enun} "
        f"enunciados se miden "
        f"con el mismo indicador: el de violencia contra las mujeres (SFD3) y el de embarazo adolescente (SFD7). Los "
        f"demás se alejan de su indicador por razones distintas. El de tolerancia social (SFD1) comparte la línea de "
        f"base, pero su indicador lo denomina de otra forma; el de ministerios que incorporan el enfoque (SFD2) se "
        f"compara con un indicador que mide entidades públicas; los de congresistas y alcaldesas (SFD4) y el de "
        f"ingresos (SFD5), con índices de empoderamiento que no los miden por separado, y el del Sistema Nacional de "
        f"Cuidado (SFD6) no tiene indicador. Esto limita la posibilidad de valorar la situación futura con el "
        f"seguimiento de la Política, como se verá en el numeral 5.4.")

    d.h2("1.4 Objetivos prioritarios, lineamientos y servicios")
    d.p(f"La Política se organiza en seis objetivos prioritarios (OP), {C['lin_total']} lineamientos y {n} servicios a "
        f"cargo de entidades de los tres niveles de gobierno, con {C['op_ind']} indicadores de OP y "
        f"{C['serv_ind']} indicadores de servicio vigentes. La cadena de valor va del problema público a los servicios: "
        f"cada OP atiende causas o efectos del modelo y se despliega en lineamientos y servicios, que se siguen con "
        f"indicadores. La tabla siguiente resume esa estructura por OP; el catálogo completo está en el Anexo 9.")
    obj = P["obj"]
    d.tabla(["OP", "Objetivo prioritario", "Lineamientos", "Servicios", "Con indicador aprobado"],
            [[o, obj.loc[o, "Objetivo prioritario"], str(int(obj.loc[o, "Lineamientos"])),
              str(int(obj.loc[o, "Servicios"])), str(int(obj.loc[o, "Servicios con indicador vigente"]))]
             for o in obj.index], [700, 4626, 1400, 1100, 1200], size=16,
            leyenda="Estructura de la PNIG por objetivo prioritario",
            fuente=f"Fuente: {PNIG} y matriz de consistencia (Anexo 5 de la PNIG); OGMEPGD (2026). Elaboración propia.",
            nota="El lineamiento 3.2 no tiene servicio.",
            rastro="REV_ALI_coherencia_alineamiento.xlsx, hoja «REV_ALI.1 Por objetivo»", nombrar=False)
    serv_op = obj["Servicios"].astype(int)
    con_ind = obj["Servicios con indicador vigente"].astype(int)
    mayor, menor = serv_op.idxmax(), serv_op.idxmin()
    d.p(f"La tabla muestra una Política muy concentrada: el {mayor} reúne {int(serv_op[mayor])} de los {n} servicios, "
        f"mientras que el {menor} tiene {int(serv_op[menor])}. En todos los OP hay servicios sin indicador aprobado, "
        f"porque los indicadores de servicio solo existen cuando el servicio tiene protocolo aprobado: en conjunto, "
        f"{int(con_ind.sum())} de los {n} servicios tienen al menos uno. La figura siguiente une esa estructura con el "
        f"modelo del problema en una sola cadena de valor.")
    fig(d, FI, "cadena_valor")
    d.p("La cadena de valor está completa en su estructura formal: cada objetivo prioritario se vincula con causas o "
        "efectos del modelo, con un grosor proporcional al número de servicios que los unen, y cuenta con lineamientos y "
        "servicios. La cadena se adelgaza, en cambio, en la medición: los servicios con indicador aprobado son el "
        "eslabón que más se reduce frente al total de servicios de cada objetivo. El seguimiento pierde capacidad "
        "justamente donde la Política debía mostrar lo que producen sus servicios.")

    d.h2("1.5 Indicadores")
    tp, vig = P["tp"], P["vig_op"]
    tipos = tp.tipo.value_counts()
    sin_vig = [f"OP{o[-1]}" for o, v in vig.items() if v == 0]
    d.p(f"La Política se sigue con {C['op_ind']} indicadores de OP, definidos en el anexo del decreto supremo, y con "
        f"los indicadores de servicio que se aprueban con los protocolos: al cierre, {C['serv_ind']} indicadores "
        f"vigentes en {n - C['serv_sin_ind_vig']} servicios. El PEMIG había asignado {C['pemig_ind']} indicadores a "
        f"los {C['pemig_serv']} servicios, pero no tuvieron seguimiento propio (numeral 3.3). La figura siguiente "
        f"compara lo que se planificó medir con lo que efectivamente se midió.")
    fig(d, FI, "indicadores")
    op_max_vig = max(vig.items(), key=lambda x: x[1])
    sa = S._leer(S.PSI, "Serie anual")   # observación 08: indicadores de servicio con valor reportado, por año
    sa = sa[sa.Nivel.astype(str).str.startswith("Serv")]
    con_valor = sa[pd.to_numeric(sa.Valor, errors="coerce").notna()].groupby("Año")["Código"].nunique()
    con_valor.index = con_valor.index.astype(int)
    d.p(f"El PEMIG planificó medir sobre todo la cobertura de los servicios ({int(tipos.get('Cobertura', 0))} de "
        f"{len(tp)} indicadores) y en mucho menor medida su calidad ({int(tipos.get('Calidad', 0))}). Lo que "
        f"efectivamente se midió fue distinto: los indicadores de objetivo y los {C['serv_ind']} indicadores aprobados "
        f"de servicio, distribuidos de forma muy desigual; el OP{op_max_vig[0][-1]} concentra {int(op_max_vig[1])} y "
        f"{lista_es(sin_vig)} {S.pl(len(sin_vig), 'no tiene', 'no tienen')} ninguno. Como los indicadores de "
        f"servicio se aprobaron con cada protocolo, los que tuvieron valor reportado fueron "
        f"{lista_es([f'{int(con_valor.get(a, 0))} en {a}' for a in range(2019, 2026)])}. La distancia entre lo "
        f"planificado y lo medido anticipa una de las conclusiones del informe: la Política no contó con un sistema "
        f"para seguir lo que producían sus servicios.")

    d.h2("1.6 Actores")
    d.p(f"La rectoría corresponde al MIMP. La DGIGND conduce la Política y su Dirección de Políticas (DPIGND) la "
        f"formula, la diseña y brinda asistencia técnica a las entidades; la OGPPM es el punto focal ante el CEPLAN y "
        f"la OGMEPGD, por medio de la OMEP, hace el seguimiento y la evaluación. La CONAIG reúne a las entidades para el "
        f"seguimiento de la implementación, y el CEPLAN, como rector del Sinaplan, emite las guías de seguimiento y "
        f"evaluación. Los {n} servicios están a cargo de {P['n_entidades']} entidades. La figura siguiente muestra "
        f"cómo se distribuyen esos papeles.")
    fig(d, FI, "actores")
    d.p("La figura muestra una Política multisectorial cuyo rector es también proveedor: el MIMP conduce, da "
        "seguimiento y presta servicios en cuatro de los seis objetivos prioritarios, mientras que la mayoría de los "
        "servicios depende de otras entidades, que se coordinan con el rector a través de la CONAIG y de los reportes "
        "anuales. En el OP4, por ejemplo, la responsabilidad se reparte entre más de una decena de entidades. Esta "
        "dispersión explica el peso de la articulación en el análisis de la implementación (numeral 4.4).")

    d.h2("1.7 Alineamiento")
    d.p("La Política se vincula verticalmente con las Políticas de Estado, el Plan Bicentenario y el PEDN, las "
        "políticas generales de gobierno (PGG) y el PESEM del sector, y en la transición con los instrumentos vigentes "
        "en 2026. Se vincula también con la Ley N.° 32535 y con los compromisos internacionales del Perú. El análisis "
        "de esas relaciones y de las duplicidades con otras políticas nacionales está en el numeral 3.5; aquí se "
        "presentan las dos miradas por separado.")
    vp, vt = C["vert_periodo"], C["vert_transicion"]
    d.p("La figura siguiente presenta el alineamiento vertical: la relación de cada objetivo prioritario con los "
        "instrumentos del planeamiento nacional del periodo 2019–2025 y de la transición.")
    fig(d, FI, "alin_vertical")
    d.p(f"En el periodo, el alineamiento es amplio: {vp[0]} relaciones directas, {vp[1]} indirectas y {vp[2]} sin "
        f"correspondencia. En la transición, las relaciones directas bajan a {vt[0]} y las relaciones sin "
        f"correspondencia suben a {vt[2]}, casi todas con la PGG 2026–2031, que no corresponde a "
        f"{lista_es(C['pgg127_sin_op'])}. Es la primera señal de que la nueva política deberá explicitar su "
        f"contribución a la política general del gobierno.")
    i4, i4n = P["int4"], P["int4_n"]
    d.p("La figura siguiente contrasta los objetivos prioritarios con la Ley N.° 32535, que regirá la nueva política, "
        "y con seis compromisos internacionales del Perú.")
    fig(d, FI, "alin_internacional")
    d.p(f"En este caso, el alineamiento es casi pleno: de las {i4n} relaciones evaluadas, {i4.get('Directa', 0)} son relaciones directas, {i4.get('Indirecta', 0)} indirectas y "
        f"{i4.get('Sin correspondencia', 0)} no tienen correspondencia. Los objetivos de la PNIG responden a los "
        f"compromisos que el Perú asumió, y la Ley N.° 32535 recoge los seis, de forma directa o indirecta, de modo "
        f"que el cambio de marco no deja objetivos sin respaldo normativo; los vacíos están en el planeamiento "
        f"nacional de la transición.")

    d.h2("1.8 Instrumentos de implementación y seguimiento")
    d.p("La implementación de una política nacional requiere instrumentos que traduzcan sus objetivos en acciones de "
        "las entidades y que permitan conocer su avance. En la PNIG son de cinco clases: un plan de implementación, "
        "que ordena qué debe hacer cada entidad y cuándo; los protocolos, normas operativas que fijan cómo se presta y "
        "cómo se mide cada servicio; los reportes periódicos, que recogen lo que hicieron las entidades y el avance de "
        "los indicadores; una instancia multisectorial, que reúne a las entidades para el seguimiento y la propuesta de "
        "mejoras, y las evaluaciones, que valoran el desempeño en momentos determinados. La evaluación usa todos ellos "
        "como fuente; a continuación se describe cada uno.")
    for t in (
        [("Plan Estratégico Multisectorial de Igualdad de Género (PEMIG). ", {"b": True}), "Aprobado con el D.S. "
         f"N.° 002-2020-MIMP, asigna indicadores a los {n} servicios y fija un cronograma para que cada entidad apruebe "
         "el protocolo de su servicio. Era el instrumento llamado a ordenar la implementación, pero no tuvo seguimiento "
         "propio (numeral 3.3)."],
        [("Protocolos de servicio. ", {"b": True}), f"Normas que cada entidad aprueba para fijar los estándares de "
         f"su servicio y sus indicadores, con la asistencia técnica del MIMP y los lineamientos de la R.M. N.° "
         f"116-2021-MIMP. Al 31 de diciembre de 2025 había {P['n_protocolos']} protocolos aprobados en "
         f"{C['serv_prot_si']} servicios; son la condición para que un servicio tenga indicadores aprobados."],
        [("Reportes de seguimiento y de cumplimiento. ", {"b": True}), "Reportes anuales que el MIMP elabora con el "
         "formato del CEPLAN desde el primer año de implementación (2020), con los valores de los indicadores de "
         "objetivo desde 2019 y su nivel de avance frente a la meta. Son la fuente oficial del avance de la Política "
         "(Anexo 5)."],
        [("Informes anuales de cumplimiento de la Ley N.° 28983 (informes LIO). ", {"b": True}), "Informes que el "
         "MIMP presenta cada año con lo que reportan las entidades; desde 2019 informan sobre cada servicio de la PNIG, "
         "con sus acciones, su cobertura y sus dificultades. Son la fuente más completa de la ejecución de los "
         "servicios en todo el periodo (MIMP, 2020-2026)."],
        [("Formulario de seguimiento de la OMEP. ", {"b": True}), "Reporte anual de las entidades por indicador, de "
         "2023 a 2025, con el valor, el nivel de cumplimiento y las dificultades; complementa a los anteriores en los "
         "últimos tres años (OGMEPGD, 2023-2025)."],
        [("Informes de gobernanza de la CONAIG. ", {"b": True}), "Cuatro informes anuales, de 2022 a 2025, en los que "
         "la comisión revisa la implementación y propone mejoras de gobernanza a las entidades (MIMP y CONAIG, "
         "2022-2025)."],
        [("Evaluaciones anteriores. ", {"b": True}), "La evaluación de implementación 2020, la de resultados 2021, la "
         "evaluación 2023 y el informe técnico del CEPLAN sobre los servicios del OP1, cuyas recomendaciones se siguen "
         "en el numeral 4.7."],
    ):
        d.vineta(*t)
    d.p("En conjunto, los instrumentos conforman un sistema completo en su diseño, pero desigual en su aplicación. "
        "El plan que debía articularlos no se siguió, y los protocolos, que debían dar a cada servicio estándares e "
        "indicadores, llegaron a menos de la mitad de los servicios. El peso del seguimiento recayó entonces en los "
        "reportes oficiales, que miden los objetivos, y en los informes LIO, que narran las acciones de las entidades "
        "servicio por servicio. Por eso esta evaluación toma los informes LIO de 2019 a 2025 como la fuente principal "
        "para reconstruir la ejecución de los servicios en todo el periodo.")


def seccion_3(d, C, H, V, P, F, FI, nombre_serv):
    d.h1("3. Análisis del diseño")
    d.p("Esta sección responde las preguntas de diseño D1 y D3 a D7. Cada numeral abre con la pregunta y cierra con su "
        "respuesta; la pregunta D2, sobre la situación futura deseada, se responde en el numeral 5.4, junto con los "
        "resultados.")
    d.h2("3.1 Problema público y vigencia del diagnóstico (D1)")
    abre(d, "D1")
    hallazgo(d, H, "D1-H1")
    b = C["brechas"]
    filas = []
    for _, r in b.iterrows():
        ent = 3 if r["Medida de la brecha"].startswith("Razón") else 1
        filas.append([r.Indicador, r.Encuesta, r["Causa o efecto"].split(": ")[1],
                      f"{fmt_num(r['Brecha inicial'], ent)} ({int(r['Año inicial'])})",
                      f"{fmt_num(r['Brecha final'], ent)} ({int(r['Año final'])})", r["Clasificación"]])
    d.p("La tabla siguiente presenta las brechas que describe el diagnóstico, medidas con los microdatos de las "
        "encuestas del INEI al inicio y al final del periodo. En las brechas medidas como diferencia, un valor negativo "
        "indica que las mujeres están por debajo de los hombres; en las medidas como razón mujeres/hombres, un valor "
        "menor que 1 indica lo mismo, y 1, la paridad.")
    d.tabla(["Brecha", "Encuesta", "Causa o efecto del modelo", "Inicial", "Final", "Evolución"], filas,
            [2626, 900, 1900, 1150, 1150, 1300], size=15, leyenda="Evolución de las brechas del diagnóstico",
            rastro="D1_D2_brechas_situacion_futura.xlsx, hoja «Evolución de las brechas»",
            fuente="Fuente: INEI (2026), microdatos de la ENAHO, la ENDES, la ENARES, la ENUT, la EPEN y la ENA; INEI "
                   "(2025c). Elaboración propia.",
            nota="Brecha: mujeres menos hombres, en puntos porcentuales; para ingresos y trabajo no remunerado, razón "
                 "mujeres/hombres. Evolución: cambio significativo al 95 % hacia la paridad (se reduce) o en contra (se "
                 "amplía). Todas las brechas siguen siendo significativas en el año final. La comparación del trabajo no "
                 "remunerado es indicativa, porque la ENUT cambió de instrumento.", nombrar=False)
    d.p(f"La tabla muestra que ninguna brecha desapareció: {len(C['brechas_reduce'])} se redujeron, "
        f"{len(C['brechas_sincambio'])} no cambiaron de forma significativa y la del trabajo no remunerado se amplió en "
        f"una comparación indicativa. Las brechas económicas (participación laboral, ingresos propios e ingresos "
        f"laborales) siguen siendo las de mayor magnitud. Para la evaluación, esto significa que el diagnóstico de la "
        f"PNIG conserva vigencia empírica: los problemas que describió en 2019 seguían presentes en 2025. La figura "
        f"siguiente muestra la evolución anual de las brechas laborales, económicas y educativas.")
    figura_brecha(d, F, "series_enaho", alt="Series anuales 2016 a 2025 de indicadores laborales, económicos y "
                  "educativos por sexo, con su intervalo de confianza.")
    d.p("Las series muestran trayectorias paralelas de mujeres y hombres: los niveles se mueven con el ciclo económico "
        "(con la caída de 2020 por la pandemia), pero la distancia entre ambos casi no cambia. Por eso las brechas "
        "persisten aunque los indicadores de las mujeres mejoren.")
    for i in ("D1-H2", "D1-H3", "D1-H4", "D1-H5"):
        hallazgo(d, H, i)
    d.p("La tabla siguiente resume, para cada indicador mapeado, en cuántos departamentos la estimación es confiable, "
        "en cuántos la brecha es significativa en desventaja de las mujeres y en cuántos cambió desde 2019.")
    r_dep = tabla_departamental(d, C)
    col = "Brecha significativa en desventaja de las mujeres"
    br_dep = r_dep[~r_dep.Tipo.str.startswith("Nivel") & r_dep[col].notna()]
    top = br_dep.sort_values(col, ascending=False).head(3)
    d.p("La tabla confirma que las brechas son territoriales y no solo nacionales: en "
        + lista_es([f"{str(x.Indicador).lower()} ({int(x[col])} departamentos)" for _, x in top.iterrows()]) +
        " la desventaja de las mujeres es significativa en la mayoría del país. En casi todos los indicadores predomina "
        "la ausencia de cambio significativo: entre 2019 y 2025 las brechas departamentales cambiaron poco y, donde "
        "cambiaron, lo hicieron en pocos departamentos. Los mapas siguientes presentan esa información para cuatro "
        "indicadores centrales del diagnóstico, con 2019 y 2025 lado a lado; la trama identifica las estimaciones "
        "referenciales, cuyo coeficiente de variación supera el 15 %.")
    dr = C["dep_resumen"].set_index("Clave")
    entradas_mapa = {
        "mapa_sin_ingresos": "El primer mapa presenta la brecha en la población sin ingresos propios, la brecha "
                             "económica de mayor peso en el diagnóstico.",
        "mapa_tasa_actividad": "La menor disponibilidad de ingresos propios de las mujeres se relaciona con su menor "
                               "participación en el mercado de trabajo, que presenta el mapa siguiente.",
        "mapa_informal": "El mapa siguiente presenta la brecha en el empleo informal, que caracteriza las condiciones de "
                         "trabajo de las mujeres ocupadas.",
        "mapa_violencia_fisica_o_sexual_12m": "El último mapa corresponde a la violencia de pareja. Como la ENDES solo "
                                              "encuesta a mujeres, presenta el nivel del indicador y no una brecha.",
    }
    cierres_mapa = {
        "mapa_sin_ingresos": "Es la única de las cuatro brechas que se redujo en varios departamentos, pero la "
                             "desventaja persiste en casi todo el país: la reducción avanzó sin cerrar la distancia.",
        "mapa_tasa_actividad": "La brecha de participación es la más uniforme del territorio: está presente en todos "
                               "los departamentos y prácticamente no varió en siete años.",
        "mapa_informal": "La desventaja es significativa en algo más de la mitad del país y no cambió en casi ningún "
                         "departamento, de modo que la informalidad sigue siendo una condición del trabajo de las "
                         "mujeres, cuya evolución no puede atribuirse causalmente a la Política con este análisis.",
        "mapa_violencia_fisica_o_sexual_12m": "La tendencia es de mejora, pero ninguna estimación de 2025 es "
                                              "confiable por departamento: el mapa sirve como referencia y la lectura "
                                              "firme es la nacional (numeral 5.2).",
    }
    for clave in MAPAS_CUERPO:
        d.p(entradas_mapa[clave])
        figura_brecha(d, F, clave, alt=f"Mapas del Perú por departamento, 2019 y 2025: {F.loc[clave, 'titulo'].lower()}.")
        x = dr.loc[clave.replace("mapa_", "")]
        d.p(f"{frase_mapa(x)} {cierres_mapa[clave]}")
    d.p("Los mapas de los demás indicadores y las series de la ENA, la EPEN, la ENARES, la ENUT y la ENDES están en el "
        "Anexo 6.")
    cierra(d, "D1", H, V)

    d.h2("3.2 Marco conceptual del enfoque de igualdad de oportunidades (D7)")
    abre(d, "D7")
    d.p("La Ley N.° 32535 y el marco conceptual de la DGIGND (MIMP, DGIGND, 2026b) reemplazan el enfoque de género de "
        "la PNIG por el enfoque "
        "de igualdad de oportunidades entre mujeres y hombres. Para determinar qué elementos de la PNIG se mantienen "
        "bajo el nuevo marco, cada concepto del documento de la Política se comparó con la categoría que le corresponde en la "
        "Ley o en el marco conceptual, en cuatro planos: conceptual, problema público, estructura de la Política y "
        "servicios. Las dos figuras siguientes presentan esa comparación: la primera, por plano; la segunda, concepto "
        "por concepto en el plano conceptual.")
    rc = pd.read_excel(MAC, sheet_name="REV_ALI.5 Conceptual", dtype=str)
    conc = rc[rc.Plano == "Conceptual"]["Tipo de relación"].value_counts()
    efe = rc["Efecto"].value_counts() if "Efecto" in rc else pd.Series(dtype=int)
    srv_c = rc[rc.Plano == "Servicios"]["Tipo de relación"].value_counts()
    # el párrafo de la segunda figura nombra los conceptos que divergen: si la matriz cambia, se reescribe
    assert set(rc[(rc.Plano == "Conceptual") & (rc["Tipo de relación"] == "Divergencia")]["Código"]) == \
        {"C01", "C02", "C11", "C12", "C13", "C16", "C20"}, "cambiaron las divergencias conceptuales"
    fig(d, FI, "conceptual_planos")
    d.p(f"De las {len(rc)} relaciones registradas, {efe.get('Se mantiene', 0)} se mantienen, "
        f"{efe.get('Se ajusta', 0)} se ajustan y {efe.get('Se rediseña', 0)} se rediseñan. Los servicios, el plano con "
        f"más relaciones, muestran casi solo continuidad y precisión ({srv_c.get('Divergencia', 0)} divergencias en "
        f"{int(srv_c.sum())} relaciones), mientras que el plano conceptual concentra las divergencias. La figura "
        f"siguiente detalla ese plano concepto por concepto.")
    fig(d, FI, "conceptual_conceptos")
    d.p(f"De los {int(conc.sum())} conceptos de la PNIG, {conc.get('Divergencia', 0)} divergen de las categorías de la "
        f"Ley N.° 32535 o del marco conceptual. El enfoque de género y "
        f"la noción de género como construcción social se reemplazan por el enfoque de igualdad de oportunidades y la "
        f"discriminación basada en el sexo; la autonomía física, la salud sexual y reproductiva y la educación sexual "
        f"integral pasan a tratarse como derecho a la salud, maternidad segura y educación sexual con base científica; "
        f"la referencia a las mujeres en su diversidad cambia por la de los grupos que requieren especial protección, y "
        f"la familia, sin categoría propia en la PNIG, entra como enfoque. Para la nueva política, esto significa que "
        f"el cambio de enfoque afecta más la justificación y el modelo del problema que el portafolio de servicios: "
        f"buena parte de lo que la PNIG presta puede trasladarse, pero tendrá que explicarse con otro marco.")
    mc, ms = C["mic_conc"], C["mic_serv"]
    ev, mx = "Clasificación de la evaluación (REV_ALI.5)", "Clasificación según la matriz"
    efectos = ["Se mantiene", "Se ajusta", "Se rediseña"]
    d.p("Esa clasificación de los servicios se contrastó con la matriz técnica de incompatibilidades que la DPIGND "
        "elaboró después, a pedido de la OGAJ (MIMP, 2026g). Para cada grupo de servicios, la matriz indica con qué "
        "base de la Ley N.° 32535 o de la Ley N.° 30364 su contenido puede seguir ejecutándose, o por qué es "
        "incompatible. Con una regla fija, cada observación de la matriz se tradujo a las mismas tres categorías: la "
        "referencia a la educación sexual integral equivale a «se rediseña»; la continuidad sin elementos del enfoque "
        "de género o con contenidos por adecuar, a «se ajusta», y la continuidad sin condiciones, a «se mantiene». La "
        "tabla siguiente cruza las dos clasificaciones.")
    d.tabla(["Según la evaluación", "Matriz: se mantiene", "Matriz: se ajusta", "Matriz: se rediseña", "Total"],
            [[e] + [str(int(((ms[ev] == e) & (ms[mx] == f)).sum())) for f in efectos] + [str(int((ms[ev] == e).sum()))]
             for e in efectos]
            + [["Total"] + [str(int((ms[mx] == f).sum())) for f in efectos] + [str(len(ms))]],
            [2626, 1600, 1600, 1600, 1600], size=16,
            leyenda="Clasificación de los servicios según la evaluación y según la matriz de incompatibilidades",
            fuente="Fuente: MIMP (2026g); Ley N.° 32535; MIMP, DGIGND (2026b). Elaboración propia.",
            nota="Las filas son la clasificación de esta evaluación; las columnas, la que resulta de la matriz de la "
                 "DPIGND con la regla descrita.",
            rastro="REV_DOC_matriz_incompatibilidades.xlsx, hoja «Tabla cruzada»", nombrar=False)
    disc = ms[ms[ev] != ms[mx]]
    aj_ev = disc[disc[ev] == "Se ajusta"]["Código"].tolist()
    aj_mx = disc[disc[mx] == "Se ajusta"]["Código"].tolist()
    # los códigos van sin su nombre: la lista de diez servicios con nombres haría ilegible el párrafo
    d._sin(d.p, f"Las dos clasificaciones coinciden en {int(mc['Coinciden'])} de los {int(mc['Servicios'])} servicios "
        f"(acuerdo de {fmt_num(mc['Acuerdo simple (%)'], 1)} % y kappa de {fmt_num(mc['Kappa de Cohen'], 2)}, una "
        f"concordancia moderada). Coinciden en los dos servicios que se rediseñan, los de educación sexual integral "
        f"({lista_es(C['mac5_serv_rediseña'])}), y las {len(disc)} diferencias son todas entre mantener y ajustar: en "
        f"{len(aj_ev)} servicios ({lista_es(aj_ev)}) esta evaluación registra una precisión del nuevo marco que la "
        f"matriz no exige, como el paso de la salud sexual y reproductiva al derecho a la salud, y en {len(aj_mx)} "
        f"({lista_es(aj_mx)}) la matriz pide prestarlos sin elementos del enfoque de género, que esta evaluación no "
        f"había marcado. Para la nueva política, el contraste confirma que el portafolio puede trasladarse casi "
        f"completo y que los cambios se concentran en la justificación y en los contenidos de un grupo acotado de "
        f"servicios.")
    for i in ("D7-H1", "D7-H2"):
        hallazgo(d, H, i)
    cierra(d, "D7", H, V)

    d.h2("3.3 Coherencia interna (D3)")
    abre(d, "D3")
    d.p("La coherencia interna se examina sobre la cadena de valor del numeral 1.4: si cada lineamiento tiene "
        "servicio, si cada servicio tiene una causa o un efecto del modelo y si cuenta con indicadores para seguirlo. "
        "La plausibilidad de esa cadena, es decir, si la evidencia respalda que los servicios actúan sobre las causas, "
        "se analiza en el numeral 3.4.")
    for i in ("D3-H1", "D3-H2"):
        hallazgo(d, H, i)

    d.h2("3.4 Revisión de la evidencia: plausibilidad de los servicios (D3)")
    evidencia_34(d, C, H, P, FI, nombre_serv)
    cierra(d, "D3", H, V)

    d.h2("3.5 Articulación con políticas y planes (D4)")
    abre(d, "D4")
    for i in ("D4-H1", "D4-H2"):
        hallazgo(d, H, i)
    vo = C["vert_por_op"]
    d.p("La tabla siguiente resume, por objetivo prioritario, cuántas relaciones con los instrumentos de planeamiento "
        "son directas, indirectas o sin correspondencia, en el periodo y en la transición.")
    d.tabla(["OP", "Periodo 2019–2025", "Transición (2026)"],
            [[r.OP, r["Valoración (periodo)"], r["Valoración (transición)"]] for _, r in vo.iterrows()],
            [1500, 3763, 3763], leyenda="Alineamiento vertical por objetivo prioritario",
            rastro="REV_ALI_coherencia_alineamiento.xlsx, hoja «REV_ALI.2 Por objetivo»",
            fuente="Fuente: Políticas de Estado, Plan Bicentenario, PEDN, PGG y PESEM (Anexo 10). Elaboración propia.",
            nota="Periodo: Políticas de Estado, Plan Bicentenario, PEDN, cuatro PGG y PESEM 2018–2026. Transición: "
                 "Políticas de Estado, PEDN actualizado, PGG 2026, PGG 2026–2031 y PESEM 2024–2030.", nombrar=False)
    vp, vt = C["vert_periodo"], C["vert_transicion"]
    d.p(f"La tabla muestra que las relaciones directas pasan de {vp[0]} en el periodo a {vt[0]} en la transición, y que "
        f"las relaciones sin correspondencia pasan de {vp[2]} a {vt[2]}. El cambio se explica sobre todo por la PGG "
        f"2026–2031, que no corresponde a {lista_es(C['pgg127_sin_op'])}: la nueva política tendrá que explicitar cómo "
        f"contribuye a la política general del gobierno. La figura siguiente presenta las relaciones de los servicios "
        f"con otras políticas nacionales.")
    fig(d, FI, "alin_horizontal")
    d.p(f"Los servicios y lineamientos de la PNIG se relacionan con {C['hor_pol']} políticas nacionales afines. Se "
        f"registran {C['hor_dup']} duplicidades, es decir, servicios con el mismo proveedor que figuran en la PNIG y en "
        f"otra política, además de {C['hor_comp']} complementariedades y "
        f"{C['hor_vac']} vacíos de atención. Las duplicidades más visibles son las de los Centros Emergencia Mujer, "
        f"presentes en {len(C['hor_cem_pol'])} políticas: un mismo servicio se reporta varias veces, lo que dispersa su "
        f"seguimiento y obliga a la nueva política a definir dónde se registra.")
    for i in ("D4-H3", "D4-H4"):
        hallazgo(d, H, i)
    cierra(d, "D4", H, V)

    d.h2("3.6 Calidad de los indicadores (D5)")
    abre(d, "D5")
    hallazgo(d, H, "D5-H1")
    d.p("La figura siguiente muestra, para cada atributo que exige la guía de indicadores del CEPLAN, cuántas fichas lo "
        "cumplen, lo cumplen en parte o no lo cumplen.")
    fig(d, FI, "fichas")
    d.p(f"Sobre un total de {C['fichas_total']} fichas, los atributos críticos son los que menos se cumplen: la línea de base con año y "
        f"valor, que falta en {C['ficha_lb_nc']} fichas, y la desagregación, que ninguna prevé. Sin ellos, un indicador "
        f"no permite medir el cambio desde un punto de partida ni seguir los enfoques de la Política por territorio o "
        f"grupo de población.")
    for i in ("D5-H2", "D5-H3", "D5-H4", "D5-H5"):
        hallazgo(d, H, i)
    cierra(d, "D5", H, V)

    d.h2("3.7 Enfoques transversales (D6)")
    abre(d, "D6")
    hallazgo(d, H, "D6-H1")
    d.p("La figura siguiente ubica cada enfoque transversal en una escala de cinco niveles, de ausente a "
        "operacionalizado con alcance amplio, según su incorporación en los lineamientos, los servicios y los "
        "indicadores de la Política.")
    fig(d, FI, "enfoques")
    ef = C["enfoques"]
    niv = ef.nivel.astype(str).value_counts()
    d.p(f"La incorporación de los enfoques es desigual: solo el enfoque de género llega a operacionalizarse con "
        f"amplitud, {niv.get('Operacionalizado', 0)} de los siete se operacionalizan en algún grado y "
        f"{niv.get('Ausente', 0)} están ausentes del diseño; la figura indica además cuántos protocolos aprobados "
        f"mencionan cada enfoque. La tabla siguiente precisa en qué elemento se verificó cada nivel.")
    d.tabla(["Enfoque", "Nivel", "Alcance", "Elemento donde se verifica", "Protocolos"],
            [[r.enfoque, r.nivel, r.alcance, r.elementos_donde_se_verifica, r.protocolos] for _, r in ef.iterrows()],
            [1500, 1500, 1000, 3826, 1200], size=15, leyenda="Lista de verificación de los enfoques transversales",
            rastro="D6_enfoques_transversales.xlsx, hoja «Lista de verificación»",
            fuente=f"Fuente: {PNIG} y sus anexos; anexo 5 del PEMIG (D.S. N.° 002-2020-MIMP); protocolos de los "
                   "servicios (Anexo 10). Elaboración propia.",
            nota="Nivel: ausente, mencionado, desarrollado u operacionalizado (numeral 2.4). Protocolos: protocolos "
                 "aprobados que mencionan el enfoque.", nombrar=False)
    d.p("La tabla confirma la lectura de la figura y añade de dónde viene cada nivel: el enfoque de género es el único "
        "que se traduce de forma amplia en lineamientos, estándares e indicadores, mientras que los demás aparecen sobre "
        "todo en los protocolos aprobados. Los enfoques entraron, entonces, más por las entidades que por el diseño de "
        "la Política, y por eso su alcance depende de cuántos servicios llegaron a tener protocolo.")
    hallazgo(d, H, "D6-H2")
    cierra(d, "D6", H, V)

    d.h2("3.8 Síntesis del diseño")
    d.p("La tabla siguiente reúne la valoración de las preguntas de diseño y su sustento.")
    d.tabla(["Pregunta", "Criterio", "Valoración", "Sustento"],
            [[p, crit, VALOR_TXT[V[p][0]].capitalize(), V[p][1]] for p, (crit, _, _) in S.PREGUNTAS.items()
             if p.startswith("D")],
            [900, 1700, 1200, 5226], size=16, leyenda="Valoración de las preguntas de diseño",
            fuente="Fuente: hallazgos de esta evaluación (Anexo 2). Elaboración propia.",
            rastro="matriz_triangulacion_preguntas.xlsx, hoja «Valoración por pregunta»", nombrar=False)
    d.p("La síntesis muestra un diseño con una cadena causal casi completa y, en su mayoría, respaldada por la "
        "evidencia, pero con dos debilidades que lo atraviesan: la medición, que es la única pregunta insuficiente, y "
        "la correspondencia con el nuevo marco normativo y conceptual, que obliga a revisar el problema público.")


def seccion_4(d, C, H, V, P, FI, g, por_op, anios, n, cs, dev):
    d.h1("4. Análisis de la implementación")
    d.p("Esta sección responde las preguntas de implementación I1, I3, I4 e I5. La pregunta I2, sobre el avance de los "
        "indicadores, se responde en la sección 5.")
    d.h2("4.1 Grado de implementación de los servicios")
    abre(d, "I1")
    hallazgo(d, H, "I1-H1")
    cols = ["Implementado", "Implementado parcialmente", "No implementado"]
    filas = []
    for opc, r in por_op.iterrows():
        filas.append([opc.replace("OP.0", "OP")] + [str(int(r.get(c, 0))) for c in cols] +
                     [str(int(sum(r.get(c, 0) for c in cols)))])
    filas.append(["Total"] + [str(int(por_op[c].sum())) if c in por_op else "0" for c in cols] + [str(n)])
    d.p("La tabla siguiente presenta el estado de implementación de los servicios por objetivo prioritario, con las "
        "reglas del numeral 2.4.")
    d.tabla(["OP"] + cols + ["Servicios"], filas, [1500, 1882, 1882, 1882, 1880],
            leyenda="Estado de implementación de los servicios por objetivo prioritario", grupos=(len(filas) - 1,),
            rastro="REV_DOC_estado_implementacion_servicios.xlsx, hoja «Estado por servicio»",
            fuente="Fuente: protocolos de los servicios (Anexo 10); MEF (2026); OGMEPGD (2026); MIMP (2020-2026). "
                   "Elaboración propia.",
            nota="El detalle por servicio, con la justificación de su estado, está en el Anexo 3.", nombrar=False)
    impl_op = por_op["Implementado"] if "Implementado" in por_op else pd.Series(dtype=int)
    op_mas = impl_op.idxmax().replace("OP.0", "OP") if len(impl_op) else ""
    op2 = int(por_op.loc["OP.02", "Implementado"]) if "OP.02" in por_op.index else 0
    parc_op = por_op["Implementado parcialmente"] if "Implementado parcialmente" in por_op else pd.Series(dtype=int)
    nom = lambda ix: [o.replace("OP.0", "OP") for o in ix]
    domina_p = nom([o for o in por_op.index if parc_op.get(o, 0) > impl_op.get(o, 0)])
    domina_i = nom([o for o in por_op.index if impl_op.get(o, 0) > parc_op.get(o, 0)])
    g2 = g[g["OP"] == "OP.02"]
    op2_sin = (op2 == 0 and (g2["Protocolo aprobado"] == "No").all() and g2["Años con valor reportado"].isna().all())
    assert domina_p and domina_i, "la tabla ya no tiene objetivos de ambos tipos: revisar el texto"
    d.p(f"La tabla muestra que la implementación parcial predomina en {lista_es(domina_p)}, mientras que los servicios "
        f"implementados son mayoría en {lista_es(domina_i)}. El {op_mas} reúne la mayor cantidad de servicios "
        f"implementados"
        + (", y el OP2, a cargo sobre todo del Minsa, no tiene ninguno: sus servicios carecen de protocolo aprobado y "
           "de reporte de indicadores, dos de los tres componentes que exige la escala." if op2_sin else ".")
        + " Las figuras siguientes presentan cada servicio con su estado, su protocolo y su indicador.")
    gg = g.assign(opn=g["OP"].astype(str).str[-1].astype(int))
    def resumen_ops(ops):
        x = gg[gg.opn.isin(ops)]
        e_ = x["Estado asignado (preliminar)"].value_counts()
        return x, int(e_.get("Implementado", 0)), int(e_.get("Implementado parcialmente", 0)), \
            int(e_.get("No implementado", 0)), int((x["Protocolo aprobado"] == "Sí").sum())
    d.p("Las seis figuras siguientes presentan, por objetivo prioritario, el estado de cada servicio agrupado por "
        "lineamiento, e indican si el servicio cuenta con protocolo aprobado y con indicador aprobado.")
    obj_op = P["obj"]
    falta = lambda x, comp: int(x["Componentes faltantes"].fillna("").str.contains(comp).sum())
    for k in range(1, 7):
        x, i_, p_, n_, pr_ = resumen_ops((k,))
        ok = f"OP{k}"
        nombre_op = str(obj_op.loc[ok, "Objetivo prioritario"]) if ok in obj_op.index else ""
        nombre_op = nombre_op[0].lower() + nombre_op[1:] if nombre_op else ""
        n_lin = int(obj_op.loc[ok, "Lineamientos"]) if ok in obj_op.index else 0
        d.p(f"El OP{k} busca {nombre_op.rstrip('.')}. Tiene {len(x)} {S.pl(len(x), 'servicio', 'servicios')} en "
            f"{n_lin} {S.pl(n_lin, 'lineamiento', 'lineamientos')}; la figura siguiente {S.pl(len(x), 'lo', 'los')} "
            f"muestra con su estado.")
        fig(d, FI, f"mapa_servicios_op{k}", 16.0)
        px = x[x["Estado asignado (preliminar)"] == "Implementado parcialmente"]
        impl_cod = x.loc[x["Estado asignado (preliminar)"] == "Implementado", "Código"].astype(str).tolist()
        reporta = int(x["Años con valor reportado"].notna().sum())
        cu = lambda m, sing, plur, cero: cero if m == 0 else f"{m} {S.pl(m, sing, plur)}"
        if p_ == len(x):          # todos parciales
            estado_txt = f"los {len(x)} servicios están implementados parcialmente"
        elif i_ == len(x):        # todos implementados
            estado_txt = f"los {len(x)} servicios están implementados ({lista_es(impl_cod)})"
        else:
            estado_txt = ((f"ninguno de los {len(x)} servicios está implementado" if i_ == 0 else
                           f"{i_} de los {len(x)} servicios {S.pl(i_, 'está implementado', 'están implementados')}"
                           f" ({lista_es(impl_cod)})")
                          + ((f" y {p_} {S.pl(p_, 'lo está', 'lo están')} parcialmente" if not n_ else
                              f", {p_} {S.pl(p_, 'lo está', 'lo están')} parcialmente") if p_ else "")
                          + (f" y {n_} no se {S.pl(n_, 'implementó', 'implementaron')}" if n_ else ""))
        t = (f"En el OP{k}, " + estado_txt
             + f"; {cu(pr_, 'tiene', 'tienen', 'ninguno tiene')} protocolo aprobado y "
             f"{cu(reporta, 'reporta', 'reportan', 'ninguno reporta')} valores de sus indicadores. ")
        COMP = {"protocolo": ("no tiene protocolo", "no tienen protocolo", "protocolo aprobado"),
                "asignación": ("no tiene recursos identificados en el MEF", "no tienen recursos identificados en el MEF",
                               "recursos identificados en el MEF"),
                "reporte": ("no reporta indicadores", "no reportan indicadores", "valores de sus indicadores")}
        # lo que ya dijo la primera frase (ningún protocolo, ningún reporte) no se repite
        ya = {"protocolo": pr_ == 0, "asignación": False, "reporte": reporta == 0}
        if len(px) == 1:
            le_faltan = [v[2] for c, v in COMP.items() if falta(px, c)]
            t += f"El servicio parcial ({px['Código'].iloc[0]}) carece de {lista_es(le_faltan)}. "
        elif len(px) > 1:
            partes = [("ninguno " + v[0][3:]) if falta(px, c) == len(px) else f"{falta(px, c)} {S.pl(falta(px, c), v[0], v[1])}"
                      for c, v in COMP.items() if falta(px, c) and not ya[c]]
            if partes:
                t += f"De los {len(px)} parciales, {lista_es(partes)}. "
        sp = px[px["Componentes faltantes"].fillna("").str.contains("protocolo")]
        sp_r = falta(sp, "reporte")
        ni = x[x["Estado asignado (preliminar)"] == "No implementado"]
        if pr_ == 0 and reporta == 0:
            t += ("Al carecer de protocolo y de reporte de indicadores, ninguno reúne los dos componentes que exige el "
                  "estado de implementado, aunque todos reportan ejecución en algún año del periodo.")
        elif i_ >= len(x) / 2:
            t += ("La mitad o más de sus servicios están implementados, es decir, reúnen al menos dos de los tres "
                  "componentes y seguían funcionando en 2024 o 2025.")
        elif len(sp) and sp_r == len(sp):
            t += (f"La falta de protocolo es lo que más pesa: "
                  f"{'el parcial' if len(sp) == 1 else f'los {len(sp)} parciales'} sin protocolo tampoco "
                  f"{S.pl(len(sp), 'reporta', 'reportan')} indicadores, porque el protocolo trae consigo los estándares "
                  f"y la medición del servicio.")
        if len(ni):
            t += (f" El servicio {lista_es([f'{c} «{n}»' for c, n in zip(ni['Código'], ni['Servicio'])])} no llegó a "
                  f"ejecutarse en ningún año del periodo.")
        sin = [c for c in C["lin_sin_serv"] if c.startswith(f"{k}.")]
        if sin:
            t += f" El lineamiento {lista_es(sin)} no tiene servicio, por lo que su avance no puede seguirse."
        d.p(t)
    parc = g[g["Estado asignado (preliminar)"] == "Implementado parcialmente"]
    parc_prot = int((parc["Protocolo aprobado"] == "Sí").sum())
    d.p(f"En conjunto, el protocolo es el componente que más pesa en el estado de los servicios, en parte porque la "
        f"regla le asigna ese peso y de él depende el indicador: de los {len(parc)} servicios "
        f"implementados parcialmente, solo {parc_prot} {S.pl(parc_prot, 'tiene', 'tienen')} protocolo, y la mayoría de "
        f"los demás reporta ejecución en varios años sin haberlo aprobado. Esos servicios funcionaron, pero sin los "
        f"estándares ni los indicadores que permitirían conocer cómo se prestaron y con qué resultado.")

    d.h2("4.2 Asignación y ejecución de los recursos")
    hallazgo(d, H, "I1-H2")
    abre(d, "I3")
    d.p("La figura siguiente presenta el devengado anual de las metas presupuestales de los órganos que conducen la "
        "Política, en millones de soles.")
    fig(d, FI, "recursos", 15.0)
    pico0 = max(anios, key=lambda a: cs[a])
    valle = min([a for a in anios if a > pico0] or [pico0], key=lambda a: cs[a])
    d.p(f"Hasta 2022, el gasto de la conducción corresponde a la meta de la DGIGND; desde 2023 se reparte también "
        f"entre la DPIGND y la DPPDM, creadas por el ROF de ese año. El devengado alcanzó su máximo en {pico0} "
        f"({monto(cs[pico0])}), bajó hasta {monto(cs[valle])} en {valle}, ya con la meta distribuida entre tres "
        f"órganos, y se recuperó en 2025 ({monto(cs[2025])}). La ejecución de la meta de la DGIGND, indicada bajo cada "
        f"año, se mantuvo alta: estos porcentajes describen la ejecución presupuestal, pero no permiten por sí solos determinar la suficiencia de los recursos ni la capacidad de gestión. La tabla siguiente "
        f"presenta las cifras, junto con la meta de transversalización del OP5.")
    filas = []
    for k, nombre in (("DGIGND", "DGIGND"), ("DPIGND", "DPIGND (desde 2023)"), ("DPPDM", "DPPDM (desde 2023)")):
        filas.append([nombre] + [millones(dev[a][k]) for a in anios])
    filas.append(["Conducción (suma)"] + [millones(cs[a]) for a in anios])
    filas.append(["DGTEG o DTEG (transversalización, OP5)"] + [millones(dev[a]["DGTEG"]) for a in anios])
    d.tabla(["Meta del órgano"] + [str(a) for a in anios], filas, [2826] + [886] * 7, size=16, grupos=(3,),
            leyenda="Devengado de las metas de los órganos de la conducción (millones de soles)",
            fuente="Fuente: MEF (2026), pliego 039, Ministerio de la Mujer y Poblaciones Vulnerables; MIMP, DGIGND "
                   "(2026a). Elaboración propia.", nota="s. m.: sin meta en el año.",
            rastro="REV_IND_3_verificacion_presupuestal.xlsx, hoja «Conducción por órgano»", nombrar=False)
    pico = max(anios, key=lambda a: cs[a])
    d.p(f"La tabla muestra que la conducción alcanzó su mayor devengado en {pico} ({monto(cs[pico])}) y que, "
        f"desde 2023, el gasto se reparte entre la DGIGND y las dos direcciones creadas por el ROF de ese año. La suma "
        f"bajó a {monto(cs[2024])} en 2024 y se recuperó a {monto(cs[2025])} en 2025. Como la meta de "
        f"la dirección general financia también otros compromisos, estas cifras, en soles corrientes y solo con las "
        f"metas de los órganos de conducción (sin otras fuentes, como la cooperación), son el techo del gasto de la "
        f"conducción, no su costo exacto.")
    pdg = f"{EVAL}/04_analisis/REV_IND_3_presupuesto_dgignd.xlsx"
    conc_mef = pd.read_excel(pdg, sheet_name="Conciliación MEF")
    comp = pd.read_excel(pdg, sheet_name="Composición (%)").set_index("Genérica")
    personas = {a: float(comp.loc[["Personal y CAS", "Locación de servicios de personas naturales"], f"Ejecución {a}"].sum())
                for a in range(2021, 2026)}
    por_org = pd.read_excel(pdg, sheet_name="Por órgano").set_index("Órgano")
    pim26 = float(por_org.loc[[o for o in ("DGIGND", "DPIGND", "DPPDM") if o in por_org.index], "PIM 2026"].sum())
    coinciden = int((conc_mef.Coincide == "Sí").sum())
    d.p(f"Estas cifras se contrastaron con la ejecución presupuestal de la actividad 5001072 «Promoción de la mujer» "
        f"de 2019 a 2026, por finalidad y específica de gasto, que la DGIGND remitió para esta evaluación (MIMP, DGIGND, "
        f"2026a). Agrupadas por el órgano que ejecuta cada finalidad, ambas fuentes coinciden al sol en {coinciden} de "
        f"{len(conc_mef)} combinaciones de órgano y año del periodo. Esa información permite, además, conocer la "
        f"composición del gasto: entre 2021 y 2025, entre {fmt_num(min(personas.values()), 0)} % y "
        f"{fmt_num(max(personas.values()), 0)} % del gasto de la conducción se destinó a personal, contratos "
        f"administrativos de servicios y locación de servicios de personas naturales; el resto financió bienes y "
        f"servicios para las actividades. En 2019 y 2020, la mayor parte del gasto distinto del de personal se registró "
        f"como «servicios diversos», sin un detalle que permita desagregarlo. La conducción de la Política consistió, "
        f"sobre todo, en un equipo de personas. Para 2026, fuera del periodo evaluado, las metas de "
        f"los tres órganos suman un PIM de {monto(pim26)}.")
    for i in ("I3-H1", "I3-H2", "I3-H3"):
        hallazgo(d, H, i)
    cierra(d, "I3", H, V)

    d.h2("4.3 Servicios que no funcionaron y lineamientos sin servicio")
    gi = g.set_index("Código")
    k_no = len(C["serv_noimpl_cod"])
    d.p(f"{S.pl(k_no, 'Un servicio no se implementó', f'{k_no} servicios no se implementaron')} en ningún año del "
        f"periodo. La revisión documental {S.pl(k_no, 'lo', 'los')} clasificó así por lectura cualitativa de la "
        f"evidencia de ejecución de 2019 a 2025:", after=60)
    for c in C["serv_noimpl_cod"]:
        d.vineta((f"Servicio {c}. ", {"b": True}), (str(gi.loc[c, "Servicio"]).rstrip(". ") + ". ", {"i": True}),
                 str(gi.loc[c, "Ajuste por lectura cualitativa"]))
    d.p(f"El "
        f"lineamiento {lista_es(C['lin_sin_serv'])} no tiene servicio, por lo que su avance no puede seguirse.")

    d.h2("4.4 Mecanismos de coordinación de actores")
    abre(d, "I4")
    hallazgo(d, H, "I4-H1")
    ses = FIGI.SESIONES_CONAIG
    ses_per = [x for x in ses if x[0] <= 2025]
    n_ses = sum(len(f) for _, f, _ in ses_per)
    con_prot = [a for a, _, ac in ses_per if any("protocolo" in t.lower() for t in ac)]
    d.p(f"La CONAIG es el espacio de coordinación de las entidades para la implementación de la Política. Su ayuda "
        f"memoria (MIMP y CONAIG, 2026) registra, por año, las sesiones y los acuerdos de la comisión, que la figura "
        f"siguiente resume.")
    fig(d, FI, "conaig_sesiones")
    d.p(f"Entre 2021 y 2025, la comisión celebró {n_ses} sesiones ordinarias, además de la de instalación, y desde "
        f"2022 cumplió el mínimo de tres sesiones anuales que fija su reglamento interno, aprobado en 2021. El rasgo más "
        f"notorio es la reiteración de un mismo acuerdo, resaltado en ámbar en la figura: "
        + (f"cada año, de {con_prot[0]} a {con_prot[-1]}," if len(con_prot) == len(ses_per) else
           f"en {len(con_prot)} de los {len(ses_per)} años ({lista_es([str(a) for a in con_prot])})")
        + f" la CONAIG volvió a acordar que las entidades culminaran y aprobaran "
        f"los protocolos pendientes, y al cierre de 2025 más de la mitad de los servicios seguía sin protocolo. La "
        f"comisión produjo sus informes anuales de gobernanza y sucesivos paquetes de medidas urgentes, pero su acuerdo "
        f"más reiterado no se tradujo en resultados. El hallazgo siguiente lo relaciona con la asistencia y el nivel "
        f"de representación de sus integrantes.")
    for i in ("I4-H2", "I4-H3", "I4-H4"):
        hallazgo(d, H, i)

    d.h2("4.5 Factores que influyeron en el desempeño de los servicios")
    hallazgo(d, H, "I1-H3")
    cierra(d, "I1", H, V)

    d.h2("4.6 Capacidades y condiciones de sostenibilidad ante la transición")
    abre(d, "I5")
    for i in ("I5-H1", "I5-H2", "I5-H3"):
        hallazgo(d, H, i)

    d.h2("4.7 Seguimiento de las recomendaciones de evaluaciones anteriores")
    d.p("La PNIG tuvo cuatro evaluaciones anteriores: la evaluación de implementación 2020 (MIMP, 2021), la evaluación "
        "de resultados de 2021 (MIMP, 2022), la evaluación de 2023 (MIMP, 2024) y el informe técnico del CEPLAN sobre "
        "el diseño y la implementación de los servicios del OP1, una evaluación colaborativa con el MIMP coordinada en "
        "el primer semestre de 2024 (CEPLAN, s. f., p. 6). Sus recomendaciones se siguen con la revisión documental, "
        "que incluye los reportes de cumplimiento 2024 y 2025 (MIMP, 2025, 2026a) y los informes de la ruta de "
        "actualización (MIMP, 2026b, 2026c), y se contrastan en las entrevistas al nivel conductor. Para cada "
        "recomendación se responde si se implementó, cómo y con qué medio de verificación, y cómo trasladar lo "
        "pendiente (preguntas p3.4 a p3.6 de la Tabla A7.1 de CEPLAN, 2024a).")
    hallazgo(d, H, "I4-H5")
    res = C["rec_resumen"]
    estados = ["Implementada", "Parcialmente implementada", "No implementada", "Sin evidencia",
               "Sin objeto por la transición"]
    nombre_ev = {"MIMP 2020": "MIMP, evaluación 2020", "MIMP 2021": "MIMP, evaluación 2021",
                 "MIMP 2023": "MIMP, evaluación 2023", "CEPLAN": "CEPLAN, informe técnico", "Total": "Total"}
    d.p("La tabla siguiente cuenta, por evaluación, cuántas recomendaciones quedaron en cada estado al corte.")
    d.tabla(["Evaluación", "Implementada", "Parcial", "No implementada", "Sin evidencia", "Sin objeto", "Total"],
            [[nombre_ev[ev]] + [str(int(r[e])) for e in estados] + [str(int(r["Total"]))] for ev, r in res.iterrows()],
            [2226, 1400, 1000, 1400, 1100, 1000, 900],
            leyenda="Avance de las recomendaciones de evaluaciones anteriores", grupos=(len(res) - 1,),
            rastro="seguimiento_recomendaciones_previas.xlsx, hoja «Resumen por evaluación»",
            fuente="Fuente: MIMP (2021, 2022, 2024); CEPLAN (s. f.); documentos al 7 de noviembre de 2026 (Anexo 7). "
                   "Elaboración propia.",
            nota="El detalle por recomendación, con su medio de verificación, está en el Anexo 7.", nombrar=False)
    sin_tot = res.drop(index="Total")
    avance = ((sin_tot["Implementada"] + sin_tot["Parcialmente implementada"]) / sin_tot["Total"]).sort_values()
    nombre_corto = {"MIMP 2020": "la evaluación de implementación 2020", "MIMP 2021": "la evaluación de resultados 2021",
                    "MIMP 2023": "la evaluación 2023", "CEPLAN": "el informe técnico del CEPLAN"}
    de_ = lambda t: ("del " + t[3:]) if t.startswith("el ") else "de " + t   # «de el» -> «del»
    d.p(f"Las recomendaciones {de_(nombre_corto[avance.index[-1]])} "
        f"son las que más avanzaron ({fmt_num(100 * avance.iloc[-1], 0)} % implementadas en todo o en parte), y las "
        f"{de_(nombre_corto[avance.index[0]])}, las que menos ({fmt_num(100 * avance.iloc[0], 0)} %). La figura siguiente "
        f"presenta esas proporciones por evaluación.")
    fig(d, FI, "recomendaciones")
    tot = res.loc["Total"]
    d.p(f"De las {int(tot['Total'])} recomendaciones de las cuatro evaluaciones, {int(tot['Implementada'])} se implementaron y "
        f"{int(tot['Parcialmente implementada'])} lo hicieron en parte. Avanzaron sobre todo las que dependían de las "
        f"entidades responsables de los servicios; las de diseño de la Política quedaron pendientes porque dependían de "
        f"una actualización que no llegó a hacerse.")
    d.p("Para completar o trasladar lo pendiente (p3.6), la nueva política debería aprobar desde su inicio el protocolo "
        "de cada servicio con estándares de calidad e indicadores con línea de base, asegurar las fuentes de "
        "información de sus indicadores, sustentar su modelo conceptual en una revisión de evidencia documentada y "
        "acompañar las recomendaciones de esta evaluación con una matriz de compromisos cuyo cumplimiento se registre "
        "(Tabla 6 de CEPLAN, 2024a).")
    cierra(d, "I4", H, V)
    cierra(d, "I5", H, V)


def seccion_5(d, C, H, V, P, F, FI, est25, sfd, anios):
    d.h1("5. Análisis de los resultados")
    d.h2("5.1 Valoración integral del cumplimiento")
    abre(d, "I2")
    hallazgo(d, H, "I2-H1")
    rs = C["resumen_semaforo"]
    op = rs[rs.Nivel == "OP"]
    d.p("La tabla siguiente cuenta, por año, cuántos indicadores de OP quedaron en cada nivel de avance según el "
        "reporte oficial. El nivel compara el valor del año con su meta (numeral 2.4).")
    d.tabla(["Nivel de avance"] + [str(a) for a in anios],
            [[r["Semáforo PNIG"]] + [str(int(r[a])) for a in anios] for _, r in op.iterrows()],
            [3426] + [800] * 7, leyenda="Indicadores de OP por nivel de avance y año", size=16,
            fuente="Fuente: reportes de seguimiento y de cumplimiento de la PNIG (Anexo 5). Elaboración propia.",
            rastro="REV_IND_base_indicadores.xlsx, hoja «Resumen PNIG»", nombrar=False)
    sv = op.set_index("Semáforo PNIG")
    sinv = sv.loc[[i for i in sv.index if i.startswith("Sin valor")]].sum()
    d.p(f"La tabla muestra que el número de indicadores sin valor creció hacia el final del periodo: pasó de "
        f"{int(sinv[2019])} en 2019 a {int(sinv[2025])} en 2025. En los años con más datos, la mayoría de los "
        f"indicadores con valor estaba en nivel alto; el problema principal al cierre no es el bajo avance, sino la "
        f"falta de información para medirlo.")

    d.h2("5.2 Logros por objetivo prioritario")
    hallazgo(d, H, "I2-H2")
    d.p("La figura siguiente presenta el nivel de avance 2025 de cada indicador de OP en una sola columna: el del "
        "reporte oficial y, cuando el reporte no tiene valor, el calculado en esta evaluación, marcado con un asterisco.")
    fig(d, FI, "indicadores_2025")
    op_t = C["op_tabla"].set_index("Código")
    ofi_n = int(op_t["Semáforo 2025"].astype(str).str.startswith(("Alto", "Medio", "Bajo")).sum())
    calc = [c for c in op_t.index if c in est25.index
            and not str(op_t.loc[c, "Semáforo 2025"]).startswith(("Alto", "Medio", "Bajo"))]
    niv = pd.Series([str(op_t.loc[c, "Semáforo 2025"]).split(" ")[0] if c not in calc else str(est25.loc[c, "Nivel"])
                     for c in op_t.index]).str.capitalize().value_counts()
    sin_n = len(op_t) - ofi_n - len(calc)
    d.p(f"De los {len(op_t)} indicadores de objetivo prioritario, {ofi_n} tienen nivel en el reporte oficial y {len(calc)} lo tienen solo por el cálculo de esta "
        f"evaluación" + (f"; {sin_n} {S.pl(sin_n, 'queda', 'quedan')} sin nivel, por falta de valor o porque el valor "
                         f"no permite calcular el avance" if sin_n else "") + ". Considerando ambas fuentes, "
        f"{int(niv.get('Alto', 0))} quedan en nivel alto frente a su meta 2025. Que {len(calc)} de los {len(op_t)} "
        f"niveles dependan de un cálculo de esta evaluación es, en sí mismo, un resultado: el reporte oficial no "
        f"permite cerrar el periodo con información propia. Ese resultado debe leerse, además, con la advertencia del "
        f"párrafo siguiente: el nivel de avance no considera la línea de base.")
    s_ex = est25.loc["IOP.02.01"]
    lb21, mt21 = float(op_t.loc["IOP.02.01", "Valor LB"]), float(op_t.loc["IOP.02.01", "Meta 2025"])
    dist = 100 * (float(s_ex.Estimación) - lb21) / (mt21 - lb21)
    v43, lb43 = float(op_t.loc["IOP.04.03", "Último valor"]), float(op_t.loc["IOP.04.03", "Valor LB"])
    assert v43 < lb43 and str(op_t.loc["IOP.04.03", "Semáforo 2025"]).startswith("Medio")
    d.p(("Cómo leer el nivel de avance. ", {"b": True}),
        f"El nivel compara el valor obtenido con la meta del año (valor/meta, o meta/valor si el sentido es "
        f"descendente) y no considera la línea de base. Por eso un indicador puede quedar en nivel alto aunque haya "
        f"recorrido poco desde su punto de partida, o en nivel medio aunque haya empeorado. Por ejemplo, el uso de "
        f"métodos anticonceptivos (IOP.02.01) queda en nivel alto con {fmt_num(s_ex.Estimación, 1)} % en 2025, aunque "
        f"solo recorrió el {fmt_num(dist, 0)} % de la distancia entre su línea de base ({fmt_num(lb21, 1)} %) y su meta "
        f"({fmt_num(mt21, 1)} %); el índice de empoderamiento social de niñas y adolescentes (IOP.04.03) queda en nivel "
        f"medio con {fmt_num(v43, 2)}, aunque está por debajo de su línea de base ({fmt_num(lb43, 3)}). Por eso este "
        f"informe presenta, junto al nivel, la comparación con la línea de base.", size=20)
    filas = []
    for _, r in C["op_tabla"].iterrows():
        e = est25.loc[r["Código"]] if r["Código"] in est25.index else None
        s25 = str(r["Semáforo 2025"])
        if s25.startswith(("Alto", "Medio", "Bajo")) or e is None:
            nivel = s25
        else:   # sin nivel oficial: el calculado en esta evaluación, con asterisco
            val = num(e.Estimación) if r["Código"] in S.INDICES else fmt_num(e.Estimación, 1)
            nivel = f"{e.Nivel} ({val})*"
        filas.append([r["Código"], r["Indicador"], f"{num(r['Valor LB'])} ({r['Año LB']})",
                      f"{num(r['Último valor'])} ({r['Año del último valor']})", num(r["Meta 2025"]),
                      nivel, r["Respecto de la LB"]])
    d.p("La tabla siguiente presenta cada indicador de OP con su línea de base, su último valor, su meta 2025, su nivel "
        "2025 y su comparación con la línea de base. El nivel es el del reporte oficial; cuando el reporte no tiene "
        "valor 2025, es el calculado en esta evaluación y lleva un asterisco.")
    h23 = H["I2-H3"]
    d.tabla(["Código", "Indicador", "Línea de base", "Último valor", "Meta 2025", "Nivel 2025", "Frente a la LB"],
            filas, [1000, 2300, 1000, 1000, 800, 1926, 1000], leyenda="Indicadores de los objetivos prioritarios",
            size=15, rastro="REV_IND_base_indicadores.xlsx, hoja «Base REV_IND.1»; D1_D2_brechas_situacion_futura."
                            "xlsx, hoja «Manifestaciones ENDES»; REV_IND_4_recalculo_OP3.xlsx, hoja «Resultado 2025»",
            fuente="Fuente: reporte de seguimiento 2025 (Anexo 5); OGMEPGD (2026); INEI (2026), microdatos de la ENDES "
                   "y la ENAHO 2025; JNE (2022a); ONPE (2026a). Elaboración propia.",
            nota="* Nivel calculado en esta evaluación, porque el reporte no tiene valor 2025; entre paréntesis, el "
                 "valor estimado. No sustituye al reporte oficial (hallazgo I2-H3). La estimación de IOP.03.02 es "
                 "una alternativa de comparabilidad limitada, porque SERVIR cambió su metodología en 2023. "
                 + h23["nota"] + " " + h23["nota_tabla"].replace("(INEI, s. f.)", "(SERVIR, 2025, 2026)")
                 + " s. d.: sin dato.", nombrar=False)
    d.p(f"La tabla confirma que {C['op_mejora']} de los {C['op_ind']} indicadores mejoran frente a su línea de base. "
        f"Los vacíos se concentran en el OP6, cuyo único indicador depende de la ENARES, que no se aplicó en 2025, y en "
        f"la brecha salarial del sector público (IOP.05.02), cuyo valor 2025 no permite calcular el avance.")
    br = pd.read_excel(f"{EVAL}/04_analisis/D1_D2_brechas_situacion_futura.xlsx",
                       sheet_name="Brechas por año")
    gp = br[br.indicador == "ingreso_publico"].set_index("anio")
    ser = pd.read_excel(PSI, sheet_name="Serie anual")
    ofi = ser[ser["Código"] == "IOP.05.02"].set_index("Año")["Valor"]
    comp = [(a, 100 * (gp.loc[a, "Hombres"] - gp.loc[a, "Mujeres"]) / gp.loc[a, "Mujeres"], ofi.get(a))
            for a in (2019, 2020, 2021, 2023) if a in gp.index and pd.notna(ofi.get(a))]
    assert all(abs(v - o) > 0.5 for a, v, o in comp if a != 2020), "la ENAHO reproduce IOP.05.02: revisar el texto"
    d.p(("Por qué no se estimó la brecha salarial del sector público (IOP.05.02). ", {"b": True}),
        "La ficha fija la ENAHO como fuente, y el numeral 3.1 usa la ENAHO para medir el ingreso de los asalariados "
        "públicos. Pero esa serie, con la fórmula de la ficha, no reproduce los valores que reportó el indicador: "
        + lista_es([f"{fmt_num(v, 1)} frente a {fmt_num(o, 1)} en {a}" for a, v, o in comp]) +
        ". Como esta evaluación solo admite estimaciones que reproducen los valores oficiales de años anteriores, el "
        "indicador no se estimó. Su valor 2025 (−1,03) proviene de la planilla electrónica de SERVIR y no de la ENAHO, "
        "por lo que no es comparable con la serie ni con la meta. Un valor negativo expresa que la brecha cambió de "
        "signo; lo que impide usarlo es el cambio de fuente y que, con la fórmula de avance (meta sobre valor), un "
        "valor negativo no da un nivel.", size=20)
    hallazgo(d, H, "I2-H3")

    d.h2("5.3 Otros resultados")
    hallazgo(d, H, "I2-H4")
    cierra(d, "I2", H, V)
    hallazgo(d, H, "I5-H4")

    d.h2("5.4 Perspectivas para la situación futura deseada (D2)")
    abre(d, "D2")
    hallazgo(d, H, "D2-H1")
    fv = lambda x: fmt_num(x, 1) if pd.notna(x) else "s. d."
    d.p("La tabla siguiente compara, para cada enunciado de la situación futura, su línea de base, su meta al 2030, su "
        "último valor del periodo 2019–2025 y, por separado, el valor que alcanzaría al 2030 si mantuviera el ritmo "
        "observado (escenario lineal) o el dato observado en 2026.")
    lineal = sfd["Método del valor 2030"].fillna("").str.lower().str.contains("lineal|proyecc")
    ALC = {"Probable": "Alcanza (escenario lineal)", "Improbable": "No alcanza"}
    d.tabla(["Enunciado de la situación futura (p. 20)", "Indicador del OP", "Base", "Meta 2030",
             "Último valor 2019–2025", "Escenario lineal 2030", "Dato 2026", "Resultado"],
            [[[(r["Enunciado (cita literal)"], {"i": True})],
              (r['Indicador del OP'] if r['Indicador del OP'] == r['Correspondencia con el indicador del OP'] else
               f"{r['Indicador del OP']} ({r['Correspondencia con el indicador del OP'].lower()})")
              + {"SFD4a": "; meta de congresistas", "SFD4b": "; meta de alcaldesas"}.get(r["Código"], ""),
              fv(r["Línea de base"]), fv(r["Meta 2030"]), fv(r["Último valor"]),
              (fv(r["Valor 2030"]) + (" [POR VERIFICAR]" if r["Valor 2030 verificado"] == "No" else "")) if es_l else "",
              "" if es_l else fv(r["Valor 2030"]),
              ALC.get(r.Alcanzabilidad, r.Alcanzabilidad)] for (_, r), es_l in zip(sfd.iterrows(), lineal)],
            [2326, 1400, 650, 750, 900, 1000, 800, 1200], size=15, leyenda="Situación futura deseada al 2030",
            rastro="D1_D2_brechas_situacion_futura.xlsx, hoja «Situación futura deseada»",
            fuente="Fuente: reportes de seguimiento de la PNIG (Anexo 5); INEI (2026), microdatos de la ENDES; JNE y "
                   "ONPE, compilados por INEI (2025a); ONPE (2026a, 2026b). Elaboración propia.",
            nota="Último valor: oficial o calculado en esta evaluación con microdatos. Escenario lineal: el valor al "
                 "2030 con el ritmo observado, que no es un pronóstico ni expresa una probabilidad. Dato 2026: para "
                 "congresistas y alcaldesas, la composición elegida en 2026, fuera del periodo evaluado. " + NOTA_2026
                 + ".", nombrar=False)
    alc = sfd.Alcanzabilidad.value_counts()
    d.p("La tabla muestra que solo una meta se alcanzaría al 2030 con el ritmo observado (embarazo adolescente) y que "
        "las demás no se alcanzarían o no pueden evaluarse por falta de indicador o de dato ("
        + lista_es([f"{v}: {str(ALC.get(k, k)).lower()}" for k, v in alc.items()]) + "). La meta de congresistas, que se alcanzó "
        "en 2021, retrocede con el Congreso elegido en 2026. La figura siguiente muestra la serie histórica de la "
        "participación política de las mujeres.")
    figura_brecha(d, F, "series_participacion_politica", alt="Series históricas del porcentaje de mujeres en el "
                  "Congreso, las alcaldías, las regidurías y los gobiernos regionales.")
    d.p("La serie muestra avances sostenidos en las regidurías y avances mucho menores en las alcaldías y los gobiernos "
        "regionales, cargos que elige una sola persona. Los mapas siguientes llevan esa diferencia al territorio.")
    pp = pd.read_csv(f"{EVAL}/04_analisis/d1_brechas/participacion_politica_departamental.csv")
    alc = pp[pp.cargo.str.startswith("Alcaldías")]
    nac = alc.groupby("periodo")[["mujeres", "total"]].sum()
    per = [x for x in ("2019/2022", "2023/2026", "2027/2030") if x in nac.index]
    pct = {x: 100 * nac.loc[x, "mujeres"] / nac.loc[x, "total"] for x in per}
    ult = alc[alc.periodo == per[-1]]
    bajo5 = int((ult.porcentaje_mujeres < 5).sum())
    d.p("El primer mapa muestra, por departamento, el porcentaje de mujeres entre las alcaldías provinciales y "
        "distritales elegidas en cada periodo.")
    figura_brecha(d, F, MAPAS_D2[0], alt=f"Mapas del Perú por departamento: {F.loc[MAPAS_D2[0], 'titulo'].lower()}.")
    assert pct[per[-1]] < max(pct.values()), "la serie de alcaldesas ya muestra un avance sostenido: revisar texto"
    d.p("A escala nacional, la proporción de mujeres en las alcaldías fue de "
        + lista_es([f"{fmt_num(pct[x], 1)} % en {x}" for x in per]) + f", y en {per[-1]} hay {bajo5} departamentos "
        f"con menos de 5 % de alcaldesas. La proporción no muestra un avance sostenido y es desigual en el territorio: "
        f"la meta de la situación futura sigue lejos en casi todo el país.")
    gr = pp[pp.cargo == "Gobiernos regionales"]
    gob = {x: int((gr[gr.periodo == x].mujeres > 0).sum()) for x in sorted(gr.periodo.unique())}
    d.p("El segundo mapa muestra dónde una mujer ganó el gobierno regional en cada periodo.")
    figura_brecha(d, F, MAPAS_D2[1], alt=f"Mapas del Perú por departamento: {F.loc[MAPAS_D2[1], 'titulo'].lower()}.")
    d.p("Las mujeres electas para un gobierno regional son excepcionales en los tres periodos: "
        + lista_es([(f"ninguna en {k}" if v == 0 else f"{v} en {k}") for k, v in gob.items()])
        + ". Junto con el mapa de alcaldías, el resultado confirma que los cargos que se ganan por elección "
        "unipersonal siguen siendo el espacio donde la participación política de las mujeres menos avanza.")
    for i in ("D2-H2", "D2-H3"):
        hallazgo(d, H, i)
    cierra(d, "D2", H, V)

    d.h2("5.5 Factores que influyeron en los resultados")
    d.p("Los factores que las entidades asocian a los resultados se tratan con los de implementación en el numeral 4.5. "
        "En la medición, el principal factor es la disponibilidad de las fuentes: varias encuestas nacionales no se "
        "publicaron o no se aplicaron en los años previstos, y algunas entidades calcularon sus indicadores con "
        "fuentes distintas de las que fija la ficha (hallazgo D5-H3).")
    d.h2("5.6 Elementos críticos al cierre")
    criticos = []
    for _, r in C["op_tabla"].iterrows():
        s25 = str(r["Semáforo 2025"])
        if not s25.startswith(("Alto", "Medio")):
            if r["Código"] in est25.index:
                e = est25.loc[r["Código"]]
                s25 += f"; estimación de esta evaluación: {S.fmt_est(r['Código'], e.Estimación)} ({e.Nivel.lower()})"
            criticos.append([f"{r['Código']} ({r['OP']})", s25, "Indicadores y recursos"])
    for c in C["serv_noimpl_cod"]:
        criticos.append([f"Servicio {c}", "No implementado", "Revisión documental"])
    for c in C["lin_sin_serv"]:
        criticos.append([f"Lineamiento {c}", "Sin servicio", "Coherencia interna"])
    d.p("La tabla siguiente reúne los elementos de la Política que cierran el periodo con alerta: indicadores de OP sin "
        "valor 2025, no interpretables o en nivel bajo; servicios no implementados, y lineamientos sin servicio.")
    d.tabla(["Elemento", "Alerta al cierre", "Análisis"], criticos, [2800, 4426, 1800],
            leyenda="Objetivos prioritarios y servicios con alerta crítica",
            fuente="Fuente: hallazgos de esta evaluación (Anexos 2 y 3). Elaboración propia.")
    d.p("La lista muestra que las alertas se concentran en la medición (indicadores sin valor) más que en el avance, y "
        "que los servicios no implementados son pocos. Son los elementos que la transición debe resolver primero.")


def metodologia(d, C, H, V, P, FI):
    """Sección 2 (ronda 33): sin fases, cronograma, independencia como numeral, disponibilidad de datos ni
    correspondencia con la Tabla A7.1; la independencia pasa a las limitaciones; las escalas de implementación y de
    avance son diagramas; cada figura y tabla con su entrada, su lectura y su análisis."""
    rev = P["rev"]
    n = C["serv_total"]
    d.h1("2. Metodología")
    d.h2("2.1 Tipo de evaluación y enfoque")
    d.p("Conforme a la Guía de seguimiento y evaluación del CEPLAN, esta es una evaluación de cierre (CEPLAN, 2024a, "
        "p. 73): valora la trayectoria de la Política durante todo su periodo de implementación y orienta sus "
        "recomendaciones a la formulación de la política que la sustituye. La guía lo expresa así:")
    d.cita("Las recomendaciones de este informe deben retroalimentar el nuevo proceso de formulación", "CEPLAN, 2024a, p. 73")
    d.p("La evaluación la condujo la Oficina de Monitoreo y Evaluación de Políticas (OMEP) de la OGMEPGD y comprende "
        "dos dimensiones:")
    d.vineta(("Evaluación de diseño. ", {"b": True}), "Verifica la vigencia del diagnóstico y de la situación futura "
             "deseada, la coherencia interna y la plausibilidad causal de la cadena de resultados, el alineamiento "
             "vertical y horizontal, la calidad técnica de los indicadores, la incorporación de los enfoques "
             "transversales y las implicancias del nuevo marco conceptual.")
    d.vineta(("Evaluación de implementación. ", {"b": True}), "Determina el grado de implementación de los servicios, "
             "la trayectoria de los indicadores, los recursos que el MIMP asignó a la conducción de la Política, el "
             "funcionamiento de la rectoría y de la articulación, y las capacidades y condiciones de sostenibilidad.")
    d.p("La evaluación abarca todo el periodo, desde la aprobación de la PNIG en abril de 2019 hasta el 31 de diciembre "
        "de 2025, último año de vigencia de la PNIG; cada análisis usa la información de todos los años "
        "disponibles en ese rango. Se distinguen tres fechas: el periodo evaluado (2019–2025), el corte de la evidencia "
        "(noviembre de 2026) y la transición (2026), que se presenta como contexto y solo entra en las preguntas que la "
        "abordan (D4, D7 e I5). Su alcance es nacional; el análisis territorial usa la desagregación departamental "
        "de las encuestas y de los registros. La unidad de análisis es la Política en su conjunto: la evaluación pone a "
        "prueba su teoría de cambio, la cadena que vincula el problema público y sus causas con los objetivos, los "
        "lineamientos y los servicios, para distinguir las fallas de diseño de las de implementación.")
    d.p("El enfoque es mixto: combina el análisis documental, la revisión de evidencia, el análisis cuantitativo de "
        "indicadores, encuestas y presupuesto, y el análisis cualitativo de lo que reportaron las entidades y de las "
        "entrevistas. Quedan fuera del alcance la medición de impacto y la atribución causal de los cambios en las "
        "brechas a la Política, el recojo de información con la población usuaria, la solicitud de información nueva "
        "a las entidades, la eficiencia en sentido estricto (costo por resultado) y el análisis jurídico de la Ley N.° "
        "32535.")

    d.h2("2.2 Preguntas y criterios de evaluación")
    d.p("La evaluación aplica seis criterios: cinco de los criterios del Comité de Ayuda al Desarrollo de la OCDE "
        "(OCDE, 2019), sin el de impacto, y el de implementación, que recoge la dimensión de procesos que la Guía del "
        "CEPLAN prevé para las políticas nacionales. La tabla siguiente define cada criterio tal como se aplicó y "
        "señala las preguntas que lo responden. Dos preguntas responden a más de un criterio: D7, a pertinencia y a "
        "coherencia externa, y D5, que valora la calidad técnica de los indicadores como parte de la coherencia entre la "
        "cadena de resultados y su medición.")
    criterios = [
        ("Pertinencia", "Grado en que el diagnóstico, la situación futura deseada y los servicios responden al problema "
         "público vigente y al marco normativo y conceptual actual.", "D1, D2 y D7"),
        ("Coherencia", "Consistencia interna de la cadena de resultados y alineamiento externo con el planeamiento "
         "nacional, otras políticas nacionales y los compromisos internacionales; incluye la calidad técnica de los "
         "indicadores (D5).", "D3 a D7"),
        ("Eficacia", "Grado en que los servicios fueron implementados y en que los indicadores de los objetivos "
         "prioritarios avanzaron respecto de sus metas.", "I1, I2"),
        ("Eficiencia (alcance acotado)", "Recursos presupuestales que el MIMP asignó y ejecutó en la conducción de la "
         "Política.", "I3"),
        ("Implementación", "Funcionamiento de la rectoría del MIMP y de los mecanismos de articulación previstos para la "
         "Política.", "I4"),
        ("Sostenibilidad", "Condiciones normativas, institucionales y financieras para preservar los avances ante la "
         "transición.", "I5"),
    ]
    d.tabla(["Criterio", "Definición aplicada", "Preguntas"], [list(c) for c in criterios], [2000, 5826, 1200], size=16,
            leyenda="Criterios de evaluación", fuente="Fuente: OCDE (2019) y CEPLAN (2024a). Elaboración propia.",
            nombrar=False)
    d.p("La tabla muestra dónde pone el acento la evaluación. La coherencia reúne cinco de las doce preguntas, porque "
        "una evaluación de cierre que alimenta una nueva formulación necesita saber, sobre todo, si la lógica de la "
        "Política se sostiene. La eficiencia se acota a la conducción, porque los datos públicos no identifican el "
        "gasto de los servicios por política. La tabla siguiente presenta el texto completo de las doce preguntas.")
    d.tabla(["Pregunta", "Criterio", "Texto de la pregunta"],
            [[p, crit, [(texto, {"i": True})]] for p, (crit, texto, _) in S.PREGUNTAS.items()],
            [900, 1900, 6226], size=16, leyenda="Preguntas de evaluación", fuente="Fuente: elaboración propia.",
            nombrar=False)
    d.p("Las preguntas siguen el recorrido de la cadena de valor: las de diseño (D1 a D7) van del problema público y "
        "la situación futura a la coherencia, el alineamiento, los indicadores, los enfoques y el nuevo marco "
        "conceptual; las de implementación (I1 a I5) van de los servicios y los indicadores a los recursos, la "
        "coordinación y la sostenibilidad. Cada pregunta se responde con más de un instrumento, como muestra la figura "
        "siguiente.")
    fig(d, FI, "matriz_evaluacion")
    cob = pd.read_excel(f"{EVAL}/04_analisis/matriz_triangulacion_preguntas.xlsx", sheet_name="Cobertura",
                        dtype=str).set_index("Pregunta")
    usados = {p: [x.strip() for x in str(cob.loc[p, "Instrumentos con evidencia en la matriz"]).split(",") if x.strip()]
              for p in S.PREGUNTAS}
    por_preg = pd.Series({p: len(u) for p, u in usados.items()})
    por_instr = pd.Series([i for u in usados.values() for i in u]).value_counts()
    nombre_instr = {"REV_DOC": "la revisión documental", "ENTREV": "las entrevistas",
                    "REP_ENT": "la información de las entidades", "REV_ALI": "el análisis de coherencia y alineamiento",
                    "REV_IND": "la revisión de indicadores y recursos", "REV_EVI": "la revisión de evidencia"}
    i1, i2 = por_instr.index[0], por_instr.index[1]
    imin = por_instr.index[-1]
    d.p(f"Ninguna pregunta descansa en un solo instrumento: cada una se sostiene en entre {por_preg.min()} y {por_preg.max()}. "
        f"Los instrumentos más transversales son {nombre_instr[i1]} ({por_instr[i1]} preguntas) y "
        f"{nombre_instr[i2]} ({por_instr[i2]}); {nombre_instr[imin]}, en cambio, se concentra en "
        f"{por_instr[imin]} {S.pl(int(por_instr[imin]), 'pregunta', 'preguntas')}. Esa combinación es la que permite "
        f"contrastar cada hallazgo con fuentes distintas (numeral 2.5).")

    d.h2("2.3 Instrumentos y procedimientos")
    d.p("La evaluación usó seis instrumentos. Cada uno tiene un procedimiento propio, que se resume a continuación; los "
        "documentos de cada fuente están en el Anexo 10.")
    d.h3("Revisión documental")
    d.p(f"Sistematiza la evidencia documental sobre el diseño de la Política y determina el estado de implementación de "
        f"cada servicio. El corpus reúne {P['n_corpus']} documentos: normas, documentos de la Política y del PEMIG, "
        f"protocolos, reportes de seguimiento y de cumplimiento, informes LIO y de la CONAIG, evaluaciones anteriores, "
        f"instrumentos de planeamiento y políticas afines. Para cada servicio se registró, en todo el periodo "
        f"2019–2025, si tiene protocolo aprobado, si su presupuesto puede identificarse en los datos abiertos del MEF, "
        f"si reporta indicadores y qué ejecución reportó en los informes LIO de cada año; con esos componentes se "
        f"clasificó su estado (numeral 2.4). La revisión incluye una lista de verificación de los siete enfoques "
        f"transversales en el documento de la Política, sus anexos y los protocolos.")
    d.h3("Entrevistas al nivel conductor")
    g_ = C["entrev"]
    d.p(f"Se entrevistó a profesionales de la OGPPM, la OGMEPGD y la DGIGND, los tres órganos del MIMP que conducen la "
        f"Política o intervienen en su seguimiento, a los que este informe llama «nivel conductor». La guía tenía "
        f"preguntas orientadoras por bloque y preguntas de contraste de los hallazgos preliminares. Las entrevistas se "
        f"grabaron con consentimiento verbal, se transcribieron, se verificaron en el audio los pasajes dudosos y se "
        f"anonimizaron, de modo que cada intervención se atribuye al órgano y no a la persona. Las respuestas se "
        f"dividieron en {len(g_)} segmentos, codificados con el mismo libro de códigos que la información de las "
        f"entidades; el especialista evaluador que condujo las entrevistas revisó y validó {C['entrev_validados']} de "
        f"ellos, incluidos los {C['entrev_criticos']} que podían cambiar una valoración, con el libro de códigos como "
        f"pauta, y registró la validación segmento por segmento. Es un control de consistencia, no una doble "
        f"codificación independiente (numeral 2.6). La aplicación se detalla en el Anexo 8.")
    d.h3("Información reportada por las entidades")
    d.p(f"Analiza lo que las entidades responsables reportaron al MIMP entre 2019 y 2025, sin pedirles información "
        f"nueva. La base son los reportes de seguimiento y de cumplimiento de la PNIG, que el MIMP elabora cada año "
        f"desde el primer año de implementación (2020), con los valores de los indicadores de objetivo desde 2019: de "
        f"ellos salen la trayectoria oficial de los indicadores y su nivel de avance (Anexo 5). Les siguen los informes "
        f"anuales de cumplimiento de la Ley N.° 28983 (informes LIO) de 2019 a 2025, que recogen, servicio por "
        f"servicio, las acciones, la cobertura y las dificultades que reportó cada entidad, los informes de "
        f"gobernanza de la CONAIG (2022 a 2025) y la ayuda memoria de sus sesiones de 2021 a 2026, con los acuerdos de "
        f"cada año. Como insumo adicional, el formulario de seguimiento de la OMEP (2023 a "
        f"2025) aporta, por indicador, el valor, el nivel de cumplimiento y las dificultades de los últimos tres años.")
    d.p(f"Las secciones por servicio de los informes y los campos abiertos del formulario se dividieron en "
        f"{fmt_num(C['seg_total'])} segmentos ({fmt_num(C['seg_lio'])} de los informes LIO, "
        f"{fmt_num(C['seg_form'])} del formulario y {C['seg_conaig']} de los de la CONAIG), cada uno codificado con un "
        f"libro de códigos deductivo derivado de las preguntas de evaluación. Los {C['dc_n']} segmentos que podían "
        f"cambiar una valoración se codificaron dos veces, de forma independiente y ciega: el segundo codificador solo "
        f"veía el texto, sus datos y el libro de códigos. El acuerdo se midió con el kappa de Cohen, que descuenta el "
        f"acuerdo esperado por azar (1 es acuerdo perfecto); el kappa medio por código fue de "
        f"{fmt_num(C['dc_kappa_cod'], 2)}, y las discrepancias se resolvieron caso por caso con el texto del segmento. "
        f"El kappa describe el acuerdo en esos segmentos, no la confiabilidad de todo el corpus; el libro de códigos, la "
        f"matriz de los segmentos con ambas codificaciones y el acuerdo por código están en el anexo técnico.")
    d.h3("Análisis de coherencia y alineamiento")
    d.p("Contrasta la cadena de resultados de la Política (coherencia interna: cada lineamiento con servicio, cada "
        "servicio con causa o efecto e indicador); el alineamiento vertical de cada OP con las Políticas de Estado, el "
        "plan de desarrollo nacional, la PGG y el PESEM, en el periodo y en la transición; el alineamiento horizontal "
        "con las políticas nacionales afines del registro del CEPLAN; la correspondencia con la Ley N.° 32535 y con seis "
        "instrumentos internacionales, y la correspondencia conceptual con el nuevo enfoque de igualdad de "
        "oportunidades. Cada relación se registra con su cita literal y su página.")
    d.h3("Revisión de indicadores y recursos")
    d.p("Procesa las series 2019–2025 de los indicadores de OP y de servicio con su meta anual, su nivel de avance y su "
        "comparación con la línea de base; verifica la calidad de las fichas técnicas frente a la guía de indicadores "
        "del CEPLAN (CEPLAN, 2024c), y analiza los recursos de la conducción con los datos abiertos del gasto del MEF "
        "(MEF, 2026) por meta presupuestal de cada órgano, contrastados con la ejecución por finalidad y específica de "
        "gasto que remitió la DGIGND para la actividad 5001072 «Promoción de la mujer» (MIMP, DGIGND, 2026a). Los "
        "montos que las entidades reportaron para sus servicios de "
        "2019 a 2025 se cotejaron con las metas, actividades y productos de esos datos: un monto cuenta como recurso "
        "identificado si coincide con una estructura de la entidad responsable; la coincidencia identifica la "
        "estructura que lo financia, pero no prueba que ese gasto sea exclusivo del servicio.")
    d.p("Las brechas por sexo y los indicadores sin valor oficial se estimaron con los microdatos de las encuestas del "
        "INEI (INEI, 2026), respetando su diseño muestral complejo: estratos, conglomerados y factores de expansión. "
        "Cada estimación lleva su intervalo de confianza al 95 %. La brecha se estima como diferencia o razón dentro "
        "del mismo diseño, de modo que su error estándar incorpora la covarianza entre las estimaciones de mujeres y de "
        "hombres, y su significancia se prueba con ese error. Los cambios entre años se prueban como diferencias entre "
        "muestras independientes; como la ENAHO tiene un componente de panel que no se aprovechó, esa prueba es "
        "conservadora. Una estimación por sexo es confiable si su coeficiente de variación es de 15 % o menos, "
        "con el criterio del INEI (INEI, 2025a, p. 178); si lo supera, es referencial y se marca. Un indicador sin valor "
        "oficial solo se estimó si el método reproduce los valores oficiales de años anteriores; así se recalcularon "
        "los indicadores de los OP1 y OP2 con la ENDES y los del OP3 con la ENAHO y los registros electorales. El "
        "empleo informal de 2024 y 2025 se reconstruyó con la regla del INEI para la ENAHO (INEI, 2025c, pp. 87 a 90), "
        "que reproduce la variable oficial de 2019 a 2023.")
    d.h3("Revisión de evidencia")
    d.p(f"Revisión rápida de literatura y evaluaciones para valorar si los tipos de intervención de los servicios "
        f"producen los resultados que la PNIG espera de ellos. Los {n} servicios se agruparon en "
        f"{C['rev_tipos']} tipos de intervención. Para cada tipo se buscaron revisiones sistemáticas y evaluaciones de "
        f"impacto con estrategias replicables: {rev['academica']} estrategias en bases académicas con sus tesauros "
        f"({rev['bases']}), {rev['dep']} en el portal de evidencia del 3ie, {rev['manual']} manuales en SciELO, LILACS, "
        f"Scopus y JSTOR y {rev['web']} búsquedas en la web y en portales de organismos y de evaluaciones peruanas. Cada "
        f"estrategia registra su cadena exacta, sus descriptores, sus filtros, su fecha y su número de resultados. Las "
        f"búsquedas se ejecutaron el 31 de octubre y el 1 de noviembre de 2026; de cada estrategia académica se "
        f"revisaron entre 9 y 33 registros por relevancia, y 176 de los 2219 registros recuperados se eliminaron como "
        f"duplicados. El detalle de cada estrategia y de la selección está en el anexo técnico.")
    d.p(f"La selección de documentos tuvo dos etapas: primero, los títulos de los primeros registros por relevancia de "
        f"cada estrategia y, después, la lectura del resumen de los seleccionados. Se incluyeron, sin límite de año, las "
        f"evaluaciones experimentales y cuasiexperimentales del Perú o de países de ingreso bajo o medio, y las "
        f"revisiones y estudios descriptivos por su alcance; se excluyeron opiniones, tesis, resúmenes de congreso y "
        f"documentos sin método declarado. Cada fuente se verificó en Crossref, PubMed o la página de la revista; en "
        f"{P['resumen_res']} de las {len(P['ref'])} fuentes únicas registradas (una misma fuente puede respaldar más "
        f"de un tipo) el resultado se tomó del resumen, no del texto completo: el plazo de la evaluación no permitió "
        f"revisar el texto completo de esas fuentes, por lo que la solidez de los tipos que dependen de ellas se lee con "
        f"esa reserva. La solidez, la dirección y la pertinencia de la evidencia de cada tipo se clasifican con las reglas "
        f"del numeral 2.4.")

    d.h2("2.4 Reglas de clasificación")
    d.p("Cada juicio del informe se apoya en una escala explícita, con sus categorías y sus criterios, para que la misma "
        "evidencia conduzca siempre a la misma clasificación. A continuación se presenta cada escala, para qué sirve y "
        "cómo resultó al aplicarla.")
    d.p("La primera escala clasifica el estado de implementación de cada servicio con tres componentes, protocolo, "
        "recursos e indicadores, y con la ejecución que reportó en todo el periodo. La figura siguiente la presenta "
        "como un árbol de decisión.")
    fig(d, FI, "escala_implementacion")
    sens3 = C["grd"]["Estado con tres de tres componentes (sensibilidad)"].value_counts()
    d.p(f"La primera condición del árbol separa los servicios que nunca se ejecutaron de los que reportaron ejecución "
        f"en algún año de 2019 a 2025. Un año tiene ejecución reportada cuando el informe LIO de ese año trae "
        f"información detallada del servicio, es decir, una sección que describe sus acciones sin declarar falta de "
        f"información, o cuando alguno de sus indicadores tiene valor; las menciones breves no cuentan, y la lectura de "
        f"cada sección descarta las que solo describen acciones preparatorias. La segunda condición considera "
        f"implementado el servicio que reúne al menos dos de los tres componentes y sigue operando al cierre, con "
        f"ejecución reportada en 2024 o 2025. El servicio que se ejecutó, pero carece de dos o más componentes o dejó "
        f"de reportar ejecución antes de 2024, queda como implementado parcialmente. Aplicada a los {n} servicios, la escala da {C['serv_impl']} implementados, "
        f"{C['serv_parc']} parciales y {C['serv_noimpl']} {S.pl(C['serv_noimpl'], 'no implementado', 'no implementados')}. "
        f"Con la regla anterior, que exigía los tres componentes, serían {sens3.get('Implementado', 0)} implementados, "
        f"{sens3.get('Implementado parcialmente', 0)} parciales y {sens3.get('No implementado', 0)} no implementados: "
        f"la regla pesa en el resultado, porque el protocolo habilita el indicador y esos dos componentes no son "
        f"evidencias independientes de la implementación. "
        f"Considerar todo el periodo, y no solo los últimos años, modifica la clasificación: dos servicios que no "
        f"reportaron provisión entre 2023 y 2025 sí la reportaron en 2021 y 2022, y uno reunió recursos identificables "
        f"antes de 2023.")
    rs = C["resumen_semaforo"]
    op_rs = rs[rs.Nivel == "OP"].set_index("Semáforo PNIG")[2025]
    con_nivel = int(op_rs[[i for i in op_rs.index if not str(i).startswith("Sin valor")]].sum())
    d.p("La segunda escala mide el avance de cada indicador frente a su meta del año, con la misma regla de los "
        "reportes de cumplimiento de la PNIG, y la figura siguiente la presenta como una regla graduada.")
    fig(d, FI, "escala_avance")
    d.p(f"La regla divide el avance (valor sobre meta para indicadores ascendentes y meta sobre valor para "
        f"indicadores descendentes, en porcentaje) en cinco tramos: bajo, medio y alto, más dos "
        f"extremos con trama que advierten una posible falla de planeación, porque un avance muy por debajo o muy por "
        f"encima de la meta suele indicar una meta mal fijada antes que un desempeño excepcional. Los casos que no "
        f"pueden medirse, como un indicador sin valor, sin meta o con valor o meta igual a cero, no se clasifican y se registran como hallazgo de "
        f"calidad técnica. La escala tiene un límite que el informe tiene presente: no mira la línea de base, de modo "
        f"que un indicador puede quedar en nivel alto aunque haya avanzado poco desde su punto de partida (numeral "
        f"5.2). En 2025, según el reporte oficial, solo {con_nivel} de los {C['op_ind']} indicadores de objetivo tienen "
        f"un valor que permita clasificarlos.")
    d.p("Las dos escalas siguientes ordenan el razonamiento de la evaluación: la confianza mide cuánto respaldo tiene "
        "cada hallazgo, y la valoración, qué tan bien responde la evidencia a cada pregunta.")
    d.tabla(["Nivel", "Criterio"], [
        ["Sólido", "Al menos dos fuentes de origen independiente coinciden."],
        ["Con respaldo parcial", "Fuentes de origen independiente coinciden, con discrepancias que se explican."],
        ["Indicio", "Se sustenta en un solo origen independiente."],
    ], [2400, 6626], size=16, leyenda="Confianza de los hallazgos", fuente="Fuente: elaboración propia.",
       nota="Origen independiente: productor primario distinto del dato. Dos documentos que reproducen el dato de un "
            "mismo productor cuentan como uno; así, los informes LIO, el formulario de seguimiento, los informes de la "
            "CONAIG y los reportes de seguimiento recogen lo que reportan las entidades, y los datos del MEF y la "
            "información presupuestal de la DGIGND provienen del mismo registro presupuestal. La entrevista a la OGMEPGD "
            "no cuenta como fuente independiente (numeral 2.6). En la revisión de evidencia, cada estudio incluido "
            "cuenta como una fuente.", nombrar=False)
    confs = pd.Series([h["confianza"] for h in H.values()]).value_counts()
    d.p(f"La confianza depende del número de fuentes independientes que respaldan el hallazgo, no de la opinión de "
        f"quien lo formula. De los {len(H)} hallazgos del informe, {confs.get('Sólido', 0)} son sólidos, "
        f"{confs.get('Con respaldo parcial', 0)} tienen respaldo parcial y {confs.get('Indicio', 0)} son indicios; "
        f"cada uno lleva su nivel en la nota al pie.")
    d.tabla(["Valoración", "Criterio"], [
        ["Satisfactoria", "El atributo que evalúa la pregunta se cumple, según evidencia suficiente."],
        ["Parcial", "El atributo se cumple en parte: la evidencia muestra avances junto con vacíos o deficiencias "
                    "identificadas."],
        ["Insuficiente", "El atributo no se cumple en sus aspectos críticos, según la evidencia disponible, o la "
                         "evidencia es ausente o contradictoria."],
    ], [2400, 6626], size=16, leyenda="Valoración de las preguntas de evaluación", fuente="Fuente: elaboración propia.",
       nombrar=False)
    vv = pd.Series({p: V[p][0] for p in S.PREGUNTAS}).value_counts()
    d.p(f"La valoración se asigna a cada pregunta a partir del conjunto de sus hallazgos. Ninguna de las doce alcanza la "
        f"categoría satisfactoria: {vv.get('Parcial', 0)} son parciales y {vv.get('Insuficiente', 0)} "
        f"{S.pl(vv.get('Insuficiente', 0), 'es insuficiente', 'son insuficientes')}, como se detalla en el numeral 2.5 "
        f"y en las conclusiones.")
    d.p("El análisis de diseño usa cuatro escalas más, para los enfoques transversales, el alineamiento, la "
        "articulación con otras políticas y la relación con el nuevo marco conceptual. La clasificación de los "
        "servicios según el nuevo marco se contrastó, además, con la matriz de incompatibilidades de la DPIGND, "
        "traducida a las mismas categorías con una regla fija y con su acuerdo medido por el kappa de Cohen "
        "(numeral 3.2).")
    d.tabla(["Escala", "Categorías y criterio"], [
        ["Incorporación de enfoques", "Ausente; mencionado (se enuncia sin desarrollo); desarrollado (orienta el "
                                      "diagnóstico o los objetivos); operacionalizado (se traduce en lineamientos, "
                                      "servicios o indicadores)."],
        ["Alineamiento vertical e internacional", "Directo: el instrumento enuncia el mismo cambio sobre la misma "
                                                  "población. Indirecto: el OP contribuye a un objetivo más amplio. Sin "
                                                  "correspondencia: ningún objetivo, eje o lineamiento guarda relación."],
        ["Articulación horizontal", "Duplicidad: el mismo servicio, con el mismo proveedor, figura en ambas políticas. "
                                    "Complementariedad: la política afín atiende otra población, otro componente o el "
                                    "mismo problema con otro abordaje. Vacío de atención: una población o necesidad "
                                    "queda sin servicio en ninguna de las dos, o quedará sin él en la transición."],
        ["Relación conceptual", "Continuidad: la categoría se mantiene sin variación sustantiva. Precisión: el nuevo "
                                "marco delimita su contenido. Divergencia: ambos marcos no son equivalentes. Efecto "
                                "sobre el elemento: se mantiene, se ajusta o se rediseña."],
    ], [2400, 6626], size=16, leyenda="Escalas del análisis de diseño", fuente="Fuente: elaboración propia.",
       nombrar=False)
    rel_c = pd.read_excel(MAC, sheet_name="REV_ALI.5 Conceptual", dtype=str)["Tipo de relación"].value_counts()
    vp = C["vert_periodo"]
    d.p(f"Cada escala se aplicó relación por relación, con su cita literal. Así, el alineamiento vertical del periodo "
        f"registra {vp[0]} relaciones directas, {vp[1]} indirectas y {vp[2]} sin correspondencia; la articulación "
        f"horizontal, {C['hor_dup']} duplicidades, {C['hor_comp']} complementariedades y {C['hor_vac']} vacíos de "
        f"atención, y la relación conceptual, {rel_c.get('Continuidad', 0)} continuidades, "
        f"{rel_c.get('Precisión', 0)} precisiones y {rel_c.get('Divergencia', 0)} divergencias. Los resultados se "
        f"analizan en los numerales 3.2, 3.5 y 3.7.")
    d.p("La última escala valora la evidencia de la literatura sobre cada tipo de intervención, con tres atributos. "
        "Mide el respaldo que la literatura da a la lógica del servicio, por lo que combina la calidad de los estudios "
        "con la dirección de sus resultados: la evidencia rigurosa con efectos nulos o adversos no respalda el servicio "
        "y se registra como dirección nula o mixta.")
    d.tabla(["Atributo", "Categorías y criterio"], [
        ["Solidez", "Sólida: al menos una revisión sistemática favorable cuya evidencia no sea de calidad baja, o dos "
                    "evaluaciones experimentales o cuasiexperimentales, todas favorables. Moderada: evaluaciones "
                    "aisladas o revisiones con resultados mixtos o de calidad baja. Limitada: solo estudios "
                    "descriptivos, revisiones narrativas o estudios de contextos poco comparables. Sin evidencia "
                    "identificada."],
        ["Contexto comparable", "La evidencia de países de ingreso alto y la de revisiones cuyo resumen no informa el "
                                "ámbito cuentan para la solidez limitada, no para la moderada ni para la sólida."],
        ["Dirección", "Favorable: todos los resultados van en el sentido esperado. Mixta: resultados favorables junto "
                      "con resultados nulos o no concluyentes."],
        ["Pertinencia", "Directa: la evidencia mide el resultado que persigue el servicio, en su población. Parcial: mide "
                        "otro resultado o cubre solo una parte de la población del servicio."],
    ], [2400, 6626], size=16, leyenda="Reglas de la revisión de evidencia", fuente="Fuente: elaboración propia.",
       nota="El Anexo 4 aplica estas reglas a cada tipo de intervención y explica su clasificación.", nombrar=False)
    sol = P["mt"].Solidez.value_counts()
    d.p(f"La regla de solidez es deliberadamente exigente con el contexto: una revisión de países de ingreso alto no "
        f"basta para calificar de sólida la evidencia de un servicio peruano. Aplicada a los {C['rev_tipos']} tipos de "
        f"intervención, dio {sol.get('Sólida', 0)} con evidencia sólida, {sol.get('Moderada', 0)} con evidencia "
        f"moderada y {sol.get('Limitada', 0)} con evidencia limitada (numeral 3.4).")

    d.h2("2.5 Triangulación")
    d.p("Cada hallazgo se construye contrastando la evidencia en tres planos: de fuentes (documentos normativos, "
        "registros administrativos, estadística oficial y literatura), de métodos (análisis cuantitativo, documental y "
        "temático) y de actores (la lectura del nivel conductor frente a lo que reportaron las entidades). La "
        "confianza del hallazgo depende del número de fuentes independientes que coinciden (numeral 2.4), y la "
        "valoración de cada pregunta se apoya en el conjunto de sus hallazgos.")
    d.p("En las secciones 3 a 5, cada hallazgo es un párrafo del texto, en un recuadro. Su nota al pie indica su "
        "identificador (por ejemplo, D1-H1: hallazgo 1 de la pregunta D1), su nivel de confianza y las fuentes que lo "
        "respaldan; el detalle está en el Anexo 2. Cada numeral abre con la pregunta de evaluación que responde y "
        "cierra con la respuesta. La figura siguiente adelanta el resultado de ese proceso: la valoración de las doce "
        "preguntas y la confianza de los hallazgos que sostienen cada una.")
    fig(d, FI, "valoracion", alt="Valoración de cada pregunta de evaluación y, por pregunta, un punto de color por "
        "cada hallazgo según su nivel de confianza: verde, sólido; ámbar, con respaldo parcial; rojo, indicio.")
    insuf = [p for p in S.PREGUNTAS if V[p][0] == "Insuficiente"]
    d.p(f"Ninguna pregunta alcanza una valoración satisfactoria: {vv.get('Parcial', 0)} son parciales y "
        f"{vv.get('Insuficiente', 0)} {S.pl(vv.get('Insuficiente', 0), 'insuficiente', 'insuficientes')} "
        f"({lista_es(insuf)}, calidad de los indicadores). La valoración parcial predomina tanto en las preguntas de "
        f"diseño como en las de implementación, lo que indica que los problemas no se concentran en una sola etapa de "
        f"la Política. El predominio de los hallazgos sólidos (en verde) muestra que esas valoraciones descansan sobre "
        f"todo en dos o más fuentes independientes.")

    d.h2("2.6 Limitaciones")
    lim, _, _ = _listas(C)
    for t in lim:
        d.vineta(t)


def evidencia_34(d, C, H, P, FI, nombre_serv):
    """Numeral 3.4: la evidencia como marco teórico, por objetivo prioritario y tipo de intervención."""
    tip, mt, me = P["tip"], P["mt"], P["me"]
    d.p(f"La revisión de evidencia responde si los tipos de intervención de los servicios producen los resultados que "
        f"la PNIG espera de ellos. Los {C['serv_total']} servicios se agruparon en {C['rev_tipos']} tipos de "
        f"intervención, y para cada uno se buscó evidencia con el procedimiento del numeral 2.3. La figura siguiente "
        f"resume la búsqueda y la selección.")
    fig(d, FI, "flujo_evidencia")
    d.p(f"La figura muestra el embudo de la revisión: de los registros recuperados se revisaron los primeros por "
        f"relevancia de cada estrategia, y se incluyeron {C['rev_fuentes']} fuentes. Las búsquedas académicas aportaron "
        f"la mayoría de las fuentes incluidas; la búsqueda web y en portales, las evaluaciones peruanas.")
    hallazgo(d, H, "D3-H3")
    d.p("Tres atributos describen la evidencia de cada tipo. La solidez indica cuánto respaldo tiene: sólida, moderada o "
        "limitada, según el diseño de los estudios y lo comparable de su contexto. La dirección indica el sentido de los "
        "resultados: favorable, si todos van en el sentido esperado, o mixta, si hay resultados nulos o no concluyentes. "
        "La pertinencia indica si la evidencia mide el resultado que persigue el servicio (directa) o solo uno "
        "relacionado (parcial). Las reglas completas están en el numeral 2.4. La figura siguiente presenta los tres "
        "atributos por tipo.")
    fig(d, FI, "solidez")
    d.p(f"La figura muestra que la mayoría de los tipos tiene evidencia sólida o moderada y que la dirección mixta es la "
        f"regla: incluso las intervenciones con respaldo sólido tienen estudios con resultados nulos. Los "
        f"{C['rev_limitada']} tipos con evidencia limitada son los de menor respaldo para la nueva política.")
    filas = []
    for t, gg in me[me.incluida].groupby("Tipo", sort=True):
        dis = gg["Diseño del estudio"].value_counts()
        clave = EV.estudios_clave(me[me.Tipo == t])
        otros = len(gg) - dis.get("Revisión sistemática", 0) - dis.get("Experimental", 0) - dis.get("Cuasi experimental", 0)
        filas.append([f"{t} {mt.loc[t, 'Tipo de intervención']}", str(dis.get("Revisión sistemática", 0)),
                      str(dis.get("Experimental", 0)), str(dis.get("Cuasi experimental", 0)), str(otros),
                      "; ".join(EV.cita(f) for f in clave.Fuente)])
    d.p("La tabla siguiente resume, por tipo de intervención, la composición de la evidencia incluida y sus estudios "
        "principales. La lista completa, con el resultado de cada estudio, está en el Anexo 4.")
    d.tabla(["Tipo de intervención", "Rev. sist.", "Exp.", "Cuasi.", "Otros", "Estudios principales"],
            filas, [2400, 650, 650, 700, 650, 3976], size=14, leyenda="Evidencia incluida por tipo de intervención",
            rastro="REV_EVI_matriz_evidencia.xlsx, hoja «Matriz de evidencia»",
            fuente="Fuente: estudios incluidos en la revisión de evidencia (Referencias). Elaboración propia.",
            nota="Una fila por fuente y tipo; una fuente puede servir a más de un tipo. Rev. sist.: revisiones "
                 "sistemáticas; Exp.: experimentales; Cuasi.: cuasiexperimentales; Otros: revisiones narrativas y "
                 "estudios descriptivos.", nombrar=False)
    inc = me[me.incluida]
    dis_t = inc["Diseño del estudio"].value_counts()
    rs_, ex_, cu_ = (int(dis_t.get(k, 0)) for k in ("Revisión sistemática", "Experimental", "Cuasi experimental"))
    n_t = inc.groupby("Tipo").size()
    t_max, t_min_sol = n_t.idxmax(), n_t[mt.loc[n_t.index, "Solidez"] == "Sólida"].idxmin()
    contraste = (f"El número de estudios no decide la solidez: el {t_max} "
                 f"({str(mt.loc[t_max, 'Tipo de intervención']).lower()}) reúne la mayor cantidad ({int(n_t[t_max])}) y "
                 f"su evidencia es {str(mt.loc[t_max, 'Solidez']).lower()}, mientras que el {t_min_sol} "
                 f"({str(mt.loc[t_min_sol, 'Tipo de intervención']).lower()}) tiene evidencia sólida con "
                 f"{int(n_t[t_min_sol])}. Lo que pesa es el diseño de los estudios y la calidad de su evidencia, según "
                 f"las reglas del numeral 2.4. " if mt.loc[t_max, "Solidez"] != "Sólida" else "")
    d.p(f"La tabla muestra de dónde viene el respaldo de cada tipo. De las {len(inc)} filas de evidencia incluidas, "
        f"{rs_} son revisiones sistemáticas, {ex_} evaluaciones experimentales y {cu_} cuasiexperimentales; las "
        f"{len(inc) - rs_ - ex_ - cu_} restantes son revisiones narrativas y estudios descriptivos. " + contraste +
        "Para la nueva política, los estudios principales de cada fila son el punto de partida para sustentar la "
        "lógica causal de los servicios que se mantengan.")
    peru = me[me.incluida & me["Ámbito"].astype(str).str.contains("Perú")]
    d.p(f"La evidencia peruana es escasa: {peru.Id.nunique()} fuentes del Perú, la mayoría cuasiexperimentales. La tabla "
        f"siguiente las presenta con su resultado principal.")
    d.tabla(["Tipo", "Fuente", "Diseño", "Resultado principal", "Dirección"],
            [[r.Tipo, EV.cita(r.Fuente), str(r["Diseño del estudio"]), str(r["Resultado principal"]), str(r["Dirección"])]
             for _, r in peru.iterrows()],
            [600, 2000, 1100, 4326, 1000], size=14, leyenda="Evaluaciones y estudios del Perú incluidos",
            fuente="Fuente: estudios del Perú incluidos en la revisión de evidencia (Referencias). Elaboración propia.",
            rastro="REV_EVI_matriz_evidencia.xlsx, hojas «Matriz de evidencia» y «Referencias»", nombrar=False)
    tipos_peru = lista_es(sorted({str(mt.loc[t, "Tipo de intervención"]).lower() for t in peru.Tipo}))
    d.p(f"Los estudios peruanos cubren pocos tipos de intervención ({tipos_peru}) y casi todos son evaluaciones "
        f"cuasiexperimentales; para los demás tipos, la aplicabilidad al Perú depende de estudios de otros países.")
    d.p("La evidencia se presenta a continuación por tipo de intervención, ordenada por objetivo prioritario. Para "
        "cada tipo se describe qué es, qué servicios lo implementan y cómo se espera que actúe sobre el modelo del "
        "problema; luego, qué encontraron los estudios, de los de mayor peso a los del Perú, y, por último, el "
        "respaldo del conjunto y su implicancia para los servicios. Los resultados son los que registró la revisión, "
        "tomados del resumen de cada estudio.")
    op_de = {}
    for t, x in tip.iterrows():
        op_de.setdefault(int(str(x.servicios).split(".")[0]), []).append(t)
    for opn in sorted(op_de):
        d.h3(f"Objetivo prioritario {opn}")
        for t in op_de[opn]:
            p1, p2, p3 = EV.parrafos_tipo(t, tip, mt, me, nombre_serv, None)
            d._sin(d.p, (f"{t}. {mt.loc[t, 'Tipo de intervención']}. ", {"b": True}), p1)
            d._sin(d.p, p2)
            d._sin(d.p, p3)
    # los párrafos por tipo nombran todos los servicios: las menciones posteriores ya no repiten el nombre
    d.N.vistos |= {("serv", s) for s in nombre_serv}
    d.p("La revisión es rápida y no sistemática: no evalúa el riesgo de sesgo de cada estudio y su aplicabilidad al "
        "Perú se valora caso por caso. Por eso, la solidez indica cuánto respaldo tiene la lógica causal de cada tipo "
        "de servicio, no el efecto esperado de los servicios de la PNIG.")


def referencias(d, E, refx):
    """Referencias en APA 7: las del informe y las 172 fuentes de la revisión de evidencia, en orden alfabético."""
    def enlace(doc):
        return E.loc["docs/fuentes/pnig/" + doc, "enlace_archivo"]

    items = []
    for texto, url in R.REFERENCIAS:
        if url == "EVAL_2020":
            url = EVAL_2020
        elif url == "LIO_COLECCION":
            url = LIO_COLECCION
        elif url and not url.startswith("http"):
            url = enlace(url)
        items.append((texto, url))
    for _, r in refx.iterrows():
        url = r.Enlace if isinstance(r.Enlace, str) and r.Enlace.startswith("http") else None
        items.append((re.sub(r"\s+", " ", str(r.Referencia)).strip(), url))
    clave = lambda t: re.sub(r"[^a-z0-9 ]", "", t[0].lower().translate(str.maketrans("áéíóúüñ", "aeiouun")))
    for texto, url in sorted(items, key=clave):
        piezas = [texto + (" " if url else "")] + ([(url, {"url": url})] if url else [])
        d._sin(d.p, *piezas, jc="left", size=20)


def anexos(d, C, H, P, F, E, nombre_serv):
    d.N.activo = False
    # Anexo 1. Matriz de evaluación
    d.h1("Anexo 1. Matriz de evaluación")
    cob = pd.read_excel(f"{EVAL}/04_analisis/matriz_triangulacion_preguntas.xlsx",
                        sheet_name="Cobertura", dtype=str).set_index("Pregunta")
    d.p("La matriz vincula cada pregunta de evaluación con su criterio, los instrumentos previstos, los instrumentos "
        "que efectivamente aportaron evidencia y los hallazgos que la responden.")
    d.tabla(["Pregunta", "Criterio", "Instrumentos previstos", "Instrumentos con evidencia", "Hallazgos"],
            [[p, crit, ", ".join(S.CORTOS[i] for i in prev),
              ", ".join(S.CORTOS[i.strip()] for i in cob.loc[p, "Instrumentos con evidencia en la matriz"].split(",")),
              ", ".join(h["id"] for h in H.values() if h["pregunta"] == p)]
             for p, (crit, _, prev) in S.PREGUNTAS.items()],
            [900, 2000, 1900, 2226, 2000], size=16, fuente="Fuente: elaboración propia.",
            rastro="matriz_triangulacion_preguntas.xlsx, hojas «Cobertura» y «Matriz de triangulación»")

    # Anexo 2. Hallazgos y nivel de confianza
    d.h1("Anexo 2. Hallazgos y nivel de confianza")
    d.p("La tabla presenta cada hallazgo con su detalle, su nivel de confianza y su respaldo documental. El número de "
        "fuentes independientes cuenta los documentos, bases o informantes distintos que lo sostienen; no cuenta la "
        "entrevista a la OGMEPGD (numeral 2.6), y en la revisión de evidencia cuenta cada estudio incluido. El respaldo "
        "nombra cada documento con su página y cada libro de análisis con su hoja; los libros se entregan con el anexo "
        "técnico de la evaluación.")
    d.p(*leyenda_confianza(), size=18)
    orden = {p: i for i, p in enumerate(S.PREGUNTAS)}
    hs = sorted(H.values(), key=lambda h: (orden[h["pregunta"]], int(h["id"].split("H")[1])))
    filas, rastros = [], []
    for h in hs:
        resp = S.respaldo(h) + (f". Salvedad: {h['nota']}" if h["nota"] else "")
        filas.append([h["id"], h["pregunta"], h["detalle"], marcador(h["confianza"]),
                      str(S.fuentes_independientes(h)), resp])
        rastros.append({"hallazgo": h["id"], "pregunta": h["pregunta"], "confianza": h["confianza"],
                        "fuentes_independientes": S.fuentes_independientes(h), "rastro": S.rastro(h),
                        "respaldo": S.respaldo(h), "nota": h["nota"]})
    d.tabla(["Hallazgo", "Pregunta", "Detalle", "Confianza", "N.° de fuentes independientes", "Respaldo"], filas,
            [850, 900, 2600, 1100, 1450, 2126], size=15, fuente="Fuente: elaboración propia.", nombrar=False)
    pd.DataFrame(rastros).to_csv(RASTRO_HALLAZGOS, index=False, encoding="utf-8-sig")

    # Anexo 3. Estado de implementación por servicio
    d.h1("Anexo 3. Estado de implementación por servicio")
    d.p("La tabla presenta cada servicio con su entidad responsable, los tres componentes que intervienen en su estado "
        "(protocolo aprobado, recursos presupuestales identificados y reporte de indicadores) y la justificación del "
        "estado en todo el periodo 2019–2025. Los recursos se identifican cotejando los montos que la entidad reportó "
        "en los informes LIO de 2019 a 2025 o en el formulario de seguimiento 2023 con las metas, actividades y "
        "productos de los datos abiertos del MEF (MEF, 2026).")
    g = C["grd"].copy()
    g["Código"] = g["Código"].astype(str)
    mon = P["montos"]
    filas = []
    for _, r in g.iterrows():
        c = r["Código"]
        prot = r["Protocolo aprobado"] == "Sí"
        rec = r["Asignación presupuestal identificable"] == "Sí"
        anios_v = str(r["Años con valor reportado"]) if pd.notna(r["Años con valor reportado"]) else ""
        reporta = bool(anios_v)
        partes = [f"Protocolo aprobado: {str(r['Norma del protocolo']).replace(' | ', '; ')}." if prot
                  else "Sin protocolo aprobado."]
        m = mon[mon.Servicio == c].sort_values("Año", ascending=False)
        if rec and len(m):
            x = m.iloc[0]
            pag = f", p. {int(x['Página'])}" if pd.notna(x["Página"]) else ""
            fuente_m = str(x.Fuente).replace("Informe LIO", "informe LIO").replace("Formulario de seguimiento",
                                                                                   "formulario de seguimiento")
            estructura = str(x["Descripción MEF"]).split(" | ")[-1].strip()
            estructura = estructura[0] + estructura[1:].lower()
            partes.append(f"Se identificó un monto de {monto(x['Monto reportado'])} reportado en el "
                          f"{fuente_m}{pag}, que coincide con el {str(x['Concepto MEF'])} {int(x['Año'])} de la "
                          f"{str(x['Nivel']).split(' (')[0].lower()} «{estructura}» en los datos del MEF"
                          + (f"; otros {len(m) - 1} montos también coinciden." if len(m) > 1 else "."))
        elif rec:
            partes.append("Recursos identificados en los datos del MEF.")
        else:
            partes.append("No se identificó un monto que coincida con una estructura presupuestal del MEF.")
        partes.append(f"Reporta indicadores con valor en {anios_v}." if reporta else
                      ("Sus indicadores no tienen valor reportado." if prot else "No tiene indicador aprobado."))
        if pd.notna(r["Ajuste por lectura cualitativa"]):
            partes.append(str(r["Ajuste por lectura cualitativa"]).rstrip(".") + ".")
        elif r["Estado asignado (preliminar)"] == "Implementado parcialmente" and pd.notna(r["Componentes faltantes"]):
            partes.append(f"Componentes que le faltan: {r['Componentes faltantes'].replace('asignación identificable', 'recursos presupuestales identificados')}.")
        filas.append([c, [(nombre_serv[c], {"i": True})], R.entidad_completa(r["Entidad u órgano responsable"]),
                      "Sí" if prot else "No", "Sí" if rec else "No", "Sí" if reporta else "No",
                      [(r["Estado asignado (preliminar)"] + ". ", {"b": True}), (" ".join(partes), {})]])
    d.tabla(["Código", "Servicio", "Entidad responsable", "Protocolo", "Recursos presupuestales identificados",
             "Reporta indicadores", "Estado y justificación"], filas,
            [700, 1600, 1150, 960, 1450, 1250, 1916], size=15,
            fuente="Fuente: protocolos de los servicios (Anexo 10); MEF (2026); OGMEPGD (2026); MIMP (2020-2026). "
                   "Elaboración propia.",
            rastro="REV_DOC_estado_implementacion_servicios.xlsx, hoja «Estado por servicio»; "
                   "REV_IND_3_verificacion_presupuestal.xlsx, hoja «Cifras verificadas»")

    # Anexo 4. Evidencia por tipo de intervención
    tip, mt, me = P["tip"], P["mt"], P["me"]
    d.h1("Anexo 4. Evidencia por tipo de intervención")
    d.p("Para cada tipo de intervención, el anexo presenta los servicios de la PNIG que lo implementan, sus entidades "
        "responsables, la clasificación de su evidencia con el motivo que la explica y todos los estudios incluidos, "
        "con su resultado principal. La regla de solidez es la del numeral 2.4: sólida, con al menos una revisión "
        "sistemática favorable cuya evidencia no sea de calidad baja, o con dos evaluaciones experimentales o "
        "cuasiexperimentales, todas favorables; moderada, con evaluaciones aisladas o revisiones con resultados mixtos "
        "o de calidad baja; limitada, con solo estudios descriptivos, revisiones narrativas o estudios de contextos poco "
        "comparables. Solo cuentan para la solidez moderada o sólida los estudios de contextos comparables: los de "
        "países de ingreso alto y las revisiones cuyo resumen no informa el ámbito cuentan solo para la limitada. Por "
        "eso un tipo con muchas revisiones puede quedar en «limitada» si todas son de contextos poco comparables. Las "
        "referencias completas están en Referencias.")
    gi = g.set_index("Código")
    for t, x in tip.iterrows():
        servs = [s.strip() for s in str(x.servicios).split(";")]
        ents = sorted({R.entidad_completa(gi.loc[s, "Entidad u órgano responsable"]) for s in servs})
        d.h3(f"{t}. {mt.loc[t, 'Tipo de intervención']}")
        d.p(("Servicios: ", {"b": True}), "; ".join(f"{s} ({nombre_serv[s]})" for s in servs) + ". ",
            ("Entidades responsables: ", {"b": True}), "; ".join(ents) + ". ",
            ("Clasificación: ", {"b": True}), EV.motivo(me[me.Tipo == t], mt.loc[t, "Solidez"]) +
            f" Dirección: {mt.loc[t, 'Dirección']}; pertinencia: {str(mt.loc[t, 'Pertinencia del resultado']).lower()}.",
            size=18)
        gg = me[(me.Tipo == t) & me.incluida]
        d.tabla(["Estudio", "Ámbito", "Diseño", "Resultado principal", "Dirección"],
                [[EV.cita(r.Fuente), str(r["Ámbito"]), str(r["Diseño del estudio"]), str(r["Resultado principal"]),
                  str(r["Dirección"])] for _, r in gg.iterrows()],
                [1800, 1400, 1100, 3826, 900], size=13, nombrar=False)
    d.fuente("Fuente: estudios incluidos en la revisión de evidencia (Referencias). Elaboración propia.")

    # Anexo 5. Reporte de seguimiento
    d.h1("Anexo 5. Reporte de seguimiento de la política nacional")
    rep = E[E.index.str.contains(r"04_seguimiento/20\d\d_mimp_reporte")]
    assert len(rep) == 13 and (rep.acceso == "público").all(), "reportes de seguimiento y cumplimiento sin verificar"
    d.p("El reporte de seguimiento del último año del periodo evaluado es el de 2025, que el MIMP publicó en su formato "
        "y en el del aplicativo del CEPLAN. Los reportes de seguimiento y de cumplimiento de 2020 a 2025 están "
        "publicados en gob.pe; la tabla da la página de publicación de cada uno.")
    filas = []
    for doc in sorted(rep.index, key=lambda x: (x.split("/")[-1][:4], "cumplimiento" in x, "aplicativo" in x)):
        anio = doc.split("/")[-1][:4]
        tipo = ("Reporte de cumplimiento" if "cumplimiento" in doc else "Reporte de seguimiento"
                + (" (aplicativo del CEPLAN)" if "aplicativo" in doc else " (formato del MIMP)" if anio == "2025" else ""))
        url = rep.loc[doc, "pagina_publicacion"]
        filas.append([anio, tipo, [(url, {"url": url})]])
    d.tabla(["Año", "Reporte", "Página de publicación en gob.pe"], filas, [700, 2700, 5626], size=16,
            leyenda="Reportes de seguimiento y de cumplimiento de la PNIG, 2020–2025",
            fuente="Fuente: gob.pe, MIMP (consulta del 6 de noviembre de 2026).",
            rastro="docs/fuentes/pnig/04_seguimiento/; docs/fuentes/enlaces_publicos_2026-11-06.tsv")

    # Anexo 6. Brechas por departamento y series por encuesta
    d.h1("Anexo 6. Brechas por departamento y series por encuesta")
    ok, tot = C["dep_validacion"]
    d.p(f"Las brechas se estimaron por departamento donde el diseño muestral lo permite: la ENAHO, la ENDES, la EPEN y la "
        f"ENA declaran inferencia departamental en sus fichas técnicas (INEI, 2025b), y la ENARES solo en 2024; la ENUT "
        f"2024 no la tiene. Los departamentos son los 26 ámbitos que publica el INEI: Lima se divide en Lima "
        f"Metropolitana y Región Lima, y el Callao se cuenta aparte; la EPEN solo distingue Lima Metropolitana con el "
        f"Callao y Región Lima, y la ENA publica Lima entera. Las estimaciones reproducen {ok} de {tot} cifras por "
        f"departamento que el INEI publica para 2024 (INEI, 2025a); las que no coinciden son cinco valores de hombres "
        f"del gráfico de uso de internet (Arequipa, Tumbes, La Libertad, Junín y Amazonas, con diferencias de 3,4 a 9,4 "
        f"puntos). Con los mismos microdatos de la ENAHO 2024 y la misma definición, el valor nacional y las cifras de "
        f"mujeres coinciden, por lo que la diferencia no se explica y no se declara una validación integral de la base. Se mapean todos los indicadores estimables: "
        f"las estimaciones referenciales llevan trama y la nota de cada mapa las lista con su coeficiente de variación. "
        f"La cartografía es la capa provincial de Peru maps (2026), con los límites del INEI, de uso solo estadístico.")
    d.p("Las series por encuesta muestran la evolución nacional de los indicadores que no van en el cuerpo del informe. "
        "La primera reúne las brechas que miden la ENA, la EPEN, la ENARES y la ENUT, con mujeres y hombres en cada "
        "año y su intervalo de confianza.")
    figura_brecha(d, F, "series_otras", alt=f"Series nacionales: {F.loc['series_otras', 'titulo'].lower()}.")
    d.p("Las brechas de estas encuestas se mantienen o cambian poco en el periodo: las líneas de mujeres y hombres se "
        "mueven en paralelo y la distancia entre ellas apenas varía. La segunda serie presenta los indicadores de la "
        "ENDES, que solo encuesta a mujeres, por lo que muestra niveles y no brechas.")
    figura_brecha(d, F, "series_endes", alt=f"Series nacionales: {F.loc['series_endes', 'titulo'].lower()}.")
    d.p("Los indicadores de salud sexual y reproductiva cambian poco entre 2019 y 2025, y la violencia de pareja baja, "
        "con la advertencia de que 2020 se levantó en buena parte por teléfono. Los mapas siguientes llevan estos "
        "indicadores y los de la ENAHO al territorio.")
    dr = C["dep_resumen"].set_index("Clave")
    grupos = [
        ("Mercado laboral y protección social", ("mapa_tasa_desempleo", "mapa_tfnr", "mapa_trabajo_hogar",
                                                 "mapa_afiliado_pension")),
        ("Ingresos", ("mapa_ingreso_laboral", "mapa_ingreso_publico")),
        ("Educación y acceso a internet", ("mapa_analfabetismo", "mapa_internet", "mapa_nini_15_29")),
        ("Producción agropecuaria", ("mapa_productora_mujer", "mapa_credito_solicito", "mapa_propietario",
                                     "mapa_titulo")),
        ("Salud sexual y reproductiva y autonomía", ("mapa_uso_anticonceptivo", "mapa_uso_moderno",
                                                    "mapa_demanda_insatisfecha", "mapa_embarazo_adolescente",
                                                    "mapa_esposo_decide_salud")),
        ("Tolerancia social a la violencia", ("mapa_tolerancia_10",)),
        ("Participación política", ("mapa_participacion_regidurias", "mapa_participacion_congreso")),
    ]
    vistos = set(MAPAS_CUERPO + MAPAS_D2)
    for nombre_g, claves in grupos:
        claves = [k for k in claves if k in F.index]
        if not claves:
            continue
        d.h3(nombre_g)
        frases = []
        for k in claves:
            ind = k.replace("mapa_", "")
            if ind in dr.index:
                x = dr.loc[ind]
                conf = f"{int(x['Confiables (último año)'])} de {int(x['Ámbitos con estimación (último año)'])}"
                des = x.get("Brecha significativa en desventaja de las mujeres")
                frases.append(f"{str(x.Indicador).lower()}: estimación confiable en {conf} departamentos"
                              + (f" y brecha significativa en desventaja de las mujeres en {int(des)}" if pd.notna(des)
                                 else ""))
        if frases:
            d.p("Los mapas de este grupo comparan el primer y el último año disponibles en cada departamento; la trama "
                "marca las estimaciones referenciales.")
        else:
            d.p("Los mapas de este grupo muestran el porcentaje de mujeres electas por departamento en cada periodo; son "
                "registros completos, sin error muestral, pero con pocos cargos en algunos departamentos.")
        for k in claves:
            titulo_k = str(F.loc[k, "titulo"])
            d.p(f"El mapa siguiente muestra {titulo_k[0].lower() + titulo_k[1:]}.")
            figura_brecha(d, F, k, ancho={"mapa_tolerancia_10": 9.0, "mapa_participacion_regidurias": 12.0}.get(k, 16.0),
                          alt=f"Mapas del Perú por departamento: {titulo_k.lower()}.")
            ind = k.replace("mapa_", "")
            if ind in dr.index:
                d.p(frase_mapa(dr.loc[ind]))
            else:
                d.p("La proporción de mujeres electas aumenta entre periodos en casi todos los departamentos, pero "
                    "ninguno se acerca a la paridad.")
            vistos.add(k)
    for k in F.index:
        if k.startswith("mapa_") and k not in vistos:
            figura_brecha(d, F, k, alt=f"Mapas del Perú por departamento: {F.loc[k, 'titulo'].lower()}.")

    # Anexo 7. Seguimiento de las recomendaciones
    d.h1("Anexo 7. Seguimiento de las recomendaciones de evaluaciones anteriores")
    rec = C["rec"]
    d.p("La tabla presenta cada recomendación de las cuatro evaluaciones anteriores con el inicio de su texto literal, "
        "su estado al corte, cómo se ejecutó y su medio de verificación. Las recomendaciones se agrupan por evaluación y "
        "se numeran en el orden en que aparecen en el informe de origen.")

    def inicio(texto, k=28):
        palabras = str(texto).split()
        return " ".join(palabras[:k]) + ("…" if len(palabras) > k else "")

    nombre_ev = {"MIMP 2020": "Evaluación de implementación 2020 (MIMP, 2021)",
                 "MIMP 2021": "Evaluación de resultados 2021 (MIMP, 2022)", "MIMP 2023": "Evaluación 2023 (MIMP, 2024)",
                 "CEPLAN": "Informe técnico del CEPLAN sobre los servicios del OP1 (CEPLAN, s. f.)"}
    filas, grupos_idx = [], []
    for ev, gg in rec.groupby("Evaluación", sort=False):
        grupos_idx.append(len(filas))
        filas.append(["", nombre_ev.get(ev, ev), "", ""])
        for _, r in gg.iterrows():
            nro = str(int(re.search(r"-(\d+)$", str(r["N.°"])).group(1)))
            servicio = (f"; servicio {r['Servicio']}" if pd.notna(r["Servicio"]) and str(r["Servicio"]).strip() else "")
            filas.append([nro,
                          [(f"«{inicio(r['Recomendación (texto literal)'])}»", {}),
                           (f"\n{r['Cita']}, p. {r['Página (PDF)']}{servicio}", {"i": True, "color": "595959"})],
                          r["Estado al corte (p3.4)"],
                          [(str(r["Cómo se ejecutó (p3.5)"]), {}), (f"\n{r['Medio de verificación']}",
                                                                    {"color": "595959"})]])
    d.tabla(["N.°", "Recomendación (inicio del texto literal) y fuente", "Estado",
             "Cómo se ejecutó y medio de verificación"], filas, [600, 3676, 1250, 3500], size=15,
            grupos=tuple(grupos_idx),
            fuente="Fuente: MIMP (2021, 2022, 2024); CEPLAN (s. f.); documentos al 7 de noviembre de 2026. Elaboración "
                   "propia.",
            nota="Páginas del PDF de cada informe. El estado y la forma de ejecución responden a las preguntas p3.4 y "
                 "p3.5 de la Tabla A7.1 de CEPLAN (2024a). Estado: implementada, si la acción recomendada consta en un "
                 "documento posterior al informe; parcialmente implementada, si consta una parte; no implementada, si "
                 "consta que el elemento recomendado no cambió; sin evidencia, si no consta medio de verificación (no se "
                 "presume), y sin objeto, si la Ley N.° 32535 deja sin objeto la recomendación.",
            rastro="seguimiento_recomendaciones_previas.xlsx, hoja «Matriz» (columna «Rastro del medio de "
                   "verificación»)", nombrar=False)

    # Anexo 8. Entrevistas
    d.h1("Anexo 8. Aplicación de las entrevistas al nivel conductor")
    A = pd.read_csv(APLICACION, dtype=str).fillna("")
    Q = pd.read_csv(f"{GES}/ENTREV_preguntas.csv", dtype=str).fillna("").set_index("id")
    nombre = {"OGPPM": "OGPPM", "OGMEPGD": "OGMEPGD (OMEP)", "DGIGND": "DGIGND (DPIGND)"}
    filas = []
    for inf in ("OGPPM", "OGMEPGD", "DGIGND"):
        a = A[A.informante == inf]
        tipo = ["Surgida" if i.startswith("PS-") else "Orientadora" if Q.loc[i, "tipo"] in ORIENTADORAS else
                "Repregunta" for i in a.id_pregunta]
        ap = APLICADA[inf]
        filas.append([nombre[inf], f"{ap['fecha'][8:10]}/{ap['fecha'][5:7]}/{ap['fecha'][:4]}, {ap['inicio']} h",
                      ap["modalidad"].replace("entrevistadores del equipo evaluador", "entrevistadores de la OMEP"),
                      ap["duracion"][2:] if ap["duracion"].startswith("0:") else ap["duracion"],
                      str(tipo.count("Orientadora")), str(tipo.count("Repregunta")), str(tipo.count("Surgida")),
                      str(int((C["entrev"].informante == inf).sum())), str(C["entrev_val_inf"].get(inf, 0))])
    d.p("La tabla presenta cada entrevista con su fecha, su modalidad, su duración y el número de preguntas que se "
        "formularon, por tipo, y de segmentos codificados y validados.")
    d.tabla(["Órgano", "Fecha", "Modalidad", "Duración", "Orient.", "Repreg.", "Surg.", "Segm.", "Valid."],
            filas, [1150, 1100, 2076, 900, 800, 800, 650, 800, 750], size=15, leyenda="Entrevistas al nivel conductor",
            rastro="05_entrevistas/ENTREV_aplicacion.csv; 05_entrevistas/ENTREV_registro_codificacion.xlsx",
            fuente="Fuente: registro de aplicación y de codificación de las entrevistas. Elaboración propia.",
            nota="Solo se cuentan las preguntas que se formularon (literal, reformulada o en parte). Orient.: "
                 "orientadoras; Repreg.: repreguntas; Surg.: surgidas en la entrevista; Segm.: segmentos codificados; "
                 "Valid.: segmentos revisados y validados.", nombrar=False)
    d.p("De las respuestas se dedujo el contenido de dos temas que no se preguntaron completos:", after=60)
    for _, x in C["entrev_ded"].iterrows():
        organos = sorted({o for o in re.findall(r"ENTREV-(\w+)-\d+", str(x.segmentos))})
        d.vineta((f"{x.tema}. ", {"b": True}), x.deduccion,
                 (f" Respaldo: {lista_es([f'entrevista a la {o}' for o in organos])}; hallazgo {x.hallazgo}.",
                  {"color": "595959"}))
    d.p(f"Hallazgos cuyo contraste estaba previsto y que no se contrastaron con el nivel conductor: "
        f"{lista_es(C['entrev_sin_contraste'])}. Las transcripciones anonimizadas se entregan con el anexo técnico; las "
        f"grabaciones y las versiones con nombres no se entregan, conforme a la Ley N.° 29733.")

    # Anexo 9. Catálogo de servicios e indicadores
    d.h1("Anexo 9. Catálogo de servicios e indicadores")
    d.p("La primera tabla presenta los servicios de la PNIG por objetivo prioritario y lineamiento, con su entidad "
        "responsable; la segunda, los indicadores de los objetivos prioritarios y de los servicios vigentes al cierre.")
    filas, grupos_idx, op_prev = [], [], None
    obj = P["obj"]
    for _, r in g.iterrows():
        opc = str(r["OP"]).replace("OP.0", "OP")
        if opc != op_prev:
            grupos_idx.append(len(filas))
            filas.append([opc, str(obj.loc[opc, "Objetivo prioritario"]) if opc in obj.index else "", "", ""])
            op_prev = opc
        filas.append([r["Código"], [(nombre_serv[r["Código"]], {"i": True})], str(r["Lineamiento"]),
                      R.entidad_completa(r["Entidad u órgano responsable"])])
    d.tabla(["Código", "Servicio", "Lineamiento", "Entidad responsable"], filas, [800, 3400, 2626, 2200], size=13,
            leyenda="Servicios de la PNIG", grupos=tuple(grupos_idx),
            fuente=f"Fuente: {PNIG}; protocolos de los servicios (Anexo 10). Elaboración propia.", nombrar=False)
    b = P["base_ind"]

    def entidad_ind(e):
        partes = []
        for e0 in re.split(r"\s*/\s*", re.sub(r"\s*-\s*PCM$", "", str(e)).strip()):
            try:
                partes.append(R.entidad_completa(e0))
            except AssertionError:
                partes.append(e0)
        return "; ".join(partes)

    filas = [[r["Código"], "OP" if r["Nivel"] == "OP" else "Servicio", str(r["Indicador"]),
              str(r["OP"]).replace("OP.0", "OP"), entidad_ind(r["Entidad"])] for _, r in b.iterrows()]
    d.tabla(["Código", "Nivel", "Indicador", "OP", "Entidad responsable"], filas, [1500, 900, 3726, 700, 2200], size=13,
            leyenda="Indicadores de la PNIG vigentes al cierre", fuente="Fuente: OGMEPGD (2026). Elaboración propia.",
            nombrar=False)

    # Anexo 10. Documentos de respaldo
    d.h1("Anexo 10. Documentos de respaldo")
    d.p("La tabla lista los documentos del corpus de la evaluación que respaldan sus hallazgos, con su norma o versión, "
        "su fecha y su enlace de publicación cuando es público. Los que no tienen enlace público pueden solicitarse a la "
        "OGMEPGD del MIMP.")
    # los documentos sobre la propia nota metodológica no respaldan hallazgos y no se listan
    docs = [x for x in R.documentos_respaldo(REGISTRO, INVENTARIO, ENLACES)
            if not S.INTERNOS.search(x["documento"] + " " + x["norma"])]
    filas, grupos_idx, sec_prev = [], [], None
    for x in docs:
        if x["seccion"] != sec_prev:
            grupos_idx.append(len(filas))
            filas.append([re.sub(r"^\d+[a-z]?\.\s*", "", x["seccion"]), "", "", ""])
            sec_prev = x["seccion"]
        filas.append([x["documento"], x["norma"], x["fecha"], [(x["url"], {"url": x["url"]})] if x["url"] else x["via"]])
    d.tabla(["Documento", "Norma o versión", "Fecha", "Enlace o vía de consulta"], filas, [3400, 1700, 1000, 2926],
            size=12, grupos=tuple(grupos_idx), fuente="Fuente: elaboración propia.", nombrar=False)
    d.p("Además, el anexo técnico entrega los libros de análisis que cita el Anexo 2, con estos nombres: "
        + lista_es(sorted(set(S.LIBROS_PUBLICOS.values()))) + ".", size=18)

    # Anexo 11. Glosario
    d.h1("Anexo 11. Glosario")
    d.tabla(["Término", "Definición"], [[t, df] for t, df in sorted(R.GLOSARIO, key=lambda x: x[0].lower())],
            [2400, 6626], size=16, nombrar=False)


def paginas_pdf(docx_path):
    """Convierte el documento a PDF con LibreOffice y devuelve el texto de cada página."""
    tmp = tempfile.mkdtemp(prefix="informe_pdf_")
    # llamada directa a soffice.exe (sin PowerShell, que pierde las comillas de la ruta con espacios)
    subprocess.run([SOFFICE, "--headless", "--convert-to", "pdf", "--outdir", tmp, os.path.abspath(docx_path)],
                   capture_output=True, timeout=900, check=True)
    pdf = os.path.join(tmp, os.path.basename(docx_path).replace(".docx", ".pdf"))
    # soffice.exe puede volver antes de que termine la conversión: se espera a que el PDF exista y deje de crecer
    import time
    previo, estable = -1, 0
    for _ in range(240):
        tam = os.path.getsize(pdf) if os.path.exists(pdf) else -1
        estable = estable + 1 if tam == previo and tam > 0 else 0
        if estable >= 3:
            break
        previo = tam
        time.sleep(1)
    assert os.path.exists(pdf), "LibreOffice no generó el PDF del informe"
    texto = subprocess.run(["pdftotext", "-enc", "UTF-8", "-layout", pdf, "-"], capture_output=True, text=True,
                           encoding="utf-8").stdout
    return [re.sub(r"\s+", " ", p) for p in texto.split("\f")]


def ubicar(paginas, titulos, leyendas):
    """Página de cada título y leyenda: la primera aparición después de los índices (las páginas de los índices
    contienen los mismos textos)."""
    norm = lambda t: re.sub(r"\s+", " ", t).strip()[:55]
    inicio = 0
    if leyendas:
        rot, numero, t = leyendas[-1]
        marca = norm(f"{rot} {numero}. {t}")
        inicio = next((i for i, p in enumerate(paginas) if marca in p), 0) + 1
    salida, k = {}, inicio
    for t, _ in titulos:
        m = norm(t)
        j = next((i for i in range(k, len(paginas)) if m in paginas[i]), None)
        if j is None:  # títulos anteriores a los índices, como la ficha técnica
            j = next((i for i in range(0, inicio) if m in paginas[i]), None)
            if j is not None:
                salida[("t", t)] = j + 1
            continue
        if j is not None:
            salida[("t", t)] = j + 1
            k = j
    for rot, numero, t in leyendas:
        m = norm(f"{rot} {numero}. {t}")
        j = next((i for i in range(inicio, len(paginas)) if m in paginas[i]), None)
        if j is not None:
            salida[(rot, numero)] = j + 1
    return salida


def barrido(docx_path):
    """Términos internos que no pueden quedar en el informe publicable; los nombres .xlsx solo en los Anexos 2 y 10."""
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    doc = Document(docx_path)
    extra = re.compile(r"\brepositorio de la evaluación\b|\bhash\b|Elaboración: equipo evaluador|Tabla 9 de la nota")
    seccion, malos = "", []
    for el in doc.element.body.iterchildren():
        tag = el.tag.split("}")[1]
        textos = []
        if tag == "p":
            p = Paragraph(el, doc)
            if p.style is not None and p.style.name.startswith("Heading"):
                seccion = p.text
            textos = [p.text]
        elif tag == "tbl":
            textos = [c.text for r in Table(el, doc).rows for c in r.cells]
        for t in textos:
            for rx in (S.INTERNOS, extra):
                for m in rx.finditer(t):
                    g0 = m.group(0).lower()
                    if g0 == ".xlsx" and seccion.startswith(("Anexo 2", "Anexo 10")):
                        continue
                    if re.match(r"libros? de análisis", g0) and seccion.startswith(("2.9", "Anexo 2", "Anexo 10")):
                        continue
                    malos.append(f"[{seccion[:28]}] «{t[max(0, m.start() - 50):m.end() + 30]}»")
    for parte in doc.part.package.parts:
        if str(parte.partname).endswith("footnotes.xml"):
            txt = re.sub(r"<[^>]+>", "", parte.blob.decode("utf-8"))
            for m in S.INTERNOS.finditer(txt):
                malos.append(f"[notas al pie] «{txt[max(0, m.start() - 50):m.end() + 30]}»")
    return malos


def revisar_formato(docx_path):
    """Comprobaciones de formato del archivo final: sin updateFields, letra mínima de 7,5 pt fuera de la portada,
    leyendas con estilo de descripción e índices con números de página distintos."""
    import zipfile
    problemas = []
    with zipfile.ZipFile(docx_path) as z:
        settings = z.read("word/settings.xml").decode("utf-8")
        doc = z.read("word/document.xml").decode("utf-8")
        estilos = z.read("word/styles.xml").decode("utf-8")
    if "updateFields" in settings:
        problemas.append("settings.xml conserva updateFields")
    cuerpo = doc[doc.find("Ficha técnica"):]
    chicos = [int(x) for x in re.findall(r'<w:sz w:val="(\d+)"/>', cuerpo) if int(x) < 15]
    if chicos:
        problemas.append(f"{len(chicos)} tramos con letra menor de 7,5 pt")
    sid = re.search(r'w:styleId="([^"]+)"[^>]*>\s*<w:name w:val="caption"/>', estilos)
    n_ley = len(re.findall(r'<w:pStyle w:val="%s"/>' % sid.group(1), doc)) if sid else 0
    if n_ley < len(LEYENDAS):
        problemas.append(f"solo {n_ley} de {len(LEYENDAS)} leyendas con estilo de descripción")
    paginas = re.findall(r'<w:pStyle w:val="(?:TDC1|TDC2|TOC1|TOC2)"/>.*?</w:p>', doc)
    nums = [re.findall(r"<w:t[^>]*>([^<]*)</w:t>", x)[-1] for x in paginas if "<w:t" in x]
    if len(set(nums)) < max(3, len(nums) // 3):
        problemas.append(f"el índice general tiene pocos números de página distintos: {sorted(set(nums))[:8]}")
    return problemas


if __name__ == "__main__":
    os.makedirs(os.path.dirname(SALIDA), exist_ok=True)
    # primera vuelta: títulos y leyendas; segunda: índices completos; luego Word actualiza campos y páginas
    d, H, RR = construir()
    TITULOS[:] = d.titulos
    LEYENDAS[:] = d.leyendas
    d, H, RR = construir()
    assert d.titulos == TITULOS and d.leyendas == LEYENDAS, "la segunda vuelta cambió los títulos o las leyendas"
    d.guardar(SALIDA, comillas_inglesas=True)
    rastros = d.rastros
    ok, salida_word = actualizar_con_word(SALIDA, PDF_WORD)
    if ok:
        print("Fuentes incrustadas:", ", ".join(quitar_fuentes_de_office(SALIDA)) or "ninguna")
    if ok:
        print("Word:", salida_word.splitlines()[-1] if salida_word else "", "· PDF:", PDF_WORD)
    else:
        # respaldo sin Word: páginas leídas del PDF de LibreOffice, escritas en caché
        print("Word no respondió; se usan las páginas de LibreOffice.", salida_word[-300:])
        pags = ubicar(paginas_pdf(SALIDA), TITULOS, LEYENDAS)
        d, H, RR = construir(pags)
        d.guardar(SALIDA, comillas_inglesas=True)
    pd.DataFrame(rastros).to_csv(RASTRO, index=False, encoding="utf-8-sig")
    # número, leyenda y archivo de cada figura, para copiarlas numeradas al anexo técnico (ronda 36)
    pd.DataFrame(d.figuras).to_csv(os.path.join(os.path.dirname(SALIDA), "indice_figuras.csv"), index=False,
                                   encoding="utf-8-sig")
    usados = sorted({(lib, hoja) for x in rastros for lib, hojas in
                     re.findall(r"([\w\-]+\.xlsx)((?:, hojas? (?:«[^»]+»(?:, | y )?)+)?)", x["rastro"])
                     for hoja in (re.findall(r"«([^»]+)»", hojas) or [""])})
    pd.DataFrame([{"libro_repositorio": lib, "libro_publicado": S.LIBROS_PUBLICOS.get(lib, lib),
                   "hoja_repositorio": hoja, "hoja_publicada": S.hoja_publica(hoja) if hoja else ""}
                  for lib, hoja in usados]).to_csv(CORRESPONDENCIA, index=False, encoding="utf-8-sig")
    malos = barrido(SALIDA)
    formato = revisar_formato(SALIDA)
    print("Informe generado:", SALIDA, "con", len(RR), "recomendaciones;", len(malos), "términos internos;",
          len(formato), "problemas de formato")
    for m in malos[:40] + formato:
        print("  ", m)
