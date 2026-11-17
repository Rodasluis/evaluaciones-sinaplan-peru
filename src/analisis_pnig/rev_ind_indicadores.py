"""REV_IND.1: base de indicadores de la PNIG con avance anual 2019–2025 y semaforización.

Fuentes:
  - consolidado de indicadores de la OMEP: metas de la ficha vigente y valores de la serie del reporte de seguimiento
    2025 (RS 2025);
  - reportes de seguimiento (RS) 2021–2025: tablas de logros esperados, valores obtenidos y sentido, para verificar las
    metas usadas (los RS pasan por la revisión de la dirección conductora);
  - reportes de cumplimiento (RC) 2022–2025: línea «IOP.xx.xx. Meta AAAA: … // Valor AAAA: … // Cumplimiento: …»;
  - panel del formulario de seguimiento 2023–2025;
  - protocolos y actas de validación ratificadas, para verificar las metas donde las fuentes difieren.

Cálculo:
  PNIG: razón sin acotar (ascendente VO/LE × 100; descendente LE/VO × 100) con la semaforización adaptada del RC 2025
        (Anexo 01, p. 16), que es la que fija la nota metodológica (Tabla 9). Es la única escala de la evaluación.
  Equipo evaluador:
    - sentido aplicado: el vigente del consolidado, que es el que siguen las metas. La hoja «Sentido» valora si es
      coherente con lo que mide el indicador y registra las correcciones que propusieron las actas de 2022 no
      ratificadas (hallazgos para D5);
    - metas: las de la ficha vigente. Donde el consolidado, el formulario o los RS difieren se verifican en el protocolo
      o en el acta ratificada: si el consolidado está errado prevalece METAS_VERIFICADAS, si lo está el formulario se
      anota en NOTAS_FORMULARIO y si lo está el RS se anota en NOTAS_RS;
    - valores: los de la serie del RS 2025. Si falta el valor y el RC del año lo informa, se toma el del RC; si ambos
      existen y difieren, queda el del RS 2025 y se anota. Para los indicadores que mide SERVIR (IOP.03.02 e IOP.05.03)
      prevalece SERIE_SERVIR: el valor publicado por SERVIR para el año de referencia con la fuente de la ficha (ENAHO),
      porque los reportes asignaron a cada año el último dato publicado o mezclaron la ENAHO con registros administrativos;
    - sin meta, sin valor o sin sentido definido: se clasifican aparte (nota metodológica, numeral 9). También una meta
      ≤ 0, o un valor < 0 (ascendente) o ≤ 0 (descendente), porque la razón no es interpretable;
    - correcciones: CORRECCIONES. Los años con fuente no comparable (NO_COMPARABLE) se semaforizan con advertencia.
  El recálculo uniforme importa: los RC no aplicaron siempre la misma fórmula (el RC 2021, p. 2, da 120,0 % al IOP.01.01
  con meta 7,8 y valor 7,6, cuando la razón descendente da 102,6 %).

Uso (desde la raíz del repositorio; requiere pdftotext):
    uv run --no-project --with openpyxl --with pandas --with pdfplumber python src/analisis_pnig/rev_ind_indicadores.py
"""
import re
import subprocess
from pathlib import Path

import pandas as pd
import pdfplumber
from openpyxl.styles import Alignment, Font, PatternFill
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
from nucleo.config import insumo  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

RAIZ = EVAL
CONSOLIDADO = f"{RAIZ}/03_datos/raw/2026_omep_consolidado-indicadores-pnig.xlsx"
PANEL = f"{RAIZ}/03_datos/procesado/panel_formulario_seguimiento_pnig_2023_2025.csv"
RC = "docs/fuentes/pnig/04_seguimiento/{}_mimp_reporte-cumplimiento-pnig.pdf"
RS = "docs/fuentes/pnig/04_seguimiento/{}_mimp_reporte-seguimiento-pnig.pdf"
SALIDA = f"{RAIZ}/04_analisis/REV_IND_base_indicadores.xlsx"
ANIOS = range(2019, 2026)
P = "docs/fuentes/pnig"
ACTAS = f"{P}/03_indicadores_protocolos/actas_modificacion_fichas/"

# Metas verificadas en el documento primario que corrigen la hoja de metas del consolidado: {indicador: ({año: meta}, fuente)}
_VER = insumo("indicadores_verificaciones.yaml")   # evaluaciones/<id>/02_insumos (Fase 2)
METAS_VERIFICADAS = {i: ({int(a): m for a, m in x["metas"].items()}, x["fuente"])
                     for i, x in _VER["metas_verificadas"].items()}
# Líneas de base del consolidado que no coinciden con la ficha: {indicador: (valor de la ficha, nota)}
LB_VERIFICADA = {i: (x["valor"], x["nota"]) for i, x in _VER["lb_verificada"].items()}
# Metas o valores del formulario que no se usan: (indicador, año): nota
NOTAS_FORMULARIO = {(x["indicador"], x["anio"]): x["nota"] for x in _VER["notas_formulario"]}
# Metas de los RS que difieren de la meta usada, con la explicación verificada: (indicador, año): nota
NOTAS_RS = {(x["indicador"], x["anio"]): x["nota"] for x in _VER["notas_rs"]}
SV = "docs/fuentes/pnig/04_seguimiento/fuentes_indicadores/"
# Valores verificados en los informes de SERVIR «La mujer en el servicio civil peruano», por año de referencia y con la
# fuente que fija la ficha (ENAHO). Prevalecen sobre el consolidado: {(indicador, año): (valor, fuente)}
SERIE_SERVIR = {(x["indicador"], x["anio"]): (x["valor"], x["fuente"]) for x in _VER["serie_servir"]}
# (indicador, año): (valor corregido, sustento)
CORRECCIONES = {(x["indicador"], x["anio"]): (x["valor"], x["sustento"]) for x in _VER["correcciones"]}
# (indicador, año): advertencia con la que se semaforiza el avance, porque la fuente no es la que fija la ficha
NO_COMPARABLE = {
    ("IOP.05.02", 2024): "Ruptura de serie: SERVIR mide con registros administrativos y no con la ENAHO que fija la ficha; "
                         "la meta se calculó sobre la ENAHO.",
    ("IOP.05.02", 2025): "Ruptura de serie: el valor (−1,03) proviene de la planilla electrónica (Informe SERVIR «Mujer en el "
                         "servicio civil 2025»), no de la ENAHO. SERVIR 2025 y 2026 solo publican la brecha con la planilla.",
    ("IOP.03.02", 2020): "Fuente distinta: el 43 % es la categoría «funcionarios» de la planilla electrónica del MTPE 2020 "
                         "(SERVIR 2022, Gráfico 8, p. 11), no la ENAHO que fija la ficha. SERVIR no publica la ENAHO 2020.",
    ("IOP.03.02", 2022): "Fuente distinta: el 32,7 % es la participación en cargos directivos del AIRHSP al IV trimestre de "
                         "2022 (SERVIR 2024, Gráfico 23, p. 35), no la ENAHO que fija la ficha.",
    ("IOP.05.03", 2025): "Fuente distinta: el 50,1 % proviene de la planilla electrónica (SERVIR 2026, p. 10: 50,0 % al corte "
                         "de julio de 2025), no de la ENAHO; la ENAHO 2025 no está publicada.",
}
# Valores del RC distintos de la serie del RS 2025, con la explicación verificada: (indicador, año): nota
NOTAS_RC = {
    ("IOP.05.03", 2024): "El RC 2024 informó 49,3, que es la planilla electrónica 2024 parcial (SERVIR 2025, p. 8); el RC "
                         "2025 y el RS 2025 usan 47,3, que es la ENAHO 2022 (SERVIR 2024, p. 5). La ENAHO 2024 da 48,2 "
                         "(SERVIR 2026, p. 10), que es el valor usado.",
}
# Valoración del equipo evaluador sobre el sentido de los indicadores que no se resuelven con la regla general:
# {indicador: (sentido coherente con lo que mide, valoración, hallazgo)}
SENTIDO_OBS = {
    "ISS.01.01.05.01": ("Ascendente", "Sentido no definido en la ficha",
                        "La ficha consigna «Promedio», que es un tipo de agregación y no un sentido. El indicador es un "
                        "porcentaje de casos resueltos dentro del tiempo estándar, por lo que su sentido coherente es "
                        "ascendente. El RS 2022 lo registró como ascendente y los RS 2023 a 2025 como «promedio»."),
    "ISS.04.02.02.01": ("Ascendente", "Sentido no definido en la ficha",
                        "La ficha consigna «N.D. / Indeterminado». El indicador es la tasa de variación anual de mujeres "
                        "con título de propiedad inscrito; el resultado deseado es su aumento, por lo que su sentido "
                        "coherente es ascendente. Al ser una tasa de variación admite valores negativos, que impiden "
                        "calcular la razón de avance."),
    "IOP.05.02": ("Descendente hacia 0", "Coherente, con precisión pendiente",
                  "El sentido descendente es coherente mientras la brecha sea positiva. La ficha no fija el valor de "
                  "llegada: con un valor negativo, como el de 2025 (−1,03, planilla electrónica), el sentido descendente "
                  "premiaría una brecha inversa. Corresponde precisar que el resultado deseado es la brecha nula."),
    "IOP.05.03": ("Ascendente hasta la paridad", "Coherente, con precisión pendiente",
                  "El sentido ascendente es coherente mientras la participación de las mujeres sea menor que la paridad "
                  "(50 %); la serie ya está entre 46 % y 50 %. El RS 2023 (p. 5) lo registró como descendente y calculó "
                  "un avance de 107,4 % donde la razón ascendente da 93,1 %. La Estructura PNIG 2022 ya recomendaba "
                  "precisar el sentido."),
    "IOP.03.01": ("Ascendente hasta la paridad", "Coherente, con precisión pendiente",
                  "El índice promedia ratios de mujeres sobre hombres, cuyo valor de paridad es 1; el sentido ascendente "
                  "es coherente hasta ese valor."),
    "ISS.03.01.02.01": ("Ascendente hasta la paridad", "Coherente, con precisión pendiente",
                        "Es una tasa de paridad, cuyo valor de llegada es 1; el sentido ascendente es coherente hasta ese "
                        "valor."),
    "ISS.04.02.03.01": ("Ascendente hasta la paridad", "Coherente, con precisión pendiente",
                        "Es una tasa de paridad, cuyo valor de llegada es 1; el sentido ascendente es coherente hasta ese "
                        "valor."),
}
SIN_VALOR = {"ND": "No disponible", "L": "Limitaciones para medir", "NR": "No reportado"}
PNIG = [(150, "Alto con posible falla de planeación"), (95, "Alto"), (75, "Medio"), (15, "Bajo"),
        (float("-inf"), "Bajo con posible falla de planeación")]
COLOR = {"Alto con posible falla de planeación": "548235", "Alto": "A9D08E", "Medio": "FFE699", "Bajo": "F4B084",
         "Bajo con posible falla de planeación": "C00000"}
ADV = " (fuente no comparable)"
LINEA_RC = re.compile(r"(I(?:OP|SS)\.[\d.]*\d)\.?\s*Meta (\d{4}):\s*(-?[\d.]+)[^/]*//\s*Valor \d{4}:\s*(-?[\d.]+|ND)[^/]*"
                      r"//\s*Cumplimiento:\s*([^\n]+)")
CODIGO = re.compile(r"^I(?:OP|SS)\.[\d.]*\d$")


def numero(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None
    if isinstance(x, (int, float)):
        return float(x)
    try:
        return float(str(x).strip().replace(",", "."))
    except ValueError:
        return None


def iguales(a, b):
    return abs(a - b) <= max(0.051, 0.005 * abs(b))


from nucleo.escalas import avance, nivel as _nivel  # noqa: E402,F401  (razón de avance y semáforo del núcleo)


def nivel(v):
    return _nivel(v, PNIG)


def lineas_rc():
    """{(indicador, año): (meta, valor, cumplimiento, año del RC)} de los RC 2022–2025."""
    datos = {}
    for anio in range(2022, 2026):
        texto = subprocess.run(["pdftotext", "-enc", "UTF-8", "-layout", RC.format(anio), "-"], capture_output=True,
                               check=True).stdout.decode("utf-8")
        for cod, a, meta, valor, cumpl in LINEA_RC.findall(texto):
            datos[(cod, int(a))] = (float(meta), numero(valor), re.sub(r"\s+", " ", cumpl).strip(), anio)
    return datos


def tablas_rs():
    """Lee las tablas de los RS 2021–2025.

    Devuelve ({(indicador, año): {año del RS: (meta, página)}}, {indicador: {año del RS: sentido}}). Del RS 2022 en
    adelante (formato del aplicativo del CEPLAN) las columnas se ubican por la fila de cabecera «Año | valor | 2019 …»: la
    primera aparición de cada año es el logro esperado y la segunda el valor obtenido. El RS 2021 solo trae los OP, con
    dos años (2020 y 2021), y se lee por el orden de sus celdas: año y valor de la línea de base, metas y valores.
    """
    metas_rs, sentidos = {}, {}
    for anio_rs in range(2021, 2026):
        mapa = None
        with pdfplumber.open(RS.format(anio_rs)) as pdf:
            for npag, pg in enumerate(pdf.pages, 1):
                for tabla in pg.extract_tables():
                    for i, fila in enumerate(tabla):
                        cel = [re.sub(r"\s+", "", c or "") for c in fila]
                        anios = [(j, int(c)) for j, c in enumerate(cel) if re.fullmatch(r"20(?:19|2\d)", c)]
                        if len(anios) >= 4 and any(c in ("Año", "Añ o") or c.startswith("Añ") for c in cel):
                            col_meta = {}
                            for j, a in anios:
                                col_meta.setdefault(a, j)
                            previa = [re.sub(r"\s+", "", c or "").lower() for c in tabla[i - 1]] if i else []
                            col_sen = next((j for j, c in enumerate(previa) if c.startswith("sentido")), None)
                            mapa = (len(fila), col_meta, col_sen)
                            continue
                        cod = next((c for c in cel if CODIGO.match(c)), None)
                        if not cod:
                            continue
                        if anio_rs == 2021:
                            fichas_rs = [c for c in cel[cel.index(cod) + 1:] if re.fullmatch(r"-?[\d.]+|ND|NR|L", c)]
                            if len(fichas_rs) >= 6:
                                for k, a in enumerate((2020, 2021)):
                                    metas_rs.setdefault((cod, a), {})[anio_rs] = (numero(fichas_rs[2 + k]), npag)
                            continue
                        if mapa is None or len(fila) != mapa[0]:
                            continue
                        _, col_meta, col_sen = mapa
                        for a, j in col_meta.items():
                            metas_rs.setdefault((cod, a), {})[anio_rs] = (numero(cel[j]), npag)
                        if col_sen is not None and cel[col_sen]:
                            sentidos.setdefault(cod, {})[anio_rs] = cel[col_sen].lower()
    return metas_rs, sentidos


fichas = pd.read_excel(CONSOLIDADO, sheet_name="1. Indicadores PNIG", dtype=object)
fichas = fichas[fichas["Código del indicador"].notna()].set_index("Código del indicador")
metas = pd.read_excel(CONSOLIDADO, sheet_name="2. Metas y resultados", header=1, dtype=object)
metas = metas[metas["Código del indicador"].notna()].set_index("Código del indicador")
panel = pd.read_csv(PANEL, dtype={"anio": int})
panel["codigo"] = panel["id"].str.replace("PNIG_", "", regex=False)
panel = panel.set_index(["codigo", "anio"])
rc = lineas_rc()
metas_rs, sentidos_rs = tablas_rs()

filas, filas_rs = [], []
for cod, m in metas.iterrows():
    f = fichas.loc[cod]
    sen_orig, sen_vig = str(f["Sentido esperado (original)"]).strip(), str(f["Sentido esperado (vigente)"]).strip()
    sentido = sen_vig.split()[0] if sen_vig else ""   # «Ascendente (acumulado anual)» → «Ascendente»
    verif, fuente_verif = METAS_VERIFICADAS.get(cod, ({}, ""))
    for anio in ANIOS:
        meta_cons, bruto = numero(m[f"Meta {anio}"]), m[f"Valor {anio}"]
        meta = verif.get(anio, meta_cons)
        valor, fuente_valor = numero(bruto), "RS 2025 (consolidado OMEP)"
        meta_rc, valor_rc, cumpl_rc, anio_rc = rc.get((cod, anio), (None, None, "", None))
        if valor is None and valor_rc is not None:
            valor, fuente_valor = valor_rc, f"RC {anio_rc}"
        if (cod, anio) in SERIE_SERVIR:
            valor, fuente_valor = SERIE_SERVIR[(cod, anio)]
        corr = CORRECCIONES.get((cod, anio))
        if corr:
            valor, fuente_valor = corr[0], "Corrección del equipo evaluador"
        if valor is None:
            fuente_valor = ""
        razon, motivo = avance(sentido, meta, valor)
        if motivo == "Sin valor" and str(bruto).strip() in SIN_VALOR:
            motivo = f"Sin valor ({SIN_VALOR[str(bruto).strip()]})"
        advertencia = NO_COMPARABLE.get((cod, anio), "")
        semaforo = motivo or nivel(razon) + (ADV if advertencia else "")

        # Contraste de la meta usada con las de los RS
        en_rs = {a_rs: mt for a_rs, (mt, _) in metas_rs.get((cod, anio), {}).items() if mt is not None}
        rs_anio, rs_2025 = en_rs.get(anio), en_rs.get(2025)
        if not en_rs:
            estado_rs = "Sin meta en los RS" if meta is not None else ""
        elif meta is None:
            estado_rs = "Solo en los RS"
        elif rs_2025 is not None and iguales(meta, rs_2025):
            estado_rs = "Coincide con el RS 2025"
        elif rs_anio is not None and iguales(meta, rs_anio):
            estado_rs = "Coincide con el RS del año"
        elif any(iguales(meta, v) for v in en_rs.values()):
            estado_rs = f"Coincide con el RS {max(a for a, v in en_rs.items() if iguales(meta, v))}"
        elif any(iguales(meta * 100, v) for v in en_rs.values()):
            estado_rs = "Coincide (el RS la expresa en porcentaje)"
        else:
            estado_rs = "Difiere"
        if estado_rs:
            filas_rs.append({
                "Código": cod, "Año": anio, "Entidad": f["Entidad responsable"], "Meta usada": meta,
                "Fuente de la meta": fuente_verif if anio in verif else "Ficha vigente (consolidado OMEP)",
                "Meta en el RS del año": rs_anio, "Meta en el RS 2025": rs_2025,
                "Metas en los RS": "; ".join(f"RS {a_rs} (p. {metas_rs[(cod, anio)][a_rs][1]}): {v:g}"
                                             for a_rs, v in sorted(en_rs.items())),
                "Contraste": estado_rs, "Nota": NOTAS_RS.get((cod, anio), ""),
            })

        p = panel.loc[(cod, anio)] if (cod, anio) in panel.index else None
        v_form, m_form = (numero(p["valor"]), numero(p["meta"])) if p is not None else (None, None)
        v_cons = numero(bruto)
        if p is None:
            contraste = "Sin registro en el formulario" if anio >= 2023 else ""
        elif v_form is None and v_cons is None:
            contraste = "Ambos sin valor"
        elif v_form is None or v_cons is None:
            contraste = "Solo en el formulario" if v_cons is None else "Solo en el consolidado"
        else:
            contraste = "Coincide" if iguales(v_form, v_cons) else "Difiere"
        if anio_rc is None:
            contraste_rc = ""
        elif valor_rc is None:
            contraste_rc = "El RC no informa valor"
        elif v_cons is None:
            contraste_rc = "Solo en el RC (se usa)"
        else:
            contraste_rc = "Coincide" if iguales(v_cons, valor_rc) else "Difiere"
        filas.append({
            "Código": cod, "Nivel": "OP" if cod.startswith("IOP") else "Servicio", "OP": f["Cód. OP"],
            "Indicador": f["Nombre del indicador (vigente)"], "Entidad": f["Entidad responsable"], "Año": anio,
            "Sentido original": sen_orig, "Sentido aplicado": sen_vig,
            "Meta": meta, "Meta consolidado": meta_cons,
            "Fuente de la meta": fuente_verif if anio in verif else ("Ficha vigente (consolidado OMEP)" if meta is not None
                                                                     else ""),
            "Contraste de la meta con los RS": estado_rs,
            "Valor": valor, "Fuente del valor": fuente_valor, "Valor consolidado": bruto,
            "Corrección": corr[1] if corr else "",
            "Avance PNIG (%)": razon, "Semáforo PNIG": semaforo,
            "Motivo de no cálculo": motivo or "", "Advertencia de comparabilidad": advertencia,
            "Meta RC": meta_rc, "Valor RC": valor_rc, "Cumplimiento RC": cumpl_rc, "Contraste RC": contraste_rc,
            "Nota sobre el RC": NOTAS_RC.get((cod, anio), ""),
            "Meta formulario": m_form, "Valor formulario": v_form, "Contraste valor formulario": contraste,
            "Contraste meta formulario": "" if p is None or m_form is None or meta is None else
            ("Coincide" if iguales(m_form, meta) else "Difiere"),
            "Nota sobre el formulario": NOTAS_FORMULARIO.get((cod, anio), ""),
        })
serie = pd.DataFrame(filas)
contraste_rs = pd.DataFrame(filas_rs)

# Coherencia del sentido con lo que mide el indicador (valoración del equipo evaluador)
filas_sen = []
for cod, f in fichas.iterrows():
    sen_orig, sen_vig = str(f["Sentido esperado (original)"]).strip(), str(f["Sentido esperado (vigente)"]).strip()
    detalle = str(f["DETALLE DE LAS MODIFICACIONES"])
    acta = re.search(r"Acta de Validación N\.° ([\w-]+)", detalle)
    en_rs = sentidos_rs.get(cod, {})
    registro_rs = "; ".join(f"RS {a}: {s}" for a, s in sorted(en_rs.items()))
    if sen_orig.startswith("No aplica"):
        coherente, valoracion = sen_vig.split()[0], "Coherente (indicador nuevo)"
        hallazgo, documento = "Indicador aprobado en 2025 con el protocolo del servicio; su sentido es coherente.", \
            str(f["Norma de aprobación del protocolo"])
    elif sen_orig != sen_vig:
        coherente, valoracion = sen_vig.split()[0], "Incoherente en la ficha original; corregido por acta no ratificada"
        documento = str(f["Acta / documento de modificación"])
        if acta:
            op = str(f["Cód. OP"]).lower().replace(".", "")   # la numeración de actas se repite entre OP
            archivos = ", ".join(x.name for x in sorted(Path(ACTAS).glob(f"*acta-{acta.group(1)[:7]}_*")) if op in x.name)
            documento = f"Acta de Validación N.° {acta.group(1)} ({f['Estado del acta']}); {archivos}"
        hallazgo = (f"La ficha original fija «{sen_orig}» para un indicador cuyo resultado deseado es su reducción. El "
                    f"acta lo corrigió a «{sen_vig}», pero la corrección no se formalizó; los RS desde 2022 y las metas "
                    "ya siguen el sentido corregido.")
    else:
        coherente, valoracion, hallazgo, documento = sen_vig.split()[0], "Coherente", "", ""
    if cod in SENTIDO_OBS:   # la valoración propia prevalece sobre «Coherente» y se suma a la corrección por acta
        coherente, obs_valoracion, obs_hallazgo = SENTIDO_OBS[cod]
        valoracion = obs_valoracion if valoracion == "Coherente" else f"{valoracion}; {obs_valoracion.lower()}"
        hallazgo = f"{hallazgo} {obs_hallazgo}"
    if en_rs and len({s[:4] for s in en_rs.values()}) > 1 and cod not in SENTIDO_OBS:
        hallazgo += " Los RS registran sentidos distintos entre años."
    filas_sen.append({"Código": cod, "Indicador": f["Nombre del indicador (vigente)"], "Entidad": f["Entidad responsable"],
                      "Sentido original": sen_orig, "Sentido vigente (aplicado)": sen_vig, "Sentido en los RS": registro_rs,
                      "Sentido coherente con lo que mide": coherente, "Valoración": valoracion,
                      "Documento que identificó la mejora": documento, "Hallazgo": hallazgo.strip()})
sentido_df = pd.DataFrame(filas_sen)

# Base en formato REV_IND.1: un indicador por fila con la serie completa
base = []
for cod, g in serie.groupby("Código", sort=False):
    f, m, g = fichas.loc[cod], metas.loc[cod], g.set_index("Año")
    fila = {"Código": cod, "Nivel": g["Nivel"].iloc[0], "OP": f["Cód. OP"], "Indicador": f["Nombre del indicador (vigente)"],
            "Entidad": f["Entidad responsable"], "Sentido original": f["Sentido esperado (original)"],
            "Sentido aplicado": f["Sentido esperado (vigente)"], "Año LB": m["Año LB"],
            "Valor LB": LB_VERIFICADA[cod][0] if cod in LB_VERIFICADA else m["Valor LB"],
            "Estado del acta": f["Estado del acta"]}
    fila |= {f"Meta {a}": g.at[a, "Meta"] for a in ANIOS}
    fila |= {f"Valor {a}": g.at[a, "Valor"] if g.at[a, "Valor"] is not None else g.at[a, "Valor consolidado"]
             for a in ANIOS}
    fila |= {f"Avance {a} (%)": g.at[a, "Avance PNIG (%)"] for a in ANIOS}
    fila |= {f"Semáforo {a}": g.at[a, "Semáforo PNIG"] for a in ANIOS}
    valores = [g.at[a, "Valor"] for a in ANIOS]
    sen = sentido_df.set_index("Código").loc[cod]
    alertas = []
    if numero(fila["Valor LB"]) is None:
        alertas.append("sin línea de base numérica")
    if cod in LB_VERIFICADA:
        alertas.append("línea de base del consolidado distinta de la ficha")
    if all(g.at[a, "Meta"] is None for a in ANIOS):
        alertas.append("sin metas 2019–2025")
    if all(g.at[a, "Avance PNIG (%)"] is None for a in ANIOS):
        alertas.append("sin avance calculable en el periodo")
    if sen["Valoración"] != "Coherente":
        alertas.append(f"sentido: {sen['Valoración'].lower()}")
    if cod in METAS_VERIFICADAS:
        alertas.append("metas del consolidado corregidas con el documento primario")
    if (g["Contraste de la meta con los RS"] == "Difiere").any():
        alertas.append("meta del RS distinta de la ficha (ver contraste con los RS)")
    if any(c == cod for c, _ in NOTAS_FORMULARIO):
        alertas.append("meta o valor del formulario no usado (ver nota)")
    if (g["Fuente del valor"].str.startswith("RC")).any():
        alertas.append("valor tomado del RC por faltar en el RS 2025")
    if (g["Contraste RC"] == "Difiere").any():
        alertas.append("valor revisado entre el RC y el RS 2025")
    servir = [(a, v) for (c, a), (v, _) in SERIE_SERVIR.items() if c == cod]
    if any(numero(m[f"Valor {a}"]) is None or not iguales(numero(m[f"Valor {a}"]), v) for a, v in servir):
        alertas.append("valores del consolidado corregidos con los informes de SERVIR (año de referencia o fuente)")
    if (g["Advertencia de comparabilidad"] != "").any():
        alertas.append("años con fuente no comparable, semaforizados con advertencia")
    if any(a is not None and a == b for a, b in zip(valores, valores[1:])):
        alertas.append("valor idéntico al del año anterior")
    if any(g["Semáforo PNIG"].str.contains("posible falla")):
        alertas.append("posible falla de planeación en algún año")
    if (g["Contraste valor formulario"] == "Difiere").any():
        alertas.append("valor distinto entre formulario y consolidado")
    fila["Alertas (D5)"] = "; ".join(alertas)
    fila["Avance 2025 consolidado (%)"] = numero(m["Avance 2025 (%)"])
    fila["Notas del consolidado"] = m["Notas sobre metas y resultados"]
    base.append(fila)
base = pd.DataFrame(base)

resumen = (serie.pivot_table(index=["Nivel", "Semáforo PNIG"], columns="Año", values="Código", aggfunc="count",
                             fill_value=0).reset_index())
contraste = serie[serie["Contraste valor formulario"].isin(["Difiere", "Solo en el formulario", "Solo en el consolidado"]) |
                  (serie["Contraste meta formulario"] == "Difiere") |
                  serie["Contraste RC"].isin(["Difiere", "Solo en el RC (se usa)"])][
    ["Código", "Año", "Entidad", "Meta", "Meta consolidado", "Meta formulario", "Contraste meta formulario", "Meta RC",
     "Valor consolidado", "Valor formulario", "Contraste valor formulario", "Valor RC", "Contraste RC", "Valor",
     "Fuente del valor", "Nota sobre el formulario", "Nota sobre el RC"]]
correcciones = pd.DataFrame(
    [{"Código": c, "Año": a, "Tratamiento": "Meta del consolidado corregida", "Valor": v, "Sustento": s}
     for c, (metas_c, s) in METAS_VERIFICADAS.items() for a, v in metas_c.items()] +
    [{"Código": c, "Año": None, "Tratamiento": "Línea de base del consolidado corregida", "Valor": v, "Sustento": s}
     for c, (v, s) in LB_VERIFICADA.items()] +
    [{"Código": c, "Año": a, "Tratamiento": "Dato del formulario no usado", "Valor": None, "Sustento": s}
     for (c, a), s in NOTAS_FORMULARIO.items()] +
    [{"Código": c, "Año": a, "Tratamiento": "Meta del RS no usada", "Valor": None, "Sustento": s}
     for (c, a), s in NOTAS_RS.items()] +
    [{"Código": c, "Año": a, "Tratamiento": "Valor corregido", "Valor": v, "Sustento": s}
     for (c, a), (v, s) in CORRECCIONES.items()] +
    [{"Código": c, "Año": a, "Tratamiento": "Semaforizado con advertencia de comparabilidad", "Valor": None, "Sustento": s}
     for (c, a), s in NO_COMPARABLE.items()] +
    [{"Código": r["Código"], "Año": r["Año"], "Tratamiento": "Valor tomado del RC", "Valor": r["Valor"],
      "Sustento": f"El RS 2025 no trae el valor; lo informa el {r['Fuente del valor']}."}
     for _, r in serie[serie["Fuente del valor"].str.startswith("RC")].iterrows()] +
    [{"Código": c, "Año": a, "Tratamiento": "Valor revisado entre reportes", "Valor": None, "Sustento": s}
     for (c, a), s in NOTAS_RC.items()] +
    [{"Código": c, "Año": a, "Tratamiento": "Valor verificado en SERVIR", "Valor": v,
      "Sustento": f"{f}, en {SV}. Consolidado: {metas.at[c, f'Valor {a}']}."}
     for (c, a), (v, f) in SERIE_SERVIR.items()])
criterios = pd.DataFrame([
    ("PNIG", "Avance", "Ascendente: VO/LE × 100; descendente: LE/VO × 100; sin acotar.", "RC de la PNIG 2023–2025"),
    ("PNIG", "Semáforo", "Alto con posible falla de planeación ≥ 150; alto [95–150>; medio [75–95>; bajo [15–75>; bajo con "
     "posible falla de planeación < 15; plomo cuando hay limitaciones para reportar.",
     "RC 2025, Anexo 01, p. 16; nota metodológica, Tabla 9"),
    ("Equipo evaluador", "Sentido aplicado", "El vigente del consolidado, coherente con la trayectoria de las metas. La hoja "
     "«Sentido» valora su coherencia con lo que mide el indicador.", "Consolidado OMEP, hoja 1; actas de 2022 no ratificadas"),
    ("Equipo evaluador", "Metas", "Ficha vigente; donde las fuentes difieren, la del protocolo o del acta ratificada. Se "
     "contrastan con los logros esperados de los RS 2021–2025.", "Protocolos; actas N.° 003 y 004-2023-OMEP-PNIG; RS"),
    ("Equipo evaluador", "Valores", "Serie del RS 2025; si falta, el valor del RC del año; si difieren, el del RS 2025 con "
     "nota.", "Consolidado OMEP; RC 2022–2025"),
    ("Equipo evaluador", "Casos aparte", "Sin meta, sin valor, sin sentido definido y no interpretable (meta o valor ≤ 0); "
     "no se semaforizan.", "Nota metodológica, numeral 9"),
    ("Equipo evaluador", "Fuente no comparable", "Se semaforiza con la advertencia «(fuente no comparable)» cuando el valor "
     "no proviene de la fuente que fija la ficha.", "Nota metodológica, numeral 9"),
    ("Equipo evaluador", "Tolerancia del contraste", "0,051 o 0,5 % del valor, para absorber el redondeo de los reportes.",
     "Decisión del equipo"),
], columns=["Plano", "Elemento", "Regla", "Fuente"])

with pd.ExcelWriter(SALIDA, engine="openpyxl") as xw:
    for nombre, df in [("Base REV_IND.1", base), ("Serie anual", serie), ("Resumen PNIG", resumen),
                       ("Sentido", sentido_df), ("Contraste metas RS", contraste_rs),
                       ("Contraste de fuentes", contraste), ("Correcciones", correcciones), ("Criterios", criterios)]:
        df.to_excel(xw, sheet_name=nombre, index=False)
        ws = xw.sheets[nombre]
        ws.freeze_panes = "B2"
        for c in ws[1]:
            c.font = Font(bold=True)
            c.alignment = Alignment(wrap_text=True, vertical="top")
        for col in ws.columns:
            ancho = max(len(str(c.value or "")) for c in col[:60])
            ws.column_dimensions[col[0].column_letter].width = min(max(ancho, 8), 50)
            if str(col[0].value).startswith("Semáforo"):
                for c in col[1:]:
                    etq = str(c.value or "").removesuffix(ADV)
                    if etq in COLOR:
                        c.fill = PatternFill("solid", fgColor=COLOR[etq])
                    elif c.value:
                        c.fill = PatternFill("solid", fgColor="D9D9D9")

print(f"OK {SALIDA}: {len(base)} indicadores, {len(serie)} indicador-año; {len(rc)} líneas de RC leídas; "
      f"{len(metas_rs)} indicador-año en los RS")
print(resumen.to_string(index=False))
print(f"\nContraste de metas con los RS: {contraste_rs['Contraste'].value_counts().to_dict()}")
print(contraste_rs[contraste_rs["Contraste"].isin(["Difiere", "Solo en los RS"])][
    ["Código", "Año", "Meta usada", "Metas en los RS"]].to_string(index=False))
print(f"\nSentido: {sentido_df['Valoración'].value_counts().to_dict()}")
print(f"Contraste con el formulario: {serie['Contraste valor formulario'].value_counts().to_dict()}")
print(f"Contraste con los RC: {serie['Contraste RC'].value_counts().to_dict()}")
dif = base.dropna(subset=["Avance 2025 consolidado (%)"])
dif = dif[(dif["Avance 2025 (%)"].fillna(-1e9) - dif["Avance 2025 consolidado (%)"]).abs() > 1.5]
print(f"Avance 2025 distinto del consolidado (> 1,5 pp): {len(dif)}")
print(dif[["Código", "Avance 2025 (%)", "Avance 2025 consolidado (%)", "Semáforo 2025"]].to_string(index=False))
