"""REV_IND.2: verificación de la calidad técnica de las fichas de los 62 indicadores vigentes de la PNIG.

Contrasta cada ficha con la Guía para la elaboración de indicadores de políticas nacionales y planes estratégicos (CEPLAN,
2024c) en los diez atributos del Anexo 3 de la nota metodológica (REV_IND.2), y añade dos verificaciones de las secciones
críticas del Anexo A-9 de la guía (pp. 134–135): nivel en la cadena de resultados y tipo de indicador de servicio. Cada
observación lleva severidad (Crítica, Mayor o Menor), sustento en la guía con página y recomendación.

Fuentes:
  - consolidado de indicadores de la OMEP, hojas 1 (fichas transcritas), 2 (metas 2026–2030) y 5 (servicios sin
    indicador);
  - REV_IND.1 (04_analisis/REV_IND_base_indicadores.xlsx): metas 2019–2025 verificadas, valores, sentido y su coherencia;
  - metas 2026–2030 verificadas en las actas ratificadas y en los protocolos (METAS_2026_2030), donde el consolidado
    arrastra las metas anteriores a la modificación.
  - anexo 5 del PEMIG (matriz de los 85 indicadores de servicio) y su clasificación por tipo (cobertura, calidad u
    otro) del equipo evaluador, en 04_analisis/pemig_indicadores/tipo_indicadores_pemig.csv: hojas «Tipo PEMIG» y
    «Tipo PEMIG resumen».

Planos:
  CEPLAN: criterios y páginas de la guía (columna «Sustento»). La guía es de 2024 y las fichas se aprobaron entre 2019 y
          2025 con el formato anterior (Anexo N.° 4 del D.S. 008-2019-MIMP): la verificación mide la distancia frente al
          estándar vigente, que es el que regirá la nueva política.
  Equipo evaluador: las valoraciones de OBS (nombre, fórmula, responsable y tipo de servicio) y las reglas automáticas.
  Límite: la verificación usa la transcripción de la OMEP. Los campos que esta no recoge (justificación, limitaciones,
  proceso de recolección, valores absolutos, segundo nivel de responsable) no se valoran indicador por indicador; los
  rasgos comunes del formato se registran en la hoja «Conjunto».

Uso (desde la raíz del repositorio):
    uv run --no-project --with openpyxl --with pandas python src/analisis_pnig/rev_ind_2_calidad_fichas.py
"""
import re

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

RAIZ = EVAL
CONSOLIDADO = f"{RAIZ}/03_datos/raw/2026_omep_consolidado-indicadores-pnig.xlsx"
PSI051 = f"{RAIZ}/04_analisis/REV_IND_base_indicadores.xlsx"
SALIDA = f"{RAIZ}/04_analisis/REV_IND_2_calidad_fichas.xlsx"
VIGENCIA, HORIZONTE = 2019, 2030
G = "Guía de indicadores (CEPLAN, 2024c)"
ACTAS = "docs/fuentes/pnig/03_indicadores_protocolos/actas_modificacion_fichas/"
PROT = "docs/fuentes/pnig/01_normativa/protocolos/"

# Metas 2026–2030 verificadas en el documento primario (el consolidado conserva las metas anteriores a la modificación)
METAS_2026_2030 = {
    "ISS.01.01.02.01": ([40, 50, 60, 80, 100], f"Protocolo R.M. 1110-2023-MINSA, p. 22 del PDF, en {PROT}"),
    "ISS.01.01.03.01": ([47.3, 50.3, 53.3, 56.3, 59.3], f"Acta N.° 004-2023-OMEP-PNIG, anexo, p. 6 del PDF, en {ACTAS}"),
    "ISS.01.02.02.01": ([38.8, 42.8, 46.8, 50.8, 53.5], f"Acta N.° 003-2023-OMEP-PNIG, anexo, p. 6 del PDF, en {ACTAS}"),
    "ISS.01.02.02.02": ([84.7, 84.9, 85.5, 85.6, 85.7], f"Acta N.° 003-2023-OMEP-PNIG, anexo, p. 8 del PDF, en {ACTAS}"),
}
PARAMETROS = ("porcentaje", "proporción", "tasa", "ratio", "razón", "promedio", "índice", "número")
SUBTIPOS = ("oportunidad", "accesibilidad", "percepción", "precisión", "fiabilidad", "continuidad")
NOTA_CONSOLIDADO = re.compile(r"\s*\((?:el indicador original|propuesta 2023|medición suspendida)[^)]*\)", re.I)
SUSPENDIDA = re.compile(r"suspendid", re.I)
# Fuentes que dejaron de producirse en el periodo, según el consolidado: {texto en la fuente: sustento}
FUENTES_CAIDAS = {"ENARES": "la ENARES no se ejecuta desde 2019 y ya no se publica; sin datos desde 2020 (consolidado de la "
                            "OMEP, observaciones del IOP.06.01)"}
SUST = {
    "nombre": f"{G}, Paso 3, pp. 38–44 (parámetro + sujeto + característica; parámetros admitidos)",
    "formula": f"{G}, Paso 7, Método de cálculo, pp. 91–92; Anexo A-9, p. 134 (coherencia nombre y método)",
    "unidad": f"{G}, Paso 3, pp. 38–43 (parámetro de medición)",
    "lb": f"{G}, Paso 7, p. 93 (periodo previo a la vigencia, el más cercano disponible) y p. 94 (sin valores vacíos)",
    "metas": f"{G}, Paso 7, p. 94 (logros anuales, coherentes con el sentido, sin valores vacíos); Anexo A-5, p. 122",
    "fuente": f"{G}, Paso 7, p. 93 (fuente y base de datos); Paso 4, p. 46 (medible y confiable)",
    "periodicidad": f"{G}, Paso 4, p. 46 (temporal: frecuencia de medición coherente con logros anuales)",
    "responsable": f"{G}, Paso 7, p. 88 (dos niveles de responsabilidad)",
    "desagregacion": f"{G}, Paso 4, p. 46 (capacidad de desagregación territorial y sociodemográfica)",
    "sentido": f"{G}, Paso 7, p. 93 (sentido esperado) y nota 35",
    "tipo": f"{G}, Paso 5, Recuadro 2, pp. 61–62 (cobertura y calidad con su atributo); Anexo A-9, p. 135",
    "nivel": f"{G}, Paso 5, pp. 50–52 y Recuadro 2, p. 61 (OP: resultado; servicio: producto)",
}
ATRIBUTOS = [("nombre", "Denominación precisa y unívoca"), ("formula", "Fórmula de cálculo explícita"),
             ("unidad", "Unidad de medida definida"), ("lb", "Línea de base con año y valor"),
             ("metas", "Meta anual hasta el fin del horizonte"), ("fuente", "Fuente de datos identificada y disponible"),
             ("periodicidad", "Periodicidad de medición establecida"), ("responsable", "Responsable del reporte identificado"),
             ("desagregacion", "Desagregaciones previstas por territorio y grupo poblacional"),
             ("sentido", "Sentido esperado definido y coherente con lo que mide el indicador"),
             ("nivel", "Nivel en la cadena de resultados (A-9)"), ("tipo", "Tipo de indicador de servicio (A-9)")]

# Valoraciones del equipo evaluador: {(indicador, atributo): (valoración, severidad, observación, recomendación)}
OBS = {
    ("IOP.01.01", "nombre"): ("Cumple parcialmente", "Mayor",
        "El nombre «Porcentaje de mujeres víctimas de violencia hacia las mujeres» es redundante y más amplio que lo que "
        "mide la fórmula: violencia física y sexual ejercida por la pareja en los últimos 12 meses, en mujeres de 15 a 49 "
        "años alguna vez unidas.",
        "Denominar «Porcentaje de mujeres de 15 a 49 años alguna vez unidas que sufrieron violencia física o sexual por "
        "parte de su pareja en los últimos 12 meses»."),
    ("IOP.01.01", "formula"): ("Cumple parcialmente", "Mayor",
        "El numerador se restringe a mujeres unidas y alguna vez unidas, y el denominador comprende a todas las mujeres de "
        "15 a 49 años. Además, la fórmula dice «física y sexual» y las especificaciones, «física y/o sexual».",
        "Definir B como «Número de mujeres de 15 a 49 años alguna vez unidas» y unificar la condición «física o sexual»."),
    ("IOP.02.01", "formula"): ("No cumple", "Crítica",
        "El nombre tiene como sujeto a las «mujeres actualmente unidas», pero el denominador es el «número total de mujeres "
        "con demanda de planificación familiar». Nombre y método miden poblaciones distintas.",
        "Alinear el denominador con el sujeto del nombre («Número de mujeres actualmente unidas») o reformular el nombre "
        "según el denominador que efectivamente se calcula."),
    ("IOP.02.03", "nombre"): ("Cumple parcialmente", "Mayor",
        "El nombre «Porcentaje de embarazo adolescente» no identifica el sujeto ni el rango de edad.",
        "Denominar «Porcentaje de mujeres adolescentes de 15 a 19 años que son madres o están embarazadas por primera "
        "vez», o el rango que use la fuente."),
    ("IOP.02.03", "formula"): ("Cumple parcialmente", "Mayor",
        "El denominador «Total de mujeres adolescentes» no precisa el rango de edad que permite reproducir el cálculo con "
        "la ENDES.",
        "Precisar en la fórmula y en las especificaciones el rango de edad del numerador y del denominador."),
    ("IOP.05.01", "formula"): ("No cumple", "Crítica",
        "El numerador cuenta las entidades que incorporan el enfoque de género «en al menos X dimensiones»: el umbral X no "
        "está definido, por lo que un tercero no puede reproducir el cálculo.",
        "Fijar el número mínimo de dimensiones y el criterio de incorporación de cada una en las especificaciones técnicas."),
    ("IOP.05.02", "nombre"): ("No cumple", "Mayor",
        "«Brecha» no es un parámetro admitido. La fórmula ((A − B) / B) × 100 es una diferencia porcentual.",
        "Denominar «Diferencia porcentual entre el salario promedio de los servidores públicos y el de las servidoras "
        "públicas» e indicar si los salarios están en soles corrientes."),
    ("IOP.06.01", "unidad"): ("No cumple", "Mayor",
        "El nombre y la fórmula corresponden a un índice ponderado, pero el parámetro de medición consignado es "
        "«Porcentaje».",
        "Consignar «Índice» como parámetro y precisar su rango de valores."),
    ("ISS.01.01.02.01", "formula"): ("Cumple parcialmente", "Mayor",
        "El nombre mide a los profesionales que «culminan satisfactoriamente» el programa, pero el numerador cuenta a los "
        "que «cuentan con capacitaciones especializadas».",
        "Definir el numerador como profesionales que culminan y aprueban el programa, con el criterio de aprobación."),
    ("ISS.01.01.03.01", "metas"): ("Cumple parcialmente", "Mayor",
        "Las metas 2021 y 2022 (51,0 y 55,7) son de la ficha original y las de 2023 a 2030 (41,3 a 59,3), del Acta N.° "
        "004-2023-OMEP-PNIG, que redefinió el indicador: la serie mezcla dos definiciones y desciende en 2023.",
        "Separar las series o recalcular 2021 y 2022 con la definición vigente."),
    ("ISS.01.01.05.01", "responsable"): ("Cumple parcialmente", "Mayor",
        "Se consignan las «Presidencias de las Cortes Superiores de Justicia», que son varias unidades sin una que integre "
        "y reporte el dato nacional.",
        "Identificar la unidad del Poder Judicial que integra y reporta el indicador."),
    ("ISS.01.01.08.01", "tipo"): ("Cumple parcialmente", "Mayor",
        "Es un indicador de cobertura cuyo denominador son las víctimas «que acuden al servicio» (la demanda) y no la "
        "población objetivo.",
        "Si el servicio es a demanda, declararlo y justificar el denominador; si no, usar la población objetivo."),
    ("ISS.03.01.01.01", "nombre"): ("Cumple parcialmente", "Mayor",
        "El nombre indica «Porcentaje de satisfacción», pero el método describe un índice de satisfacción construido a "
        "partir de encuestas.",
        "Denominar «Porcentaje de participantes satisfechos con los eventos de educación electoral» y definir el umbral "
        "de satisfacción, o adoptar el parámetro índice."),
    ("ISS.03.01.02.01", "nombre"): ("Cumple parcialmente", "Mayor",
        "El nombre usa «tasa de paridad», pero la fórmula A/B es un ratio entre mujeres y hombres.",
        "Denominar «Ratio de mujeres a hombres miembros de organizaciones sociales y políticas que participan en las "
        "actividades de formación y asistencia técnica»."),
    ("ISS.03.01.02.02", "tipo"): ("Cumple parcialmente", "Menor",
        "El tipo consigna dos atributos («pertinencia / accesibilidad»).",
        "Elegir el atributo esencial que mide la fórmula."),
    ("ISS.04.03.01.01", "formula"): ("No cumple", "Crítica",
        "La fórmula consignada, RCCL = (D / C) × 100, divide la demanda (D) entre las evaluadas (C): invierte el cociente "
        "que describe el nombre y da valores mayores que 100.",
        "Corregir a (C / D) × 100: mujeres evaluadas entre mujeres que demandan el servicio."),
    ("ISS.04.03.01.02", "nombre"): ("No cumple", "Mayor",
        "«Nivel» no es un parámetro admitido; la fórmula es un porcentaje.",
        "Denominar «Porcentaje de mujeres evaluadas satisfechas con el servicio de certificación de competencias "
        "laborales»."),
    ("ISS.04.03.01.02", "tipo"): ("Cumple parcialmente", "Menor",
        "Se consigna el atributo «pertinencia»; la fórmula mide la satisfacción de las usuarias.",
        "Consignar el atributo «percepción»."),
    ("ISS.04.03.05.05", "tipo"): ("Cumple parcialmente", "Mayor",
        "Es un indicador de cobertura cuyo denominador es el total de personas capacitadas, no la población objetivo: "
        "mide la composición por sexo de las participantes.",
        "Usar como denominador la población objetivo de mujeres o reclasificar el indicador."),
    ("ISS.04.03.05.05", "responsable"): ("Cumple parcialmente", "Mayor",
        "Se consignan «Direcciones Generales, Unidades Ejecutoras y Organismos Públicos que intervienen», sin una unidad "
        "que integre el dato.",
        "Identificar la unidad de PRODUCE que integra y reporta el indicador."),
    ("ISS.04.03.05.06", "tipo"): ("Cumple parcialmente", "Mayor",
        "El atributo «culminación» no corresponde a lo que mide la fórmula: la proporción de mujeres entre las personas que "
        "formalizaron su empresa.",
        "Reformular el indicador para medir un atributo del servicio (por ejemplo, porcentaje de usuarias asesoradas que "
        "formalizan su empresa) y consignar el atributo correspondiente."),
    ("ISS.04.04.01.01", "tipo"): ("Cumple parcialmente", "Mayor",
        "Se consigna «Calidad (continuidad)», pero la fórmula (hogares que reciben acompañamiento sobre el total de hogares "
        "con adolescentes) mide cobertura.",
        "Consignar el tipo «Cobertura» o reformular el indicador para medir la continuidad del acompañamiento."),
    ("ISS.05.01.01.01", "formula"): ("Cumple parcialmente", "Mayor",
        "El nombre se refiere a las «entidades públicas priorizadas», pero el denominador considera a las entidades que "
        "culminan el proceso de asistencia técnica.",
        "Alinear nombre y denominador."),
    ("ISS.06.01.02.01", "tipo"): ("Cumple parcialmente", "Mayor",
        "El atributo «enfoque de género» no es un atributo de calidad del servicio en los términos de la guía.",
        "Consignar el atributo que mide la fórmula o tratar el enfoque como estándar de cumplimiento."),
    ("ISS.06.01.02.02", "tipo"): ("Cumple parcialmente", "Mayor",
        "El atributo «enfoque de género» no es un atributo de calidad del servicio en los términos de la guía.",
        "Consignar el atributo que mide la fórmula o tratar el enfoque como estándar de cumplimiento."),
    ("ISS.06.01.02.02", "responsable"): ("Cumple parcialmente", "Mayor",
        "Se consignan cinco unidades (DIGEDD, DIFODS, DIGC, DIGE y DIF) sin una que integre el dato.",
        "Identificar la unidad del MINEDU que integra y reporta el indicador."),
    ("ISS.06.01.02.03", "tipo"): ("Cumple parcialmente", "Mayor",
        "Se consigna el atributo «fiabilidad», pero el nombre mide la asistencia técnica recibida «oportunamente».",
        "Consignar el atributo «oportunidad» y definir el plazo que la hace oportuna."),
    ("ISS.06.01.02.03", "responsable"): ("Cumple parcialmente", "Mayor",
        "Se consignan varias unidades (DIGEDD, DIFOID, DIFODS, DIGC y DIGE) sin una que integre el dato.",
        "Identificar la unidad del MINEDU que integra y reporta el indicador."),
    ("ISS.06.03.01.01", "nombre"): ("Cumple parcialmente", "Menor",
        "El nombre no identifica la encuesta («incluidas en la encuesta»).",
        "Denominar «Porcentaje de variables de desagregación interseccional priorizadas incluidas en la ENARES»."),
}
# Fuentes de los indicadores que mide SERVIR con la ENAHO: hallazgo de REV_IND.1
for c in ("IOP.03.02", "IOP.05.02", "IOP.05.03"):
    OBS[(c, "fuente")] = ("Cumple parcialmente", "Mayor",
        "La ficha consigna «SERVIR (ENAHO - INEI)», pero los reportes del periodo usaron también la planilla electrónica y "
        "el AIRHSP según la publicación disponible (REV_IND.1, hoja «Correcciones»).",
        "Fijar como base de datos la ENAHO e identificar la publicación de SERVIR y el año de referencia de cada valor.")


def texto(x):
    return "" if pd.isna(x) else str(x).strip()


def numero(x):
    """Primer número del texto («0.334 (antes 19.9 %)» → 0.334); None si no hay número o es ND."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None
    if isinstance(x, (int, float)):
        return float(x)
    m = re.match(r"\s*(-?\d+(?:[.,]\d+)?)", str(x))
    return float(m.group(1).replace(",", ".")) if m else None


def anio(x):
    m = re.search(r"(19|20)\d{2}", texto(x))
    return int(m.group(0)) if m else None


fichas = pd.read_excel(CONSOLIDADO, sheet_name="1. Indicadores PNIG", dtype=object)
fichas = fichas[fichas["Código del indicador"].notna()].set_index("Código del indicador")
metas_c = pd.read_excel(CONSOLIDADO, sheet_name="2. Metas y resultados", header=1, dtype=object)
metas_c = metas_c[metas_c["Código del indicador"].notna()].set_index("Código del indicador")
sin_ind = pd.read_excel(CONSOLIDADO, sheet_name="5. Servicios sin indicador", dtype=object)
base = pd.read_excel(PSI051, sheet_name="Base REV_IND.1").set_index("Código")
sentido = pd.read_excel(PSI051, sheet_name="Sentido").set_index("Código")

nombres = {c: NOTA_CONSOLIDADO.sub("", texto(f["Nombre del indicador (vigente)"])) for c, f in fichas.iterrows()}
repetidos = {n for n in nombres.values() if list(nombres.values()).count(n) > 1}

matriz, obs = [], []
for cod, f in fichas.iterrows():
    servicio = cod.startswith("ISS")
    b, sen = base.loc[cod], sentido.loc[cod]
    nombre, metodo, param = nombres[cod], texto(f["MÉTODO DE CÁLCULO (vigente)"]), texto(f["Parámetro / unidad de medida"])
    tipo, fuente, organo = texto(f["Tipo de indicador"]), texto(f["Fuente y base de datos"]), texto(f["Órgano / unidad responsable de la medición"])
    notas = " ".join([texto(metas_c.at[cod, "Notas sobre metas y resultados"])] +
                     [texto(f[c]) for c in ("Observaciones", "Especificaciones técnicas", "Nombre del indicador (vigente)")])
    val = {}

    def poner(atr, valoracion, severidad="", observacion="", recomendacion=""):
        """Registra la valoración; una valoración del equipo (OBS) prevalece sobre la regla automática."""
        if (cod, atr) in OBS:
            valoracion, severidad, observacion, recomendacion = OBS[(cod, atr)]
        val[atr] = valoracion
        if observacion:
            obs.append({"id_ficha": cod, "elemento": texto(f["Cód. servicio"]) if servicio else texto(f["Cód. OP"]),
                        "indicador": nombre, "entidad": texto(f["Entidad responsable"]),
                        "campo": dict(ATRIBUTOS)[atr], "valoracion": valoracion, "severidad": severidad,
                        "observacion": observacion, "sustento": SUST[atr], "recomendacion": recomendacion})

    # 1. Denominación
    primera = nombre.split()[0].lower() if nombre else ""
    if primera not in PARAMETROS and not nombre.lower().startswith("tasa de variación"):
        poner("nombre", "No cumple", "Mayor", f"El nombre empieza con «{nombre.split()[0]}», que no es un parámetro admitido.",
              "Reformular el nombre con un parámetro admitido (porcentaje, tasa, ratio, promedio, índice o número).")
    elif nombre in repetidos:
        otros = ", ".join(c for c, n in nombres.items() if n == nombre and c != cod)
        poner("nombre", "Cumple parcialmente", "Mayor", f"El nombre es idéntico al de {otros}, de otra entidad: no es unívoco.",
              "Incluir en el nombre el programa o la entidad que presta el servicio.")
    else:
        poner("nombre", "Cumple")
    # 2. Fórmula
    if "imagen" in metodo:
        poner("formula", "No consta", "Menor", "La transcripción de la OMEP indica que la fórmula figura como imagen en la "
              "ficha; no se pudo verificar su correspondencia con el nombre.", "Transcribir la fórmula y definir sus variables.")
    else:
        poner("formula", "Cumple")
    # 3. Unidad
    if "→" in param:
        poner("unidad", "Cumple parcialmente", "Mayor", f"El parámetro cambió de «{param.split('→')[0].strip()}» a "
              f"«{param.split('→')[1].strip()}» por un acta de 2022 no ratificada; la ficha aprobada conserva el original.",
              "Formalizar el cambio de parámetro con el procedimiento de modificación de la Guía de Políticas Nacionales "
              "(numeral 4.2.1).")
    else:
        poner("unidad", "Cumple")
    # 4. Línea de base
    anio_lb, valor_lb = anio(b["Año LB"]), numero(b["Valor LB"])
    lb_cons = texto(f["Valor línea de base"])
    prot = anio(f["Año de aprobación del protocolo / del indicador"])
    if valor_lb is None:
        detalle = (f"el valor que registra el consolidado («{lb_cons}») no corresponde a la ficha (REV_IND.1, hoja "
                   "«Correcciones»)" if numero(lb_cons) is not None else f"consolidado: «{lb_cons}»")
        poner("lb", "No cumple", "Crítica", f"La ficha no consigna valor de línea de base ({detalle}).",
              "Calcular la línea de base con la fuente del indicador para el año previo más cercano a su aprobación.")
    elif anio_lb and anio_lb > VIGENCIA - 1:
        nota = " El año 2020 o 2021 es atípico por la pandemia." if anio_lb in (2020, 2021) else ""
        aprob = f"; el indicador se aprobó en {prot} con el protocolo del servicio" if prot else ""
        poner("lb", "Cumple parcialmente", "Menor", f"La línea de base es de {anio_lb}, posterior a la entrada en vigencia "
              f"de la PNIG (abril de 2019){aprob}.{nota}",
              "En la nueva política, fijar líneas de base anteriores a su entrada en vigencia.")
    else:
        poner("lb", "Cumple")
    # 5. Metas 2019–2030
    metas = {a: numero(b[f"Meta {a}"]) for a in range(2019, 2026)}
    extra = METAS_2026_2030.get(cod, ([numero(metas_c.at[cod, f"Meta {a}"]) for a in range(2026, 2031)], ""))[0]
    metas |= dict(zip(range(2026, 2031), extra))
    con_meta = [a for a, m in metas.items() if m is not None]
    sen_apl = texto(b["Sentido aplicado"]).split()[0].lower() if texto(b["Sentido aplicado"]) else ""
    problemas = []
    if not con_meta:
        poner("metas", "No cumple", "Crítica", "La ficha no tiene metas en el horizonte 2019–2030.",
              "Establecer logros anuales hasta 2030 con su sustento.")
    else:
        faltan = [a for a in range(con_meta[0], HORIZONTE + 1) if metas[a] is None]
        serie = [metas[a] for a in con_meta]
        if faltan:
            problemas.append(("Crítica", f"faltan metas en {', '.join(map(str, faltan))}"))
        if "porcentaje" in param.lower() and "→" not in param and any(m > 100 or m < 0 for m in serie):
            problemas.append(("Crítica", "hay metas fuera del rango de un porcentaje"))
        pasos = list(zip(serie, serie[1:]))
        if sen_apl == "ascendente" and any(y < x for x, y in pasos) or sen_apl == "descendente" and any(y > x for x, y in pasos):
            problemas.append(("Crítica", f"las metas no siguen el sentido {sen_apl}"))
        elif sen_apl in ("ascendente", "descendente") and any(y == x for x, y in pasos):
            problemas.append(("Menor", "las metas se repiten en años consecutivos sin nota que lo explique"))
        if valor_lb is not None and sen_apl in ("ascendente", "descendente"):
            if (sen_apl == "ascendente" and serie[-1] <= valor_lb) or (sen_apl == "descendente" and serie[-1] >= valor_lb):
                problemas.append(("Crítica", "la meta de 2030 no supera la línea de base en el sentido esperado"))
        if problemas:
            sev = "Crítica" if any(s == "Crítica" for s, _ in problemas) else "Menor"
            nota_c = texto(metas_c.at[cod, "Notas sobre metas y resultados"])
            poner("metas", "No cumple" if sev == "Crítica" else "Cumple parcialmente", sev,
                  f"Metas {con_meta[0]}–{HORIZONTE}: " + "; ".join(t for _, t in problemas) + "."
                  + (f" Nota del consolidado: {nota_c}" if nota_c else ""),
                  "Completar y corregir los logros anuales según el sentido esperado, con su sustento.")
        else:
            poner("metas", "Cumple")
    # 6. Fuente
    caida = next((s for k, s in FUENTES_CAIDAS.items() if k in fuente), "")
    frase = next((s.strip() for s in re.split(r"(?<=\.)\s+|\(|\)", notas) if SUSPENDIDA.search(s)), "")
    if caida:
        poner("fuente", "No cumple", "Crítica", f"La fuente dejó de producir el dato durante el periodo: {caida}.",
              "Sustituir la fuente o el indicador por uno medible con una fuente vigente.")
    elif "no especificada" in fuente:
        poner("fuente", "Cumple parcialmente", "Mayor", "La base de datos del denominador no está especificada.",
              "Identificar la base de datos del denominador.")
    else:
        poner("fuente", "Cumple")
    # 7. Periodicidad (temporal): valores efectivamente disponibles desde el primer año con meta hasta 2025
    anios = [a for a in range(max(con_meta[0] if con_meta else VIGENCIA, VIGENCIA, prot or VIGENCIA), 2026)]
    sin_dato = [a for a in anios if numero(b[f"Valor {a}"]) is None]
    if caida:
        val["periodicidad"] = "No cumple"   # la observación se registra en el atributo fuente
    elif frase:
        poner("periodicidad", "No cumple", "Crítica", f"La medición se suspendió durante el periodo (consolidado: «{frase}»).",
              "Reanudar la medición o sustituir el indicador; la suspensión deja al servicio sin seguimiento.")
    elif len(anios) < 2:
        val["periodicidad"] = "No aplica"   # indicador aprobado al final del periodo
    elif len(sin_dato) == len(anios):
        poner("periodicidad", "No cumple", "Crítica", f"No hay valores en ninguno de los años {anios[0]}–2025.",
              "Asegurar una fuente de producción anual o declarar la periodicidad real como limitación.")
    elif len(sin_dato) >= 2:
        poner("periodicidad", "Cumple parcialmente", "Mayor", f"Sin valor en {len(sin_dato)} de {len(anios)} años "
              f"({', '.join(map(str, sin_dato))}).", "Verificar que la fuente produzca el dato cada año; si no, declararlo como "
              "limitación y ajustar los logros a la periodicidad real.")
    else:
        poner("periodicidad", "Cumple")
    # 8. Responsable
    poner("responsable", "Cumple" if organo else "No consta")
    # 9. Desagregación: el formato de ficha no tiene el campo (ver hoja «Conjunto»)
    val["desagregacion"] = "No consta"
    # 10. Sentido (REV_IND.1, hoja «Sentido»)
    v = texto(sen["Valoración"])
    if v == "Coherente" or v.startswith("Coherente (indicador nuevo"):
        poner("sentido", "Cumple")
    elif v.startswith("Sentido no definido"):
        poner("sentido", "No cumple", "Crítica", texto(sen["Hallazgo"]), "Definir el sentido esperado como "
              f"{texto(sen['Sentido coherente con lo que mide']).lower()}.")
    elif v.startswith("Incoherente"):
        poner("sentido", "Cumple parcialmente", "Mayor", texto(sen["Hallazgo"]), "Formalizar el sentido corregido con el "
              "procedimiento de modificación de la Guía de Políticas Nacionales (numeral 4.2.1).")
    else:
        poner("sentido", "Cumple parcialmente", "Menor", texto(sen["Hallazgo"]), "Precisar en la ficha el valor de llegada.")
    # A-9: nivel en la cadena y tipo de servicio
    if servicio:
        poner("nivel", "Cumple")
        sub = re.search(r"\(([^)]*)\)", tipo)
        if tipo.startswith("Cobertura"):
            poner("tipo", "Cumple")
        elif tipo.startswith("Calidad") and sub and any(s in sub.group(1).lower() for s in SUBTIPOS):
            poner("tipo", "Cumple")
        elif tipo.startswith("Calidad"):
            poner("tipo", "Cumple parcialmente", "Menor", f"El atributo de calidad consignado, «{sub.group(1) if sub else ''}», "
                  "no es uno de los que nombra la guía (oportunidad, accesibilidad, percepción, precisión o continuidad).",
                  "Precisar el atributo esencial que mide la fórmula.")
        else:
            poner("tipo", "No cumple", "Crítica", f"La ficha de un servicio consigna el tipo «{tipo}», que no es cobertura ni calidad.",
                  "Consignar el tipo de indicador de servicio (cobertura o calidad con su atributo).")
    else:
        poner("nivel", "Cumple" if tipo.startswith("Resultado") else "No cumple")
        val["tipo"] = "No aplica"
    matriz.append({"Código": cod, "Nivel": "Servicio" if servicio else "OP", "Indicador": nombre,
                   "Entidad": texto(f["Entidad responsable"])} | {etq: val[a] for a, etq in ATRIBUTOS})

matriz = pd.DataFrame(matriz)
obs = pd.DataFrame(obs)
dictamen = matriz[["Código", "Nivel", "Indicador", "Entidad"]].copy()
dictamen["Críticas"] = dictamen["Código"].map(lambda c: (obs[(obs.id_ficha == c)].severidad == "Crítica").sum())
dictamen["Mayores"] = dictamen["Código"].map(lambda c: (obs[(obs.id_ficha == c)].severidad == "Mayor").sum())
dictamen["Menores"] = dictamen["Código"].map(lambda c: (obs[(obs.id_ficha == c)].severidad == "Menor").sum())
dictamen["Dictamen"] = dictamen.apply(lambda r: "Observada" if r["Críticas"] or r["Mayores"] else
                                      ("Conforme con observaciones menores" if r["Menores"] else "Conforme"), axis=1)

# Revisión del conjunto
conj = []
srv = fichas[fichas.index.str.startswith("ISS")].copy()
srv["clase"] = srv["Tipo de indicador"].map(lambda t: "cobertura" if str(t).startswith("Cobertura") else
                                            ("calidad" if str(t).startswith("Calidad") else "otro"))
for s, g in srv.groupby("Cód. servicio"):
    clases = set(g["clase"])
    if not {"cobertura", "calidad"} <= clases:
        falta = " y ".join(sorted({"cobertura", "calidad"} - clases))
        conj.append({"Elemento": s, "Severidad": "Mayor", "Observación": f"El servicio tiene {len(g)} indicador(es) "
                     f"({', '.join(g.index)}) y le falta un indicador de {falta}.", "Sustento": SUST["tipo"],
                     "Recomendación": f"Incorporar un indicador de {falta} en la nueva política."})
serv_sin = sin_ind[sin_ind["Código del servicio"].astype(str).str.startswith("SS.")]
lin_sin = sin_ind[~sin_ind.index.isin(serv_sin.index)]
assert len(lin_sin) == 1 and str(lin_sin["Código del servicio"].iloc[0]).startswith("Lin.")
conj.append({"Elemento": "Servicios sin indicador", "Severidad": "Crítica",
             "Observación": f"{len(serv_sin)} servicios no tienen indicador vigente porque no tienen protocolo aprobado, "
                            f"y el lineamiento 3.2 no tiene servicio (consolidado de la OMEP, hoja 5): no hay forma de "
                            "valorar su provisión con indicadores.", "Sustento": SUST["tipo"],
             "Recomendación": "Definir en la nueva política al menos un indicador de cobertura y uno de calidad por servicio."})
conj.append({"Elemento": "Formato de ficha", "Severidad": "Menor",
             "Observación": "El formato de ficha de la PNIG (Anexo N.° 4 del D.S. 008-2019-MIMP) no tiene campo de "
                            "desagregación; ninguna ficha transcrita la prevé. Para la PNIG, que adopta el enfoque "
                            "interseccional, la desagregación por territorio, edad, etnia y discapacidad es relevante.",
             "Sustento": SUST["desagregacion"], "Recomendación": "Incluir el nivel de desagregación en las fichas de la nueva "
             "política (campo adicional admitido por la guía, p. 86)."})
conj.append({"Elemento": "Formato de ficha", "Severidad": "Menor",
             "Observación": "El formato de ficha de la PNIG tiene un solo campo «Responsable del indicador»; la guía de 2024 "
                            "distingue a quien integra el dato y a quien responde por el objetivo o el servicio.",
             "Sustento": SUST["responsable"], "Recomendación": "Consignar ambos niveles en las fichas de la nueva política."})
conj.append({"Elemento": "Modificación de fichas", "Severidad": "Mayor",
             "Observación": "Las fichas se modificaron mediante actas de validación (2022–2023), varias no ratificadas, "
                            "porque la Guía de Políticas Nacionales de 2018 no regulaba la modificación de indicadores. Los "
                            "reportes aplicaron cambios no formalizados (sentido, parámetro, metas).",
             "Sustento": "Guía de Políticas Nacionales (CEPLAN, 2024b), numeral 4.2.1", "Recomendación": "Modificar los "
             "indicadores de la nueva política con el procedimiento formal de la guía vigente."})
conj = pd.DataFrame(conj)

resumen = (matriz.melt(id_vars=["Código", "Nivel"], value_vars=[e for _, e in ATRIBUTOS], var_name="Atributo",
                       value_name="Valoración")
           .pivot_table(index="Atributo", columns="Valoración", values="Código", aggfunc="count", fill_value=0)
           .reindex([e for _, e in ATRIBUTOS]).reset_index())
criterios = pd.DataFrame([
    ("Valoración", "Cumple; cumple parcialmente; no cumple; no consta (el dato no figura en la transcripción de la OMEP); "
                   "no aplica.", "Nota metodológica, Anexo 3 (REV_IND.2)"),
    ("Severidad", "Crítica: el indicador no es evaluable o mide otra cosa (secciones críticas del Anexo A-9). Mayor: "
                  "incumple un criterio y se corrige sin rediseñar el indicador. Menor: forma o precisión.",
     "Skill de revisión de fichas técnicas; Anexo A-9 de la guía, pp. 134–135"),
    ("Dictamen", "Observada si tiene al menos una observación crítica o mayor; conforme con observaciones menores si solo "
                 "tiene menores; conforme si no tiene observaciones.", "Skill de revisión de fichas técnicas"),
    ("Estándar", "Guía de 2024. Las fichas se aprobaron con el formato anterior; la verificación mide la distancia frente "
                 "al estándar vigente, que regirá la nueva política.", "Equipo evaluador"),
    ("Metas", "2019–2025 de REV_IND.1 (verificadas); 2026–2030 del consolidado, salvo METAS_2026_2030 (actas y protocolo "
              "del MINSA). Se verifica completitud anual desde el primer año con meta, rango y sentido.", "Equipo evaluador"),
    ("Periodicidad", "Años con valor efectivamente disponible entre el primer año con meta (o de aprobación del "
                     "protocolo) y 2025 (REV_IND.1). Con menos de dos años en el periodo, no aplica. Si la fuente dejó de "
                     "producirse, la observación se registra en el atributo fuente.",
     "Equipo evaluador"),
    ("Límite", "No se valoran por indicador la justificación, las limitaciones, el proceso de recolección, los valores "
               "absolutos ni el segundo nivel de responsable, que la transcripción de la OMEP no recoge.", "Equipo evaluador"),
], columns=["Elemento", "Regla", "Fuente"])

# Tipo de los indicadores de servicio del PEMIG (anexo 5): cobertura, calidad u otro, con el criterio del Recuadro 2 de
# la guía (pp. 61-62). La clasificación es del equipo evaluador (pemig_indicadores/tipo_indicadores_pemig.csv); aquí
# se comprueba que cada indicador coincida con la matriz del anexo 5 y se resume.

PEMIG_A5 = "docs/fuentes/pnig/02_pemig/2020_mimp_pemig_anexo-5_fichas-indicadores-servicios.xlsx"
TIPO_PEMIG = f"{RAIZ}/04_analisis/pemig_indicadores/tipo_indicadores_pemig.csv"
a5 = [(str(r[0]).strip().rstrip("."), re.sub(r"\s+", " ", str(r[2])).strip())
      for r in list(openpyxl.load_workbook(PEMIG_A5, read_only=True, data_only=True).worksheets[0]
                    .iter_rows(values_only=True))[8:]
      if r[0] and re.match(r"\d\.\d\.\d", str(r[0]).strip()) and r[2]]
tipo_pemig = pd.read_csv(TIPO_PEMIG, dtype={"servicio": str})
assert list(zip(tipo_pemig.servicio, tipo_pemig.indicador)) == a5, "la clasificación no coincide con el anexo 5"
assert set(tipo_pemig.tipo) <= {"Cobertura", "Calidad", "Otro"}
tipo_pemig.insert(1, "OP", tipo_pemig.servicio.str[0].map(lambda x: f"OP{x}"))
n_pemig = len(tipo_pemig)
res_pemig = (tipo_pemig.groupby(["tipo", "atributo"]).size().rename("Indicadores").reset_index()
             .assign(**{"% de los indicadores": lambda d: (100 * d.Indicadores / n_pemig).round(1)}))
por_tipo = tipo_pemig.tipo.value_counts()
serv_cal = tipo_pemig.loc[tipo_pemig.tipo == "Calidad", "servicio"].nunique()
res_pemig = pd.concat([res_pemig, pd.DataFrame(
    [{"tipo": f"Total {t}", "atributo": "", "Indicadores": int(por_tipo.get(t, 0)),
      "% de los indicadores": round(100 * por_tipo.get(t, 0) / n_pemig, 1)} for t in ("Cobertura", "Calidad", "Otro")]
    + [{"tipo": "Servicios con al menos un indicador de calidad", "atributo": "",
        "Indicadores": serv_cal, "% de los indicadores": None},
       {"tipo": "Servicios con indicador en el PEMIG", "atributo": "", "Indicadores": tipo_pemig.servicio.nunique(),
        "% de los indicadores": None}])], ignore_index=True)
res_pemig = res_pemig.rename(columns={"tipo": "Tipo", "atributo": "Atributo"})
criterios = pd.concat([criterios, pd.DataFrame([
    ("Tipo de indicador del PEMIG", "Cobertura: mide la provisión del servicio respecto de su población objetivo o en "
     "unidades territoriales o institucionales. Calidad: mide un atributo del servicio (oportunidad, accesibilidad, "
     "percepción, precisión, continuidad, cumplimiento del estándar o del protocolo, aprobación o conclusión). Otro: "
     "participación de las mujeres entre los beneficiarios, paridad, resultado en la población o gestión de la "
     "información. Se clasifica por el nombre y la fórmula; si no coinciden, por el nombre.",
     f"{G}, Paso 5, Recuadro 2, pp. 61–62 (criterios de cobertura y de calidad); clasificación del equipo evaluador"),
], columns=criterios.columns)], ignore_index=True)

COLOR = {"Cumple": "C6EFCE", "Cumple parcialmente": "FFEB9C", "No cumple": "FFC7CE", "No consta": "D9D9D9",
         "Crítica": "FFC7CE", "Mayor": "FFEB9C", "Menor": "DDEBF7", "Observada": "FFC7CE",
         "Conforme con observaciones menores": "DDEBF7", "Conforme": "C6EFCE"}
with pd.ExcelWriter(SALIDA, engine="openpyxl") as xw:
    for hoja, df in [("Matriz de atributos", matriz), ("Observaciones", obs), ("Dictamen", dictamen),
                     ("Conjunto", conj), ("Resumen", resumen), ("Tipo PEMIG", tipo_pemig),
                     ("Tipo PEMIG resumen", res_pemig), ("Criterios", criterios)]:
        df.to_excel(xw, sheet_name=hoja, index=False)
        ws = xw.sheets[hoja]
        ws.freeze_panes = "B2"
        for c in ws[1]:
            c.font = Font(bold=True)
            c.alignment = Alignment(wrap_text=True, vertical="top")
        for col in ws.columns:
            ancho = max(len(str(c.value or "")) for c in col[:80])
            ws.column_dimensions[col[0].column_letter].width = min(max(ancho, 9), 60)
            for c in col[1:]:
                c.alignment = Alignment(wrap_text=True, vertical="top")
                if c.value in COLOR:
                    c.fill = PatternFill("solid", fgColor=COLOR[c.value])

print(f"OK {SALIDA}: {len(matriz)} fichas, {len(obs)} observaciones, {len(conj)} del conjunto")
print(dictamen["Dictamen"].value_counts().to_string())
print(obs["severidad"].value_counts().to_string())
print(resumen.to_string(index=False))
