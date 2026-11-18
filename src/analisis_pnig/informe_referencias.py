"""Elementos de consulta del informe de cierre de la PNIG: abreviaturas, nombres de entidades, glosario, referencias y
documentos de respaldo (ADR-010).

- ABREVIATURAS y ENTIDADES: siglas con su nombre completo. Los nombres de las unidades del MIMP se tomaron de los
  documentos del corpus (informes LIO, informes de la CONAIG y protocolos); los de las entidades, de su denominación
  oficial.
- GLOSARIO: definiciones operativas del equipo evaluador. Cuando una definición sigue una guía o una norma, la cita
  con su página verificada; las demás son definiciones de uso de esta evaluación.
- REFERENCIAS: obras citadas en el texto, en formato APA 7 (en español). Las 172 fuentes de la revisión de evidencia
  se agregan desde la hoja «Referencias» de la matriz de evidencia (informe_cierre.py).
- documentos_respaldo(): documentos del corpus para el Anexo 10, con su título tal como figura en el registro de
  fuentes (docs/fuentes/registro_fuentes.md) y su enlace público o la vía de consulta.

Lo importa informe_cierre.py.
"""
import re

import pandas as pd

ENTIDADES = {
    "ANA": "Autoridad Nacional del Agua", "COFOPRI": "Organismo de Formalización de la Propiedad Informal",
    "FONDEPES": "Fondo Nacional de Desarrollo Pesquero", "INEI": "Instituto Nacional de Estadística e Informática",
    "INPE": "Instituto Nacional Penitenciario", "JNE": "Jurado Nacional de Elecciones",
    "MEF": "Ministerio de Economía y Finanzas", "MIDAGRI": "Ministerio de Desarrollo Agrario y Riego",
    "MIDIS": "Ministerio de Desarrollo e Inclusión Social", "MIMP": "Ministerio de la Mujer y Poblaciones Vulnerables",
    "MINAM": "Ministerio del Ambiente", "MINCETUR": "Ministerio de Comercio Exterior y Turismo",
    "MINEDU": "Ministerio de Educación", "MININTER": "Ministerio del Interior",
    "MINJUSDH": "Ministerio de Justicia y Derechos Humanos", "MINSA": "Ministerio de Salud",
    "MPFN": "Ministerio Público, Fiscalía de la Nación", "MTPE": "Ministerio de Trabajo y Promoción del Empleo",
    "MVCS": "Ministerio de Vivienda, Construcción y Saneamiento", "ONPE": "Oficina Nacional de Procesos Electorales",
    "PJ": "Poder Judicial", "PRODUCE": "Ministerio de la Producción",
    "PROMPERU": "Comisión de Promoción del Perú para la Exportación y el Turismo",
    "SERVIR": "Autoridad Nacional del Servicio Civil",
}
UNIDADES_MIMP = {
    "DFF": "Dirección de Fortalecimiento de Familias", "DGCVG": "Dirección General contra la Violencia de Género",
    "DGFC": "Dirección General de la Familia y la Comunidad",
    "DGPDAEM": "Dirección General de Promoción y Desarrollo de la Autonomía Económica de las Mujeres",
    "DPIGND": "Dirección de Políticas de Igualdad de Género y No Discriminación",
    "DTEG": "Dirección de Transversalización del Enfoque de Género",
    "INABIF": "Programa Integral Nacional para el Bienestar Familiar",
    "PN Warmi Ñan": "Programa Nacional para la Prevención y Erradicación de la Violencia contra las Mujeres e "
                    "Integrantes del Grupo Familiar (Warmi Ñan)",
}


def entidad_completa(texto):
    """«MIMP (DTEG y DPIGND)» -> «Ministerio de la Mujer y Poblaciones Vulnerables (Dirección de Transversalización …
    y Dirección de Políticas …)»; «MININTER» -> «Ministerio del Interior». Varias entidades van separadas por «; »."""
    salida = []
    for parte in (x.strip() for x in str(texto).split(";")):
        m = re.match(r"^(\w+)(?: \((.+)\))?$", parte)
        assert m and m.group(1) in ENTIDADES, f"entidad sin nombre completo: {parte}"
        nombre = ENTIDADES[m.group(1)]
        if m.group(2):
            unidades = [u.strip() for u in re.split(r" y ", m.group(2))]
            assert all(u in UNIDADES_MIMP for u in unidades), f"unidad sin nombre completo: {parte}"
            nombre += " (" + " y ".join(UNIDADES_MIMP[u] for u in unidades) + ")"
        salida.append(nombre)
    return "; ".join(salida)


ABREVIATURAS = {
    "CEDAW": "Convención sobre la eliminación de todas las formas de discriminación contra la mujer",
    "CEM": "Centro Emergencia Mujer",
    "CEPLAN": "Centro Nacional de Planeamiento Estratégico",
    "CONAIG": "Comisión Nacional de Alto Nivel para la Igualdad de Género",
    "DGIGND": "Dirección General de Igualdad de Género y No Discriminación (MIMP)",
    "DGTEG": "Dirección General de Transversalización del Enfoque de Género (MIMP, hasta el ROF de 2023)",
    "DPPDM": "Dirección de Promoción y Protección de los Derechos de las Mujeres (MIMP)",
    "D.S.": "Decreto supremo",
    "DVMM": "Despacho Viceministerial de la Mujer",
    "ENA": "Encuesta Nacional Agropecuaria",
    "ENAHO": "Encuesta Nacional de Hogares",
    "ENARES": "Encuesta Nacional sobre Relaciones Sociales",
    "ENDES": "Encuesta Demográfica y de Salud Familiar",
    "ENUT": "Encuesta Nacional de Uso del Tiempo",
    "EPEN": "Encuesta Permanente de Empleo Nacional",
    "ERM": "Elecciones Regionales y Municipales",
    "IOP": "Indicador de objetivo prioritario",
    "PCM": "Presidencia del Consejo de Ministros",
    "ISS": "Indicador de servicio",
    "LIO": "Ley de Igualdad de Oportunidades entre Mujeres y Hombres (Ley N.° 28983); informe LIO: informe anual de "
           "cumplimiento de esa ley",
    "OCDE": "Organización para la Cooperación y el Desarrollo Económicos",
    "ODS": "Objetivo de Desarrollo Sostenible",
    "OGMEPGD": "Oficina General de Monitoreo, Evaluación de Políticas y Gestión Descentralizada (MIMP)",
    "OGPPM": "Oficina General de Planeamiento, Presupuesto y Modernización (MIMP)",
    "OMEP": "Oficina de Monitoreo y Evaluación de Políticas (OGMEPGD)",
    "OP": "Objetivo prioritario",
    "PEDN": "Plan Estratégico de Desarrollo Nacional",
    "PEI": "Plan Estratégico Institucional",
    "PEMIG": "Plan Estratégico Multisectorial de Igualdad de Género",
    "PESEM": "Plan Estratégico Sectorial Multianual",
    "PGG": "Política General de Gobierno",
    "PIA": "Presupuesto institucional de apertura",
    "PIM": "Presupuesto institucional modificado",
    "PNIG": "Política Nacional de Igualdad de Género",
    "POI": "Plan Operativo Institucional",
    "RC": "Reporte de cumplimiento de la PNIG",
    "R.M.": "Resolución ministerial",
    "ROF": "Reglamento de Organización y Funciones",
    "RPCD": "Resolución de Presidencia de Consejo Directivo (CEPLAN)",
    "RS": "Reporte de seguimiento de la PNIG",
    "SFD": "Situación futura deseada",
    "Sinaplan": "Sistema Nacional de Planeamiento Estratégico",
}
ABREVIATURAS.update(ENTIDADES)
ABREVIATURAS.update({k: v for k, v in UNIDADES_MIMP.items() if k != "PN Warmi Ñan"})

# definiciones operativas de esta evaluación; las que siguen una guía la citan (página verificada en el corpus)
GLOSARIO = [
    ("Avance de un indicador", "Cociente entre el valor obtenido y la meta del año, por cien, si el sentido es "
     "ascendente; entre la meta y el valor obtenido, si es descendente. No considera la línea de base (numeral 2.4)."),
    ("Brecha de género", "Diferencia entre mujeres y hombres en un indicador. En este informe se expresa como "
     "diferencia en puntos porcentuales (mujeres menos hombres) o como razón mujeres/hombres, según el indicador."),
    ("Coeficiente de variación", "Error estándar de una estimación dividido entre la estimación, en porcentaje. Con "
     "el criterio del INEI, una estimación con coeficiente de variación mayor a 15 % es referencial (INEI, 2025a, "
     "p. 178)."),
    ("Confianza del hallazgo", "Respaldo de un hallazgo según el número de fuentes independientes que coinciden: "
     "sólido, con respaldo parcial o indicio (numeral 2.4)."),
    ("Derogación tácita", "Pérdida de vigencia de una norma por su incompatibilidad con una norma posterior, sin "
     "declaración expresa (artículo I del Título Preliminar del Código Civil). Según la OGAJ, es el caso de la PNIG "
     "desde el 25 de diciembre de 2025 (MIMP, 2026e)."),
    ("Devengado", "Etapa de la ejecución del gasto en la que se reconoce la obligación de pago por un bien o servicio "
     "recibido. Es la medida de ejecución que usa este informe."),
    ("Diseño muestral complejo", "Diseño de las encuestas del INEI, con estratos, conglomerados y factores de "
     "expansión. Las estimaciones y sus errores estándar lo toman en cuenta, porque tratar la muestra como aleatoria "
     "simple subestimaría el error."),
    ("Doble codificación", "Segunda codificación, independiente y ciega, de los segmentos que podían cambiar una "
     "valoración, para medir el acuerdo entre codificadores."),
    ("Fuente independiente", "Fuente cuyo dato tiene un productor primario distinto del de las demás que respaldan un "
     "hallazgo: dos documentos que reproducen el dato de un mismo productor cuentan como una sola fuente. La entrevista "
     "a la OGMEPGD no se cuenta como fuente independiente (numeral 2.6)."),
    ("Intervalo de confianza al 95 %", "Intervalo construido con un procedimiento que, en muestras repetidas con el "
     "mismo diseño, contiene el valor poblacional en el 95 % de los casos."),
    ("Kappa de Cohen", "Medida del acuerdo entre dos codificadores que descuenta el acuerdo esperado por azar. Vale 1 "
     "con acuerdo perfecto y 0 cuando el acuerdo es el que daría el azar."),
    ("Línea de base", "Valor del indicador en el año que se toma como punto de partida para medir su evolución."),
    ("Meta", "Valor que el indicador debe alcanzar en un año, según su ficha técnica."),
    ("Nivel conductor", "Órganos del MIMP que condujeron la Política o intervinieron en su seguimiento: la DGIGND, la OGPPM "
     "y la OGMEPGD. Son los entrevistados de esta evaluación."),
    ("Nivel de avance", "Clasificación del avance de un indicador: alto con posible falla de planeación, alto, medio, "
     "bajo o bajo con posible falla de planeación (numeral 2.4)."),
    ("Plausibilidad", "Respaldo que la evidencia da a la relación entre un tipo de servicio y la causa o el efecto del "
     "problema público que atiende."),
    ("PIM", "Presupuesto institucional modificado: presupuesto actualizado de una entidad o una meta, después de las "
     "modificaciones del año."),
    ("Protocolo de servicio", "Norma de la entidad responsable que fija los estándares de cumplimiento y los "
     "indicadores de un servicio de la Política, conforme a la R.M. N.° 116-2021-MIMP."),
    ("Razón mujeres/hombres", "Cociente entre el valor de las mujeres y el de los hombres. Menor que 1 indica que las "
     "mujeres están por debajo de los hombres; 1, paridad."),
    ("Segmento", "Fragmento de texto de un reporte o de una entrevista que trata un solo tema y recibe uno o más "
     "códigos del libro de códigos."),
    ("Servicio", "Prestación que una entidad entrega a la población para cumplir un lineamiento de la Política."),
    ("Situación futura deseada", "Situación que la Política se proponía alcanzar al 2030, expresada en enunciados con "
     "metas."),
    ("Solidez de la evidencia", "Grado de respaldo de la evidencia sobre un tipo de intervención: sólida, moderada, "
     "limitada o sin evidencia identificada (numeral 2.4 y Anexo 4)."),
    ("Triangulación", "Contraste de un hallazgo entre fuentes, métodos y actores distintos."),
]


def documentos_respaldo(registro, inventario, enlaces):
    """Documentos del corpus para el Anexo 10: título, entidad o norma, fecha y enlace o vía de consulta. Lee las
    tablas del registro de fuentes (secciones 2, 3, 4, 4b y 4c) y toma el enlace público del inventario y del
    registro de enlaces verificados."""
    inv = pd.read_csv(inventario)
    inv["archivo"] = inv.ruta.str.split("/").str[-1]
    inv = inv.drop_duplicates("archivo").set_index("archivo")
    enl = pd.read_csv(enlaces, sep="\t")
    enl["archivo"] = enl.documento.str.split("/").str[-1]
    enl = enl.drop_duplicates("archivo").set_index("archivo")
    filas, seccion, cols = [], "", {}
    for linea in open(registro, encoding="utf-8"):
        if linea.startswith("## "):
            seccion, cols = linea[3:].strip(), {}
            continue
        if not (linea.startswith("| ") and seccion[:2] in ("2.", "3.", "4.", "4b", "4c")):
            continue
        celdas = [c.strip() for c in linea.strip().strip("|").split("|")]
        if set(celdas[0]) <= {"-"}:
            continue
        if not cols:
            # cabecera de la tabla: columnas por su nombre (las de 4c no tienen ID ni fecha)
            for k, c in enumerate(celdas):
                clave = ("documento" if c in ("Documento", "Política") else "norma" if c.startswith("Norma") or
                         c.startswith("Símbolo") else "fecha" if c == "Fecha" else
                         "ubicacion" if c.startswith(("Ubicación", "Archivo")) else None)
                if clave:
                    cols[clave] = k
            continue
        val = lambda k: celdas[cols[k]] if k in cols and cols[k] < len(celdas) else ""
        documento, norma, fecha, ubicacion = val("documento"), val("norma"), val("fecha"), val("ubicacion")
        if not documento:
            continue
        archivos = re.findall(r"`([^`]+)`", ubicacion)
        archivo = archivos[0].split("/")[-1] if archivos else ""
        # el registro abrevia los nombres largos con «…»: se busca el archivo por su inicio y su final
        ini, _, fin = archivo.partition("…")
        archivo = next((a for a in inv.index if archivo and a.startswith(ini) and a.endswith(fin)), archivo)
        url, via = "", "A solicitud de la OGMEPGD"
        if archivo in enl.index and str(enl.loc[archivo, "acceso"]).startswith("público"):
            url = enl.loc[archivo, "pagina_publicacion"] if pd.notna(enl.loc[archivo, "pagina_publicacion"]) else \
                enl.loc[archivo, "enlace_archivo"]
        elif archivo in inv.index and str(inv.loc[archivo, "acceso"]).startswith("público"):
            u = inv.loc[archivo, "pagina_publicacion"]
            url = u if pd.notna(u) and str(u).startswith("http") else inv.loc[archivo, "enlace_publico"]
        if isinstance(url, str) and url.startswith("http"):
            via = ""
        else:
            url = ""
        documento = re.sub(r"\s*\((?:p\.|pp\.)[^)]*\)", "", documento.replace("**", "").replace("`", ""))
        filas.append({"seccion": seccion, "documento": documento, "norma": norma.replace("`", ""),
                      "fecha": fecha.replace("`", ""), "url": url, "via": via})
    return filas


# Referencias citadas en el texto (APA 7, en español). Cada entrada: (texto, enlace o None). Las de la revisión de
# evidencia se agregan aparte. El orden final es alfabético.
REFERENCIAS = [
    ("Autoridad Nacional del Servicio Civil [SERVIR]. (2023). La mujer en el servicio civil peruano 2023. SERVIR.",
     None),
    ("Autoridad Nacional del Servicio Civil [SERVIR]. (2025). La mujer en el servicio civil peruano 2025 (1.ª ed. "
     "electrónica). SERVIR.", None),
    ("Autoridad Nacional del Servicio Civil [SERVIR]. (2026). La mujer en el servicio civil peruano 2026 (1.ª ed. "
     "electrónica). SERVIR.", None),
    ("Centro Nacional de Planeamiento Estratégico [CEPLAN]. (s. f.). Informe técnico. Evaluación de la Política "
     "Nacional de Igualdad de Género (PNIG). Dirección Nacional de Seguimiento y Evaluación.", None),
    ("Centro Nacional de Planeamiento Estratégico [CEPLAN]. (2024a). Guía para el seguimiento y evaluación de "
     "políticas nacionales y planes del SINAPLAN. Aprobada por Resolución de Presidencia de Consejo Directivo N.° "
     "0056-2024-CEPLAN/PCD.", None),
    ("Centro Nacional de Planeamiento Estratégico [CEPLAN]. (2024c). Guía para la elaboración de indicadores de "
     "políticas nacionales y planes estratégicos.", None),
    ("Centro Nacional de Planeamiento Estratégico [CEPLAN]. (2026). Informe N.° 000006-2026-CEPLAN-DNCP-LTF: "
     "asistencia técnica para la implementación de la Segunda Disposición Complementaria Final de la Ley N.° 32535. "
     "10 de marzo de 2026.", None),
    ("Decreto Supremo N.° 002-2020-MIMP, que aprueba el Plan Estratégico Multisectorial de Igualdad de Género de la "
     "Política Nacional de Igualdad de Género. (2020, 7 de marzo). El Peruano.", None),
    ("Decreto Supremo N.° 018-2021-MIMP, que crea la Comisión Multisectorial de naturaleza permanente denominada "
     "«Comisión Nacional de Alto Nivel para la Igualdad de Género, CONAIG». (2021, 22 de julio). El Peruano.", None),
    ("Instituto Nacional de Estadística e Informática [INEI]. (2025a). Perú: Brechas de género, 2025. Avances hacia "
     "la igualdad de mujeres y hombres. INEI.", None),
    ("Instituto Nacional de Estadística e Informática [INEI]. (2025b). Fichas técnicas de la Encuesta Nacional de "
     "Hogares 2025, de la ENDES 2025 y de la Encuesta Nacional Agropecuaria 2024; Perú: Indicadores del mercado "
     "laboral a nivel departamental, 2022-2024 (EPEN); Informe de principales resultados de la ENARES 2024; "
     "Principales resultados de la ENUT 2024. INEI.", None),
    ("Instituto Nacional de Estadística e Informática [INEI]. (2025c). Producción y empleo informal en el Perú. "
     "Cuenta satélite de la economía informal 2022-2024. INEI.", None),
    ("Instituto Nacional de Estadística e Informática [INEI]. (2026). Microdatos de la Encuesta Nacional de Hogares "
     "(2016 a 2025), la Encuesta Demográfica y de Salud Familiar (2019 a 2025), la Encuesta Nacional sobre Relaciones "
     "Sociales (2015, 2019 y 2024), la Encuesta Nacional de Uso del Tiempo (2010 y 2024), la Encuesta Permanente de "
     "Empleo Nacional (2023 a 2025) y la Encuesta Nacional Agropecuaria (2019 a 2025) [conjuntos de datos]. "
     "Consultados del 1 al 6 de noviembre de 2026.", "https://proyectos.inei.gob.pe/microdatos/"),
    ("Jurado Nacional de Elecciones [JNE]. (2022a). Perfil electoral N.° 4: El perfil de las candidaturas a las "
     "Elecciones Regionales y Municipales 2022. Dirección Nacional de Educación y Formación Cívica Ciudadana.", None),
    ("Jurado Nacional de Elecciones [JNE]. (2022b). Perfil electoral N.° 5: Participación política de las mujeres en "
     "las ERM 2022. Dirección Nacional de Educación y Formación Cívica Ciudadana.", None),
    ("Jurado Nacional de Elecciones [JNE]. (2026a). Resolución N.° 1411-2026-JNE, proclamación de los diputados "
     "electos para el periodo 2026-2031. 19 de junio de 2026.", None),
    ("Jurado Nacional de Elecciones [JNE]. (2026b). Resolución N.° 1412-2026-JNE, proclamación de los senadores "
     "electos para el periodo 2026-2031. 19 de junio de 2026.", None),
    ("Jurado Nacional de Elecciones [JNE]. (2026c). El 6 de diciembre será la Segunda Elección Regional [nota de "
     "prensa]. 10 de octubre de 2026.", None),
    ("Ley N.° 28983, Ley de Igualdad de Oportunidades entre Mujeres y Hombres. (2007, 16 de marzo). El Peruano.",
     None),
    ("Ley N.° 29470, Ley que modifica diversos artículos de la Ley N.° 27683, Ley de Elecciones Regionales. (2009, "
     "14 de diciembre). El Peruano.", None),
    ("Ley N.° 32535, Ley de Igualdad de Oportunidades entre Mujeres y Hombres. (2025, 24 de diciembre). El Peruano.",
     None),
    ("Ministerio de Economía y Finanzas [MEF]. (2026). Presupuesto y ejecución de gasto, 2019 a 2025 [conjunto de "
     "datos]. Portal de Datos Abiertos del MEF. Consultado del 29 de octubre al 5 de noviembre de 2026.",
     "https://datosabiertos.mef.gob.pe"),
    ("Ministerio de la Mujer y Poblaciones Vulnerables [MIMP]. (2019). Política Nacional de Igualdad de Género. "
     "Aprobada por Decreto Supremo N.° 008-2019-MIMP.", None),
    ("Ministerio de la Mujer y Poblaciones Vulnerables [MIMP]. (2021). Informe de evaluación de implementación 2020. "
     "Política Nacional Multisectorial de Igualdad de Género.", "EVAL_2020"),
    ("Ministerio de la Mujer y Poblaciones Vulnerables [MIMP]. (2022). Informe de evaluación de resultados de la "
     "Política Nacional de Igualdad de Género, periodo 2021.", None),
    ("Ministerio de la Mujer y Poblaciones Vulnerables [MIMP]. (2024). Informe de evaluación de la Política Nacional "
     "de Igualdad de Género, 2023.", "05_evaluaciones_previas/2023_mimp_informe-evaluacion-pnig.pdf"),
    ("Ministerio de la Mujer y Poblaciones Vulnerables [MIMP]. (2025). Reporte de cumplimiento de la Política "
     "Nacional de Igualdad de Género, 2024.", "04_seguimiento/2024_mimp_reporte-cumplimiento-pnig.pdf"),
    ("Ministerio de la Mujer y Poblaciones Vulnerables [MIMP]. (2026a). Reporte de cumplimiento de la Política "
     "Nacional de Igualdad de Género, 2025.", "04_seguimiento/2025_mimp_reporte-cumplimiento-pnig.pdf"),
    ("Ministerio de la Mujer y Poblaciones Vulnerables [MIMP]. (2026b). Informe N.° D000039-2026-MIMP-DPIGND. "
     "Dirección de Políticas de Igualdad de Género y No Discriminación. 13 de julio de 2026.", None),
    ("Ministerio de la Mujer y Poblaciones Vulnerables [MIMP]. (2026c). Informe N.° D000112-2026-MIMP-OP. Oficina de "
     "Planeamiento. 20 de julio de 2026.", None),
    ("Ministerio de la Mujer y Poblaciones Vulnerables [MIMP]. (2026d). Oficio N.° D000289-2026-MIMP-OGPPM. Oficina "
     "General de Planeamiento, Presupuesto y Modernización. 20 de julio de 2026.", None),
    ("Ministerio de la Mujer y Poblaciones Vulnerables [MIMP]. (2026e). Informe N.° D000926-2026-MIMP-OGAJ: sustento "
     "técnico sobre la vigencia de las normas e instrumentos aprobados en el marco de la Ley N.° 28983 tras la entrada "
     "en vigor de la Ley N.° 32535. Oficina General de Asesoría Jurídica. 11 de noviembre de 2026.", None),
    ("Ministerio de la Mujer y Poblaciones Vulnerables [MIMP]. (2026f). Informe N.° D000810-2026-MIMP-OGAJ: opinión "
     "legal sobre la vigencia de las normas e instrumentos aprobados en el marco de la Ley N.° 28983, Ley de igualdad "
     "de oportunidades entre mujeres y hombres, tras la entrada en vigor de la Ley N.° 32535. Oficina General de "
     "Asesoría Jurídica. 16 de setiembre de 2026.", None),
    ("Ministerio de la Mujer y Poblaciones Vulnerables [MIMP]. (2026g). Informe N.° D000041-2026-MIMP-DGIGND: sustento "
     "técnico sobre la vigencia de las normas e instrumentos aprobados en el marco de la Ley N.° 28983 tras la entrada "
     "en vigor de la Ley N.° 32535, y remisión de la matriz técnica de incompatibilidades. Dirección General de "
     "Igualdad de Género y No Discriminación. 9 de noviembre de 2026.", None),
    ("Ministerio de la Mujer y Poblaciones Vulnerables [MIMP]. (2026h). Informe N.° D000067-2026-MIMP-DPIGND: sustento "
     "técnico sobre la vigencia de las normas e instrumentos aprobados en el marco de la Ley N.° 28983 tras la entrada "
     "en vigor de la Ley N.° 32535. Dirección de Políticas de Igualdad de Género y No Discriminación. 9 de noviembre "
     "de 2026.", None),
    ("Ministerio de la Mujer y Poblaciones Vulnerables [MIMP]. (2026i). Informe Técnico N.° D000040-2026-MIMP-DTEG: "
     "sustento técnico sobre la vigencia de las normas e instrumentos aprobados en el marco de la Ley N.° 28983 tras "
     "la entrada en vigor de la Ley N.° 32535. Dirección de Transversalización del Enfoque de Género. 9 de noviembre "
     "de 2026.", None),
    ("Ministerio de la Mujer y Poblaciones Vulnerables [MIMP]. (2020-2026). Informes anuales de cumplimiento de la "
     "Ley N.° 28983, Ley de Igualdad de Oportunidades entre Mujeres y Hombres, 2019 a 2025.", "LIO_COLECCION"),
    ("Ministerio de la Mujer y Poblaciones Vulnerables [MIMP], Comisión Nacional de Alto Nivel para la Igualdad de Género [CONAIG]. "
     "(2022-2025). Informes I a IV de propuestas de mejora de la gobernanza multisectorial para la implementación de "
     "la Política Nacional de Igualdad de Género.", None),
    ("Ministerio de la Mujer y Poblaciones Vulnerables [MIMP], Comisión Nacional de Alto Nivel para la Igualdad de Género [CONAIG]. "
     "(2026). Ayuda memoria [documento interno con las sesiones y los acuerdos de la comisión de 2021 a junio de "
     "2026].", None),
    ("Ministerio de la Mujer y Poblaciones Vulnerables [MIMP], Dirección General de Igualdad de Género y No "
     "Discriminación [DGIGND]. (2026a). Ejecución presupuestal 2019 al 2026 de la actividad 5001072 «Promoción de la "
     "mujer», por finalidad y específica de gasto [base de datos remitida para esta evaluación].", None),
    ("Ministerio de la Mujer y Poblaciones Vulnerables [MIMP], Dirección General de Igualdad de Género y No "
     "Discriminación [DGIGND]. (2026b). Marco conceptual del enfoque de igualdad de oportunidades entre mujeres y "
     "hombres (versión final). 16 de setiembre de 2026.", None),
    ("Oficina General de Monitoreo, Evaluación de Políticas y Gestión Descentralizada [OGMEPGD], MIMP. (2023-2025). "
     "Formulario de seguimiento de la Política Nacional de Igualdad de Género: reportes anuales de las entidades "
     "responsables [base de datos]. Sistema de Seguimiento de Políticas de la OMEP.", None),
    ("Oficina General de Monitoreo, Evaluación de Políticas y Gestión Descentralizada [OGMEPGD], MIMP. (2026). "
     "Consolidado de indicadores de la Política Nacional de Igualdad de Género: fichas, metas y valores 2019-2025 "
     "[base de datos].", None),
    ("Oficina Nacional de Procesos Electorales [ONPE]. (2026a). Analítica electoral 3: Resultados electorales EG "
     "2026: votos, escaños y perfil de autoridades. ONPE.", None),
    ("Oficina Nacional de Procesos Electorales [ONPE]. (2026b). Presentación de resultados de las Elecciones "
     "Regionales y Municipales 2026 [portal de resultados]. Consultado el 2 de noviembre de 2026.",
     "https://resultadoelectoral.onpe.gob.pe"),
    ("Organización para la Cooperación y el Desarrollo Económicos [OCDE]. (2019). Better criteria for better "
     "evaluation: Revised evaluation criteria definitions and principles for use. OCDE, Red de Evaluación del Comité "
     "de Ayuda al Desarrollo.", None),
    ("Peru maps. (2026). Límites administrativos del Perú en GeoJSON [repositorio de datos].",
     "https://github.com/Rodasluis/Peru-maps"),
    ("Resolución de Presidencia de Consejo Directivo N.° 042-2026/CEPLAN/PCD. (2026, 20 de junio). El Peruano.", None),
    ("Resolución Ministerial N.° 116-2021-MIMP, que aprueba los lineamientos para la elaboración de los protocolos o "
     "normas equivalentes de implementación de los servicios de la PNIG. (2021, 21 de abril).", None),
]
