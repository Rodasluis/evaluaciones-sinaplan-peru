"""ENTREV: guías de entrevista semiestructurada al nivel conductor y registro de codificación.

Desarrolla el Anexo 4 de la nota metodológica (bloques 0 a 5 y preguntas orientadoras, que se copian sin cambios) con
preguntas de profundización que contrastan hallazgos concretos de REV_DOC, REV_ALI, REV_IND, REP_ENT y REV_EVI. Cada
pregunta declara la pregunta de evaluación que atiende, los códigos del libro común (libro_codigos_pnig.csv), los
hallazgos que contrasta (sintesis_hallazgos.py) y la evidencia con su fuente y página.

Las preguntas se diseñan en preguntas(C) y, después de las entrevistas, el catálogo, las guías y los cuestionarios
quedan con lo que se preguntó (ADR-008): las preguntas de la guía que no se formularon se eliminan y las surgidas se
agregan. La versión diseñada, con los cuestionarios enviados, se conserva en la carpeta administrativa.

Salidas en evaluaciones/pnig_2026/05_entrevistas/:
  - ENTREV_preguntas.csv: catálogo de las preguntas formuladas (lo lee sintesis_triangulacion.py);
  - ENTREV_guia_DGIGND.docx, ENTREV_guia_OGPPM.docx, ENTREV_guia_OGMEPGD.docx: guía aplicada de cada informante, con
    la pregunta de la guía, cómo se formuló, el minuto y la evidencia de contraste;
  - ENTREV_cuestionario_DGIGND.docx, …_OGPPM.docx, …_OGMEPGD.docx: cuestionario aplicado, con las preguntas en el orden
    de la entrevista y su redacción real, y los documentos ofrecidos;
  - ENTREV_registro_codificacion.xlsx: registro codificado (producto del Anexo 4). Se rehace en cada corrida; lo que el
    equipo evaluador anota en las columnas amarillas («validacion», «observacion_evaluador») se conserva por
    id_segmento.

Entradas de la aplicación (ADR-007 y ADR-008), escritas a mano sobre las transcripciones anonimizadas:
  - ENTREV_aplicacion.csv: cada pregunta formulada a cada informante (literal, reformulada o en parte) y las surgidas
    (S-<órgano>-NN), con su minuto y su redacción real;
  - ENTREV_codificacion.csv: segmentos de las respuestas, sin nombres, con códigos del libro común, sentido, hallazgo
    que contrastan y resultado del contraste. Primera codificación: Claude (1.ª); valida el usuario;
  - ENTREV_deducciones.csv: lo que el equipo evaluador deduce de las respuestas para temas que no se preguntaron
    completos, con los segmentos que lo sustentan (decisión del usuario, 2026-11-06).

Planos: el Anexo 4 (bloques, preguntas orientadoras, duración, informantes) es de la nota; la pregunta sobre las
recomendaciones de evaluaciones anteriores viene de la Tabla A7.1 del CEPLAN (p3.4 a p3.6, Guía de seguimiento y
evaluación, p. 185); las preguntas de profundización, la prioridad y la evidencia de contraste son propuesta del equipo
evaluador. También lo es el bloque 4 adaptado a la función de cada órgano (ADR-006): la DGIGND responde sobre el
PEMIG, la OGPPM sobre el planeamiento y la coordinación con el CEPLAN y la OGMEPGD sobre el seguimiento y la
evaluación; las dos orientadoras genéricas del bloque 4 y «¿Qué recomendación priorizaría?» pasan a repreguntas B.
Datos personales (Ley N.° 29733): ni las guías ni el registro llevan nombres. Las grabaciones, las notas en bruto y la
relación entre el código de informante y la persona se guardan en la carpeta administrativa, fuera del repositorio.

Uso (desde la raíz del repositorio):
    uv run --no-project --with pandas --with openpyxl --with python-docx python src/analisis_pnig/entrev_guias_entrevista.py
"""
import os
import re
import sys

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, os.path.dirname(__file__))
import sintesis_hallazgos as S  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx_mimp import Documento, lista_es, soles  # noqa: E402
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

SALIDA = f"{EVAL}/05_entrevistas"
VALORACION_REC = f"{EVAL}/04_analisis/seguimiento_recomendaciones/valoracion_recomendaciones.csv"
TODOS = ("DGIGND", "OGPPM", "OGMEPGD")
INFORMANTE = {
    "DGIGND": "Profesional designado por la DGIGND, por su rol en la conducción técnica de la Política",
    "OGPPM": "Profesional designado por la OGPPM, por su rol en la programación presupuestal y el planeamiento",
    "OGMEPGD": "Especialista de seguimiento designado internamente por la OGMEPGD",
}
BLOQUES = {  # Anexo 4: contenido, pregunta de evaluación e informantes
    0: ("Apertura", "Rol de la persona entrevistada y de su órgano; consentimiento informado.", "Todas", TODOS),
    1: ("Marco conceptual", "Marco conceptual del enfoque de igualdad de oportunidades e implicancias sobre el diseño "
        "de la Política.", "D7", ("DGIGND",)),
    2: ("Rectoría y articulación",
        {"DGIGND": "Rectoría del MIMP, funcionamiento de la CONAIG y mecanismos de articulación.",
         "OGPPM": "Rectoría del MIMP, papel del PEMIG en la articulación y mecanismos de articulación.",
         "OGMEPGD": "Rectoría del MIMP, papel del PEMIG en la articulación y mecanismos de articulación."},
        "I4", TODOS),  # Anexo 4; la CONAIG solo con la DGIGND (ADR-006)
    3: ("Recursos de la conducción", "Recursos asignados a la conducción de la Política y su programación.", "I3",
        ("OGPPM", "DGIGND")),
    4: ({"DGIGND": "Implementación y PEMIG", "OGPPM": "Planeamiento y coordinación con el CEPLAN",
         "OGMEPGD": "Seguimiento y evaluación"},
        {"DGIGND": "Uso del PEMIG en la implementación de los servicios y su destino en la nueva política; "
                   "capacidades y sostenibilidad ante la transición.",
         "OGPPM": "Articulación con el planeamiento institucional, coordinación con el CEPLAN y condiciones para "
                  "formular la nueva política; capacidades y sostenibilidad ante la transición.",
         "OGMEPGD": "Limitaciones del seguimiento y de la evaluación y condiciones para la nueva política; capacidades "
                    "y sostenibilidad ante la transición."},
        "I5", TODOS),  # Anexo 4: «Capacidades y sostenibilidad»; adaptado a la función de cada órgano (ADR-006)
    5: ("Lecciones y temas vigentes", "Lecciones aprendidas y temas vigentes para la nueva política.", "Transversal",
        TODOS),
}
DIRIGIDO = {  # destinatario del cuestionario (nombres de los órganos como en la nota metodológica)
    "DGIGND": "profesional designado por la Dirección General de Igualdad de Género y No Discriminación (DGIGND)",
    "OGPPM": "profesional designado por la Oficina General de Planeamiento, Presupuesto y Modernización (OGPPM)",
    "OGMEPGD": "especialista de seguimiento designado por la OGMEPGD",
}
FECHA = {  # memorándums de la OGMEPGD a la DGIGND y a la OGPPM (carpeta administrativa, memorandos_evaluacion)
    "DGIGND": "Miércoles 4 de noviembre de 2026, 10:00 h",
    "OGPPM": "Jueves 5 de noviembre de 2026, 10:00 h",
    "OGMEPGD": "Por coordinar en la OGMEPGD",
}
DOCUMENTOS = {  # documentos que el cuestionario pide tener a mano o remitir (propuesta del equipo evaluador)
    "DGIGND": (
        "Informe técnico del equipo de trabajo que valoró el alcance de la actualización de la PNIG y su plan de "
        "trabajo, si los hubiera.",
        "Protocolos de servicios aprobados o en trámite desde 2024 y documentos que expliquen los que no se aprobaron.",
        "Cronograma de protocolos del PEMIG y registros de su seguimiento, si los hubiera.",
        "Acuerdos de la CONAIG y documentos de su seguimiento.",
    ),
    "OGPPM": (
        "Documentos de programación de la conducción de la PNIG (actividades operativas y metas presupuestales).",
        "Documentos que vinculan la PNIG y el PEMIG con el PEI y el POI del MIMP.",
        "Comunicaciones con el CEPLAN sobre el seguimiento, la evaluación, la modificación de los indicadores y la "
        "actualización de la PNIG.",
        "Documentos de las gestiones con el INEI, SERVIR u otras entidades sobre la información de los indicadores.",
    ),
    "OGMEPGD": (
        "Matriz de compromisos de la evaluación de 2023 y registros de su seguimiento.",
        "Informe técnico del equipo de trabajo que valoró el alcance de la actualización de la PNIG y su plan de "
        "trabajo, si los hubiera.",
        "Reportes o registros de seguimiento de los indicadores del PEMIG, si los hubiera.",
        "Documentos sobre la revisión de evidencia del modelo conceptual y sobre las coordinaciones con el INEI y SERVIR "
        "para la información de los indicadores.",
    ),
}
MINUTOS = {  # propuesta del equipo evaluador dentro de los 45 a 60 minutos del Anexo 4
    "DGIGND": {0: 5, 1: 14, 2: 15, 3: 8, 4: 11, 5: 6},
    "OGPPM": {0: 5, 2: 11, 3: 17, 4: 16, 5: 6},
    "OGMEPGD": {0: 5, 2: 20, 4: 20, 5: 6},
}

APLICACION = f"{SALIDA}/ENTREV_aplicacion.csv"
CODIFICACION = f"{SALIDA}/ENTREV_codificacion.csv"
DEDUCCIONES = f"{SALIDA}/ENTREV_deducciones.csv"
TRANSCRIPCION = SALIDA + "/transcripciones/ENTREV_transcripcion_{}.txt"
APLICADA = {  # entrevistas realizadas (grabaciones de Meet en la carpeta administrativa; ADR-007)
    "DGIGND": {"fecha": "2026-11-05", "inicio": "15:04", "termino": "16:29", "duracion": "1:24:59",
               "modalidad": "Grupal: tres especialistas de la DPIGND designadas por la DGIGND, en una sola cuenta",
               "acepta": "Sí, verbal («Okay») tras el aviso de participación voluntaria (00:01)",
               "grabacion": "Informada al inicio como de uso interno; sin objeción (00:01)"},
    "OGPPM": {"fecha": "2026-11-04", "inicio": "16:01", "termino": "16:35", "duracion": "0:34:07",
              "modalidad": "Individual, con dos entrevistadores del equipo evaluador",
              "acepta": "Sí, verbal (00:57), tras el aviso de participación voluntaria y de la Ley N.° 29733 (00:00)",
              "grabacion": "Sí, verbal (00:57); se pidió el consentimiento para grabar (00:00)"},
    "OGMEPGD": {"fecha": "2026-11-05", "inicio": "11:56", "termino": "12:20", "duracion": "0:24:38",
                "modalidad": "Individual",
                "acepta": "Sí, verbal (00:02), tras el aviso de protección de datos",
                "grabacion": "Informada al inicio, sin difusión de la grabación; sin objeción (00:02)"},
}
FORMULACION = ("Literal", "Reformulada", "En parte", "Surgida")  # solo lo que se preguntó (ADR-008)
SENTIDOS = ("Favorable", "Dificultad", "Mixto", "Propuesta", "Neutro")
RESULTADOS = ("Coincide", "Matiza", "Discrepa", "Aporta información nueva")
COLS_REGISTRO = ["id_segmento", "informante", "fecha", "minuto", "bloque", "pregunta_guia", "pregunta_eval",
                 "tipo_texto", "texto", "codigos", "sentido", "hallazgo_contrastado", "resultado_contraste",
                 "hallazgo_critico", "documento_ofrecido", "codificador", "validacion", "observacion_evaluador", "nota"]
EVALUADOR_REG = ("validacion", "observacion_evaluador")  # columnas amarillas que se conservan entre corridas
MARCAS = r"\s*\[(?:audio|Pinpoint|¿)[^\]]*\]"  # marcas de la transcripción que no forman parte de lo dicho

PRIORIDAD_A_MAC5 = ("P01", "C01")  # problema público y enfoque de género: divergencia con efecto «se rediseña»
TIPO_A71 = "Tabla A7.1 del CEPLAN, p3.4 a p3.6"  # pregunta orientadora que no está en el Anexo 4
TIPO_FUNCION = "Orientadora por función"  # bloque 4 adaptado a la función de cada órgano (equipo evaluador, ADR-006)
ORIENTADORAS = ("Anexo 4", "Anexo 4 (apertura)", TIPO_A71, TIPO_FUNCION)  # las únicas que van en el cuestionario
PASA_A_REPREGUNTA = "Pregunta orientadora del Anexo 4 que se formula como repregunta (ADR-006)."


def de_bloque(b, i, inf):
    """Nombre (i=0) o contenido (i=1) del bloque; el bloque 4 varía según el informante."""
    v = BLOQUES[b][i]
    return v[inf] if isinstance(v, dict) else v


def texto_de(r, inf):
    """Texto que ve el informante: la pregunta de la Tabla A7.1 nombra a su órgano."""
    return r.pregunta.replace("su órgano", f"la {inf}") if r.tipo == TIPO_A71 else r.pregunta


def prioridad_de(r, inf):
    """Prioridad de una pregunta para un informante («A», «B» o «DGIGND: A; OGMEPGD: B»)."""
    if ":" not in r.prioridad:
        return r.prioridad
    return dict(x.split(": ") for x in r.prioridad.split("; "))[inf]


def preguntas(C):
    """Catálogo de preguntas. tipo: «Anexo 4» (texto de la nota) o «Profundización» (equipo evaluador)."""
    dev, cs, vd = C["cond_dev"], C["cond_suma"], C["act5001072"]
    Q = []

    def q(bloque, texto, tipo="Profundización", inf=None, prioridad="A", preg=None, codigos="", contrasta="",
          evidencia=""):
        n = sum(1 for x in Q if x["bloque"] == bloque) + 1
        if isinstance(prioridad, dict):  # prioridad distinta según el informante
            prioridad = "; ".join(f"{k}: {v}" for k, v in prioridad.items())
        Q.append({"id": f"P{bloque}.{n}", "bloque": bloque, "tipo": tipo, "informantes": ";".join(inf or
                  BLOQUES[bloque][3]), "prioridad": prioridad, "pregunta_eval": preg or BLOQUES[bloque][2],
                  "pregunta": texto, "codigos": codigos, "contrasta": contrasta, "evidencia": evidencia})

    # Bloque 0
    q(0, "¿Cuál es su rol y el de su órgano en la conducción, la programación o el seguimiento de la PNIG, y desde "
      "cuándo participa en esas funciones?", tipo="Anexo 4 (apertura)", preg="Todas",
      evidencia="Permite ponderar el testimonio según el periodo y la función de la persona.")

    # Bloque 1 (DGIGND, D7)
    for t in ("¿Qué cambios introduce el enfoque de igualdad de oportunidades respecto del enfoque con el que se "
              "formuló la Política?",
              "¿Qué elementos del diagnóstico y de los objetivos prioritarios se mantienen vigentes y cuáles deben "
              "redefinirse en la nueva política?",
              "¿Cómo afecta ese marco a la definición de los servicios y de su población receptora?"):
        q(1, t, tipo="Anexo 4", codigos="D7.IMPL", contrasta="D7-H1; D1-H4")
    ent = C["mac5_entrevista"]
    for _, r in ent.iterrows():
        q(1, str(r["Para contrastar con la DGIGND (ENTREV)"]).strip(),
          prioridad="A" if r["Código"] in PRIORIDAD_A_MAC5 else "B", codigos="D7.IMPL",
          contrasta="D7-H2" + ("; D1-H4" if r["Código"] == "P01" else ""),
          evidencia=f"Elemento {r['Código']}, «{r['Elemento de la PNIG']}». Relación: {r['Tipo de relación'].lower()}; "
                    f"efecto: {r['Efecto'].lower()}. Fuente: REV_ALI.5 ({C.origen['mac5_entrevista']}).")
    m5 = C["mac5"]
    q(1, f"La matriz preliminar clasifica como «se rediseña» los servicios {lista_es(C['mac5_serv_rediseña'])} y "
      f"como «se ajusta» a {m5.loc['Servicios', 'Se ajusta']} servicios. ¿Comparte esa lectura? ¿Qué otros servicios "
      f"considera que deben rediseñarse?", codigos="D7.IMPL", contrasta="D7-H1",
      evidencia=f"REV_ALI.5: {m5.loc['Servicios', 'Se mantiene']} servicios se mantienen, "
                f"{m5.loc['Servicios', 'Se ajusta']} se ajustan y {m5.loc['Servicios', 'Se rediseña']} se rediseñan "
                f"({C.origen['mac5']}).")

    # Bloque 2 (todos, I4). Lo de la CONAIG solo a la DGIGND; la OGPPM y la OGMEPGD responden sobre el PEMIG (ADR-006)
    a4 = dict(tipo="Anexo 4", codigos="I4.RECT; I4.CONAIG; I4.COORD", contrasta="I4-H2; I4-H4")
    q(2, "¿Cómo valora la rectoría técnica del MIMP en la conducción de la Política?", **a4)
    q(2, "¿Con qué regularidad sesionó la CONAIG y qué acuerdos se tradujeron en decisiones efectivas?",
      inf=("DGIGND",), **a4)
    q(2, "¿Qué papel cumplió el PEMIG en la articulación de las entidades responsables de los servicios y qué "
      "limitaciones tuvo como instrumento de implementación de la Política?", tipo=TIPO_FUNCION,
      inf=("OGPPM", "OGMEPGD"), codigos="I4.COORD; I4.RECT", contrasta="I4-H4; D3-H2",
      evidencia="Pregunta orientadora por función del órgano (equipo evaluador; ADR-006). Reemplaza, para la OGPPM y "
                "la OGMEPGD, la pregunta sobre la CONAIG.")
    q(2, "¿Qué mecanismos de articulación funcionan en la práctica y cuáles existen solo formalmente?", **a4)
    q(2, "¿Qué medidas adoptó su órgano ante las recomendaciones de las evaluaciones de la PNIG de 2021 y 2023 y del "
      "informe técnico de evaluación del CEPLAN, en qué estado se encuentran y qué haría falta para completarlas o "
      "trasladarlas a la nueva política?", tipo=TIPO_A71, codigos="I4.RECT", contrasta="I4-H5",
      evidencia="Seguimiento de las recomendaciones de evaluaciones anteriores (sección 3.7 del informe).")
    q(2, "Los informes LIO registran tres sesiones ordinarias de la CONAIG en 2024 y en 2025, con participación "
      "mayoritaria de representantes alternos y una asistencia de 57,1 %, 42,9 % y 90,5 % en las sesiones de 2025. "
      "¿Qué efecto tuvo el nivel de representación en la adopción y el cumplimiento de los acuerdos?",
      inf=("DGIGND",), codigos="I4.CONAIG", contrasta="I4-H2",
      evidencia=f"{S.seg(C, 'LIO-2024-6.4.1-07')}; {S.seg(C, 'LIO-2025-6.3.3-11')}.")
    q(2, "Según el III Informe de gobernanza, solo 2 de 16 servicios incorporaron en los planes de sus mecanismos "
      "para la igualdad de género las acciones de seguimiento recomendadas, y el LIO 2025 reporta 4 de 12 medidas "
      "urgentes cumplidas. ¿Con qué medios contaba la CONAIG para acompañar o exigir el cumplimiento de sus "
      "recomendaciones?", inf=("DGIGND",), codigos="I4.CONAIG; I4.RECT", contrasta="I4-H3; I4-H2",
      evidencia=f"{S.seg(C, 'CONAIG-2024-040')}; {S.seg(C, 'LIO-2025-6.3.3-14')}.")
    q(2, "Algunas entidades respondieron que las recomendaciones de la CONAIG no les eran aplicables porque prestan sus "
      "servicios sin diferencias por sexo, como COFOPRI en el servicio 4.2.2. ¿Cómo se procesaron esas respuestas?",
      inf=("DGIGND",), prioridad="B", codigos="I4.CONAIG; I5.AJSERV", contrasta="I4-H3; I5-H3",
      evidencia=f"{S.seg(C, 'CONAIG-2023-022')}; {S.seg(C, 'CONAIG-2024-042')}.")
    q(2, "El LIO 2025 indica que el 38,7 % de los servicios reportó coordinación con entidades del nivel nacional, "
      "concentrada en el OP4, y que los servicios del OP2 no reportaron articulación nacional. ¿Qué explica la falta "
      "de articulación en salud sexual y reproductiva?", inf=("DGIGND", "OGMEPGD"), prioridad="B",
      codigos="I4.COORD; I4.DIFC",
      contrasta="I4-H4", evidencia=f"{S.seg(C, 'LIO-2025-6.3.3-21')}.")
    q(2, f"La revisión de las políticas nacionales afines identificó {C['hor_dup']} relaciones de duplicidad; el "
      f"servicio de los Centros Emergencia Mujer figura en {C['hor_cem']} políticas además de la PNIG. ¿Cómo se "
      f"coordinaban el reporte y el seguimiento de un mismo servicio registrado en varias políticas, y qué conviene "
      f"para la nueva política?", prioridad={"OGMEPGD": "A", "DGIGND": "B", "OGPPM": "B"},
      codigos="I4.COORD; I4.DIFC", contrasta="D4-H3; I4-H4",
      evidencia=f"REV_ALI.3: {C['hor_rel']} relaciones con {C['hor_pol']} políticas ({C.origen['hor_rel']}).")
    q(2, "En el formulario de seguimiento 2025, el MIMP señaló la inexistencia de convenios o protocolos de "
      "articulación y la duplicidad de funciones como causas de dificultad del indicador del OP4. ¿Qué instrumentos "
      "de articulación faltaron?", inf=("DGIGND", "OGMEPGD"), prioridad="B", codigos="I4.DIFC",
      contrasta="D4-H3", evidencia=f"{S.seg(C, 'FORM-2025-IOP.04.01-causa_dificultad')}.")
    q(2, "El LIO 2025 registra que el 30,6 % de los servicios coordinó con gobiernos regionales y locales, sobre todo "
      "en el OP4. ¿Qué papel tuvieron los mecanismos regionales y locales para la igualdad de género en la "
      "implementación?", prioridad="B", codigos="I4.SUBN", contrasta="I4-H4",
      evidencia=f"{S.seg(C, 'LIO-2025-6.3.3-21')}.")
    # Recomendaciones de evaluaciones anteriores (sección 3.7; seguimiento_recomendaciones_previas.xlsx)
    q(2, "El reporte de cumplimiento 2024 indica que el equipo de trabajo recomendado por la evaluación de 2023 "
      "concluyó que la actualización de la PNIG debía empezar en el Entregable 3. ¿Se elaboró el plan de trabajo que "
      "pedía la recomendación? ¿Qué parte de ese análisis puede aprovechar la nueva política?",
      inf=("DGIGND", "OGMEPGD"), codigos="I4.RECT", contrasta="I4-H5",
      evidencia="Reporte de cumplimiento de la PNIG 2024, p. 13; Informe N.° D000039-2026-MIMP-DPIGND, p. 11 "
                "(recomendación EV2023-02).")
    q(2, "El Informe N.° D000039-2026-MIMP-DPIGND señala que la matriz de compromisos de la PNIG se elaboró por única "
      "vez en 2023. ¿Qué compromisos contenía, quién registró su cumplimiento y se elaboró alguna para las "
      "recomendaciones del informe técnico del CEPLAN?", inf=("OGMEPGD",), codigos="I4.RECT", contrasta="I4-H5",
      evidencia="Informe N.° D000039-2026-MIMP-DPIGND, p. 11; Guía de seguimiento y evaluación, Tabla 6, p. 65.")
    q(2, "La evaluación de 2023 recomendó a la OMEP revisar la evidencia sobre el problema público y proponer ajustes al "
      "modelo conceptual en el primer trimestre de 2025. ¿Se hizo esa revisión? ¿En qué documento consta?",
      inf=("OGMEPGD", "DGIGND"), prioridad={"OGMEPGD": "A", "DGIGND": "B"}, codigos="I4.RECT", contrasta="I4-H5",
      evidencia="Informe de evaluación de la PNIG 2023 (MIMP, 2024), p. 66 (recomendación EV2023-01); no consta un "
                "documento de esa revisión en el corpus.")
    q(2, "Después del informe técnico del CEPLAN se aprobó el protocolo del servicio 1.2.6 (diciembre de 2025) y hay "
      "propuestas de protocolo para los servicios 1.1.4 y 1.2.4; los servicios 1.1.1, 1.1.6 y 1.1.9 siguen sin "
      "protocolo y no se reformularon los indicadores de los demás servicios del OP1. ¿Cómo se trasladaron esas "
      "recomendaciones a las entidades y "
      "qué asistencia técnica recibieron para incorporar los estándares de calidad recomendados?",
      inf=("DGIGND", "OGMEPGD"), prioridad={"DGIGND": "A", "OGMEPGD": "B"}, codigos="I4.RECT; I1.NORM",
      contrasta="I4-H5", evidencia="REV_DOC; Informe LIO 2025, pp. 40, 64 y 66; Informe N.° D000112-2026-MIMP-OP, "
                                   "p. 3 (recomendaciones CEPLAN-10, CEPLAN-20 y CEPLAN-62).")
    q(2, "La evaluación de 2021 recomendó coordinar con el INEI y SERVIR para cerrar la brecha de información de los "
      "indicadores; en 2025 se recopiló información de 6 de 13 indicadores de los objetivos prioritarios. ¿Qué "
      "coordinaciones se hicieron y con qué resultado?", inf=("OGMEPGD", "OGPPM"),
      prioridad={"OGMEPGD": "A", "OGPPM": "B"}, codigos="I4.RECT; I4.COORD", contrasta="I4-H5",
      evidencia="Informe de evaluación de resultados de la PNIG 2021 (MIMP, 2022), p. 42 del PDF (recomendación "
                "EV2021-01); Informe N.° D000039-2026-MIMP-DPIGND, p. 11.")
    q(2, "El informe técnico del CEPLAN recomendó, para servicios del Programa Nacional Warmi Ñan (AURORA en el "
      "informe), capacitar al personal de los CEM para zonas rurales y lenguas originarias, mejorar el SIRA y "
      "capacitar en su uso, articular las intervenciones de la estrategia comunitaria y reducir los tiempos "
      "administrativos. ¿Qué medidas conoce sobre esos puntos?", inf=("DGIGND",), prioridad="B",
      codigos="I4.RECT; I1.RRHH", contrasta="I4-H5",
      evidencia="Informe técnico del CEPLAN, pp. 45, 61, 62 y 67 (recomendaciones CEPLAN-26, CEPLAN-45 a CEPLAN-47 y CEPLAN-59); no "
                "constan en el LIO 2025 ni en el formulario 2025.")

    # Bloque 3 (OGPPM y DGIGND, I3)
    for t in ("¿Con qué recursos y en qué actividades presupuestales se financió la conducción de la Política?",
              "¿Qué limitaciones presupuestales afectaron la asistencia técnica y el seguimiento?"):
        q(3, t, tipo="Anexo 4", codigos="I3.REC; I1.PRES", contrasta="I3-H1; I3-H3")
    q(3, f"Los datos abiertos del MEF muestran que el devengado de la meta de la DGIGND pasó de "
      f"{soles(dev[2022]['DGIGND'])} en 2022 a {soles(dev[2023]['DGIGND'])} en 2023, año en que aparecen metas "
      f"propias de la DPIGND y de la DPPDM. Sumadas las tres metas, la conducción pasó de {soles(cs[2022])} a "
      f"{soles(cs[2023])}. ¿La reducción responde solo a la reorganización o también a un cambio en los recursos de "
      f"la conducción?", codigos="I3.REC", contrasta="I3-H1",
      evidencia=f"REV_IND.3 ({C.origen['cond_dev']}).")
    q(3, "El POI no indica la actividad presupuestal de cada actividad operativa, por lo que no es posible cuadrar el "
      "POI de la conducción con la ejecución del MEF. ¿Cómo se vincula, en la programación, la actividad operativa "
      "de la conducción con su meta presupuestal?", inf=("OGPPM",), codigos="I3.REC", contrasta="I3-H1",
      evidencia="REV_IND.3 (REV_IND_3_recursos_conduccion.xlsx, hoja «Cuadre POI»).")
    q(3, f"La actividad 5001072 «Promoción de la mujer» pasó de un PIM de {soles(vd.loc[2019, 'pim'])} en 2019 a "
      f"{soles(vd.loc[2025, 'pim'])} en 2025. ¿Qué parte de esa actividad financia la conducción de la PNIG y qué "
      f"parte otras funciones?", inf=("OGPPM",), prioridad="B", codigos="I3.REC", contrasta="I3-H2",
      evidencia=f"REV_IND.3 ({C.origen['act5001072']}).")
    q(3, f"De {C['mef_total']} montos que las entidades reportaron para sus servicios, {C['mef_coinc']} coinciden con "
      f"una meta, actividad o producto de los datos abiertos del MEF y {C['mef_sinc']} no coinciden. ¿Existe o se "
      f"prevé un mecanismo para identificar el gasto de los servicios de una política nacional?", inf=("OGPPM",),
      prioridad="B", codigos="I3.REC; I1.PRES", contrasta="I1-H2",
      evidencia=f"REV_IND.3 ({C.origen['mef_coinc']}).")
    q(3, "El MIMP reportó en 2025 que la falta de presupuesto impide aplicar la ENARES cada dos años, y esa encuesta "
      "es la fuente del indicador del OP6. ¿Cómo se programó el financiamiento de las fuentes de información de la "
      "Política?", codigos="I3.REC; I5.SOST", contrasta="I3-H3; D5-H3",
      evidencia=f"{S.seg(C, 'FORM-2025-IOP.06.01-contexto')}; {S.seg(C, 'LIO-2023-5.5.1-05')}.")
    q(3, "Las limitaciones presupuestales llevaron a priorizar la asistencia técnica virtual y a reemplazar las visitas "
      "presenciales de seguimiento de lactarios por un formulario virtual. ¿Cómo afectó eso la asistencia técnica a "
      "las entidades responsables?", inf=("DGIGND",), prioridad="B", codigos="I3.REC; I1.PRES",
      contrasta="I3-H3", evidencia=f"{S.seg(C, 'LIO-2025-5.1.1-03')}; {S.seg(C, 'CONAIG-2024-035')}.")

    # Bloque 4 (todos, I5): orientadoras por función de cada órgano (ADR-006)
    fun = "Pregunta orientadora por función del órgano (equipo evaluador; ADR-006)."
    q(4, "¿Cómo se usó el PEMIG para organizar con las entidades responsables la implementación de los servicios, la "
      "aprobación de sus protocolos y su seguimiento, y qué limitaciones tuvo?", tipo=TIPO_FUNCION, inf=("DGIGND",),
      preg="I4", codigos="I4.RECT; I4.COORD; I1.NORM", contrasta="I1-H1; D3-H2",
      evidencia=f"{fun} Conducción técnica de la Política por la DGIGND.")
    q(4, "¿Qué elementos del PEMIG (servicios, entidades responsables, cronograma de protocolos, indicadores) conviene "
      "conservar o cambiar en el instrumento que implemente la nueva política?", tipo=TIPO_FUNCION, inf=("DGIGND",),
      codigos="I5.SOST; I5.AJSERV; I5.AJIND", contrasta="D3-H2; I5-H2", evidencia=fun)
    q(4, "¿Cómo se articularon la PNIG y el PEMIG con el planeamiento institucional del MIMP (PESEM, PEI y POI) y con "
      "la programación presupuestal, y qué dificultades se presentaron?", tipo=TIPO_FUNCION, inf=("OGPPM",),
      preg="I3", codigos="I3.REC; I4.COORD", contrasta="I3-H1",
      evidencia=f"{fun} Planeamiento y presupuesto institucional a cargo de la OGPPM.")
    q(4, "Como órgano que coordina con el CEPLAN, ¿cómo fue esa coordinación para el seguimiento, la evaluación, la "
      "modificación de los indicadores y la actualización de la PNIG, y qué dificultades se presentaron?",
      tipo=TIPO_FUNCION, inf=("OGPPM",), preg="I4", codigos="I4.COORD; I4.DIFC; I5.AJIND",
      contrasta="I4-H5; D5-H2", evidencia=f"{fun} La OGPPM es el punto focal ante el CEPLAN (Oficios N.° "
                                         "D000054-2026-MIMP-OGPPM y D000289-2026-MIMP-OGPPM). Aporta además "
                                         "contraste a D5.")
    q(4, "¿Qué condiciones de planeamiento hacen falta para formular la nueva política en el marco de la Ley N.° 32535 "
      "y de las disposiciones del CEPLAN?", tipo=TIPO_FUNCION, inf=("OGPPM",), codigos="I5.RIESGO; I5.SOST",
      contrasta="I5-H2", evidencia=fun)
    q(4, "¿Qué limitaciones enfrentó la OGMEPGD para el seguimiento de la PNIG y del PEMIG (información que reportan las "
      "entidades, fuentes y fichas de los indicadores, herramientas, plazos, equipo) y qué medidas se tomaron frente a "
      "ellas?", tipo=TIPO_FUNCION, inf=("OGMEPGD",),
      codigos="I5.CAPI; I1.INFO; D5.FUENTE; D5.FICHA; D5.REPORTE", contrasta="I2-H1; D5-H2",
      evidencia=f"{fun} Aporta además contraste a D5.")
    q(4, "¿Qué limitaciones tuvieron las evaluaciones de la PNIG de 2021 y 2023 (información, plazos, alcance) y cómo se "
      "usaron sus resultados en la gestión de la Política?", tipo=TIPO_FUNCION, inf=("OGMEPGD",), preg="I4",
      codigos="I4.RECT", contrasta="I4-H5", evidencia=fun)
    q(4, "¿Qué condiciones de seguimiento y evaluación deben asegurarse desde el diseño de la nueva política?",
      tipo=TIPO_FUNCION, inf=("OGMEPGD",), codigos="I5.SOST; I5.AJIND", contrasta="I5-H1", evidencia=fun)
    for t in ("¿Con qué capacidades cuenta el MIMP para sostener la conducción durante la transición?",
              "¿Qué riesgos identifica para la continuidad de los avances?"):
        q(4, t, prioridad="B", codigos="I5.CAPI; I5.RIESGO; I5.SOST", contrasta="I5-H1; I5-H2",
          evidencia=PASA_A_REPREGUNTA)
    q(4, "El PEMIG dispuso que las entidades formularan o adecuaran los protocolos de sus servicios, con la asistencia "
      f"técnica del MIMP, según un cronograma que forma parte del Plan. Al 31 de diciembre de 2025, "
      f"{C['serv_prot_si']} de los {C['serv_total']} servicios tenían protocolo aprobado. ¿Qué seguimiento tuvo ese "
      f"cronograma y qué explica que los demás servicios no tengan protocolo aprobado?", inf=("DGIGND",),
      preg="I4", codigos="I1.NORM; I4.RECT", contrasta="I1-H1",
      evidencia=f"Informe N.° D000039-2026-MIMP-DPIGND, p. 8, numerales 2.3.2 y 2.3.3; REV_DOC "
                f"({C.origen['serv_prot_si']}).")
    q(4, f"El informe técnico del CEPLAN recomendó seguir los servicios 1.1.1, 1.1.4, 1.1.7 y 1.1.9 con el indicador de "
      f"cobertura del PEMIG, y {C['serv_sin_ind_vig']} de los {C['serv_total']} servicios no tienen indicador "
      f"vigente. ¿Se reportaron y siguieron los indicadores del PEMIG? ¿En qué documentos constan?",
      inf=("OGMEPGD", "DGIGND"), prioridad={"OGMEPGD": "A", "DGIGND": "B"}, preg="I4", codigos="I4.RECT; I5.AJIND",
      contrasta="D3-H2; I4-H5",
      evidencia=f"Informe técnico del CEPLAN, pp. 22, 31, 37 y 42 (recomendaciones CEPLAN-03, CEPLAN-11, CEPLAN-19 y CEPLAN-24); "
                f"REV_ALI.1 ({C.origen['serv_sin_ind_vig']}).")
    q(4, "El Oficio N.° D000289-2026-MIMP-OGPPM comunicó al CEPLAN que, de las cuatro políticas que conduce el MIMP, solo "
      "los indicadores de la PNIG se actualizarían mediante la actualización de la Política, y los de las otras tres "
      "mediante una propuesta de modificación aprobada por resolución ministerial, con opinión favorable del CEPLAN. "
      "¿Qué llevó a elegir esa ruta para la PNIG y cómo se ajusta a la sustitución que dispone la Ley N.° 32535?",
      inf=("OGPPM",), preg="I4", codigos="I4.COORD; I5.AJIND; I5.RIESGO", contrasta="I4-H5; I5-H2",
      evidencia="Oficio N.° D000289-2026-MIMP-OGPPM (20 de julio de 2026), p. 1; Informe N.° D000112-2026-MIMP-OP, "
                "p. 3.")
    q(4, "El Informe N.° 000006-2026 del CEPLAN, que respondió a la solicitud de asistencia técnica de la OGPPM sobre la "
      "Ley N.° 32535, concluyó que la «sustitución» no es un proceso previsto en el Sinaplan y que el MIMP debe evaluar "
      "la PNIG para decidir entre formular una nueva política o actualizar la vigente. ¿Cómo se recogieron esas "
      "orientaciones en la ruta que siguió el MIMP?", inf=("OGPPM",), prioridad="B", preg="I4",
      codigos="I4.COORD; I5.RIESGO", contrasta="I5-H2",
      evidencia="Informe N.° 000006-2026 del CEPLAN (10 de marzo de 2026), p. 9, numerales 3.2 y 3.5.")
    q(4, "Los informes LIO de 2021, 2023 y 2025 señalan la rotación de personal como causa de pérdida de capacidades en "
      "los mecanismos para la igualdad de género y de dificultades de la asistencia técnica en transversalización. "
      "¿Qué medidas existen o harían falta para retener esas capacidades?", codigos="I5.CAPI; I1.RRHH",
      contrasta="I5-H1",
      evidencia=f"{S.seg(C, 'LIO-2021-5.5.1-03')}; {S.seg(C, 'LIO-2023-5.3.1-03')}; {S.seg(C, 'LIO-2025-5.3.1-04')}.")
    q(4, "El LIO 2025 registra que un protocolo de servicio quedó a la espera de las acciones derivadas de la Ley N.° "
      "32535 y que la incertidumbre del proyecto de nueva ley fue la principal dificultad de la asistencia técnica "
      "en 2025. ¿Qué instrumentos quedaron en suspenso por la transición y qué destino se prevé para ellos?",
      inf=("DGIGND", "OGMEPGD"), codigos="I5.RIESGO", contrasta="I5-H2",
      evidencia=f"{S.seg(C, 'LIO-2025-6.3.3-02')}; {S.seg(C, 'LIO-2025-5.1.1-03')}.")
    q(4, f"La revisión documental clasifica como no implementados los servicios {lista_es(C['serv_noimpl_cod'])}, y "
      f"el servicio 6.4.1 seguía sin protocolo en 2025 después de estar en diseño desde 2022. ¿Qué explica que no se "
      f"pusieran en marcha y qué recomienda hacer con ellos?", inf=("DGIGND", "OGMEPGD"),
      prioridad={"DGIGND": "A", "OGMEPGD": "B"},
      codigos="I5.AJSERV; I1.NORM", contrasta="I1-H1",
      evidencia=f"REV_DOC ({C.origen['serv_noimpl_cod']}); {S.seg(C, 'LIO-2022-6.4.1-01')}; "
                f"{S.seg(C, 'LIO-2025-6.4.1-01')}.")
    q(4, "COFOPRI pidió en 2025 evaluar el retiro del servicio 4.2.2 de la PNIG y el Programa Juntos señaló desde 2019 "
      "que su contribución al servicio 4.4.1 es indirecta. ¿Cómo se procesaron las solicitudes de ajuste o de "
      "retiro de servicios durante el periodo?", inf=("DGIGND", "OGMEPGD"), prioridad="B", codigos="I5.AJSERV",
      contrasta="I5-H3",
      evidencia=f"{S.seg(C, 'FORM-2025-ISS.04.02.02.01-falla_abajo')}; {S.seg(C, 'LIO-2019-4.4.1-03')}.")
    q(4, f"De los {C['op_ind']} indicadores de los objetivos prioritarios, {C['op_sinvalor_2025']} no tienen valor "
      f"2025, y {C['serv_sin_ind_vig']} de los {C['serv_total']} servicios no tienen indicador vigente. ¿Qué "
      f"capacidades y fuentes de información de seguimiento conviene asegurar en la transición?",
      inf=("OGMEPGD", "DGIGND"), prioridad={"OGMEPGD": "A", "DGIGND": "B"}, codigos="I5.CAPI; I5.AJIND",
      contrasta="I2-H1; D3-H2",
      evidencia=f"REV_IND ({C.origen['resumen_semaforo']}); REV_ALI.1 ({C.origen['serv_sin_ind_vig']}).")
    q(4, "Las fichas técnicas se modificaron mediante actas de validación de 2022 y 2023, varias no ratificadas. ¿Qué "
      "procedimiento se siguió y qué lección deja para modificar los indicadores de la nueva política?",
      inf=("OGMEPGD", "OGPPM"), prioridad="B", preg="I5", codigos="I5.AJIND; D5.FICHA", contrasta="D5-H2",
      evidencia="REV_IND.2 (REV_IND_2_calidad_fichas.xlsx, hoja «Conjunto»). Aporta además contraste a D5.")
    q(4, "El MINCETUR reportó que el horario nocturno de las capacitaciones a artesanas recargó su jornada y afectó su "
      "descanso, lo que coincide con la consulta de la CONAIG en Piura. ¿Conoce otros efectos no previstos de los "
      "servicios?", prioridad="B", codigos="I5.EFNP", contrasta="I5-H4",
      evidencia=f"{S.seg(C, 'FORM-2025-ISS.04.03.05.01-efectos_no_previstos')}; {S.seg(C, 'CONAIG-2025-030')}.")

    # Bloque 5 (todos, transversal)
    for t in ("¿Qué avances de la PNIG deben preservarse en la nueva política?",
              "¿Qué lecciones de su diseño y de su implementación deberían orientarla?"):
        q(5, t, tipo="Anexo 4", codigos="I1.BP; I5.SOST; I5.AJSERV")
    q(5, "¿Qué recomendación priorizaría?", prioridad="B", codigos="I1.BP; I5.SOST; I5.AJSERV",
      evidencia=PASA_A_REPREGUNTA)
    lim = [t[1].lower() for t in C["rev_lim"]]
    q(5, f"La revisión de evidencia encontró respaldo sólido para {C['rev_sólida']} de {C['rev_tipos']} tipos de "
      f"servicio y respaldo limitado para {lista_es(lim)}. ¿Qué servicios considera prioritario conservar y cuáles "
      f"requerirían una evaluación antes de mantenerse?", codigos="I5.AJSERV; I1.BP", contrasta="D3-H3",
      evidencia=f"REV_EVI ({C.origen['rev_tipos']}).")
    q(5, "¿Conoce evaluaciones, estudios o sistematizaciones del MIMP o de las entidades responsables sobre los "
      "servicios de la PNIG que no estén publicados y que el equipo evaluador pueda revisar?", prioridad="B",
      codigos="EMER", contrasta="D3-H3", evidencia="Complementa la revisión de evidencia (REV_EVI) y la revisión "
      "documental (REV_DOC).")
    q(5, "¿Dónde se conserva el acervo de la PNIG (formularios, actas de validación, series de indicadores) y qué hace "
      "falta para que la nueva política lo aproveche?", inf=("OGMEPGD", "DGIGND"), prioridad="B",
      codigos="I5.SOST", evidencia="Recomendación de gestión del conocimiento (nota metodológica, numeral 3.1).")
    Q = pd.DataFrame(Q)
    Q["evidencia"] = Q.evidencia.map(S.corto)
    return Q


def _verificar(Q, C, H):
    ids_h = {h["id"] for h in H}
    libro = set(C["libro"].index)
    for _, r in Q.iterrows():
        for c in filter(None, (x.strip() for x in r.codigos.split(";"))):
            assert c in libro, f"{r.id}: código fuera del libro: {c}"
        for h in filter(None, (x.strip() for x in r.contrasta.split(";"))):
            assert h in ids_h, f"{r.id}: hallazgo inexistente: {h}"
        assert set(r.informantes.split(";")) <= set(BLOQUES[r.bloque][3]), f"{r.id}: informante fuera del bloque"
        for txt in (r.pregunta, r.evidencia):
            assert " — " not in txt and " - " not in txt, f"{r.id}: guion explicativo"
    for inf in TODOS:  # cada órgano tiene su bloque 4 propio (ADR-006)
        assert ((Q.tipo == TIPO_FUNCION) & (Q.informantes == inf)).sum() >= 2, f"{inf}: sin orientadoras por función"
    for inf, mins in MINUTOS.items():
        bloques = {b for b, (_, _, _, infs) in BLOQUES.items() if inf in infs}
        assert set(mins) == bloques, f"{inf}: bloques sin tiempo asignado"
        assert 45 <= sum(mins.values()) <= 60, f"{inf}: duración fuera de 45 a 60 minutos"


def _normal(t):
    return re.sub(r"\s+", " ", re.sub(MARCAS, "", t)).strip()


def aplicacion(Q_dis, C, H):
    """Lee y verifica la aplicación, la codificación y las deducciones (ADR-007 y ADR-008).

    Q_dis es el catálogo del diseño: la aplicación solo registra las preguntas que se formularon (literal,
    reformulada o en parte) y las surgidas."""
    A = pd.read_csv(APLICACION, dtype=str).fillna("")
    seg = pd.read_csv(CODIFICACION, dtype=str).fillna("")
    D = pd.read_csv(DEDUCCIONES, dtype=str).fillna("")
    q = Q_dis.set_index("id")
    ids_h = {h["id"] for h in H}
    preguntas_eval = {h["pregunta"] for h in H} | {"Todas", "Transversal"}
    libro = set(C["libro"].index)

    def codigos_ok(txt, donde):
        for c in filter(None, (x.strip() for x in txt.split(";"))):
            assert c in libro, f"{donde}: código fuera del libro: {c}"

    assert not A.duplicated(["informante", "id_pregunta"]).any(), "aplicación: pregunta repetida"
    assert set(A.formulacion) <= set(FORMULACION), f"aplicación: formulación no prevista {set(A.formulacion)}"
    for _, a in A.iterrows():
        donde = f"{a.informante} {a.id_pregunta}"
        assert (a.formulacion == "Surgida") == a.id_pregunta.startswith("PS-"), f"{donde}: surgida mal identificada"
        assert a.minuto and a.como_se_formulo, f"{donde}: falta el minuto o cómo se formuló"
        assert " — " not in a.como_se_formulo and " - " not in a.como_se_formulo, f"{donde}: guion explicativo"
        if a.formulacion == "Surgida":
            assert a.id_pregunta.startswith(f"PS-{a.informante}-"), f"{donde}: id de surgida"
            assert a.bloque.isdigit() and a.pregunta_eval in preguntas_eval, f"{donde}: bloque o pregunta de evaluación"
            codigos_ok(a.codigos, donde)
        else:
            assert a.id_pregunta in q.index, f"{donde}: no está en la guía"
            assert a.informante in q.loc[a.id_pregunta, "informantes"].split(";"), f"{donde}: no era de su guía"
    formuladas = {(a.informante, a.id_pregunta) for _, a in A.iterrows()}
    textos = {i: _normal(open(TRANSCRIPCION.format(i), encoding="utf-8").read()) for i in set(seg.informante)}
    assert not seg.id_segmento.duplicated().any(), "codificación: id de segmento repetido"
    for _, s in seg.iterrows():
        d = s.id_segmento
        assert d.startswith(f"ENTREV-{s.informante}-"), f"{d}: id e informante no coinciden"
        assert (s.informante, s.pregunta_guia) in formuladas, f"{d}: remite a una pregunta no formulada"
        assert s.pregunta_eval in preguntas_eval, f"{d}: pregunta de evaluación {s.pregunta_eval}"
        assert s.tipo_texto in ("Cita breve", "Paráfrasis") and s.sentido in SENTIDOS, f"{d}: tipo o sentido"
        assert s.hallazgo_critico in ("", "Candidato"), f"{d}: hallazgo crítico"
        assert bool(s.hallazgo_contrastado) == bool(s.resultado_contraste), f"{d}: contraste incompleto"
        assert not s.resultado_contraste or s.resultado_contraste in RESULTADOS, f"{d}: resultado del contraste"
        for h in filter(None, (x.strip() for x in s.hallazgo_contrastado.split(";"))):
            assert h in ids_h, f"{d}: hallazgo inexistente {h}"
        codigos_ok(s.codigos, d)
        assert " — " not in s.texto and " - " not in s.texto, f"{d}: guion explicativo"
        if s.tipo_texto == "Cita breve":  # la cita sale de la transcripción anonimizada, en orden si tiene elisiones
            pos = 0
            for frag in filter(None, (f.strip() for f in s.texto.split("…"))):
                pos = textos[s.informante].find(_normal(frag), pos)
                assert pos >= 0, f"{d}: la cita no está en la transcripción anonimizada: {frag[:60]}"
    for _, x in D.iterrows():  # contenido deducido de las respuestas por el equipo evaluador (decisión del usuario)
        citados = [c.strip() for c in x.segmentos.split(";")]
        assert set(citados) <= set(seg.id_segmento), f"{x.id}: la deducción cita segmentos inexistentes"
        assert x.hallazgo in ids_h and x.pregunta_eval in preguntas_eval, f"{x.id}: hallazgo o pregunta de evaluación"
    bloque = {(a.informante, a.id_pregunta): int(a.bloque) for _, a in A[A.formulacion == "Surgida"].iterrows()}
    seg["bloque"] = [bloque.get((s.informante, s.pregunta_guia), None) or int(q.loc[s.pregunta_guia, "bloque"])
                     for _, s in seg.iterrows()]
    seg["fecha"] = seg.informante.map(lambda i: APLICADA[i]["fecha"])
    return A, seg, D


def aplicada(Q_dis, A, seg):
    """Catálogo de lo aplicado: las preguntas de la guía que se formularon a algún informante, con los informantes y la
    prioridad restringidos a quienes se les formuló, y las surgidas (ADR-008)."""
    filas = []
    for _, r in Q_dis.iterrows():
        infs = [i for i in r.informantes.split(";") if ((A.informante == i) & (A.id_pregunta == r.id)).any()]
        if not infs:
            continue
        r = r.copy()
        if ":" in r.prioridad:
            pr = {i: prioridad_de(r, i) for i in infs}
            r.prioridad = (next(iter(pr.values())) if len(set(pr.values())) == 1
                           else "; ".join(f"{i}: {p}" for i, p in pr.items()))
        r.informantes = ";".join(infs)
        filas.append(r)
    Q = pd.DataFrame(filas)
    surg = []
    for _, a in A[A.formulacion == "Surgida"].iterrows():
        hs = seg[(seg.informante == a.informante) & (seg.pregunta_guia == a.id_pregunta)].hallazgo_contrastado
        contrasta = "; ".join(sorted({h for x in hs for h in filter(None, x.split("; "))}))
        surg.append({"id": a.id_pregunta, "bloque": int(a.bloque), "tipo": "Surgida", "informantes": a.informante,
                     "prioridad": "", "pregunta_eval": a.pregunta_eval, "pregunta": a.como_se_formulo,
                     "codigos": a.codigos, "contrasta": contrasta,
                     "evidencia": "Pregunta surgida en la entrevista, fuera de la guía."})
    Q = pd.concat([Q, pd.DataFrame(surg)], ignore_index=True)
    Q["bloque"] = Q.bloque.astype(int)
    return Q


def _verificar_aplicada(Q):
    # cada recomendación de la sección 3.7 que remite a ENTREV lo hace a una pregunta formulada que contrasta I4-H5
    val = pd.read_csv(VALORACION_REC, dtype=str).fillna("")
    for _, v in val[val.pregunta_entrev != ""].iterrows():
        assert v.pregunta_entrev in set(Q.id), f"{v.id}: pregunta ENTREV no formulada {v.pregunta_entrev}"
        assert "I4-H5" in Q.set_index("id").loc[v.pregunta_entrev, "contrasta"], f"{v.id}: la pregunta no contrasta I4-H5"


def vacios(Q_dis, A, seg):
    """Preguntas formuladas en parte, hallazgos cuyo contraste preveía la guía y que ningún segmento contrasta, y
    documentos ofrecidos pendientes."""
    q = Q_dis.set_index("id")
    filas = []
    for _, a in A[A.formulacion == "En parte"].iterrows():
        r = q.loc[a.id_pregunta]
        propios = set(seg[seg.informante == a.informante].hallazgo_contrastado.str.split("; ").explode())
        sin = [h for h in filter(None, (x.strip() for x in r.contrasta.split(";"))) if h not in propios]
        filas.append(["Pregunta formulada en parte", a.informante, a.id_pregunta,
                      f"{'Orientadora' if r.tipo in ORIENTADORAS else 'Repregunta'} ({r.pregunta_eval}). {a.nota}",
                      f"Hallazgos sin contraste de este informante: {', '.join(sin)}" if sin else ""])
    previstos = sorted({h for x in Q_dis.contrasta for h in filter(None, (y.strip() for y in x.split(";")))})
    con = set(seg.hallazgo_contrastado.str.split("; ").explode())
    for h in previstos:
        if h not in con:
            filas.append(["Hallazgo sin contraste", "Todos", h, "La guía preveía contrastarlo",
                          "Ningún segmento lo contrasta: las preguntas que lo contrastaban no se formularon"])
    for _, s in seg[seg.documento_ofrecido != ""].iterrows():
        filas.append(["Documento ofrecido o pedido", s.informante, s.id_segmento, s.documento_ofrecido,
                      "Pendiente de recepción; al llegar se registra con su hash en docs/fuentes"])
    return filas


def _previo_registro(ruta):
    """Lo anotado por el equipo evaluador en las columnas amarillas del registro anterior."""
    if not os.path.exists(ruta):
        return {}
    d = pd.read_excel(ruta, sheet_name="Registro", dtype=str).fillna("")
    if "id_segmento" not in d:
        return {}
    return {r.id_segmento: {c: r.get(c, "") for c in EVALUADOR_REG} for _, r in d.iterrows() if r.id_segmento}


def _segundos(m):
    p = [int(v) for v in m.strip().split(":")]
    return p[0] * 60 + p[1] if len(p) == 2 else p[0] * 3600 + p[1] * 60 + p[2]


def formuladas_de(Q, A, inf):
    """Preguntas formuladas al informante en el orden de la entrevista: (fila de la aplicación, fila del catálogo)."""
    a = A[A.informante == inf].copy()
    a["orden"] = a.minuto.map(lambda m: _segundos(m.split(";")[0]))
    q = Q.set_index("id")
    return [(x, q.loc[x.id_pregunta]) for _, x in a.sort_values("orden", kind="stable").iterrows()]


MESES = ("enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "setiembre", "octubre", "noviembre",
         "diciembre")


def _ficha(inf):
    ap = APLICADA[inf]
    a, m, d = (int(x) for x in ap["fecha"].split("-"))
    h, mi, se = (int(x) for x in ap["duracion"].split(":"))
    dur = (f"{h} h {mi} min" if h else f"{mi} min") + (f" {se} s" if se else "")
    return f"{d} de {MESES[m - 1]} de {a}, de {ap['inicio']} a {ap['termino']} h ({dur}), por videollamada"


def guia(Q, A, inf):
    """Guía aplicada: solo las preguntas que se formularon, con su redacción real, su minuto y su evidencia (ADR-008)."""
    d = Documento(pie="Guía de entrevista ENTREV (aplicada) · Evaluación de la PNIG · OMEP–OGMEPGD",
                  titulo_doc=f"ENTREV: guía de entrevista aplicada ({inf})")
    d.titulo("GUÍA DE ENTREVISTA ENTREV (APLICADA)",
             "Evaluación de cierre de la Política Nacional de Igualdad de Género (2019–2025)",
             f"Informante: {INFORMANTE[inf][0].lower() + INFORMANTE[inf][1:]}")
    pares = formuladas_de(Q, A, inf)
    bloques = sorted({int(r.bloque) for _, r in pares})
    ap = APLICADA[inf]

    def pregs_de(rs):
        return sorted({r.pregunta_eval for r in rs} - {"Todas", "Transversal"})

    tipos = pd.Series([("Surgida" if x.formulacion == "Surgida" else
                        "Orientadora" if r.tipo in ORIENTADORAS else "Repregunta") for x, r in pares]).value_counts()
    d.tabla(["Campo", "Contenido"], [
        ["Instrumento", "ENTREV, Guía de entrevista semiestructurada al nivel conductor (nota metodológica, Anexo 4)"],
        ["Informante", INFORMANTE[inf]],
        ["Código de registro", f"{inf} (la relación con la persona se guarda fuera del repositorio)"],
        ["Fecha y hora", _ficha(inf)],
        ["Modalidad", ap["modalidad"]],
        ["Preguntas de evaluación", ", ".join(pregs_de([r for _, r in pares])) + " y bloque transversal de lecciones"],
        ["Preguntas formuladas", lista_es([f"{tipos[t]} {t.lower()}{'s' if tipos[t] > 1 else ''}"
                                           for t in ("Orientadora", "Repregunta", "Surgida") if t in tipos])],
        ["Versión diseñada", "La guía previa a la entrevista, con todas las preguntas previstas, se conserva en la "
                             "carpeta administrativa (ADR-008)."],
    ], [2400, 6626], leyenda="Datos de la entrevista")

    d.h1("1. Propósito")
    d.p("Recoger la perspectiva del nivel conductor sobre las implicancias del nuevo marco conceptual, la conducción "
        "presupuestal, el funcionamiento de la rectoría y la articulación, y las condiciones de sostenibilidad ante la "
        "transición (nota metodológica, Anexo 4), y el avance de las recomendaciones de evaluaciones anteriores (Tabla "
        "A7.1 del CEPLAN, p3.4 a p3.6). El bloque 4 se adapta a la función del órgano: "
        f"{de_bloque(4, 0, inf).lower().replace('pemig', 'PEMIG').replace('ceplan', 'CEPLAN')}. Esta versión registra "
        "solo las preguntas que se formularon en la entrevista, con su redacción real y el minuto de la grabación.")
    d.h1("2. Indicaciones que se siguieron")
    for t in (
        [("Contraste. ", {"b": True}), "La evidencia de cada pregunta se presentó como lo que registran los documentos, "
         "no como conclusión de la evaluación."],
        [("Neutralidad. ", {"b": True}), "No se adelantaron valoraciones ni se mencionaron las respuestas de otros "
         "informantes."],
        [("Registro. ", {"b": True}), "La grabación y la transcripción con nombres se guardan en la carpeta "
         "administrativa. En el repositorio quedan la transcripción anonimizada y el registro de codificación, con el "
         "código de registro, el bloque, la pregunta, el texto (cita breve o paráfrasis) y los códigos."],
    ) + ((
        [("Independencia. ", {"b": True}), "La persona entrevistada no integra el equipo evaluador. Lo que declaró sobre "
         "el cumplimiento de recomendaciones de evaluaciones anteriores se contrasta con documentos."],)
         if inf == "OGMEPGD" else ()) + ((
        [("Entrevista grupal. ", {"b": True}), "Respondieron juntas tres especialistas de la DPIGND desde una sola cuenta; "
         "lo dicho se atribuye al órgano (ADR-007)."],) if inf == "DGIGND" else ()):
        d.vineta(*t)
    inicio = {b: min(_segundos(x.minuto.split(";")[0]) for x, r in pares if int(r.bloque) == b) for b in bloques}
    d.tabla(["Bloque", "Contenido (Anexo 4; el bloque 4, según la función)", "Pregunta", "Primera pregunta"],
            [[str(b), de_bloque(b, 1, inf), ", ".join(pregs_de([r for _, r in pares if int(r.bloque) == b]))
              or BLOQUES[b][2], f"{inicio[b] // 3600}:{inicio[b] % 3600 // 60:02d}:{inicio[b] % 60:02d}"]
             for b in bloques], [800, 5626, 1500, 1100], leyenda="Bloques tratados")

    d.h1("3. Apertura y consentimiento informado")
    d.p("Al inicio se informó el propósito de la evaluación, el carácter voluntario de la participación, el uso de la "
        "información y su tratamiento conforme a la Ley N.° 29733, y se pidió la autorización para grabar.", after=60)
    d.tabla(["Registro del consentimiento", ""], [["Acepta participar", ap["acepta"]],
                                                  ["Autoriza la grabación", ap["grabacion"]],
                                                  ["Hora de inicio y de término", f"{ap['inicio']} h / {ap['termino']} h"]],
            [3000, 6026])

    d.h1("4. Preguntas formuladas por bloque")
    for b in bloques:
        sub = [(x, r) for x, r in pares if int(r.bloque) == b]
        d.h2(f"Bloque {b}. {de_bloque(b, 0, inf)}")
        d.p(f"{de_bloque(b, 1, inf)} Pregunta de evaluación: "
            f"{lista_es(pregs_de([r for _, r in sub])) or BLOQUES[b][2]}.", after=80)
        filas = []
        for x, r in sub:
            guia_txt = texto_de(r, inf) if x.formulacion != "Surgida" else ""
            celda = ([(guia_txt, {}), (f"\nFormulada ({x.minuto}): ", {"b": True, "color": "595959"})] if guia_txt
                     else [(f"Formulada ({x.minuto}): ", {"b": True, "color": "595959"})])
            celda += [(x.como_se_formulo, {"i": True, "color": "595959"})]
            vinc = r.pregunta_eval + (f"\n{r.contrasta}" if r.contrasta else "")
            filas.append([x.id_pregunta, celda, x.formulacion, vinc, r.evidencia or "Pregunta orientadora del Anexo 4."])
        d.tabla(["N.°", "Pregunta de la guía y cómo se formuló", "Formulación", "Vincula con",
                 "Evidencia para el contraste"], filas, [1100, 3450, 1250, 1000, 2226], size=17)

    d.h1("5. Cierre")
    d.p("Se agradeció la participación y se acordó que los documentos mencionados se compartirían por correo "
        "institucional. Los hallazgos se presentarán en las reuniones de validación con la DGIGND y el Despacho "
        "Viceministerial de la Mujer, y en la reunión técnica con el CEPLAN.")
    d.guardar(f"{SALIDA}/ENTREV_guia_{inf}.docx")


def cuestionario(Q, A, seg, inf):
    """Cuestionario aplicado: las preguntas que se formularon, en el orden de la entrevista y con su redacción real.
    El cuestionario remitido antes de la entrevista se conserva en la carpeta administrativa (ADR-008)."""
    d = Documento(pie="Cuestionario ENTREV (aplicado) · Evaluación de la PNIG · OMEP–OGMEPGD",
                  titulo_doc=f"ENTREV: cuestionario aplicado ({inf})")
    d.titulo("CUESTIONARIO DE ENTREVISTA (APLICADO)",
             "Evaluación de cierre de la Política Nacional de Igualdad de Género (2019–2025)",
             f"Dirigido a: {DIRIGIDO[inf]}")
    ap = APLICADA[inf]
    d.tabla(["Campo", "Contenido"], [
        ["Fecha y hora", _ficha(inf)],
        ["Modalidad", ap["modalidad"]],
        ["A cargo de", "Oficina de Monitoreo y Evaluación de Políticas (OMEP) de la OGMEPGD"],
    ], [2400, 6626])
    d.h1("1. Propósito")
    d.p("La OMEP conduce la evaluación de cierre de la Política Nacional de Igualdad de Género (PNIG), que la Ley N.° "
        "32535 sustituirá por una nueva política. Este documento registra las preguntas que se formularon en la "
        "entrevista, en el orden en que se hicieron y con su redacción. Reemplaza al cuestionario remitido antes de la "
        "entrevista, que se conserva en la carpeta administrativa.")
    d.h1("2. Uso de la información")
    d.p("La participación fue voluntaria. El informe atribuye lo dicho al nivel conductor de la Política y no a personas, "
        "y los datos personales se tratan conforme a la Ley N.° 29733, Ley de Protección de Datos Personales. La "
        "grabación se guarda de forma segura y no forma parte del informe; la evaluación usa una transcripción "
        "anonimizada.")
    d.h1("3. Preguntas formuladas")
    n, vistos = 0, set()
    for x, r in formuladas_de(Q, A, inf):
        if x.como_se_formulo in vistos:  # preguntas que se formularon juntas
            continue
        vistos.add(x.como_se_formulo)
        n += 1
        mins = "minutos" if ";" in x.minuto else "minuto"
        d.p((f"{n}. ", {"b": True}), x.como_se_formulo, (f" ({mins} {x.minuto})", {"color": "595959"}), after=80)
    d.h1("4. Documentos ofrecidos o pedidos")
    docs = seg[(seg.informante == inf) & (seg.documento_ofrecido != "")].documento_ofrecido
    if docs.empty:
        d.p("En la entrevista no se ofrecieron ni se pidieron documentos.")
    for t in docs:
        d.vineta(t)
    d.p("Agradecemos su colaboración.", after=60)
    d.guardar(f"{SALIDA}/ENTREV_cuestionario_{inf}.docx")
    return n


def registro(Q, Q_dis, C, A, seg, D, ruta):
    wb = Workbook()
    azul = PatternFill("solid", fgColor="1F3864")
    amarillo = PatternFill("solid", fgColor="FFF2CC")
    blanco = Font(bold=True, color="FFFFFF")

    def hoja(ws, cols, anchos, filas=()):
        ws.append(cols)
        for c in ws[1]:
            c.fill, c.font = azul, blanco
            c.alignment = Alignment(wrap_text=True, vertical="center")
        for f in filas:
            ws.append(list(f))
        for i, a in enumerate(anchos):
            ws.column_dimensions[chr(65 + i)].width = a
        for fila in ws.iter_rows(min_row=2):
            for c in fila:
                c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.freeze_panes = "A2"

    ws = wb.active
    ws.title = "Registro"
    previo = _previo_registro(ruta)
    filas = []
    for _, s in seg.iterrows():
        fila = {c: s.get(c, "") for c in COLS_REGISTRO}
        fila.update(previo.get(s.id_segmento, {}))
        filas.append([fila[c] for c in COLS_REGISTRO])
    hoja(ws, COLS_REGISTRO, [15, 12, 11, 8, 7, 12, 9, 12, 60, 24, 11, 16, 16, 11, 26, 12, 12, 30, 34], filas)
    letra = {c: get_column_letter(i + 1) for i, c in enumerate(COLS_REGISTRO)}
    ultima = len(seg) + 100
    for c in EVALUADOR_REG:
        for f in range(2, ultima + 1):
            ws[f"{letra[c]}{f}"].fill = amarillo
    listas = {"informante": ",".join(TODOS), "tipo_texto": "Cita breve,Paráfrasis",
              "sentido": ",".join(SENTIDOS), "resultado_contraste": ",".join(RESULTADOS),
              "hallazgo_critico": "Candidato", "validacion": "Conforme,Corregido"}
    for col, valores in listas.items():
        dv = DataValidation(type="list", formula1=f'"{valores}"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"{letra[col]}2:{letra[col]}{ultima}")
    hoja(wb.create_sheet("Consentimiento"), ["informante", "fecha", "inicio", "termino", "duracion", "modalidad",
                                            "acepta_participar", "autoriza_grabacion"],
         [14, 11, 8, 8, 9, 40, 44, 44],
         [[i] + [APLICADA[i][k] for k in ("fecha", "inicio", "termino", "duracion", "modalidad", "acepta",
                                                   "grabacion")] for i in TODOS])
    q = Q.set_index("id")
    n_seg = seg.groupby(["informante", "pregunta_guia"]).size()
    aplic = []
    for _, a in A.iterrows():
        inf = a.informante
        if not a.id_pregunta.startswith("PS-"):
            r = q.loc[a.id_pregunta]
            base = [r.bloque, r.tipo, prioridad_de(r, inf), texto_de(r, inf), r.pregunta_eval, r.codigos]
        else:
            base = [int(a.bloque), "Surgida", "", a.como_se_formulo, a.pregunta_eval, a.codigos]
        aplic.append([a.informante, a.id_pregunta, base[0], base[1], base[2], a.formulacion, a.minuto, base[3],
                      a.como_se_formulo if not a.id_pregunta.startswith("PS-") else "", base[4], base[5],
                      int(n_seg.get((a.informante, a.id_pregunta), 0)), a.nota])
    hoja(wb.create_sheet("Aplicación"), ["informante", "id_pregunta", "bloque", "tipo", "prioridad", "formulacion",
                                        "minuto", "pregunta", "como_se_formulo", "pregunta_eval", "codigos",
                                        "segmentos", "nota"],
         [13, 12, 7, 18, 8, 13, 16, 60, 60, 10, 22, 10, 50], aplic)
    hoja(wb.create_sheet("Deducciones"), list(D.columns), [9, 10, 40, 80, 40, 10], D.values.tolist())
    form = A.set_index(["informante", "id_pregunta"]).formulacion
    preg = Q.copy()
    for inf in TODOS:
        preg[f"formulacion_{inf}"] = [form.get((inf, i), "") for i in preg.id]
    hoja(wb.create_sheet("Preguntas"), list(preg.columns), [8, 7, 18, 24, 8, 10, 70, 24, 18, 60, 14, 14, 14],
         preg.values.tolist())
    hoja(wb.create_sheet("Vacíos"), ["tipo", "informante", "referencia", "detalle", "efecto"],
         [24, 13, 14, 70, 60], vacios(Q_dis, A, seg))
    lib = C["libro"].reset_index()
    hoja(wb.create_sheet("Libro de códigos"), list(lib.columns), [12, 9, 7, 30, 30, 60, 50, 40, 20],
         lib.astype(str).values.tolist())
    hoja(wb.create_sheet("Criterios"), ["Elemento", "Regla", "Plano"], [22, 90, 22], [
        ["Unidad de registro", "Un segmento por idea: una respuesta puede dar varios segmentos. Cada segmento lleva la "
         "pregunta de la guía que lo originó.", "Equipo evaluador"],
        ["Anonimización", "Solo el código de registro (órgano). Sin nombres, cargos personales ni datos de contacto. "
         "La relación entre código y persona, las grabaciones y las notas en bruto se guardan fuera del repositorio.",
         "Nota metodológica, numeral 10; Ley N.° 29733"],
        ["Texto", "Cita breve (hasta dos oraciones) o paráfrasis fiel; se indica cuál en «tipo_texto».",
         "Equipo evaluador"],
        ["Códigos y sentido", "Libro común con REP_ENT (hoja «Libro de códigos») y las mismas categorías de sentido.",
         "Nota metodológica, numeral 9"],
        ["Contraste", "«hallazgo_contrastado» usa los identificadores de la matriz de triangulación (p. ej., I4-H2). "
         "«resultado_contraste»: coincide, matiza, discrepa o aporta información nueva.", "Equipo evaluador"],
        ["Atribución en el informe", "Por nivel de informante («nivel conductor»), no por persona.",
         "Nota metodológica, numeral 10"],
        ["Hallazgos críticos", "Los segmentos que puedan cambiar la valoración de una pregunta («hallazgo_critico» = "
         "Candidato) pasan a doble codificación ciega, con el mismo procedimiento de REP_ENT.",
         "Nota metodológica, numeral 9; decisión del usuario, 2026-10-31"],
        ["Entrevista grupal", "La DGIGND designó a tres especialistas de la DPIGND, que respondieron juntas desde una "
         "sola cuenta. Se codifica como entrevista del órgano (DGIGND), sin distinguir voces; el Anexo 4 preveía "
         "entrevistas individuales.", "Decisión del usuario, 2026-11-06 (ADR-007)"],
        ["Transcripciones", "Anonimizadas en 05_entrevistas/transcripciones/ (órgano, [informante], [evaluador], "
         "[especialista DPIGND], [cargo]). Cada cita breve se comprueba contra la transcripción de su informante.",
         "Decisión del usuario, 2026-11-06 (ADR-007); Ley N.° 29733"],
        ["Formulación de las preguntas", "Hoja «Aplicación»: solo las preguntas que se formularon, literal, "
         "reformulada o en parte, y las surgidas (fuera de la guía, id PS-órgano-NN), con el minuto de la grabación y "
         "su redacción real. Las preguntas de la guía que no se formularon se eliminaron del catálogo, de las guías y "
         "de los cuestionarios; la versión diseñada se conserva en la carpeta administrativa.",
         "Decisión del usuario, 2026-11-06 (ADR-008)"],
        ["Primera codificación", "Claude (1.ª), sobre la transcripción anonimizada. El usuario valida cada segmento en "
         "«validacion» (Conforme o Corregido) y anota la corrección en «observacion_evaluador»; ambas columnas se "
         "conservan al volver a correr el script.", "Equipo evaluador"],
        ["Vacíos", "Hoja «Vacíos»: preguntas formuladas en parte, con los hallazgos que dejan sin contraste; hallazgos "
         "cuyo contraste preveía la guía y que ningún segmento contrasta; documentos ofrecidos pendientes. Se consignan "
         "como hallazgo, no se estiman.", "CLAUDE.md, regla 4"],
        ["Deducción", "Hoja «Deducciones»: lo que el equipo evaluador deduce de las respuestas para temas que no se "
         "preguntaron completos, con los segmentos que lo sustentan y el hallazgo al que se incorpora. Así se trataron "
         "las medidas ante las recomendaciones (Tabla A7.1) y las limitaciones presupuestales, sin consulta adicional.",
         "Decisión del usuario, 2026-11-06"],
    ])
    wb.save(ruta)


if __name__ == "__main__":
    os.makedirs(SALIDA, exist_ok=True)
    C = S.cifras()
    H = S.hallazgos(C)
    Q_dis = preguntas(C)  # diseño de la guía (Anexo 4 y repreguntas); se verifica como se diseñó
    _verificar(Q_dis, C, H)
    A, seg, D = aplicacion(Q_dis, C, H)
    Q = aplicada(Q_dis, A, seg)  # catálogo de lo que se preguntó (ADR-008)
    _verificar_aplicada(Q)
    Q.to_csv(f"{SALIDA}/ENTREV_preguntas.csv", index=False, encoding="utf-8")
    for inf in TODOS:
        guia(Q, A, inf)
        print(inf, cuestionario(Q, A, seg, inf), "preguntas en el cuestionario aplicado")
    registro(Q, Q_dis, C, A, seg, D, f"{SALIDA}/ENTREV_registro_codificacion.xlsx")
    print(pd.crosstab(A.informante, A.formulacion).to_string())
    print(seg.groupby("informante").agg(segmentos=("id_segmento", "size"),
                                        candidatos=("hallazgo_critico", lambda x: (x == "Candidato").sum()),
                                        con_contraste=("hallazgo_contrastado", lambda x: (x != "").sum())).to_string())
    V = pd.DataFrame(vacios(Q_dis, A, seg), columns=["tipo", "informante", "referencia", "detalle", "efecto"])
    print(V.tipo.value_counts().to_string())
    print(len(Q_dis), "preguntas en el diseño;", len(Q), "en el catálogo aplicado;", len(D), "deducciones")
