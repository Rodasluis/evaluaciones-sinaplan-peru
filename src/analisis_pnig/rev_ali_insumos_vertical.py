"""REV_ALI.2, alineamiento vertical: genera vertical.csv (por objetivo prioritario, momento e instrumento: Políticas de Estado,
Plan Bicentenario, PEDN, seis PGG y PESEM) y vertical_declarado.csv (alineamiento que declaró la PNIG en 2019).
vertical_indicadores_pedn.csv y conceptual.csv no tienen generador: se escribieron directamente.
Los CSV de 04_analisis/rev_ali_insumos/ son el registro del equipo evaluador y están versionados; este script es la fuente
de la que se generaron y los reproduce idénticos. Para cambiar una valoración se edita la lista de este script, se vuelve
a correr y se revisa la diferencia del CSV con git antes de correr rev_ali_coherencia_alineamiento.py, que valida la escala
y ubica cada cita en su página. Las citas son textuales y se ubican en el documento indicado en cada fila.

Uso (desde la raíz del repositorio):
    python src/analisis_pnig/rev_ali_insumos_vertical.py [carpeta_destino]
La carpeta de destino por defecto es evaluaciones/pnig_2026/04_analisis/rev_ali_insumos.
"""
import csv
import sys
from pathlib import Path
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

DEST = Path(sys.argv[1] if len(sys.argv) > 1 else f"{EVAL}/04_analisis/rev_ali_insumos")
P, T, PT = "Periodo 2019–2025", "Transición", "Periodo y transición"
D, I, S = "Directa", "Indirecta", "Sin correspondencia"

# momento, op, instrumento, documento, referencia, cita, correspondencia, sustento
V = [
    # Políticas de Estado (la Tabla 2 de la nota usa las mismas en ambos momentos)
    (PT, "OP1", "Políticas de Estado", "pe2023", "Política de Estado N.° 7, literal c",
     "extender los mecanismos legales para combatir prácticas violentas", D,
     "Enuncia el combate del maltrato familiar y de la violación contra la integridad física y mental de las mujeres. La PNIG no declaró esta relación en 2019."),
    (PT, "OP2", "Políticas de Estado", "pe2023", "Política de Estado N.° 13, literal h",
     "promoverá la maternidad saludable y ofrecerá servicios de planificación familiar, con libre elección de los métodos y sin coerción", D,
     "Corresponde a los lineamientos 2.1 y 2.3 (servicios de salud sexual y reproductiva); no cubre el lineamiento 2.2 (educación sexual integral). La PNIG no declaró esta relación en 2019."),
    (PT, "OP3", "Políticas de Estado", "pe2023", "Política de Estado N.° 11, literal b",
     "fortalecerá la participación de las mujeres como sujetos sociales y políticos que dialogan y conciertan con el Estado", D,
     "Mismo cambio (participación de las mujeres) sobre la misma población. Relación declarada por la PNIG."),
    (PT, "OP4", "Políticas de Estado", "pe2023", "Política de Estado N.° 11, literal d",
     "dará acceso equitativo a las mujeres a recursos productivos y empleo", D,
     "Mismo cambio sobre la misma población (lineamientos 4.2 y 4.3). Relación declarada por la PNIG."),
    (PT, "OP5", "Políticas de Estado", "pe2023", "Política de Estado N.° 11, literal c",
     "fortalecerá una institución al más alto nivel del Estado en su rol rector de políticas y programas para la promoción de la igualdad de oportunidades entre hombres y mujeres", D,
     "Rectoría al más alto nivel para la igualdad de oportunidades, que el OP5 operativiza en la transversalización. Relación declarada por la PNIG."),
    (PT, "OP6", "Políticas de Estado", "pe2023", "Política de Estado N.° 11, literal a",
     "combatirá toda forma de discriminación, promoviendo la igualdad de oportunidades", I,
     "El OP6 contribuye a un objetivo más amplio: no menciona los patrones socioculturales. Relación declarada por la PNIG."),
    # Plan de desarrollo nacional del periodo: Plan Bicentenario (vigente al aprobarse la PNIG)
    (P, "OP1", "Plan Bicentenario", "bicentenario", "Eje 2, objetivo específico 6 (seguridad ciudadana), acción estratégica",
     "Implementar el Plan Nacional contra la Violencia contra la Mujer", D,
     "Mismo cambio sobre la misma población. Relación declarada por la PNIG."),
    (P, "OP2", "Plan Bicentenario", "bicentenario", "Eje 2, objetivo específico 2 (salud), acción estratégica",
     "Ampliar el acceso de las mujeres a los partos institucionales con adecuación cultural", I,
     "Se limita a la atención del parto; el OP2 abarca los derechos a la salud sexual y reproductiva. Relación declarada por la PNIG."),
    (P, "OP3", "Plan Bicentenario", "bicentenario", "Eje 1, objetivo específico 1, acción estratégica",
     "Desarrollar la normativa, las políticas públicas y las acciones afirmativas necesarias para garantizar la participación en las decisiones públicas de los grupos más vulnerables", I,
     "Participación en decisiones públicas de grupos vulnerables, sin referencia expresa a las mujeres. La PNIG declaró esta acción para el OP4, no para el OP3 (ver hoja del alineamiento declarado)."),
    (P, "OP4", "Plan Bicentenario", "bicentenario", "Eje 2, objetivo específico 4, acción estratégica",
     "Mejorar y extender los servicios de guardería infantil para facilitar el trabajo de las madres", D,
     "Cuidado infantil para el trabajo de las mujeres (lineamiento 4.1). La PNIG declaró esta acción para el OP6."),
    (P, "OP5", "Plan Bicentenario", "bicentenario", "Eje 1, objetivo específico 1, acción estratégica",
     "Poner en práctica los compromisos adquiridos en los tratados y convenios sobre derechos humanos suscritos por el Perú", I,
     "Acción genérica que la PNIG declaró para los seis OP; no enuncia el cambio institucional del OP5."),
    (P, "OP6", "Plan Bicentenario", "bicentenario", "Eje 1, objetivo específico 1, acción estratégica",
     "Fortalecer la educación en materia de derechos humanos con relevancia suficiente en los programas de estudio escolares", I,
     "Contribuye al lineamiento 6.1 sin mencionar estereotipos ni patrones de género. Relación declarada por la PNIG."),
    # PEDN al 2050 (D.S. 095-2022-PCM), vigente desde 2022 dentro del periodo
    (P, "OP1", "PEDN al 2050 (2022)", "pedn2022", "OE 1.5, AE 1.5.2",
     "Reducir todas las formas de violencia hacia las mujeres, niñas, niños y adolescentes", D,
     "Mismo cambio sobre la misma población."),
    (P, "OP2", "PEDN al 2050 (2022)", "pedn2022", "OE 1.2, AE 1.2.4",
     "Garantizar el ejercicio de los derechos a la salud sexual y reproductiva de las mujeres", D,
     "Mismo enunciado que el OP2 e incluye la educación sexual integral en el sistema educativo (lineamiento 2.2)."),
    (P, "OP3", "PEDN al 2050 (2022)", "pedn2022", "OE 1.5, AE 1.5.3",
     "Elevar los niveles de participación social y política de las mujeres", D,
     "Mismo cambio sobre la misma población."),
    (P, "OP4", "PEDN al 2050 (2022)", "pedn2022", "OE 3.2, AE 3.2.5",
     "Asegurar la igualdad de condiciones para el ejercicio de los derechos fundamentales laborales de los trabajadores, con énfasis en las mujeres", D,
     "Derechos laborales con énfasis en las mujeres (lineamiento 4.3)."),
    (P, "OP5", "PEDN al 2050 (2022)", "pedn2022", "OE 1.5, AE 1.5.1",
     "Reducir la desigualdad de género que afecta a las mujeres", I,
     "Objetivo más amplio al que contribuye el OP5; la acción no enuncia la incorporación del enfoque en las entidades."),
    (P, "OP6", "PEDN al 2050 (2022)", "pedn2022", "OE 1.5, AE 1.5.1",
     "Reducir la desigualdad de género que afecta a las mujeres", I,
     "Objetivo más amplio al que contribuye el OP6; la acción no menciona los patrones socioculturales."),
    # PEDN al 2050 actualizado (D.S. 103-2023-PCM), referencia para la transición
    (T, "OP1", "PEDN al 2050 actualizado (2023)", "pedn2023", "OE 1.5, AE 1.5.2",
     "Reducir todas las formas de violencia hacia las mujeres, niñas, niños y adolescentes", D,
     "Mismo cambio sobre la misma población. Su indicador es el IOP.01.01 de la PNIG (tomo II)."),
    (T, "OP2", "PEDN al 2050 actualizado (2023)", "pedn2023", "OE 1.2, AE 1.2.4",
     "Garantizar el ejercicio de los derechos a la salud sexual y reproductiva de las mujeres, en base a una educación sexual integral en el sistema educativo", D,
     "Directa con la PNIG, pero la acción se basa en la educación sexual integral, que la Sexta DCF de la Ley N.° 32535 ordena eliminar. Tensión entre el PEDN vigente y la Ley para el diseño de la política sustituta."),
    (T, "OP3", "PEDN al 2050 actualizado (2023)", "pedn2023", "OE 1.5, AE 1.5.3",
     "Elevar los niveles de participación social y política de las niñas, niños y adolescentes, mujeres", D,
     "Mismo cambio sobre la misma población. Su indicador es el IOP.03.01 de la PNIG (tomo II)."),
    (T, "OP4", "PEDN al 2050 actualizado (2023)", "pedn2023", "OE 3.2, AE 3.2.5",
     "Asegurar la igualdad de condiciones para el ejercicio de los derechos fundamentales laborales de los trabajadores, con énfasis en las mujeres", D,
     "Su indicador es el IOP.04.01 de la PNIG (tomo II)."),
    (T, "OP5", "PEDN al 2050 actualizado (2023)", "pedn2023", "OE 1.5, AE 1.5.1",
     "Reducir la desigualdad de género que afecta a las mujeres", I,
     "Objetivo más amplio; el tomo II usa el IOP.05.01 de la PNIG como indicador de esta acción."),
    (T, "OP6", "PEDN al 2050 actualizado (2023)", "pedn2023", "OE 1.5, AE 1.5.1",
     "Reducir la desigualdad de género que afecta a las mujeres", I,
     "Objetivo más amplio; el tomo II usa el IOP.06.01 de la PNIG como indicador de esta acción."),
    # PGG al 2021 (D.S. 056-2018-PCM)
    (P, "OP1", "PGG al 2021 (D.S. 056-2018-PCM)", "pgg2018", "Lineamiento 4.6",
     "garantizar la protección de la niñez, la adolescencia y las mujeres frente a todo tipo de violencia", D,
     "Mismo cambio sobre la misma población. La PNIG cita el lineamiento 4.6 en su parte considerativa."),
    (P, "OP2", "PGG al 2021 (D.S. 056-2018-PCM)", "pgg2018", "Lineamiento 4.2",
     "Brindar servicios de salud de calidad, oportunos, con capacidad resolutiva y con enfoque territorial", I,
     "Lineamiento general de salud sin referencia a la salud sexual y reproductiva."),
    (P, "OP3", "PGG al 2021 (D.S. 056-2018-PCM)", "pgg2018", "Lineamiento 4.6",
     "Promover la igualdad y no discriminación entre hombres y mujeres", I,
     "Objetivo amplio de igualdad al que contribuye el OP3."),
    (P, "OP4", "PGG al 2021 (D.S. 056-2018-PCM)", "pgg2018", "Lineamiento 3.6",
     "Fomentar la generación de empleo formal y de calidad, con énfasis en los", I,
     "Empleo formal con énfasis en los jóvenes, sin referencia a las mujeres."),
    (P, "OP5", "PGG al 2021 (D.S. 056-2018-PCM)", "pgg2018", "Lineamiento 4.6",
     "Promover la igualdad y no discriminación entre hombres y mujeres", I,
     "Objetivo amplio de igualdad al que contribuye el OP5."),
    (P, "OP6", "PGG al 2021 (D.S. 056-2018-PCM)", "pgg2018", "Lineamiento 4.6",
     "Promover la igualdad y no discriminación entre hombres y mujeres", I,
     "Objetivo amplio de igualdad al que contribuye el OP6."),
    # PGG 2021–2026 (D.S. 164-2021-PCM)
    (P, "OP1", "PGG 2021–2026 (D.S. 164-2021-PCM)", "pgg2021", "Eje 6, línea de intervención 6.1.7",
     "Impulsar una acción multisectorial permanente para garantizar acciones de prevención, la atención, recuperación y acceso a la justicia de las mujeres", D,
     "Mismo cambio sobre la misma población; la línea 6.1.4 añade el Sistema Nacional Especializado de Justicia."),
    (P, "OP2", "PGG 2021–2026 (D.S. 164-2021-PCM)", "pgg2021", "Eje 1, lineamiento 1.2",
     "Preservar la vida y la salud", I,
     "Lineamiento general de salud sin referencia a la salud sexual y reproductiva."),
    (P, "OP3", "PGG 2021–2026 (D.S. 164-2021-PCM)", "pgg2021", "Eje 6, línea de intervención 6.1.3",
     "Priorizar la participación efectiva de mujeres en su diversidad y poblaciones vulnerables, en la toma de decisiones y espacios públicos", D,
     "Mismo cambio sobre la misma población."),
    (P, "OP4", "PGG 2021–2026 (D.S. 164-2021-PCM)", "pgg2021", "Eje 2, línea de intervención 2.5.4",
     "Promover la autonomía económica femenina, y el acceso a recursos productivos, económicos y financieros de emprendimientos liderados por mujeres", D,
     "Mismo cambio sobre la misma población (lineamiento 4.2)."),
    (P, "OP5", "PGG 2021–2026 (D.S. 164-2021-PCM)", "pgg2021", "Eje 5, línea de intervención 5.3.5",
     "Transversalizar el enfoque intercultural, de género, intergeneracional y la perspectiva de discapacidad en la gestión pública", D,
     "Transversalización del enfoque de género en la gestión pública (lineamientos 5.1 a 5.4)."),
    (P, "OP6", "PGG 2021–2026 (D.S. 164-2021-PCM)", "pgg2021", "Eje 6, línea de intervención 6.1.6",
     "en la gestión pública para garantizar el ejercicio de los derechos de toda la población del país, en particular de las mujeres", I,
     "Objetivo amplio de garantía de derechos; no menciona los patrones socioculturales."),
    # PGG 2023–2026 (D.S. 042-2023-PCM, modificado por el D.S. 075-2024-PCM, que solo añade el eje 10)
    (P, "OP1", "PGG 2023–2026 (D.S. 042-2023-PCM)", "pgg2023", "Eje 3, lineamiento 3.4",
     "Fortalecer la prevención y atención de la violencia contra las mujeres e integrantes del grupo familiar", D,
     "Mismo cambio sobre la misma población."),
    (P, "OP2", "PGG 2023–2026 (D.S. 042-2023-PCM)", "pgg2023", "Eje 8, lineamiento 8.3",
     "Promover el acceso a los servicios de salud, con énfasis en la población en situación de vulnerabilidad", I,
     "Acceso general a servicios de salud, sin referencia a la salud sexual y reproductiva."),
    (P, "OP3", "PGG 2023–2026 (D.S. 042-2023-PCM)", "pgg2023", "Eje 3, lineamiento 3.5",
     "Fortalecer la participación efectiva de mujeres en la toma de decisiones y ámbitos públicos", D,
     "Mismo cambio sobre la misma población."),
    (P, "OP4", "PGG 2023–2026 (D.S. 042-2023-PCM)", "pgg2023", "Eje 4, lineamiento 4.8",
     "Promover la generación del empleo formal, garantizando el cumplimiento de los derechos laborales", I,
     "Empleo formal sin referencia a las mujeres."),
    (P, "OP5", "PGG 2023–2026 (D.S. 042-2023-PCM)", "pgg2023", "Eje 3, lineamiento 3.1",
     "Garantizar la lucha contra la discriminación y el racismo", I,
     "Objetivo amplio de no discriminación, sin referencia a la transversalización del enfoque."),
    (P, "OP6", "PGG 2023–2026 (D.S. 042-2023-PCM)", "pgg2023", "Eje 3, lineamiento 3.1",
     "Garantizar la lucha contra la discriminación y el racismo", I,
     "Objetivo amplio de no discriminación, sin referencia a los patrones de género."),
    # PGG 2025–2026 (D.S. 141-2025-PCM)
    (P, "OP1", "PGG 2025–2026 (D.S. 141-2025-PCM)", "pgg2025", "Eje 4, lineamiento 4.2",
     "Avanzar en la reducción de los niveles de violencia contra mujeres, niños y adolescentes y otros grupos vulnerables", D,
     "Mismo cambio sobre la misma población."),
    (P, "OP2", "PGG 2025–2026 (D.S. 141-2025-PCM)", "pgg2025", "Eje 4, lineamiento 4.1",
     "facilitar el acceso universal y con calidad a servicios de salud", I,
     "Acceso general a servicios de salud para población en pobreza, sin referencia a la salud sexual y reproductiva."),
    (P, "OP3", "PGG 2025–2026 (D.S. 141-2025-PCM)", "pgg2025", "Ninguna", "", S,
     "Los lineamientos del eje 3 tratan de elecciones generales, institucionalidad y diálogo sin referencia a la participación de las mujeres."),
    (P, "OP4", "PGG 2025–2026 (D.S. 141-2025-PCM)", "pgg2025", "Eje 2, lineamiento 2.4",
     "Promover la mejora de la economía familiar facilitando el acceso a servicios", I,
     "Economía familiar y formalización, sin referencia a las mujeres."),
    (P, "OP5", "PGG 2025–2026 (D.S. 141-2025-PCM)", "pgg2025", "Ninguna", "", S,
     "Ningún lineamiento trata de igualdad entre mujeres y hombres ni de transversalización."),
    (P, "OP6", "PGG 2025–2026 (D.S. 141-2025-PCM)", "pgg2025", "Ninguna", "", S,
     "Ningún lineamiento trata de discriminación ni de patrones socioculturales."),
    # PGG de la transición: D.S. 043-2026-PCM (marzo a setiembre de 2026)
] + [
    (T, f"OP{i}", "PGG 2026 (D.S. 043-2026-PCM)", "pgg2026a", "Ninguna", "", S,
     "Sus tres ejes (seguridad ciudadana, continuidad económica y transición democrática) no contienen lineamientos sobre mujeres, igualdad ni violencia.")
    for i in range(1, 7)
] + [
    # PGG 2026–2031 (D.S. 127-2026-PCM)
    (T, "OP1", "PGG 2026–2031 (D.S. 127-2026-PCM)", "pgg2026b", "Eje 1, lineamiento L1.2",
     "Mejorar el acceso oportuno, articulado y confiable de la ciudadanía a la seguridad y la justicia", I,
     "Acceso a seguridad y justicia de la ciudadanía en general; ningún lineamiento menciona la violencia contra las mujeres."),
    (T, "OP2", "PGG 2026–2031 (D.S. 127-2026-PCM)", "pgg2026b", "Eje 3, lineamiento L3.1",
     "Fortalecer la atención en salud a lo largo del curso de vida, con énfasis en la gestación", I,
     "Salud con énfasis en la gestación; no trata la salud sexual y reproductiva."),
    (T, "OP3", "PGG 2026–2031 (D.S. 127-2026-PCM)", "pgg2026b", "Ninguna", "", S,
     "El eje 5 trata de gobernabilidad democrática y diálogo, sin referencia al acceso de las mujeres a la toma de decisiones."),
    (T, "OP4", "PGG 2026–2031 (D.S. 127-2026-PCM)", "pgg2026b", "Eje 4, lineamiento L4.2",
     "Incrementar el acceso a empleo decente formal y productivo, especialmente para los jóvenes y la población en situación de vulnerabilidad", I,
     "Empleo para jóvenes y población vulnerable, sin referencia a las mujeres."),
    (T, "OP5", "PGG 2026–2031 (D.S. 127-2026-PCM)", "pgg2026b", "Ninguna", "", S,
     "Ningún lineamiento trata de igualdad entre mujeres y hombres ni de transversalización de un enfoque."),
    (T, "OP6", "PGG 2026–2031 (D.S. 127-2026-PCM)", "pgg2026b", "Ninguna", "", S,
     "Ningún lineamiento trata de discriminación ni de patrones socioculturales."),
    # PESEM del sector Mujer y Poblaciones Vulnerables
    (P, "OP1", "PESEM 2018–2026", "pesem2022", "OES 01, AES.01.01",
     "Proteger de la violencia de género y de otras que afectan a las mujeres y poblaciones vulnerables", D,
     "Mismo cambio sobre la misma población."),
    (P, "OP2", "PESEM 2018–2026", "pesem2022", "Ninguna", "", S,
     "El PESEM del sector Mujer no aborda la salud sexual y reproductiva, que corresponde al sector Salud (su PESEM no está en la Tabla 2 de la nota)."),
    (P, "OP3", "PESEM 2018–2026", "pesem2022", "OES 02, AES.02.02",
     "Promover la igualdad de derechos y oportunidades para las mujeres y poblaciones vulnerables", I,
     "Acción amplia de igualdad de derechos, sin referencia a la toma de decisiones."),
    (P, "OP4", "PESEM 2018–2026", "pesem2022", "OES 02, AES.02.02",
     "Promover la igualdad de derechos y oportunidades para las mujeres y poblaciones vulnerables", I,
     "Acción amplia; su indicador es el índice de autonomía de las mujeres."),
    (P, "OP5", "PESEM 2018–2026", "pesem2022", "OES 02",
     "Reducir la desigualdad de género, la discriminación, la violencia y otras desigualdades", I,
     "Objetivo amplio; no hay acción estratégica sobre la incorporación del enfoque en las entidades."),
    (P, "OP6", "PESEM 2018–2026", "pesem2022", "OES 02, AES.02.01",
     "Promover el cambio de los patrones sociales y culturales que subyacen a la desigualdad", D,
     "Mismo cambio; su indicador es el índice de tolerancia social a la violencia hacia las mujeres."),
    (T, "OP1", "PESEM 2024–2030", "pesem2024", "OES 01, AES 01.1",
     "Reducir los niveles de exposición al riesgo de las mujeres", D,
     "Su indicador es el porcentaje de mujeres víctimas de violencia de género en los últimos 12 meses."),
    (T, "OP2", "PESEM 2024–2030", "pesem2024", "Ninguna", "", S,
     "El PESEM 2024–2030 no menciona la salud sexual y reproductiva."),
    (T, "OP3", "PESEM 2024–2030", "pesem2024", "OES 02, AES 02.3",
     "Promover el empoderamiento y liderazgo de las mujeres", D,
     "Su indicador es el promedio de empoderamiento político de las mujeres."),
    (T, "OP4", "PESEM 2024–2030", "pesem2024", "OES 02, AES 02.1",
     "Asegurar la igualdad de condiciones para el ejercicio de los derechos económicos, laborales y productivos de las mujeres", D,
     "Mismo cambio sobre la misma población; la AES 01.3 añade el derecho al cuidado (lineamiento 4.1)."),
    (T, "OP5", "PESEM 2024–2030", "pesem2024", "OES 03, AES 03.1",
     "Incorporar las temáticas del sector Mujer y Poblaciones Vulnerables en los documentos de gestión", D,
     "Incorporación en los documentos de gestión de las entidades; el OES 03 mide además la implementación de la PNIG."),
    (T, "OP6", "PESEM 2024–2030", "pesem2024", "OES 02, AES 02.2",
     "Reducir los patrones sociales y culturales que subyacen a conductas de discriminación", D,
     "Mismo cambio; su indicador es el porcentaje de personas que toleran la violencia contra las mujeres."),
]

# Alineamiento declarado por la PNIG (El Peruano, p. 31): instrumento, referencia, OP declarado, cita en la PNIG, valoración
DECL = [
    ("Políticas de Estado", "Política de Estado N.° 10, literal a", "OP4",
     "Promoverá la producción, el desarrollo empresarial local y el empleo", "Consistente"),
    ("Políticas de Estado", "Política de Estado N.° 10, literal f", "OP4",
     "Establecerá un sistema local de identificación, atención y promoción del empleo", "Consistente"),
    ("Políticas de Estado", "Política de Estado N.° 11, literal a", "OP6",
     "Combatirá toda forma de discriminación, promoviendo la igualdad de oportunidades", "Consistente"),
    ("Políticas de Estado", "Política de Estado N.° 11, literal b", "OP3",
     "Fortalecerá la participación de las mujeres como sujetos sociales y políticos", "Consistente"),
    ("Políticas de Estado", "Política de Estado N.° 11, literal c", "OP5",
     "Fortalecerá una institución al más alto nivel del Estado en su rol rector", "Consistente"),
    ("Políticas de Estado", "Política de Estado N.° 11, literal d", "OP4",
     "Dará acceso equitativo a las mujeres a recursos productivos y empleo", "Consistente"),
    ("Políticas de Estado", "Política de Estado N.° 11, literal e", "OP4 y OP6",
     "Desarrollará sistemas que permitan proteger a niños, niñas, adolescentes", "Consistente"),
    ("Políticas de Estado", "Política de Estado N.° 11, literal f", "OP4",
     "Promoverá y protegerá los derechos de los integrantes de las comunidades étnicas", "Consistente"),
    ("Plan Bicentenario", "Eje 1, objetivo específico 1", "OP4",
     "necesarias para garantizar la participación en las decisiones públicas de los", "Inconsistente: la acción trata de participación en decisiones públicas (OP3), no de derechos económicos y sociales"),
    ("Plan Bicentenario", "Eje 1, objetivo específico 1", "OP6",
     "Fortalecer la educación en materia de derechos humanos con relevancia", "Consistente"),
    ("Plan Bicentenario", "Eje 1, objetivo específico 1", "OP1 a OP6",
     "Poner en práctica los compromisos adquiridos en los tratados y convenios", "Consistente, genérica"),
    ("Plan Bicentenario", "Eje 2, objetivo específico 4", "OP4",
     "Incrementar los empleos y la productividad", "Consistente"),
    ("Plan Bicentenario", "Eje 2, objetivo específico 4", "OP6",
     "Mejorar y extender los servicios de guardería infantil para facilitar el trabajo", "Inconsistente: la acción corresponde al cuidado (OP4, lineamiento 4.1), no a los patrones socioculturales"),
    ("Plan Bicentenario", "Eje 2, objetivo específico 4", "OP6",
     "Fomentar que las mujeres jefas de", "Inconsistente: la acción trata del acceso a ahorro y crédito (OP4, lineamiento 4.2), no de los patrones socioculturales"),
    ("Plan Bicentenario", "Eje 2, objetivo específico 2", "OP2",
     "Ampliar el acceso de las mujeres a los partos institucionales con adecuación cultural", "Consistente"),
    ("Plan Bicentenario", "Eje 2, objetivo específico 3", "OP2",
     "Seguridad alimentaria, con énfasis en la nutrición", "Consistente, indirecta: nutrición de gestantes"),
    ("Plan Bicentenario", "Eje 2, objetivo específico 6", "OP1",
     "Implementar programas de asistencia a las víctimas de todo tipo de delitos", "Consistente"),
]

cols = ["momento", "op", "instrumento", "documento", "referencia", "cita", "correspondencia", "sustento"]
with open(DEST / "vertical.csv", "w", encoding="utf-8", newline="") as f:
    w = csv.writer(f, quoting=csv.QUOTE_ALL)
    w.writerow(cols)
    w.writerows(V)
with open(DEST / "vertical_declarado.csv", "w", encoding="utf-8", newline="") as f:
    w = csv.writer(f, quoting=csv.QUOTE_ALL)
    w.writerow(["instrumento", "referencia", "op_declarado", "cita_pnig", "valoracion"])
    w.writerows(DECL)
print(len(V), len(DECL))
