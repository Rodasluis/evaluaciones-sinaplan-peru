"""REV_ALI.4, alineamiento normativo e internacional: genera normativa_internacional.csv (cada objetivo prioritario
frente a la Ley N.° 32535, la CEDAW, las observaciones finales del Comité CEDAW de 2022, Belém do Pará, Beijing, el
ODS 5 y la Estrategia de Montevideo).
Los CSV de 04_analisis/rev_ali_insumos/ son el registro del equipo evaluador y están versionados; este script es la fuente
de la que se generaron y los reproduce idénticos. Para cambiar una valoración se edita la lista de este script, se vuelve
a correr y se revisa la diferencia del CSV con git antes de correr rev_ali_coherencia_alineamiento.py, que valida la escala
y ubica cada cita en su página. Las citas son textuales y se ubican en el documento indicado en cada fila.

Uso (desde la raíz del repositorio):
    python src/analisis_pnig/rev_ali_insumos_normativa.py [carpeta_destino]
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
D, I, S = "Directa", "Indirecta", "Sin correspondencia"
LEY, CED, CO9, BEL, BEI, ODS, MON = ("Ley N.° 32535", "CEDAW", "Observaciones finales del Comité CEDAW (2022)",
                                     "Convención de Belém do Pará", "Plataforma de Acción de Beijing", "ODS 5 (Agenda 2030)",
                                     "Estrategia de Montevideo")
# op, instrumento, documento, referencia, cita, correspondencia, sustento
R = [
    # OP1: violencia
    ("OP1", LEY, "ley", "Art. 8, literal c",
     "Desarrollar políticas, planes y programas para la prevención, atención y eliminación de la violencia en todas sus formas y en todos los espacios, en especial la ejercida contra las mujeres", D,
     "Mandato expreso al Ejecutivo; el art. 8 g (hostigamiento sexual) y el art. 9 b (violencia sexual) lo complementan."),
    ("OP1", CED, "cedaw", "Arts. 1 y 2",
     "Los Estados Partes condenan la discriminación contra la mujer en todas sus formas", I,
     "La Convención no nombra la violencia; el Comité la considera discriminación desde sus recomendaciones generales 19 y 35, que la PNIG cita y que no están en el corpus."),
    ("OP1", CO9, "cedaw_co9", "Párr. 24, literal a (seguimiento a dos años, párr. 57)",
     "Mejore los mecanismos para hacer un seguimiento de la aplicación de las leyes que tipifican como delito la violencia de género contra las mujeres", D,
     "Recomendación de seguimiento prioritario al Estado peruano."),
    ("OP1", BEL, "belem", "Art. 7",
     "políticas orientadas a prevenir, sancionar y erradicar dicha violencia", D,
     "Objeto central de la Convención."),
    ("OP1", BEI, "beijing", "Esfera D, objetivo estratégico D.1",
     "Adoptar medidas integradas para prevenir y eliminar la violencia contra", D,
     "Mismo cambio; el objetivo D.3 añade la trata."),
    ("OP1", ODS, "agenda2030", "Meta 5.2",
     "Eliminar todas las formas de violencia contra todas las mujeres y las niñas en los ámbitos público y privado", D,
     "Mismo cambio sobre la misma población."),
    ("OP1", MON, "montevideo", "Dimensión crítica i) y medida 1.h",
     "Derecho a una vida libre de toda forma de violencia y discriminación", D,
     "La medida 1.h pide protocolos de intervención integrales e intersectoriales, como los que la PNIG exigió a sus servicios."),
    # OP2: salud sexual y reproductiva
    ("OP2", LEY, "ley", "Art. 8, literales k y n; Sexta DCF",
     "Garantizar el derecho a la salud, asegurando la disponibilidad, calidad, aceptabilidad y accesibilidad de los servicios, la prevención del embarazo adolescente", I,
     "La Ley no nombra los derechos sexuales y reproductivos: cubre servicios de salud, embarazo adolescente y maternidad segura (8 k) y una educación sexual «con base científica, biológica y ética» (8 n). La Sexta DCF ordena eliminar toda referencia a la educación sexual integral, contenido del lineamiento 2.2."),
    ("OP2", CED, "cedaw", "Art. 12",
     "el acceso a servicios de atención médica, inclusive los que se refieren a la planificación de la familia", D,
     "El art. 16.1 e añade el derecho a decidir el número de hijos y el art. 10 h la información sobre planificación familiar."),
    ("OP2", CO9, "cedaw_co9", "Párr. 34, literal c, y párr. 38",
     "Refuerce la labor de impartir educación sobre salud y derechos sexuales y reproductivos", D,
     "El Comité recomienda educación sobre salud y derechos sexuales y reproductivos en los planes de estudio (34 c) y medidas de salud reproductiva (38), con seguimiento a dos años del 34 b y del 38 c. La Sexta DCF de la Ley va en sentido contrario."),
    ("OP2", BEL, "belem", "Art. 4, literal b",
     "el derecho a que se respete su integridad física, psíquica y moral", I,
     "La Convención trata la violencia; la salud sexual y reproductiva solo se vincula a través de la integridad y de la violencia sexual."),
    ("OP2", BEI, "beijing", "Esfera C, objetivo estratégico C.1",
     "Fomentar el acceso de la mujer duran", D,
     "Acceso a servicios de salud apropiados, incluida la salud sexual y reproductiva (párrafos 94 a 96 de la Plataforma)."),
    ("OP2", ODS, "agenda2030", "Meta 5.6",
     "Asegurar el acceso universal a la salud sexual y reproductiva y los derechos reproductivos", D,
     "Mismo cambio sobre la misma población."),
    ("OP2", MON, "montevideo", "Dimensión crítica ii) y medida 9.c",
     "Derechos sexuales y reproductivos con relación a: información y educación sexual integral", D,
     "La dimensión incluye expresamente la educación sexual integral, que la Ley N.° 32535 ordena eliminar."),
    # OP3: toma de decisiones
    ("OP3", LEY, "ley", "Art. 8, literal a; art. 5, literal b; art. 10.1, literal b",
     "Promover y garantizar la participación plena y efectiva de mujeres y hombres en la consolidación del sistema democrático", D,
     "El art. 5 b admite medidas de acción positiva temporales (lineamiento 3.2) y el art. 10.1 b encarga al JNE y a la ONPE promover la participación política de la mujer."),
    ("OP3", CED, "cedaw", "Art. 7; art. 4.1",
     "eliminar la discriminación contra la mujer en la vida política y pública del país", D,
     "El art. 4.1 sustenta las medidas especiales de carácter temporal del lineamiento 3.2."),
    ("OP3", CO9, "cedaw_co9", "Párr. 30, literal a",
     "que incluyan cuotas reglamentarias para la representación de las mujeres en puestos de toma de decisiones", D,
     "Recomienda cuotas en Gobierno, administración pública y servicio exterior; el lineamiento 3.2 no tiene servicio."),
    ("OP3", BEL, "belem", "Art. 4, literal j",
     "el derecho a tener igualdad de acceso a las funciones públicas de su país y a participar en los asuntos públicos, incluyendo la toma de decisiones", D,
     "Derecho expreso de las mujeres a participar en la toma de decisiones."),
    ("OP3", BEI, "beijing", "Esfera G, objetivo estratégico G.1",
     "Adoptar medidas para garantizar a la mujer igualdad de acceso y la plena", D,
     "Mismo cambio sobre la misma población."),
    ("OP3", ODS, "agenda2030", "Meta 5.5",
     "Asegurar la participación plena y efectiva de las mujeres y la igualdad de oportunidades de liderazgo", D,
     "Mismo cambio sobre la misma población."),
    ("OP3", MON, "montevideo", "Dimensión crítica iv); medidas 1.k y 3.b",
     "Adoptar leyes y normativas para garantizar el acceso igualitario de las mujeres al poder político", D,
     "La Estrategia plantea la democracia paritaria; la Ley N.° 32058 eliminó la alternancia (hallazgo de REP_ENT)."),
    # OP4: derechos económicos y sociales
    ("OP4", LEY, "ley", "Art. 8, literales d, e, f, h, i, j y m",
     "Fomentar el acceso a recursos productivos", D,
     "La Ley cubre recursos productivos, empleo, conciliación, formalización, trabajo del hogar y educación; no menciona un sistema de cuidados (lineamiento 4.1)."),
    ("OP4", CED, "cedaw", "Arts. 11, 13 y 14",
     "eliminar la discriminación contra la mujer en la esfera del empleo", D,
     "El art. 11.2 c trata los servicios de cuidado infantil y el art. 14 a la mujer rural."),
    ("OP4", CO9, "cedaw_co9", "Párr. 36, literal c; párr. 40",
     "Asigne un presupuesto adecuado para la implementación del Sistema Nacional de Cuidados", D,
     "Recomendación directa sobre el lineamiento 4.1 y sobre protección social de las mujeres (párr. 40)."),
    ("OP4", BEL, "belem", "Art. 5",
     "Toda mujer podrá ejercer libre y plenamente sus derechos civiles, políticos, económicos, sociales y culturales", I,
     "Reconoce que la violencia anula estos derechos, sin medidas económicas específicas."),
    ("OP4", BEI, "beijing", "Esfera F, objetivo estratégico F.1; esfera A",
     "Promover la independencia y los dere", D,
     "Las esferas A (pobreza) y F (economía) cubren los lineamientos 4.2 a 4.4; el objetivo F.6 trata la armonización de responsabilidades laborales y familiares."),
    ("OP4", ODS, "agenda2030", "Metas 5.4 y 5.a",
     "Reconocer y valorar los cuidados y el trabajo doméstico no remunerados", D,
     "La meta 5.4 corresponde al lineamiento 4.1 y la 5.a al 4.2."),
    ("OP4", MON, "montevideo", "Dimensión crítica iii); nudo estructural c)",
     "Derechos económicos, sociales y culturales con relación a: ingresos, empleo e inserción productiva", D,
     "El nudo c) (división sexual del trabajo e injusta organización social del cuidado) coincide con la causa «asignación desigual de roles» de la PNIG."),
    # OP5: barreras institucionales
    ("OP5", LEY, "ley", "Art. 6.2; Quinta DCF; art. 11, literal a",
     "garantizan la incorporación progresiva del enfoque de igualdad de oportunidades entre mujeres y hombres en las políticas nacionales y sectoriales", D,
     "Mandato de incorporación del enfoque en las entidades; la Quinta DCF sustituye los lineamientos de transversalización del enfoque de género (ver REV_ALI.5)."),
    ("OP5", CED, "cedaw", "Arts. 2 y 3",
     "una política encaminada a eliminar la discriminación contra la mujer", D,
     "Obligación de adoptar una política y medidas en todas las esferas, incluida la administración pública."),
    ("OP5", CO9, "cedaw_co9", "Párr. 18, literales a y b",
     "Garantice la integridad del Ministerio de la Mujer y Poblaciones Vulnerables, tanto en su nombre como en su mandato", D,
     "Pide reforzar al mecanismo nacional y un presupuesto con perspectiva de género (lineamientos 5.1 y 6.3.5)."),
    ("OP5", BEL, "belem", "Art. 8, literal c",
     "fomentar la educación y capacitación del personal en la administración de justicia, policial", I,
     "Capacitación de funcionarios (lineamiento 5.4) solo en materia de violencia."),
    ("OP5", BEI, "beijing", "Esfera H, objetivo estratégico H.2",
     "Integrar perspectivas de género en las legislaciones, políticas, programas y", D,
     "Mismo cambio: el objetivo H.1 añade el fortalecimiento de los mecanismos nacionales."),
    ("OP5", ODS, "agenda2030", "Meta 5.c",
     "Aprobar y fortalecer políticas acertadas y leyes aplicables para promover la igualdad de género", D,
     "Mismo cambio en el plano normativo e institucional."),
    ("OP5", MON, "montevideo", "Eje 2 (institucionalidad), medidas 2.a y 2.c; eje 4, medida 4.b",
     "Otorgar a los mecanismos para el adelanto de las mujeres jerarquía al más alto nivel", D,
     "Los ejes 2 y 4 corresponden a los lineamientos 5.1, 5.3 y 5.4."),
    # OP6: patrones socioculturales
    ("OP6", LEY, "ley", "Art. 4.2, literal a; art. 9, literal a; art. 8, literal ñ",
     "desterrando prácticas, concepciones y lenguajes que justifiquen la superioridad de alguno de los sexos", I,
     "Figura como principio y, para los operadores de justicia, como modificación de concepciones discriminatorias; la Ley no tiene un lineamiento para la población en general. El art. 8 ñ (estadística desagregada) sostiene el lineamiento 6.3."),
    ("OP6", CED, "cedaw", "Art. 5, literal a",
     "Modificar los patrones socioculturales de conducta de hombres y mujeres", D,
     "Mismo cambio; el art. 10 c pide eliminar los estereotipos en la enseñanza (lineamiento 6.1)."),
    ("OP6", CO9, "cedaw_co9", "Párr. 22, literales a y d (seguimiento a dos años del 22 b)",
     "eliminar los estereotipos discriminatorios sobre los roles y responsabilidades de las mujeres y los hombres", D,
     "Estrategia intersectorial contra los estereotipos y educación en género para niñas y adolescentes."),
    ("OP6", BEL, "belem", "Art. 8, literal b; art. 6, literal b",
     "modificar los patrones socioculturales de conducta de hombres y mujeres", D,
     "Incluye programas educativos formales y no formales en todo nivel."),
    ("OP6", BEI, "beijing", "Esfera J, objetivo estratégico J.2; esferas B y L",
     "Fomentar una imagen equilibrada y no estereotipada de la mujer en los", D,
     "Los objetivos B.4 (educación no discriminatoria) y L.2 (actitudes culturales que perjudican a la niña) completan el OP6."),
    ("OP6", ODS, "agenda2030", "Meta 5.1; meta 4.7",
     "Poner fin a todas las formas de discriminación contra todas las mujeres y las", I,
     "Objetivo más amplio; la meta 4.7 incluye la educación para la igualdad de género."),
    ("OP6", MON, "montevideo", "Nudo estructural b); eje 6, medida 6.c",
     "Realizar campañas de comunicación continuas, orientadas al cambio cultural para la igualdad de género", D,
     "El nudo b) (patrones culturales patriarcales) coincide con la causa principal del modelo de la PNIG; la medida 1.m añade a los hombres como aliados (lineamiento 6.4)."),
]
cols = ["op", "instrumento", "documento", "referencia", "cita", "correspondencia", "sustento"]
with open(DEST / "normativa_internacional.csv", "w", encoding="utf-8", newline="") as f:
    w = csv.writer(f, quoting=csv.QUOTE_ALL)
    w.writerow(cols)
    w.writerows(R)
print(len(R))
