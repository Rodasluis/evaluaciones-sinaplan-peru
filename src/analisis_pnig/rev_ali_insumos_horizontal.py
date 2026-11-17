"""REV_ALI.3, alineamiento horizontal: genera horizontal.csv (relaciones de complementariedad, duplicidad o vacío de
atención entre servicios de la PNIG y de las políticas afines) y horizontal_universo.csv (universo de políticas
analizadas, del registro del CEPLAN y del alineamiento declarado en 2019, con su criterio de inclusión y resultado).
Los CSV de 04_analisis/rev_ali_insumos/ son el registro del equipo evaluador y están versionados; este script es la fuente
de la que se generaron y los reproduce idénticos. Para cambiar una valoración se edita la lista de este script, se vuelve
a correr y se revisa la diferencia del CSV con git antes de correr rev_ali_coherencia_alineamiento.py, que valida la escala
y ubica cada cita en su página. Las citas son textuales y se ubican en el documento indicado en cada fila.

Uso (desde la raíz del repositorio):
    python src/analisis_pnig/rev_ali_insumos_horizontal.py [carpeta_destino]
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
A = "docs/fuentes/politicas_afines/"
NNA = (A + "2021_mimp_ds-008-2021-mimp_pnm-ninas-ninos-adolescentes-al-2030.pdf", "PNM para las Niñas, Niños y Adolescentes al 2030", "D.S. 008-2021-MIMP")
PAM = (A + "2021_mimp_ds-006-2021-mimp_pnm-personas-adultas-mayores-al-2030.pdf", "PNM para las Personas Adultas Mayores al 2030", "D.S. 006-2021-MIMP")
PCD = (A + "2021_mimp_ds-007-2021-mimp_pnm-discapacidad-para-el-desarrollo-al-2030.pdf", "PNM en Discapacidad para el Desarrollo al 2030", "D.S. 007-2021-MIMP")
TRA = (A + "2021_mininter_ds-009-2021-in_pn-trata-de-personas-al-2030.pdf", "PN frente a la Trata de Personas y sus Formas de Explotación al 2030", "D.S. 009-2021-IN")
JUV = (A + "2019_minedu_ds-013-2019-minedu_pn-juventud.pdf", "PN de Juventud", "D.S. 013-2019-MINEDU")
PED = (A + "2021_mtpe_ds-013-2021-tr_pn-empleo-decente.pdf", "PN de Empleo Decente", "D.S. 013-2021-TR")
SAL = (A + "2020_minsa_ds-026-2020-sa_pnm-salud-al-2030.pdf", "PNM de Salud al 2030", "D.S. 026-2020-SA")
PIF = (A + "2019_mef_ds-255-2019-ef_pn-inclusion-financiera.pdf", "PN de Inclusión Financiera", "D.S. 255-2019-EF")
SUP = (A + "2020_minedu_ds-012-2020-minedu_pn-educacion-superior-tecnico-productiva.pdf", "PN de Educación Superior y Técnico-Productiva", "D.S. 012-2020-MINEDU")
RUR = (A + "2018_minedu_ds-013-2018-minedu_pn-atencion-educativa-ambitos-rurales.pdf", "PN de Atención Educativa para la Población de Ámbitos Rurales", "D.S. 013-2018-MINEDU")
AFR = (A + "2022_mincul_ds-005-2022-mc_pn-pueblo-afroperuano-al-2030.pdf", "PN del Pueblo Afroperuano al 2030", "D.S. 005-2022-MC")
PPI = (A + "2026_mincul_ds-002-2026-mc_pn-pueblos-indigenas-u-originarios-al-2040.pdf", "PN de Pueblos Indígenas u Originarios al 2040", "D.S. 002-2026-MC")
DDH = (A + "2024_minjusdh_ds-018-2024-jus_pnm-derechos-humanos-al-2040.pdf", "PNM de Derechos Humanos al 2040", "D.S. 018-2024-JUS")
VIV = (A + "2021_mvcs_ds-012-2021-vivienda_pn-vivienda-urbanismo-al-2030.pdf", "PN de Vivienda y Urbanismo al 2030", "D.S. 012-2021-VIVIENDA")
ECC = (A + "2024_minam_ds-012-2024-minam_estrategia-cambio-climatico-al-2050.pdf", "PN: Estrategia Nacional ante el Cambio Climático al 2050", "D.S. 012-2024-MINAM")
SC = (A + "2022_mininter_ds-006-2022-in_pnm-seguridad-ciudadana-al-2030.pdf", "PNM de Seguridad Ciudadana al 2030", "D.S. 006-2022-IN")
DIS = (A + "2022_midis_ds-008-2022-midis_pn-desarrollo-inclusion-social-al-2030_parte-1.pdf", "PN de Desarrollo e Inclusión Social al 2030", "D.S. 008-2022-MIDIS")
DIS2 = (A + "2022_midis_ds-008-2022-midis_pn-desarrollo-inclusion-social-al-2030_parte-2.pdf", "PN de Desarrollo e Inclusión Social al 2030", "D.S. 008-2022-MIDIS")
DIS4 = (A + "2022_midis_ds-008-2022-midis_pn-desarrollo-inclusion-social-al-2030_parte-4.pdf", "PN de Desarrollo e Inclusión Social al 2030", "D.S. 008-2022-MIDIS")
PES = (A + "2025_produce_ds-012-2025-produce_pn-pesca-al-2040_parte-1.pdf", "PN de Pesca al 2040", "D.S. 012-2025-PRODUCE")
AGR = (A + "2021_midagri_ds-017-2021-midagri_pn-agraria-2021-2030.pdf", "PN Agraria 2021–2030", "D.S. 017-2021-MIDAGRI")
C, D, V = "Complementariedad", "Duplicidad", "Vacío de atención"

# (politica, elemento de la política afín, entidad, OP de la PNIG, servicio de la PNIG, cita, tipo, sustento)
R = [
    (NNA, "OP3, lineamiento 03.04: servicio de atención de la violencia familiar y sexual a NNA", "MIMP (AURORA)", "OP1", "1.1.3",
     "SERVICIO DE ATENCIÓN DE LA VIOLENCIA FAMILIAR Y SEXUAL A NIÑAS, NIÑOS Y ADOLESCENTES", D,
     "Mismo proveedor (AURORA) y misma prestación (atención en los CEM) que el servicio 1.1.3, que comprende a los integrantes del grupo familiar. El servicio se reporta en dos políticas."),
    (NNA, "Servicio de implementación de la ESI a estudiantes de educación básica", "MINEDU", "OP2", "2.2.2",
     "SERVICIO DE IMPLEMENTACIÓN DE LA EDUCACIÓN SEXUAL INTEGRAL (ESI) A ESTUDIANTES DE EDUCACIÓN BÁSICA", D,
     "Mismo proveedor y contenido que el servicio 2.2.2. En la transición, la Sexta DCF de la Ley N.° 32535 obliga a rediseñar ambos."),
    (NNA, "Análisis horizontal de la PNMNNA con el OP2 de la PNIG", "MIMP", "OP2", "2.1.x y 2.3.1",
     "La Política Nacional de Igualdad de Género constituye tanto a nivel de objetivos como de lineamientos un abordaje centrado en la mujer", C,
     "La PNMNNA declara la complementariedad por la población adolescente que ambas atienden."),
    (NNA, "Servicio de cuidado diurno (SCD)", "MIDIS", "OP4", "4.1.1", "SERVICIO DE CUIDADO DIURNO - SCD", C,
     "Cuidado diurno infantil, parte de la oferta de cuidados que el servicio 4.1.1 debía articular en un sistema."),
    (PAM, "OP01: derecho al cuidado de las personas adultas mayores", "MIMP", "OP4", "4.1.1",
     "en el marco de un Sistema Nacional de Cuidados, sistema funcional previsto en la Política Nacional de Igualdad de Género", C,
     "La PNMPAM se apoya en el Sistema Nacional de Cuidados de la PNIG. Si ese sistema no se implementa, la oferta de cuidados para personas adultas mayores queda sin marco (ver REV_DOC)."),
    (PCD, "Servicio de atención integral en los CEM a personas con discapacidad", "MIMP (AURORA)", "OP1", "1.1.3",
     "Atención integral en los Centros de Emergencia Mujer a las personas con discapacidad afectadas por hechos de violencia", D,
     "Misma prestación y proveedor que el servicio 1.1.3; difiere solo en la población priorizada."),
    (PCD, "Análisis horizontal de la PNMDD con el lineamiento 4.1 de la PNIG", "MIMP", "OP4", "4.1.1",
     "L4.1. Implementar el Sistema Nacional de Cuidado con enfoque de género para personas en situación de dependencia", C,
     "La PNMDD declara la relación con el sistema de cuidados para personas con discapacidad en situación de dependencia."),
    (TRA, "Servicio de atención a víctimas de trata con fines de explotación sexual en los CEM", "MIMP (AURORA)", "OP1", "1.1.3",
     "Servicio de atención a víctimas de trata con fines de explotación sexual a través de los Centros de Emergencia Mujer (CEM)", D,
     "La atención en los CEM se registra también como servicio de la política frente a la trata."),
    (TRA, "Servicio de salud mental especializada en los CSMC", "MINSA", "OP1", "1.1.2",
     "Servicio de atención de salud mental especializada a niñas, niños, adolescentes, jóvenes y adultos (mujeres y hombres) víctimas de violencia, en los Centros de Salud Mental Comunitaria", D,
     "Mismo proveedor y prestación que el servicio 1.1.2."),
    (TRA, "Análisis horizontal: OP1, lineamiento 1.2 de la política frente a la trata", "MININTER", "OP1", "Lineamiento 1.2",
     "Los lineamientos contribuyen en la generación de estrategias preventivas para la población en riesgo de trata de personas", C,
     "Prevención focalizada en población en riesgo de trata, que complementa la prevención de la violencia contra las mujeres."),
    (JUV, "Servicio 5.1.1 de atención integral a mujeres afectadas por la violencia", "MIMP", "OP1", "1.1.3",
     "Servicio 5.1.1. Servicio de atención integral a las mujeres e integrantes del grupo familiar afectados por la violencia", D,
     "Mismo servicio del MIMP con la población joven como segmento."),
    (JUV, "Servicio 5.2.1 de acceso a información sobre violencia", "MIMP", "OP1", "1.2.3",
     "Acceso a la información sobre violencia contra las mujeres y los integrantes del grupo familiar", D,
     "Misma prestación informativa del MIMP que el servicio 1.2.3."),
    (PED, "Servicio 2.1.1 de certificación de competencias laborales", "MTPE", "OP4", "4.3.1",
     "Servicio 2.1.1 Servicio de certificación de competencias laborales", D,
     "Mismo proveedor y prestación que el servicio 4.3.1, que solo añade el foco en mujeres."),
    (PED, "Servicio 5.1.1 de orientación y prevención frente al hostigamiento sexual laboral", "MTPE", "OP1 y OP5", "1.1.8 y 5.5.1",
     "Servicios de orientación y prevención frente al hostigamiento sexual laboral", D,
     "Misma prestación del MTPE (Trabaja sin Acoso) registrada en dos políticas."),
    (PED, "Servicio de facilidades para mujeres con hijos pequeños: lactarios y guarderías", "MIMP y Comisión Multisectorial de Lactarios", "OP4", "4.3.7",
     "Consiste en facilitar la implementación de lactarios, guarderías y otras facilidades para las mujeres con hijos/as pequeños/as", D,
     "El seguimiento de lactarios figura en la PNIG y en la PED con el mismo proveedor; la comisión que lo proveía se extinguió en 2021 (REP_ENT)."),
    (PED, "Análisis horizontal: OP1, lineamiento 1.3 de la PED", "MTPE", "OP4", "Lineamiento 4.3",
     "Los lineamientos se complementan puesto que el primero propone una alternativa para aumentar la productividad laboral", C,
     "La PED declara la complementariedad con la inserción laboral formal de las mujeres."),
    (SAL, "Servicio 3.2 de salud articulado al CEM", "MINSA", "OP1", "1.1.2 y 1.1.3",
     "Servicio de salud articulado al Centro de Emergencia Mujer para la intervención y erradicación de la violencia contra las mujeres", C,
     "Componente de salud de la atención a víctimas, articulado al servicio 1.1.3. Su ficha consigna al SENASA como responsable del indicador, un error de transcripción."),
    (SAL, "Servicio de lactario institucional", "MINSA", "OP4", "4.3.7",
     "Servicio de lactario institucional disponible para madres en periodo de lactancia", D,
     "Tercera política que registra los lactarios institucionales como servicio."),
    (PIF, "Análisis horizontal: OP1 y OP2 de la PNIF", "MEF", "OP4", "4.3.6",
     "se vincula con los lineamientos del OP1 y OP2 de la PNIF", C,
     "La PNIF incluye el enfoque de género en el acceso al sistema financiero; el servicio 4.3.6 financia emprendimientos de mujeres."),
    (SUP, "Marco normativo de la PNESTP que cita el OP4 de la PNIG", "MINEDU", "OP4", "4.3.3",
     "dispone fortalecer la inserción laboral formal de las mujeres", C,
     "El acceso equitativo a la educación superior complementa la formación de mujeres en carreras masculinizadas."),
    (RUR, "Intervención de soporte asistencial, ESI y protección para la permanencia escolar", "MINEDU", "OP2", "2.2.2",
     "Soporte asistencial, educación sexual integral y protección para la permanencia en el sistema educativo", C,
     "Componente de ESI en ámbitos rurales; en la transición queda alcanzado por la Sexta DCF de la Ley N.° 32535."),
    (AFR, "Servicio de atención integral a mujeres afectadas por la violencia", "MIMP (AURORA)", "OP1", "1.1.3",
     "Servicio de atención integral a las mujeres e integrantes del grupo familiar afectados por hechos de violencia física, psicológica, sexual y económica o patrimonial", D,
     "Mismo servicio del MIMP con la población afroperuana como segmento."),
    (AFR, "Servicio de orientación y acompañamiento frente al hostigamiento sexual en el trabajo", "MTPE", "OP1", "1.1.8",
     "Servicio de orientación y acompañamiento para los casos de hostigamiento sexual en el trabajo", D,
     "Misma prestación del MTPE que el servicio 1.1.8."),
    (AFR, "OP4: autonomía de las mujeres afroperuanas", "MINCUL", "OP1, OP2 y OP4", "Lineamientos 1.1, 1.2, 2.1 y 2.3",
     "Incrementar la autonomía de las mujeres afroperuanas, en su diversidad", C,
     "Objetivo focalizado en mujeres afroperuanas que se apoya en servicios de la PNIG."),
    (PPI, "OP6, lineamiento 6.1: atención a mujeres indígenas afectadas por violencia en los CEM", "MIMP (AURORA)", "OP1", "1.1.3",
     "Servicio de atención interdisciplinaria y con pertinencia cultural a mujeres e integrantes del grupo familiar indígenas u originarias afectadas por hechos de violencia", D,
     "Misma prestación de los CEM con pertinencia cultural para mujeres indígenas."),
    (PPI, "Propuesta de ESI de las organizaciones indígenas, desestimada por el MINEDU", "MINEDU", "OP2", "2.2.2",
     "El MINEDU señala que la propuesta viene siendo trabajada en el marco de la Política Nacional de Igualdad de Género", V,
     "El MINEDU no incluyó la ESI en la política de pueblos indígenas porque ya estaba en la PNIG. Con la Sexta DCF de la Ley N.° 32535, la demanda de las organizaciones indígenas queda sin servicio en ninguna política."),
    (DDH, "Análisis horizontal: plataforma de estadística de la PNMDH frente al servicio 6.3.4", "MINJUSDH", "OP6", "6.3.4",
     "La nueva plataforma virtual complementa los servicios existentes", C,
     "La PNMDH declara que su plataforma complementa el sistema estadístico desagregado de la PNIG."),
    (DDH, "OP3: reducir estereotipos, prejuicios y estigmas", "MINJUSDH", "OP6", "Lineamientos 6.1 a 6.4",
     "Reducir la prevalencia de estereotipos, prejuicios, estigmas y actitudes negativas en el ejercicio pleno de los derechos de las personas", C,
     "Objetivo general para todos los grupos; el OP6 de la PNIG se centra en los patrones de género."),
    (DDH, "OP4: disminuir barreras institucionales de discriminación", "MINJUSDH", "OP5", "Lineamientos 5.1 a 5.5",
     "Disminuir las barreras institucionales, sistemáticas y estructurales de discriminación", C,
     "Enunciado muy próximo al del OP5 de la PNIG, con alcance general. Riesgo de superposición en la política sustituta."),
    (VIV, "Servicios de COFOPRI en la PNVU (saneamiento físico legal de equipamiento urbano)", "MVCS (COFOPRI)", "OP4", "4.2.2",
     "Saneamiento físico legal para equipamiento urbano y espacio público", V,
     "La PNVU no tiene un servicio de formalización individual de predios urbanos para mujeres. COFOPRI pidió retirar el 4.2.2 de la PNIG (REP_ENT); si se retira, la titulación urbana de mujeres queda sin servicio en ninguna política."),
    (ECC, "Servicio de fortalecimiento de capacidades en gestión del cambio climático con enfoque de género", "MINAM", "OP4", "4.2.4",
     "considerando los enfoques de género, interculturalidad e intergeneracional", C,
     "La ENCC 2050, sucesora del marco del Plan de Acción en Género y Cambio Climático que la PNIG declaró en 2019, incorpora el enfoque de género en la gestión climática."),
    # Políticas escaneadas, leídas con OCR el 2026-10-30 (docs/fuentes/ocr_2026-10-30.tsv)
    (SC, "OP4, servicio 37", "MIMP (AURORA)", "OP1", "1.2.1",
     "Estrategia comunitaria de prevención de la violencia basada en género contra las mujeres e integrantes del grupo familiar", D,
     "La PNMSC (2022) transcribe el enunciado del servicio 1.2.1, con el mismo proveedor (AURORA). El servicio se registra en dos políticas."),
    (SC, "OP4, servicio 38", "MIMP (AURORA)", "OP1", "1.2.2",
     "Formación de redes de hombres para promover la igualdad, nuevas masculinidades y erradicar prácticas machistas y discriminatorias", D,
     "Mismo enunciado, población (hombres de 18 años a más) y proveedor que el servicio 1.2.2."),
    (SC, "OP4, servicio 39", "MIMP (AURORA)", "OP1", "1.2.3",
     "de información que promueva el ejercicio del derecho a la denuncia de la violencia contra las mujeres", D,
     "Mismo enunciado que el servicio 1.2.3, con el enfoque de género añadido, y el mismo proveedor."),
    (SC, "OP4, servicio 43 (Chat 100)", "MIMP (AURORA)", "OP1", "1.2.5",
     "Servicio Chat 100 que brinda información en plataforma virtual para la detección y prevención de la violencia contra las mujeres", D,
     "Misma plataforma (Chat 100) y proveedor que el servicio 1.2.5."),
    (SC, "OP4, servicio 44 (CEM)", "MIMP (AURORA)", "OP1", "1.1.3",
     "a mujeres e integrantes del grupo familiar afectadas por hechos de violencia física, psicológica, sexual y económica o patrimonial, así como cualquier persona afectada por violencia sexual (CEM)", D,
     "Enunciado idéntico al del servicio 1.1.3, con el mismo proveedor."),
    (SC, "OP4, servicio 45", "Ministerio Público", "OP1", "1.1.6",
     "Servicio Fiscal Especializado eficaz para la protección y sanción de la violencia contra las mujeres e integrantes del grupo familiar", D,
     "Mismo proveedor (fiscalías especializadas) y prestación que el servicio 1.1.6."),
    (SC, "OP4, servicio 40", "Poder Judicial", "OP1", "1.1.5",
     "Servicio de protección y sanción efectiva mediante la implementación del sistema especializado de justicia dirigido a las mujeres e integrantes del grupo familiar víctimas de violencia", C,
     "El sistema especializado de justicia comprende el otorgamiento de medidas de protección del servicio 1.1.5 y añade la sanción."),
    (SC, "OP4, indicador 3: violencia psicológica ejercida por el esposo o compañero", "MININTER", "OP1", "Indicador del OP1",
     "Porcentaje de mujeres entre los 15 a 49 años de edad víctimas de violencia psicológica ejercida por el esposo o compañero en los últimos 12 meses", C,
     "La PNMSC mide la violencia de pareja contra las mujeres con la misma encuesta (ENDES) que el indicador del OP1 de la PNIG. Las dos políticas reportan sobre el mismo fenómeno."),
    (DIS, "OP2, servicio S.2.2.4: transferencia secundaria del programa Juntos", "MIDIS (Juntos)", "OP4", "4.4.1",
     "Entrega oportuna de incentivos monetarios condicionados según esquema de transferencia secundaria a hogares en situación de pobreza con estudiantes en educación secundaria", D,
     "Mismo programa (Juntos) y prestación que el servicio 4.4.1, que la PNIG acota a las mujeres en edad normativa. La PNDIS lo registra como servicio nuevo."),
    (DIS, "OP1, servicio S.1.4.1: cuidado y atención integral para niñas y niños de 6 a 36 meses (Cuna Más)", "MIDIS (Cuna Más)", "OP4", "4.1.1",
     "extra familiar por ausencia de un adulto competente para su atención en el ámbito del hogar", C,
     "Cuidado diurno infantil focalizado en hogares cuya madre o padre trabaja o estudia. Es parte de la oferta de cuidados que el servicio 4.1.1 debía articular en un sistema."),
    (DIS2, "OP5, lineamiento 5.3, servicio S.5.3.1: servicios de prevención de la violencia continuos dirigidos a hombres (ficha, pp. 197 y 198 del documento)", "MIMP (AURORA)", "OP1", "1.2.2",
     "voluntarios de la comunidad en masculinidades con enfoque de género", D,
     "La ficha describe la formación de redes de hombres («Colectivo de hombres líderes voluntarios…»), con el mismo proveedor que el servicio 1.2.2 (Unidad de Articulación de Servicios de Prevención, Atención y Protección de AURORA). Cambia solo el rango de edad de los receptores (16 a 59 años frente a 18 a más)."),
    (DIS2, "OP5, lineamiento 5.3, servicio S.5.3.2: medidas para la protección efectiva accesible a las mujeres víctimas de violencia (ficha, pp. 198 y 199 del documento)", "MININTER (PNP)", "OP1", "1.1.4",
     "mecanismo para la prohibición de la compra y posesión de armas", D,
     "Duplicidad parcial. Mismo proveedor (PNP) y objeto (medidas de protección) que el servicio 1.1.4; la ficha de la PNDIS añade la prohibición de armas, la detención en flagrancia y la mejora del SIDPOL, y mide el porcentaje de mujeres que denuncian y obtienen medidas de protección efectivas."),
    (DIS2, "OP5, lineamiento 5.3, servicio S.5.3.3: servicio de protección integral oportuno a mujeres víctimas de violencia (ficha, pp. 199 y 200 del documento)", "MINJUSDH (Defensa Pública)", "OP1", "1.1.7",
     "acceden a los servicios de Asistencia Legal y Defensa de", D,
     "Duplicidad parcial. El proveedor es la Dirección de Asistencia Legal y Defensa de Víctimas del MINJUSDH, el mismo del servicio 1.1.7, y el indicador de cobertura cuenta los patrocinios de la Defensa Pública a mujeres víctimas. La ficha añade la segunda respuesta policial y la detección de casos de riesgo severo por la PNP."),
    (PES, "Diagnóstico de la participación de las mujeres; la matriz de servicios (Tabla 35) no tiene servicios dirigidos a mujeres", "PRODUCE", "OP4", "4.3.5 y 4.3.6",
     "requiere igualdad de oportunidades y condiciones equitativas para la participación de mujeres", C,
     "La PNP reconoce la participación de las mujeres en el procesamiento y la comercialización, pero ninguno de sus servicios las distingue. La atención específica depende de los servicios 4.3.5 y 4.3.6, en los que PRODUCE es proveedor; si la política sustituta no los conserva, las mujeres de la pesca artesanal quedan sin servicio específico."),
    (AGR, "OP3, lineamiento 3.1, servicio 3.1.1: saneamiento físico legal de la propiedad agraria y catastro rural", "MIDAGRI (DIGESPACR) y gobiernos regionales", "OP4", "4.2.1",
     "consiste en ofrecer el saneamiento físico legal de la propiedad agraria y catastro rural", D,
     "Mismo proveedor y prestación que el servicio 4.2.1, que la PNIG acota a las mujeres dedicadas a la actividad agraria. En la PNA el receptor son los productores agrarios familiares y empresariales, sin distinción por sexo."),
    (AGR, "OP3, lineamiento 3.1, servicio 3.1.4: formalización de usuarios de agua para fines agrarios", "ANA", "OP4", "4.2.3",
     "formalizar el uso del agua por parte de los productores agrarios, a través de la emisión de licencias de uso de agua", D,
     "Misma prestación (licencias de uso de agua) que el servicio 4.2.3, que la PNIG acota a las mujeres en áreas rurales."),
    (AGR, "Alineamiento horizontal declarado: lineamiento 3.1 de la PNA y lineamiento 4.2 de la PNIG", "MIDAGRI", "OP4", "Lineamiento 4.2",
     "Existe complementariedad de acciones en ambas políticas, relacionadas a la formalización de la propiedad agraria", C,
     "Es la única relación con la PNIG que declara la PNA: formalización de la propiedad y uso sostenible de los recursos naturales."),
    (AGR, "Población de la PNA y responsables del OP1 y del OP2; la matriz de servicios no tiene servicios dirigidos a mujeres", "MIDAGRI", "OP4", "4.3.5 y 4.3.6",
     "Incluye a la mujer productora agraria", C,
     "La PNA incluye a las productoras en su población, registra la brecha de tierras (30,4 % de los productores individuales con tierras son mujeres, p. 92 del PDF) y pone a la Dirección de Promoción de la Mujer Productora Agraria entre los responsables del OP1 y del OP2 (pp. 121 y 122). Ninguno de sus servicios menciona a las mujeres; la atención específica depende de los servicios 4.3.5 y 4.3.6, en los que el MIDAGRI es proveedor."),
    (DIS4, "Alineamiento horizontal declarado (Anexo 8): lineamiento 2.3 de la PNDIS, autoestima y confianza de niñas, niños y adolescentes", "MIDIS", "OP1 y OP2", "Lineamientos 1.1 y 2.2",
     "contribuyen a fortalecer la autoestima y confianza de niñas, niños y adolescentes a través de la construcción de la identidad", C,
     "La PNDIS declara que los lineamientos 1.1 y 2.2 (ESI) de la PNIG complementan el suyo. En la transición, la Sexta DCF de la Ley N.° 32535 alcanza a la ESI."),
    (DIS4, "Alineamiento horizontal declarado (Anexo 8): lineamiento 3.1 de la PNDIS, salud de jóvenes y adultos", "MIDIS", "OP2", "Lineamientos 2.1 y 2.3",
     "Existe complementariedad entre ambos lineamientos toda vez que el lineamiento 3.1 de la PNDIS", C,
     "La PNDIS declara que su lineamiento de salud abarca la salud sexual y reproductiva con énfasis en la población en pobreza."),
    (DIS4, "Alineamiento horizontal declarado (Anexo 8): lineamiento 5.3 de la PNDIS, violencia en el entorno de la población", "MIDIS", "OP1", "Lineamiento 1.2",
     "plantea diversos servicios para prevenir y atender dicha problemática", C,
     "La PNDIS declara complementariedad con el lineamiento 1.2, pero su matriz registra como servicios propios tres servicios de la PNIG (filas de S.5.3.1 a S.5.3.3): lo declarado no coincide con lo que la matriz registra."),
]
cols = ["codigo", "politica", "norma", "documento", "elemento_politica_afin", "entidad", "op_pnig", "servicio_pnig", "cita",
        "tipo_relacion", "sustento"]
with open(DEST / "horizontal.csv", "w", encoding="utf-8", newline="") as f:
    w = csv.writer(f, quoting=csv.QUOTE_ALL)
    w.writerow(cols)
    for i, (pol, el, ent, op, sv, cita, tipo, sus) in enumerate(R, 1):
        w.writerow([f"H{i:02d}", pol[1], pol[2], pol[0], el, ent, op, sv, cita, tipo, sus])

# Universo: política, norma, criterio de inclusión, antecesora declarada en 2019, resultado
U = [
    ("PNM para las Niñas, Niños y Adolescentes al 2030", "D.S. 008-2021-MIMP", "Sucesora de una política declarada en 2019", "PNAIA 2012–2021 y Plan Multisectorial para la Prevención del Embarazo en Adolescentes 2013–2021", "Analizada"),
    ("PNM de Seguridad Ciudadana al 2030", "D.S. 006-2022-IN", "Sucesora de una política declarada en 2019", "Plan Nacional de Seguridad Ciudadana 2013–2018", "Analizada (escaneo con OCR)"),
    ("PN frente a la Trata de Personas y sus Formas de Explotación al 2030", "D.S. 009-2021-IN", "Sucesora de una política declarada en 2019", "Política Nacional frente a la Trata de Personas y sus Formas de Explotación (versión anterior)", "Analizada"),
    ("PNM de Derechos Humanos al 2040", "D.S. 018-2024-JUS", "Sucesora de una política declarada en 2019", "Plan Nacional de Derechos Humanos 2018–2021", "Analizada"),
    ("PN: Estrategia Nacional ante el Cambio Climático al 2050", "D.S. 012-2024-MINAM", "Sucesora de una política declarada en 2019", "Plan de Acción en Género y Cambio Climático", "Analizada"),
    ("Plan Nacional contra la Violencia de Género 2016–2021", "", "Declarada en 2019", "", "Sin sucesora en el registro del CEPLAN: su contenido quedó en el OP1 de la PNIG"),
    ("Plan Nacional de Fortalecimiento a las Familias 2016–2021", "", "Declarada en 2019", "", "Sin sucesora en el registro del CEPLAN"),
    ("Proyecto Educativo Nacional al 2021", "", "Declarada en 2019", "", "Sucesor (PEN al 2036) fuera del registro de políticas nacionales"),
    ("Política Nacional para la Transversalización del Enfoque Intercultural", "", "Declarada en 2019", "", "En actualización según el registro del CEPLAN; sin documento actualizado"),
    ("PNM de Salud al 2030", "D.S. 026-2020-SA", "Registro del CEPLAN: comparte servicios", "", "Analizada"),
    ("PN de Empleo Decente", "D.S. 013-2021-TR", "Registro del CEPLAN: comparte servicios", "", "Analizada"),
    ("PN de Inclusión Financiera", "D.S. 255-2019-EF", "Registro del CEPLAN: comparte población", "", "Analizada"),
    ("PN Agraria 2021–2030", "D.S. 017-2021-MIDAGRI", "Registro del CEPLAN: comparte servicios", "", "Analizada (escaneo con OCR)"),
    ("PN de Desarrollo e Inclusión Social al 2030", "D.S. 008-2022-MIDIS", "Registro del CEPLAN: comparte servicios", "", "Analizada (escaneo con OCR, cuatro partes). El registro del CEPLAN enlaza solo la parte 1; las partes 2 a 4 se tomaron de la página del decreto en gob.pe. El D.S. 009-2024-MIDIS la modificó solo en seguimiento y fichas de indicadores"),
    ("PNM para las Personas Adultas Mayores al 2030", "D.S. 006-2021-MIMP", "Registro del CEPLAN: comparte servicios", "", "Analizada"),
    ("PNM en Discapacidad para el Desarrollo al 2030", "D.S. 007-2021-MIMP", "Registro del CEPLAN: comparte servicios", "", "Analizada"),
    ("PN de Juventud", "D.S. 013-2019-MINEDU", "Registro del CEPLAN: comparte servicios", "", "Analizada"),
    ("PN de Atención Educativa para la Población de Ámbitos Rurales", "D.S. 013-2018-MINEDU", "Registro del CEPLAN: comparte población", "", "Analizada"),
    ("PN de Educación Superior y Técnico-Productiva", "D.S. 012-2020-MINEDU", "Registro del CEPLAN: comparte servicios", "", "Analizada"),
    ("PN de Modernización de la Gestión Pública al 2030", "D.S. 103-2022-PCM", "Registro del CEPLAN: posible relación con el OP5", "", "Excluida: no comparte población ni servicios; solo menciona la desigualdad de género como tendencia"),
    ("PN del Pueblo Afroperuano al 2030", "D.S. 005-2022-MC", "Registro del CEPLAN: comparte población", "", "Analizada"),
    ("PN de Pueblos Indígenas u Originarios al 2040", "D.S. 002-2026-MC", "Registro del CEPLAN: comparte población", "", "Analizada"),
    ("PN de Vivienda y Urbanismo al 2030", "D.S. 012-2021-VIVIENDA", "Registro del CEPLAN: comparte proveedor (COFOPRI)", "", "Analizada"),
    ("PN de Pesca al 2040", "D.S. 012-2025-PRODUCE", "Registro del CEPLAN: comparte servicios", "", "Analizada (escaneo con OCR). Parte 1 del anexo, que contiene la matriz de servicios"),
]
with open(DEST / "horizontal_universo.csv", "w", encoding="utf-8", newline="") as f:
    w = csv.writer(f, quoting=csv.QUOTE_ALL)
    w.writerow(["politica", "norma", "criterio_inclusion", "antecesora_declarada_2019", "resultado"])
    w.writerows(U)
print(len(R), len(U))
