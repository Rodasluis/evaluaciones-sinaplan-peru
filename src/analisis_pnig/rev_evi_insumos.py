"""REV_EVI, revisión rápida de evidencia (nota metodológica v15, numeral 5.1 y Anexo 6): genera los insumos que registra el
equipo evaluador en evaluaciones/pnig_2026/04_analisis/rev_evi_insumos/ (las estrategias, los descriptores y los registros de
la búsqueda académica los escribe rev_evi_busqueda_academica.py en la misma carpeta):
  - tipologia.csv: los 52 servicios agrupados en 24 tipos de intervención, con la causa o el efecto del modelo que
    atienden (Anexo 5 de la PNIG, tal como lo clasifica REV_ALI.1), la población, la intervención, el resultado esperado y
    las cadenas de búsqueda en español e inglés;
  - busquedas.csv: cada búsqueda realizada (fecha, buscador, restricción de dominio, cadena, resultados revisados);
  - cribado.csv: candidatas de la búsqueda web que no entran a la matriz, con su motivo;
  - evidencia.csv: fuentes incluidas, con referencia, enlace, verificación (fecha y qué se comprobó), diseño, ámbito,
    año, resultado principal, dirección del efecto y origen (búsqueda web o registro de la búsqueda académica);
  - seleccion_academica.csv: registros de la búsqueda académica seleccionados por título y leídos por su resumen, con la
    decisión (incluida, excluida con su motivo, o ya incluida por la búsqueda web).
Toda referencia se verificó en una página de origen (repositorio, editorial, RePEc, PubMed o Cochrane) antes de
registrarse; lo que no se pudo verificar no se incluye. La regla de inclusión y la de solidez se aplican en
rev_evi_evidencia.py, no aquí.

Uso (desde la raíz del repositorio):
    python src/analisis_pnig/rev_evi_insumos.py [carpeta_destino]
La carpeta de destino por defecto es evaluaciones/pnig_2026/04_analisis/rev_evi_insumos.
"""
import csv
import sys
from pathlib import Path
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

DEST = Path(sys.argv[1] if len(sys.argv) > 1 else f"{EVAL}/04_analisis/rev_evi_insumos")
FECHA = "2026-10-31"
VIOL = "Efecto: vulneración del derecho a una vida libre de violencia"
SSR = "Efecto: vulneración de los derechos a la salud sexual y reproductiva"
PAT = "Causa: patrones socioculturales discriminatorios"
ROL = "Causa: asignación desigual de roles"
NOR = "Causa: normas y cultura institucional"
DESC = "Efecto: vulneración de los derechos económicos y sociales"
PART = "Efecto: vulneración del derecho a la participación"

# tipo, nombre, servicios, nodo del modelo, población, intervención, resultado esperado, pertinencia de la evidencia
# hallada respecto del resultado esperado, cadena en español, cadena en inglés
TIPOS = [
    ("T01", "Atención integral a víctimas de violencia", "1.1.1; 1.1.3", VIOL, "Mujeres, niñas, niños y adolescentes "
     "víctimas de violencia", "Centros de atención integral con servicios policiales, legales, psicológicos y sociales",
     "Menor violencia y revictimización", "Directa", "centros de atención integral violencia contra la mujer evaluación",
     "women's justice centers OR one-stop crisis centres intimate partner violence impact"),
    ("T02", "Salud mental para víctimas de violencia", "1.1.2", VIOL, "Mujeres víctimas de violencia",
     "Atención psicológica y de salud mental", "Menor depresión y trauma", "Directa",
     "terapias psicológicas mujeres violencia de pareja revisión", "psychological therapies women intimate partner violence"),
    ("T03", "Respuesta policial y judicial especializada y medidas de protección", "1.1.4; 1.1.5; 1.1.6; 1.1.7", VIOL,
     "Mujeres víctimas de violencia", "Comisarías, fiscalías y juzgados especializados; medidas de protección; defensa "
     "legal", "Más denuncia y sanción; menor violencia", "Directa", "comisarías de mujeres violencia evaluación de impacto",
     "women's police stations violence against women impact evaluation"),
    ("T04", "Atención y prevención del hostigamiento sexual laboral", "1.1.8; 5.5.1", VIOL, "Trabajadoras",
     "Orientación, asistencia legal y difusión sobre el hostigamiento sexual laboral", "Menor hostigamiento; más denuncia",
     "Directa", "hostigamiento sexual laboral capacitación efectividad", "sexual harassment training effectiveness"),
    ("T05", "Capacitación de operadores de justicia en violencia contra las mujeres", "1.1.9", VIOL,
     "Operadores de justicia y policía", "Formación especializada", "Mejor respuesta institucional a las víctimas", "Directa",
     "capacitación policía violencia familiar evaluación",
     "police training domestic violence response systematic review"),
    ("T06", "Prevención comunitaria de la violencia y cambio de normas", "1.2.1", PAT, "Comunidad",
     "Movilización comunitaria e intervenciones grupales", "Normas menos tolerantes con la violencia; menor violencia de pareja",
     "Directa", "movilización comunitaria prevención violencia de pareja",
     "community mobilisation group-based interventions intimate partner violence low- and middle-income countries"),
    ("T07", "Trabajo con hombres y masculinidades", "1.2.2", PAT, "Hombres de 18 años a más",
     "Redes y grupos de hombres para la igualdad", "Actitudes y conductas igualitarias; menor violencia", "Directa",
     "programas con hombres masculinidades violencia evaluación",
     "engaging men and boys gender norms violence prevention evidence"),
    ("T08", "Información y comunicación para la denuncia y la igualdad", "1.2.3; 1.2.5; 6.2.1", PAT, "Población general",
     "Campañas, medios masivos y plataformas de información", "Más denuncia; normas menos tolerantes", "Directa",
     "campañas medios masivos violencia contra la mujer experimento",
     "mass media campaign violence against women disclosure experiment"),
    ("T09", "Prevención de la violencia en la escuela", "1.2.4", PAT, "Estudiantes de educación básica",
     "Programas escolares de convivencia y prevención", "Menor violencia escolar y en el enamoramiento", "Directa",
     "programas escolares prevención violencia enamoramiento revisión sistemática",
     "school-based interventions dating sexual violence systematic review"),
    ("T10", "Tratamiento de agresores sexuales", "1.2.6", PAT, "Personas condenadas por agresión sexual",
     "Tratamiento especializado", "Menor reincidencia", "Directa", "tratamiento agresores sexuales reincidencia",
     "sexual offender treatment recidivism meta-analysis"),
    ("T11", "Servicios e información en salud sexual y reproductiva", "2.1.1; 2.1.2; 2.1.3; 2.1.4; 2.3.1", SSR,
     "Mujeres y adolescentes", "Consejería, métodos anticonceptivos, anticoncepción oral de emergencia e información",
     "Menos embarazos no planificados; ejercicio de derechos", "Directa",
     "prevención embarazo adolescente intervenciones revisión sistemática",
     "interventions preventing unintended pregnancies adolescents; adolescent friendly sexual reproductive health services"),
    ("T12", "Educación sexual integral", "2.2.1; 2.2.2", PAT, "Estudiantes, familias y docentes",
     "Educación sexual integral en la escuela", "Conocimientos y conductas sexuales seguras; menos embarazos e ITS",
     "Directa", "educación sexual integral escuela efectos",
     "school-based sexuality education gender and power evaluation low- and middle-income countries"),
    ("T13", "Participación política de las mujeres", "3.1.1; 3.1.2", PART, "Mujeres electoras, candidatas y "
     "organizaciones políticas", "Educación electoral y asistencia técnica a organizaciones políticas",
     "Mayor participación y representación de las mujeres", "Parcial",
     "capacitación candidatas mujeres evaluación de impacto",
     "training women candidates impact evaluation; voter education women turnout experiment"),
    ("T14", "Servicios de cuidado", "4.1.1", ROL, "Mujeres a cargo de personas dependientes",
     "Servicios de cuidado infantil y de personas dependientes", "Mayor participación laboral de las madres", "Parcial",
     "cuidado infantil empleo materno países de ingreso medio",
     "childcare maternal labor market outcomes lower- and middle-income countries"),
    ("T15", "Formalización de la propiedad y de derechos sobre recursos", "4.2.1; 4.2.2; 4.2.3", DESC,
     "Mujeres rurales y urbanas", "Titulación de predios y otorgamiento de derechos de agua",
     "Acceso y control de recursos por las mujeres", "Parcial", "titulación de tierras mujeres efectos",
     "land property rights interventions systematic review women"),
    ("T16", "Participación de las mujeres en la gestión de recursos naturales", "4.2.4", DESC, "Mujeres rurales",
     "Gestión de ecosistemas con participación de las mujeres", "Mejor gobernanza y conservación", "Directa",
     "participación mujeres gestión forestal pesquera", "gender composition forest fishery management groups outcomes"),
    ("T17", "Formación laboral, certificación y carreras no tradicionales", "4.3.1; 4.3.2; 4.3.3; 4.3.4", ROL,
     "Mujeres en edad de trabajar", "Capacitación vocacional, certificación de competencias y formación técnica",
     "Empleo, formalidad e ingresos", "Directa", "formación vocacional mujeres empleo revisión sistemática",
     "vocational training women labour market outcomes systematic review; recognition of prior learning impact evaluation"),
    ("T18", "Desarrollo empresarial y financiamiento para mujeres", "4.3.5; 4.3.6", ROL, "Mujeres emprendedoras",
     "Capacitación empresarial, asistencia técnica y crédito", "Ventas, ganancias y supervivencia de los negocios",
     "Directa", "capacitación empresarial microcrédito mujeres evaluación",
     "business training entrepreneurship evaluations developing world; microcredit randomized evaluations"),
    ("T19", "Lactarios institucionales", "4.3.7", ROL, "Trabajadoras en periodo de lactancia",
     "Seguimiento de lactarios en el centro de trabajo", "Permanencia laboral de las madres", "Parcial",
     "lactarios centro de trabajo efectos", "employer-based breastfeeding support working mothers systematic review"),
    ("T20", "Transferencias condicionadas y culminación de la secundaria", "4.4.1", DESC,
     "Adolescentes de hogares en pobreza", "Transferencias monetarias condicionadas a la asistencia escolar",
     "Matrícula, asistencia y culminación de la secundaria", "Directa",
     "programa Juntos evaluación de impacto educación",
     "cash transfer programmes schooling outcomes systematic review"),
    ("T21", "Transversalización del enfoque de género en la gestión pública", "5.1.1; 5.2.1; 5.3.1; 5.4.1; 6.3.5", NOR,
     "Entidades públicas y servidores", "Asistencia técnica, capacitación y presupuesto con enfoque de género",
     "Políticas, presupuestos y servicios que atienden las brechas", "Directa",
     "transversalización enfoque de género asistencia técnica evaluación",
     "gender mainstreaming technical assistance effectiveness; gender budgeting outcomes"),
    ("T22", "Estadísticas y estudios de género", "6.3.1; 6.3.2; 6.3.3; 6.3.4", PAT, "Entidades públicas y usuarios de "
     "información", "Encuestas de uso del tiempo y de relaciones sociales, estadística desagregada, estudios",
     "Uso de la información en las políticas", "Directa", "encuestas de uso del tiempo uso en políticas",
     "time use surveys gender statistics policy influence"),
    ("T23", "Educación para la igualdad en la escuela", "6.1.1; 6.1.2; 6.1.3", PAT, "Estudiantes y docentes",
     "Programas educativos y formación docente para la igualdad", "Actitudes igualitarias", "Directa",
     "educación igualdad de género escuela experimento actitudes",
     "school-based gender attitudes experiment adolescents"),
    ("T24", "Corresponsabilidad familiar en el cuidado", "6.4.1", PAT, "Familias",
     "Consejería a las familias para compartir el cuidado", "Distribución equitativa del trabajo doméstico y de cuidado",
     "Directa", "corresponsabilidad cuidado parejas intervención",
     "couples intervention men's participation childcare household tasks randomized"),
]

# Búsquedas: cadena, buscador o restricción, resultados revisados, tipos
WEB = "Buscador web general"
BUSQUEDAS = [
    ("Gender violence, enforcement, and human capital women's justice centers Peru", WEB, 10, "T01; T03"),
    ("Cochrane advocacy interventions violence women intimate partner abuse", WEB, 10, "T01"),
    ("Cochrane psychological therapies women intimate partner violence", WEB, 10, "T02"),
    ("Women's police stations intimate partner violence Brazil", WEB, 10, "T03"),
    ("Women police stations India gender crime punishment", WEB, 10, "T03"),
    ("Sexual harassment training effectiveness review", WEB, 10, "T04"),
    ("Systematic review police training response intimate partner violence", WEB, 10, "T05"),
    ("Police training family domestic sexual violence Australian Institute of Criminology", WEB, 10, "T05"),
    ("SASA! cluster randomised trial community mobilisation intimate partner violence", WEB, 10, "T06"),
    ("Prevention of violence against women and girls what does the evidence say", WEB, 10, "T06"),
    ("Violence against women intervention systematic review",
     "Restringida a campbellcollaboration.org y onlinelibrary.wiley.com", 9, "T01; T06"),
    ("Gender norms intimate partner violence women's empowerment policy insights", "Restringida a povertyactionlab.org",
     9, "T06; T07"),
    ("Work with men and boys social norms gender relations prevention", WEB, 10, "T07"),
    ("Bandebereho couples intervention Rwanda randomized controlled trial", WEB, 10, "T07; T24"),
    ("Mass media experiment violence against women disclosure Uganda", WEB, 10, "T08"),
    ("School-based experiment adolescents gender attitudes India", WEB, 10, "T23"),
    ("School-based interventions dating sexual violence systematic review", WEB, 10, "T09"),
    ("Sexual offender treatment recidivism Campbell systematic review", WEB, 10, "T10"),
    ("Interventions preventing unintended pregnancies adolescents Cochrane", WEB, 10, "T11"),
    ("Effective strategies adolescent sexual and reproductive health services", WEB, 10, "T11"),
    ("Advance provision emergency contraception Cochrane", WEB, 10, "T11"),
    ("Gender and power in sexuality and HIV education review", WEB, 10, "T12"),
    ("School based sex education HIV prevention low- and middle-income countries meta-analysis", WEB, 10, "T12"),
    ("Experimental evidence female voting behavior Pakistan", WEB, 10, "T13"),
    ("Gender quotas women's political leadership", WEB, 10, "T13"),
    ("Impact evaluation training women candidates political parties", WEB, 10, "T13"),
    ("Childcare mothers' labor market outcomes lower- and middle-income countries", WEB, 10, "T14"),
    ("Women's economic empowerment evidence gap map land rights childcare business training",
     "Restringida a 3ieimpact.org y sus portales", 9, "T14; T15; T17; T18"),
    ("Urban property rights labor supply Peru", WEB, 10, "T15"),
    ("Land property rights interventions investment agricultural productivity systematic review", WEB, 10, "T15"),
    ("Gender composition forest fishery management groups systematic map", WEB, 10, "T16"),
    ("Meta-analysis active labor market program evaluations", WEB, 10, "T17"),
    ("Youth employment programs labor market outcomes quantitative review", WEB, 10, "T17"),
    ("Occupational training gender segregation ProJoven Peru", WEB, 10, "T17"),
    ("Recognition of prior learning skills certification impact evaluation", WEB, 10, "T17"),
    ("Business training entrepreneurship evaluations developing world", WEB, 10, "T18"),
    ("Six randomized evaluations of microcredit", WEB, 10, "T18"),
    ("Gender matters economic empowerment interventions research review", WEB, 10, "T18"),
    ("Employer-based programs breastfeeding working mothers systematic review", WEB, 10, "T19"),
    ("Cash transfer programmes schooling outcomes systematic review", WEB, 10, "T20"),
    ("Evaluación de impacto programa Juntos Perú educación", WEB, 10, "T20"),
    ("Gender budgeting fiscal context outcomes", WEB, 10, "T21"),
    ("Gender mainstreaming technical assistance public sector effectiveness systematic review", WEB, 10, "T21"),
    ("Evaluación MEF programa presupuestal 0080 violencia familiar EDEP", WEB, 13, "T01; T03"),
    ("Gender statistics time-use surveys policy influence evidence", WEB, 10, "T22"),
]

# Candidatas revisadas que no entran a la matriz: referencia, tipos, enlace, motivo
CRIBADO = [
    ("O'Brien, D. Z. y Rickne, J. (2016). Gender quotas and women's political leadership. American Political Science "
     "Review, 110(1), 112–126.", "T13", "https://www.cambridge.org/core/product/0C8CC7B51CC37871A1B1FD925A8AA976",
     "Intervención distinta (cuota partidaria), en un país de ingreso alto; no comparable con la asistencia técnica."),
    ("MEF (2015). Evaluación de Diseño y Ejecución Presupuestal del Programa Presupuestal Lucha contra la Violencia "
     "Familiar. Panel independiente: V. Carrasco, I. Bustamante y A. Gushiken.", "T01; T03",
     "https://mef.gob.pe/contenidos/presu_publ/ppr/eval_indep/EDEP_violencia_familiar.pdf",
     "Evaluación de diseño y ejecución sin estimación de efectos; se ubicó el documento de difusión y no se leyó el "
     "informe final. Pendiente de lectura para el informe."),
    ("Perova, E. y Vakis, R. (2012). 5 years in Juntos: New evidence on the program's short and long-term impacts. "
     "Economía (PUCP), n.° 69, 53–82.", "T20", "https://econpapers.repec.org/RePEc:pcp:pucrev:y:2012:i:69:p:53-82",
     "Se ubicó la referencia, pero no se verificaron sus resultados; pendiente de lectura del texto completo."),
    ("Abramsky, T. et al. (2016). Seguimiento de SASA! sobre experiencias de violencia de las mujeres, Journal of "
     "Epidemiology and Community Health, 70(8), 818.", "T06", "https://jech.bmj.com/content/70/8/818.abstract",
     "Análisis secundario del mismo ensayo ya incluido; no se cuenta dos veces."),
    ("Leal (2019). Disertación de maestría sobre capacitación policial en violencia intrafamiliar, Curitiba.", "T05",
     "https://tede.utp.br/jspui/handle/tede/1659", "Tesis no arbitrada; se excluye por tipo de documento."),
    ("Ensayos sobre capacitación de candidatas en Zambia y Brasil (registros de ensayos).", "T13",
     "https://www.socialscienceregistry.org/trials/15617", "Registros de ensayos sin resultados publicados."),
]

# Evidencia incluida: id, tipos, referencia, enlace, verificación, diseño, ámbito, año, calidad según la propia fuente,
# resultado principal, dirección
V = f"{FECHA}: título, autores, año y fuente comprobados en "
EVIDENCIA = [
    ("E01", "T01; T03", "Sviatschi, M. M. y Trako, I. (2024). Gender violence, enforcement, and human capital: Evidence "
     "from women's justice centers in Peru. Journal of Development Economics. Versión previa: Banco Mundial, Policy "
     "Research Working Paper 9624 (2021).", "https://econpapers.repec.org/paper/wbkwbrwps/9624.htm",
     V + "RePEc y Banco Mundial (Open Knowledge Repository).", "Cuasi experimental", "Perú", 2024, "No informada",
     "La apertura escalonada de los Centros Emergencia Mujer redujo en cerca de 10 % la violencia doméstica "
     "autorreportada, las muertes de mujeres por agresión y las hospitalizaciones por salud mental; aumentaron las "
     "denuncias y los procesamientos por delitos de género.", "Favorable"),
    ("E02", "T01", "Rivas, C. et al. (2015). Advocacy interventions to reduce or eliminate violence and promote the "
     "physical and psychosocial well-being of women who experience intimate partner abuse. Cochrane Database of "
     "Systematic Reviews, CD005043.pub3.", "https://cochranelibrary.com/cdsr/doi/10.1002/14651858.CD005043.pub3/information",
     V + "Cochrane Library.", "Revisión sistemática", "Global", 2015, "Muy baja",
     "La abogacía intensiva redujo la violencia física a los 24 meses en mujeres que salían de albergues, pero no a los 12 "
     "ni a los 36; la abogacía breve tendría beneficios pequeños de corto plazo.", "Mixto"),
    ("E03", "T02", "Hameed, M. et al. (2020). Psychological therapies for women who experience intimate partner "
     "violence. Cochrane Database of Systematic Reviews, CD013017.", "https://pubmed.ncbi.nlm.nih.gov/32608505/",
     V + "PubMed.", "Revisión sistemática", "Global", 2020, "Moderada",
     "Las terapias psicológicas probablemente reducen la depresión (DME −0,24; IC 95 % −0,47 a −0,01; certeza moderada); "
     "la evidencia es incierta para otros resultados.", "Favorable"),
    ("E04", "T03", "Perova, E. y Reynolds, S. A. (2017). Women's police stations and intimate partner violence: "
     "Evidence from Brazil. Social Science & Medicine, 174, 188–196.",
     "https://ideas.repec.org/a/eee/socmed/v174y2017icp188-196.html", V + "RePEc.", "Cuasi experimental",
     "Ingreso medio (Brasil)", 2017, "No informada",
     "Sin asociación promedio entre las comisarías de mujeres y los homicidios de mujeres; reducción entre mujeres "
     "jóvenes de áreas metropolitanas.", "Mixto"),
    ("E05", "T03", "Amaral, S., Bhalotra, S. y Prakash, N. (2021). Gender, crime and punishment: Evidence from women "
     "police stations in India. IZA Discussion Paper 14250.", "https://docs.iza.org/dp14250.pdf",
     V + "IZA y RePEc.", "Cuasi experimental", "Ingreso medio (India)", 2021, "No informada",
     "La apertura de comisarías de mujeres aumentó en 29 % las denuncias de delitos contra mujeres, sobre todo de "
     "violencia doméstica, sin cambios en feminicidios ni en la violencia autorreportada.", "Favorable"),
    ("E06", "T04", "Roehling, M. V. y Huang, J. (2018). Sexual harassment training effectiveness: An interdisciplinary "
     "review and call for research. Journal of Organizational Behavior, 39(2), 134–150.",
     "https://www.deepdyve.com/lp/wiley/sexual-harassment-training-effectiveness-an-interdisciplinary-review-u0FexYv0Ob",
     V + "DeepDyve y Peeref.", "Revisión narrativa", "Ingreso alto (principalmente Estados Unidos)", 2018, "No informada",
     "Revisión interdisciplinaria que encuentra pocos estudios empíricos sobre la efectividad de la capacitación y "
     "propone una agenda de investigación.", "Mixto"),
    ("E07", "T05", "Dowling, C. (2024). Police training in responding to family, domestic and sexual violence. Trends & "
     "issues in crime and criminal justice n.° 689. Australian Institute of Criminology.",
     "https://aic.gov.au/sites/default/files/2024-04/ti689_police_training_in_responding_to_fdsv.pdf",
     V + "Biblioteca Nacional de Australia y KrimDok.", "Revisión narrativa", "Ingreso alto (Australia e internacional)",
     2024, "No informada", "La policía es receptiva a la capacitación y la de componentes prácticos es la más "
     "prometedora; hay poca evidencia de cambios en la respuesta o en las víctimas.", "Mixto"),
    ("E08", "T06", "Leight, J., Cullen, C., Ranganathan, M. y Yakubovich, A. (2023). Effectiveness of community "
     "mobilisation and group-based interventions for preventing intimate partner violence against women in low- and "
     "middle-income countries: A systematic review and meta-analysis. Journal of Global Health, 13, 04115.",
     "https://pmc.ncbi.nlm.nih.gov/articles/PMC10588291/", V + "PubMed Central.", "Revisión sistemática",
     "Ingreso bajo y medio", 2023, "Ensayos aleatorios",
     "En 27 ensayos aleatorios, las intervenciones comunitarias y grupales redujeron la violencia de pareja (OR ajustado "
     "0,78; IC 95 % 0,63 a 0,97), con alta heterogeneidad.", "Favorable"),
    ("E09", "T06", "Abramsky, T. et al. (2014). Findings from the SASA! Study: A cluster randomized controlled trial to "
     "assess the impact of a community mobilization intervention to prevent violence against women and reduce HIV risk "
     "in Kampala, Uganda. BMC Medicine, 12, 122.", "https://bmcmedicine.biomedcentral.com/track/pdf/10.1186/s12916-014-0122-5",
     V + "BMC Medicine y LSHTM.", "Experimental", "Ingreso bajo (Uganda)", 2014, "Ensayo aleatorio",
     "Menor aceptación social de la violencia de pareja entre mujeres (RR ajustado 0,54); la reducción de la violencia "
     "física no fue estadísticamente significativa.", "Mixto"),
    ("E10", "T06", "Ellsberg, M. et al. (2015). Prevention of violence against women and girls: What does the evidence "
     "say? The Lancet, 385, 1555–1566.", "https://hsrc.himmelfarb.gwu.edu/sphhs_global_facpubs/121",
     V + "repositorio de la George Washington University y gov.uk.", "Revisión narrativa", "Global", 2015, "No informada",
     "La evidencia de prevención en países de ingreso bajo y medio favorece intervenciones participativas que involucran "
     "a varios actores; la mayor parte de la evidencia proviene de países de ingreso alto.", "Favorable"),
    ("E11", "T07", "Jewkes, R., Flood, M. y Lang, J. (2015). From work with men and boys to changes of social norms and "
     "reduction of inequities in gender relations: A conceptual shift in prevention of violence against women and "
     "girls. The Lancet, 385(9977), 1580–1589.", "https://ro.uow.edu.au/lhapapers/2063",
     V + "repositorio de la University of Wollongong.", "Revisión narrativa", "Global", 2015, "No informada",
     "Los programas solo con hombres muestran sobre todo cambios de actitudes y poca evidencia de reducción de la "
     "violencia; recomienda trabajar con hombres y mujeres para cambiar normas.", "Mixto"),
    ("E12", "T07; T24", "Doyle, K. et al. (2018). Gender-transformative Bandebereho couples' intervention to promote male "
     "engagement in reproductive and maternal health and violence prevention in Rwanda: Findings from a randomized "
     "controlled trial. PLoS ONE.", "https://doaj.org/article/246254d631d8465b87348fa64375caf3",
     V + "DOAJ y ClinicalTrials.gov (NCT02694627).", "Experimental", "Ingreso bajo (Ruanda)", 2018, "Ensayo aleatorio",
     "Menor violencia física (OR 0,37) y sexual (OR 0,34) de pareja; mayor participación de los hombres en el cuidado "
     "infantil y las tareas domésticas.", "Favorable"),
    ("E13", "T08", "Green, D. P., Wilke, A. M. y Cooper, J. (2020). Countering violence against women by encouraging "
     "disclosure: A mass media experiment in rural Uganda. Comparative Political Studies, 53(14), 2283–2320.",
     "https://gap.hks.harvard.edu/countering-violence-against-women-encouraging-disclosure-mass-media-experiment-rural-uganda",
     V + "Harvard Kennedy School (Gender Action Portal) y vLex.", "Experimental", "Ingreso bajo (Uganda)", 2020,
     "Ensayo aleatorio", "Sin cambio en actitudes; mayor disposición a denunciar y reducción de hasta 5 puntos "
     "porcentuales en la violencia contra mujeres del hogar en seis meses.", "Favorable"),
    ("E14", "T23", "Dhar, D., Jain, T. y Jayachandran, S. (2022). Reshaping adolescents' gender attitudes: Evidence from "
     "a school-based experiment in India. American Economic Review, 112(3), 899–927.", "https://www.nber.org/papers/w25331",
     V + "NBER y CiNii.", "Experimental", "Ingreso medio (India)", 2022, "Ensayo aleatorio",
     "Actitudes más favorables a la igualdad (0,18 desviaciones estándar) y conducta autorreportada más igualitaria; los "
     "efectos persisten dos años después.", "Favorable"),
    ("E15", "T09", "De La Rue, L., Polanin, J. R., Espelage, D. L. y Pigott, T. D. (2014). School-based interventions to "
     "reduce dating and sexual violence: A systematic review. Campbell Systematic Reviews, 10(1).",
     "https://doi.org/10.4073/csr.2014.7", V + "Campbell Collaboration y RePEc.", "Revisión sistemática",
     "Ingreso alto (Estados Unidos y Canadá)", 2014, "Baja",
     "Mejoras en conocimientos y actitudes sobre la violencia en el enamoramiento, sostenidas en el seguimiento; poco "
     "efecto en la conducta.", "Mixto"),
    ("E16", "T10", "Schmucker, M. y Lösel, F. (2017). Sexual offender treatment for reducing recidivism among convicted "
     "sex offenders: A systematic review and meta-analysis. Campbell Systematic Reviews, 13(1), 1–75.",
     "https://ideas.repec.org/a/wly/camsys/v13y2017i1p1-75.html", V + "RePEc y KrimDok.", "Revisión sistemática",
     "Ingreso alto (principalmente Norteamérica)", 2017, "No informada",
     "Reincidencia sexual de 10,1 % en tratados frente a 13,7 % sin tratamiento; efectos en la comunidad y en hospitales "
     "forenses, evidencia insuficiente para el tratamiento en prisiones.", "Mixto"),
    ("E17", "T11", "Oringanje, C. et al. (2016). Interventions for preventing unintended pregnancies among adolescents. "
     "Cochrane Database of Systematic Reviews, CD005215.pub3.",
     "https://cochranelibrary.com/cdsr/doi/10.1002/14651858.CD005215.pub3/information", V + "Cochrane Library.",
     "Revisión sistemática", "Global", 2016, "No informada",
     "Solo la combinación de educación y promoción de anticonceptivos redujo el embarazo no planificado a mediano y "
     "largo plazo (RR 0,66; IC 95 % 0,50 a 0,87).", "Favorable"),
    ("E18", "T11", "Denno, D. M., Hoopes, A. J. y Chandra-Mouli, V. (2015). Effective strategies to provide adolescent "
     "sexual and reproductive health services and to increase demand and community support. Journal of Adolescent "
     "Health, 56, S22–S41.", "https://pmc.ncbi.nlm.nih.gov/articles/PMC4570008", V + "PubMed Central y 3ie.",
     "Revisión narrativa", "Ingreso bajo y medio", 2015, "Baja (según 3ie)",
     "Revisión descriptiva de iniciativas para mejorar el acceso y el uso de servicios de salud sexual y reproductiva "
     "por adolescentes.", "Favorable"),
    ("E19", "T11", "Polis, C. B. et al. (2007). Advance provision of emergency contraception for pregnancy prevention. "
     "Cochrane Database of Systematic Reviews, CD005497.pub2.",
     "https://www.cochranelibrary.com/cdsr/doi/10.1002/14651858.CD005497.pub2/pdf/CDSR/CD005497/CD005497_abstract.pdf",
     V + "Cochrane Library.", "Revisión sistemática", "Global", 2007, "No informada",
     "La provisión anticipada de anticoncepción de emergencia aumentó y adelantó su uso, sin reducir los embarazos a "
     "nivel poblacional ni aumentar conductas de riesgo.", "Mixto"),
    ("E20", "T12", "Haberland, N. A. (2015). The case for addressing gender and power in sexuality and HIV education: A "
     "comprehensive review of evaluation studies. International Perspectives on Sexual and Reproductive Health, 41(1), "
     "31–42.", "https://knowledgecommons.popcouncil.org/departments_sbsr-pgy/776", V + "Population Council y Guttmacher.",
     "Revisión sistemática", "Global", 2015, "No informada",
     "De 22 programas evaluados, 80 % de los que abordaban género y poder redujeron ITS o embarazos no planificados, "
     "frente a 17 % de los que no lo hacían.", "Favorable"),
    ("E21", "T12", "Fonner, V. A., Armstrong, K. S., Kennedy, C. E., O'Reilly, K. R. y Sweat, M. D. (2014). School based "
     "sex education and HIV prevention in low- and middle-income countries: A systematic review and meta-analysis. PLoS "
     "ONE, 9(3), e89692.", "https://econpapers.repec.org/RePEc:plo:pone00:0089692", V + "DOAJ y RePEc.",
     "Revisión sistemática", "Ingreso bajo y medio", 2014, "No informada",
     "Mayor conocimiento sobre VIH (g = 0,63), más uso de condón, menos parejas sexuales e inicio sexual más tardío.",
     "Favorable"),
    ("E22", "T13", "Giné, X. y Mansuri, G. (2018). Together we will: Experimental evidence on female voting behavior in "
     "Pakistan. American Economic Journal: Applied Economics, 10(1), 207–235.",
     "https://econpapers.repec.org/RePEc:aea:aejapp:v:10:y:2018:i:1:p:207-35", V + "RePEc y AEA.", "Experimental",
     "Ingreso medio (Pakistán)", 2018, "Ensayo aleatorio",
     "Una campaña de información electoral para mujeres aumentó en 11 puntos porcentuales su participación electoral, "
     "con efectos de difusión a mujeres no atendidas.", "Favorable"),
    ("E23", "T14", "Halim, D., Perova, E. y Reynolds, S. (2023). Childcare and mothers' labor market outcomes in lower- "
     "and middle-income countries. The World Bank Research Observer, 38(1), 73–114.",
     "https://econpapers.repec.org/RePEc:oup:wbrobs:v:38:y:2023:i:1:p:73-114.", V + "RePEc y Banco Mundial.",
     "Revisión sistemática", "Ingreso bajo y medio", 2023, "No informada",
     "De 22 estudios causales, todos menos uno hallan efectos positivos del cuidado infantil institucional en la "
     "participación laboral o las horas de trabajo de las madres.", "Favorable"),
    ("E24", "T15", "Lawry, S., Samii, C., Hall, R., Leopold, A., Hornby, D. y Mtero, F. (2017). The impact of land "
     "property rights interventions on investment and agricultural productivity in developing countries: A systematic "
     "review. Journal of Development Effectiveness, 9(1), 61–81.", "https://uwcscholar.uwc.ac.za/handle/10566/4115",
     V + "repositorio de la University of the Western Cape y Global Landscapes Forum.", "Revisión sistemática",
     "Ingreso bajo y medio", 2017, "No informada",
     "Ganancias de productividad e ingresos por el reconocimiento de la tenencia, variables por región; los estudios "
     "cualitativos advierten posibles efectos negativos sobre los derechos de las mujeres.", "Mixto"),
    ("E25", "T15", "Field, E. (2007). Entitled to work: Urban property rights and labor supply in Peru. The Quarterly "
     "Journal of Economics, 122(4), 1561–1602.", "https://econpapers.repec.org/RePEc:oup:qjecon:v:122:y:2007:i:4:p:1561-1602.",
     V + "RePEc y Duke University.", "Cuasi experimental", "Perú", 2007, "No informada",
     "La titulación urbana de COFOPRI aumentó las horas de trabajo, desplazó el trabajo del hogar al mercado y redujo el "
     "trabajo infantil.", "Favorable"),
    ("E26", "T16", "Leisher, C. et al. (2016). Does the gender composition of forest and fishery management groups affect "
     "resource governance and conservation outcomes? A systematic map. Environmental Evidence, 5, 6.",
     "https://environmentalevidencejournal.biomedcentral.com/track/pdf/10.1186/s13750-016-0057-8",
     V + "Environmental Evidence y University of Minnesota.", "Revisión sistemática", "Ingreso bajo y medio", 2016, "Baja",
     "Los 17 estudios incluidos identifican mejoras en la gobernanza local y tres en la conservación cuando participan "
     "las mujeres; solo dos son de alta calidad.", "Favorable"),
    ("E27", "T17; T18", "Chinen, M., De Hoop, T., Balarin, M., Alcázar, L., Sennett, J. y Mezarina, J. (2018). "
     "Vocational and business training to improve women's labour market outcomes in low- and middle-income countries: A "
     "systematic review. 3ie Systematic Review 40.", "https://doi.org/10.23846/SR71094",
     V + "el informe de 3ie (página de créditos y resumen).", "Revisión sistemática", "Ingreso bajo y medio", 2018,
     "Mixta (14 de 35 estudios con bajo riesgo de sesgo de selección)",
     "La formación vocacional aumentó 11 % la probabilidad de empleo y 5,5 % los ingresos de las mujeres; la formación "
     "empresarial tuvo efectos pequeños en ventas y ganancias (0,05 DE fuera de África).", "Favorable"),
    ("E28", "T17", "Kluve, J. et al. (2019). Do youth employment programs improve labor market outcomes? A quantitative "
     "review. World Development, 114, 237–253.", "https://econpapers.repec.org/RePEc:eee:wdevel:v:114:y:2019:i:c:p:237-253",
     V + "RePEc y OIT.", "Revisión sistemática", "Global", 2019, "No informada",
     "En 113 evaluaciones, efectos pequeños en empleo (g = 0,05) e ingresos (g = 0,04), mayores en países de ingreso bajo "
     "y medio y en programas que combinan servicios.", "Favorable"),
    ("E29", "T17", "Card, D., Kluve, J. y Weber, A. (2018). What works? A meta analysis of recent active labor market "
     "program evaluations. Journal of the European Economic Association, 16(3), 894–931.", "https://www.nber.org/papers/w21431",
     V + "NBER y RePEc.", "Revisión sistemática", "Global", 2018, "No informada",
     "Impactos cercanos a cero en el corto plazo y positivos a dos o tres años; mayores en programas de capital humano y "
     "para mujeres.", "Favorable"),
    ("E30", "T17", "Nakata, S., Sharma, U., Rahman, T., Rahman, M. y Aziz, M.-U. (2021). Effects of recognition of prior "
     "learning on job market outcomes: Impact evaluation in Bangladesh. Banco Mundial, Policy Research Working Paper "
     "9644. El orden de autores difiere entre el repositorio del Banco Mundial y RePEc.",
     "https://ideas.repec.org/p/wbk/wbrwps/9644.html", V + "RePEc y Banco Mundial.", "Experimental",
     "Ingreso medio (Bangladés)", 2021, "Ensayo aleatorio",
     "La certificación de competencias mejoró el empleo, los salarios y la formalidad; las mujeres ganan más que los "
     "hombres con el certificado.", "Favorable"),
    ("E31", "T17", "Ñopo, H., Robles, M. y Saavedra, J. (2008). Occupational training to reduce gender segregation: The "
     "impacts of ProJoven. Economía (PUCP), n.° 62, 33–54.",
     "https://econpapers.repec.org/RePEc:pcp:pucrev:y:2008:i:62:p:33-54", V + "GRADE y RePEc.", "Cuasi experimental",
     "Perú", 2008, "No informada", "A los 18 meses, el empleo de las mujeres aumentó cerca de 15 % y la segregación "
     "ocupacional por género cayó 30 %.", "Favorable"),
    ("E32", "T18", "McKenzie, D. y Woodruff, C. (2014). What are we learning from business training and entrepreneurship "
     "evaluations around the developing world? The World Bank Research Observer, 29(1), 48–82.",
     "https://econpapers.repec.org/RePEc:cpr:ceprdp:9564", V + "RePEc y Unión Africana.", "Revisión narrativa",
     "Ingreso bajo y medio", 2014, "Baja (muestras pequeñas y seguimiento corto)",
     "Efectos modestos en la supervivencia de empresas existentes; mejor evidencia para la creación de nuevos negocios; "
     "pocos estudios hallan efectos en ganancias.", "Mixto"),
    ("E33", "T18", "Banerjee, A., Karlan, D. y Zinman, J. (2015). Six randomized evaluations of microcredit: Introduction "
     "and further steps. American Economic Journal: Applied Economics, 7(1), 1–21.", "https://dspace.mit.edu/handle/1721.1/95940",
     V + "MIT DSpace y RePEc.", "Experimental", "Ingreso bajo y medio", 2015, "Ensayos aleatorios",
     "Seis evaluaciones aleatorias en seis países muestran efectos modestamente positivos, no transformadores.", "Mixto"),
    ("E34", "T18", "Buvinić, M. y O'Donnell, M. (2019). Gender matters in economic empowerment interventions: A research "
     "review. The World Bank Research Observer, 34(2), 309–346.",
     "https://econpapers.repec.org/RePEc:oup:wbrobs:v:34:y:2019:i:2:p:309-346.", V + "RePEc y Banco Mundial.",
     "Revisión narrativa", "Ingreso bajo y medio", 2019, "No informada",
     "Los servicios financieros y de formación no son neutrales al género; el ahorro, los programas de graduación y la "
     "formación con estipendio o cuidado infantil mejoran los resultados de las mujeres.", "Favorable"),
    ("E35", "T19", "Dinour, L. M. y Szaro, J. M. (2017). Employer-based programs to support breastfeeding among working "
     "mothers: A systematic review. Breastfeeding Medicine, 12(3), 131–141.",
     "https://digitalcommons.montclair.edu/nutr-foodstudies-facpubs/115", V + "Montclair State University.",
     "Revisión sistemática", "Global", 2017, "No informada",
     "En 22 artículos de 10 países, la mayoría reporta al menos un resultado positivo en la lactancia o en la "
     "satisfacción laboral; el lactario es la medida más frecuente.", "Favorable"),
    ("E36", "T20", "Baird, S., Ferreira, F. H. G., Özler, B. y Woolcock, M. (2014). Conditional, unconditional and "
     "everything in between: A systematic review of the effects of cash transfer programmes on schooling outcomes. "
     "Journal of Development Effectiveness, 6(1), 1–43.", "https://econpapers.repec.org/RePEc:taf:jdevef:v:6:y:2014:i:1:p:1-43",
     V + "RePEc y Banco Mundial.", "Revisión sistemática", "Ingreso bajo y medio", 2014, "No informada",
     "Las transferencias condicionadas y no condicionadas aumentan la matrícula y la asistencia; el efecto es mayor con "
     "condiciones monitoreadas y pequeño en aprendizajes.", "Favorable"),
    ("E37", "T20", "Perova, E. y Vakis, R. (2009). El impacto y potencial del Programa Juntos en Perú: evidencia de una "
     "evaluación no experimental. Banco Mundial.", "https://www.mef.gob.pe/contenidos/pol_econ/documentos/Perova_Vakis_JuntosIE.pdf",
     V + "MEF y repositorio del MINEDU.", "Cuasi experimental", "Perú", 2009, "No informada",
     "Impactos positivos en varias dimensiones de bienestar; en educación, efectos en matrícula y asistencia concentrados "
     "en los años de transición; sin impacto en nutrición a dos años.", "Favorable"),
    ("E38", "T21", "Stotsky, J. G. (2016). Gender budgeting: Fiscal context and current outcomes. IMF Working Paper "
     "WP/16/149.", "https://www.imf.org/en/Publications/WP/Issues/2016/12/31/Gender-Budgeting-Fiscal-Context-and-Current-Outcomes-44132",
     V + "FMI.", "Descriptivo", "Global", 2016, "No aplica",
     "Revisión de experiencias de presupuesto con enfoque de género por región, sin estimaciones causales de sus efectos.",
     "Sin estimación"),
    ("E39", "T22", "Data2X (2018). Invisible no more? Country case studies (Volume 2).",
     "https://data2x.org/resource-center/invisible-no-more-country-case-studies-volume-2/", V + "Data2X.",
     "Descriptivo", "Global", 2018, "No aplica",
     "Dieciocho estudios de caso muestran que los datos de uso del tiempo informaron políticas laborales, sociales y de "
     "cuidado.", "Favorable"),
]

# Búsqueda académica (rev_evi_busqueda_academica.py): fuentes incluidas tras leer el resumen de los registros
# seleccionados por título. Datos bibliográficos comprobados en Crossref, PubMed o la página de la revista o del
# repositorio; resultado tomado del resumen.
VC = f"{FECHA}: título, autores, año y fuente comprobados en Crossref (DOI); resultado tomado del resumen."
VP = f"{FECHA}: título, autores, año y fuente comprobados en PubMed y Crossref; resultado tomado del resumen."
VR = lambda donde: f"{FECHA}: título, autores, año y fuente comprobados en {donde}; resultado tomado del resumen."
EAS = "Ensayo aleatorio"
EAC = "Ensayo aleatorio por conglomerados"
ACADEMICA = [
    ("E40", "T01", "Emezue, C., Chase, J. D., Udmuangpia, T. y Bloom, T. L. (2022). Technology-based and digital "
     "interventions for intimate partner violence: A systematic review and meta-analysis. Campbell Systematic Reviews, "
     "18(3), e1271.", "https://doi.org/10.1002/cl2.1271", VC, "Revisión sistemática", "No informado en el resumen", 2022,
     "No informada", "Metaanálisis de ensayos aleatorios: las intervenciones digitales para sobrevivientes de violencia de "
     "pareja tienen efectos pequeños y favorables en salud mental, seguridad y victimización; se recomiendan como "
     "complemento de la atención presencial coordinada, no como reemplazo.", "Favorable"),
    ("E41", "T01", "Van Parys, A.-S., Verhamme, A., Temmerman, M. y Verstraelen, H. (2014). Intimate partner violence and "
     "pregnancy: A systematic review of interventions. PLoS ONE, 9(1), e85084.",
     "https://doi.org/10.1371/journal.pone.0085084", VP, "Revisión sistemática", "No informado en el resumen", 2014,
     "No informada", "Pocos ensayos aleatorios; cinco reportaron menor violencia de pareja en el periodo perinatal (OR de "
     "0,47 a 0,92) y la evidencia en salud mental es limitada; la revisión concluye que falta evidencia sólida.", "Mixto"),
    ("E42", "T01", "Bair-Merritt, M. H., Lewis-O'Connor, A., Goel, S., Amato, P., Ismailji, T., Jelley, M., Lenahan, P. y "
     "Cronholm, P. (2014). Primary care-based interventions for intimate partner violence: A systematic review. American "
     "Journal of Preventive Medicine, 46(2), 188–194.", "https://doi.org/10.1016/j.amepre.2013.10.001", VP,
     "Revisión sistemática", "No informado en el resumen", 2014, "No informada",
     "De 17 estudios, 13 muestran al menos un beneficio: seis de once que midieron la persistencia de la violencia "
     "hallaron reducciones y seis de diez, más derivaciones a servicios comunitarios.", "Favorable"),
    ("E43", "T02", "Velecela-Chumbi, E., Jaramillo-Oyervide, J. y Polo-Martínez, E. (2025). Terapia cognitivo conductual en "
     "el trastorno de estrés postraumático en mujeres víctimas de violencia de pareja: revisión sistemática. 593 Digital "
     "Publisher CEIT, 10(6), 77–92.", "https://doi.org/10.33386/593dp.2025.6.3596", VC, "Revisión sistemática",
     "No informado en el resumen", 2025, "No informada (cinco estudios, incluye estudios de caso)",
     "Entre 60 % y 80 % de las participantes mostró una mejoría clínicamente significativa del estrés postraumático, "
     "sostenida cerca de seis meses.", "Favorable"),
    ("E44", "T03", "Johnson, L. y Stylianou, A. M. (2022). Coordinated community responses to domestic violence: A "
     "systematic review of the literature. Trauma, Violence, & Abuse, 23(2), 506–522.",
     "https://doi.org/10.1177/1524838020957984", VP, "Revisión sistemática", "Ingreso alto (Estados Unidos)", 2022,
     "No informada", "Dieciocho estudios con diseño experimental o cuasi experimental; la variedad de intervenciones y de "
     "resultados impide concluir sobre la efectividad de las respuestas comunitarias coordinadas con la policía.", "Mixto"),
    ("E45", "T03", "Connor, R. A., Johnson, L., Bridgeman, M., Shokraneh, F. y Hakobyan, B. (2025). Outcomes associated "
     "with arrest for domestic violence: A systematic review and meta-analysis. Trauma, Violence, & Abuse, 26(3), "
     "528–545.", "https://doi.org/10.1177/15248380241284777", VP, "Revisión sistemática", "Ingreso alto (Estados Unidos)",
     2025, "No informada", "En 34 estudios, la detención del agresor no reduce de forma consistente la violencia repetida.",
     "Mixto"),
    ("E46", "T03", "Cordier, R., Chung, D., Wilkes-Gillan, S. y Speyer, R. (2021). The effectiveness of protection orders "
     "in reducing recidivism in domestic violence: A systematic review and meta-analysis. Trauma, Violence, & Abuse, "
     "22(4), 804–828.", "https://doi.org/10.1177/1524838019882361", VP, "Revisión sistemática",
     "No informado en el resumen", 2021, "Baja (la mayoría de los estudios con alto riesgo de sesgo, según la revisión)",
     "En 25 estudios, el incumplimiento de las órdenes de protección es menor cuando se combinan con la detención del "
     "agresor; no hay consenso sobre cómo medir su efectividad.", "Mixto"),
    ("E47", "T03", "Acurio, G., Montes, Á., Cea, B. y Novo, M. (2023). Efectos de las medidas de protección y "
     "asistenciales a víctimas de violencia de género sobre el bienestar psicológico. Acción Psicológica, 20(2).",
     "https://doi.org/10.5944/ap.20.2.39358", VR("Crossref (DOI) y el repositorio de la Universidade de Santiago de "
                                                  "Compostela"), "Descriptivo", "Ingreso medio (Ecuador)", 2023,
     "No aplica", "En 136 mujeres que concluyeron un proceso judicial, las medidas de protección y asistenciales no "
     "tuvieron efecto en el bienestar psicológico; el proceso judicial sí se asoció con mejoras grandes.", "Nulo"),
    ("E48", "T05", "Amaral, S., Borker, G., Prakash, N. y Sviatschi, M. M. (2026). Debiasing law enforcement officers: "
     "Evidence from an expressive arts intervention in India. CESifo Working Paper n.° 13050.",
     "https://doi.org/10.65864/9gm44fw644", VC, "Experimental", "Ingreso medio (India)", 2026,
     "Ensayo aleatorio por conglomerados (417 comisarías)", "Una intervención de tres días con teatro del oprimido mejoró "
     "las actitudes de los oficiales ante la violencia de género; en auditorías con víctimas simuladas fueron 15 puntos "
     "porcentuales menos propensos a culpar a la víctima, y los registros muestran 41 % menos demoras en el registro de "
     "denuncias.", "Favorable"),
    ("E49", "T04", "Tinkler, J. E., Clay-Warner, J. y Alinor, M. (2022). Sexual harassment training and men's motivation "
     "to work with women. Social Science Research, 107, 102740.", "https://doi.org/10.1016/j.ssresearch.2022.102740", VP,
     "Experimental", "No informado en el resumen (estudiantes universitarios)", 2022, "Experimento de laboratorio",
     "La capacitación sobre la política de hostigamiento no cambió la elección de pareja de trabajo, pero llevó a los "
     "hombres a percibir a las mujeres como más distintas a ellos.", "Mixto"),
    ("E50", "T04", "Glass, N., Hanson, G. C., Anger, W. K., Laharnar, N., Campbell, J. C., Weinstein, M. y Perrin, N. "
     "(2017). Computer-based training (CBT) intervention reduces workplace violence and harassment for homecare workers. "
     "American Journal of Industrial Medicine, 60(7), 635–643.", "https://doi.org/10.1002/ajim.22728", VP, "Experimental",
     "No informado en el resumen (trabajadoras de cuidado domiciliario)", 2017,
     "Dos brazos de intervención, sin grupo sin intervención", "Ambos brazos aumentaron la confianza para prevenir la "
     "violencia y redujeron los incidentes de violencia y hostigamiento a los seis meses.", "Favorable"),
    ("E51", "T09", "Liu, L., Taylor, B. G., Liu, W. y Stein, N. D. (2024). The saturation- and dose-dependent effects of a "
     "teen sexual harassment prevention program: Findings from a randomized controlled trial. Prevention Science, 25(4), "
     "616–627.", "https://doi.org/10.1007/s11121-024-01661-2", VR("ERIC y Crossref (DOI)"), "Experimental",
     "Ingreso alto (Estados Unidos, Nueva York)", 2024, EAS, "El programa Shifting Boundaries redujo la victimización por "
     "hostigamiento sexual en 2104 estudiantes de secundaria, sin efecto en la perpetración.", "Mixto"),
    ("E52", "T09", "Walsh, K., Zwi, K., Woolfenden, S. y Shlonsky, A. (2015). School-based education programmes for the "
     "prevention of child sexual abuse: A systematic review. Campbell Systematic Reviews, 11(1).",
     "https://doi.org/10.4073/csr.2015.10", VC, "Revisión sistemática",
     "Global (Estados Unidos, Canadá, China, Alemania, Taiwán y Turquía)", 2015, "Moderada",
     "En 24 ensayos, los programas mejoran el conocimiento y las conductas de protección y aumentan la revelación del "
     "abuso; ningún estudio midió la reducción de su incidencia.", "Favorable"),
    ("E53", "T09", "Kettrey, H. H., Marx, R. A. y Tanner-Smith, E. E. (2019). Effects of bystander programs on the "
     "prevention of sexual assault among adolescents and college students: A systematic review. Campbell Systematic "
     "Reviews, 15(1-2).", "https://doi.org/10.4073/csr.2019.1", VC, "Revisión sistemática", "No informado en el resumen",
     2019, "Alta (27 estudios, 21 ensayos aleatorios)", "Los programas de testigos aumentan la intervención ante "
     "situaciones de riesgo, sin efecto en la perpetración de agresión sexual; los efectos en conocimientos y actitudes "
     "son inconsistentes.", "Mixto"),
    ("E54", "T06; T07", "Christofides, N. J., Hatcher, A. M., Rebombo, D., McBride, R.-S., Munshi, S., Pino, A., "
     "Abdelatif, N., Peacock, D., Levin, J. y Jewkes, R. K. (2020). Effectiveness of a multi-level intervention to reduce "
     "men's perpetration of intimate partner violence: A cluster randomised controlled trial. Trials, 21, 359.",
     "https://doi.org/10.1186/s13063-020-4185-7", VP, "Experimental", "Ingreso medio (Sudáfrica)", 2020, EAC,
     "La intervención comunitaria Sonke CHANGE tuvo un efecto limitado en la violencia de pareja ejercida por hombres de "
     "18 a 40 años en un asentamiento periurbano.", "Nulo"),
    ("E55", "T06; T08", "James, L. E., Welton-Mitchell, C., Michael, S., Santoadi, F., Shakirah, S., Hussin, H., Anwar, "
     "M., Kilzar, L. y James, A. (2021). Development and testing of a community-based intervention to address intimate "
     "partner violence among Rohingya and Syrian refugees: A social norms-based mental health-integrated approach. "
     "International Journal of Environmental Research and Public Health, 18(21), 11674.",
     "https://doi.org/10.3390/ijerph182111674", VP, "Experimental",
     "Ingreso medio (Malasia y Líbano, población refugiada)", 2021,
     "Talleres con medición antes y después; afiches evaluados con un ensayo aleatorio y una comparación por "
     "conglomerados", "Los talleres redujeron la aceptación de la violencia y las normas de género rígidas; con los "
     "afiches, las mujeres consideraron menos aceptable la violencia de pareja.", "Favorable"),
    ("E56", "T06; T07", "Shaw, B., Stevanovic-Fenn, N., Gibson, L., Davin, C., Chipanta, N. S. K., Lubin, A. B. et al. "
     "(2023). Shifting norms in faith communities to reduce intimate partner violence: Results from a cluster randomized "
     "controlled trial in Nigeria. Journal of Interpersonal Violence, 38(19-20), 10865–10899.",
     "https://doi.org/10.1177/08862605231176799", VP, "Experimental", "Ingreso medio (Nigeria)", 2023, EAC,
     "El programa Masculinity, Faith and Peace redujo en 61 % la probabilidad de que las mujeres reportaran violencia de "
     "pareja y mejoró las normas, las actitudes y la calidad de la relación.", "Favorable"),
    ("E57", "T07", "Pérez-Martínez, V., Marcos-Marcos, J., Cerdán-Torregrosa, A., Briones-Vozmediano, E., Sanz-Barbero, "
     "B. et al. (2023). Positive masculinities and gender-based violence educational interventions among young people: "
     "A systematic review. Trauma, Violence, & Abuse, 24(2), 468–486.", "https://doi.org/10.1177/15248380211030242", VP,
     "Revisión sistemática", "Global (10 de 15 estudios en África)", 2023, "No informada",
     "La mayoría de los 15 estudios (8 ensayos aleatorios) reporta menos violencia de género física o sexual tras "
     "intervenciones educativas sobre masculinidades con jóvenes.", "Favorable"),
    ("E58", "T07", "John, N. A., Adebayo, A., Boychuk, N. A. y OlaOlorun, F. (2022). Intimate partner violence (IPV) "
     "prevention using a cross-sectoral couple-based intervention: Results from a cluster randomised control trial in "
     "Ibadan, Nigeria. BMJ Global Health, 7(2), e007192.", "https://doi.org/10.1136/bmjgh-2021-007192", VP, "Experimental",
     "Ingreso medio (Nigeria)", 2022, EAC, "La socialización de género con parejas redujo la violencia física de pareja "
     "frente al control, sin efectos claros en la violencia emocional o sexual.", "Mixto"),
    ("E59", "T07", "Dunkle, K., Stern, E., Chatterji, S. y Heise, L. (2020). Effective prevention of intimate partner "
     "violence through couples training: A randomised controlled trial of Indashyikirwa in Rwanda. BMJ Global Health, "
     "5(12), e002439.", "https://doi.org/10.1136/bmjgh-2020-002439", VP, "Experimental", "Ingreso bajo (Ruanda)", 2020,
     EAS, "A los 24 meses, menos violencia física o sexual de pareja experimentada por las mujeres (RR ajustado 0,44) y "
     "ejercida por los hombres (0,54).", "Favorable"),
    ("E60", "T07", "Santhya, K. G., Jejeebhoy, S. J., Acharya, R., Pandey, N., Gogoi, A., Joshi, M. et al. (2019). "
     "Transforming the attitudes of young men about gender roles and the acceptability of violence against women, Bihar. "
     "Culture, Health & Sexuality, 21(12), 1409–1424.", "https://doi.org/10.1080/13691058.2019.1568574", VP,
     "Experimental", "Ingreso medio (India)", 2019, EAC, "Un programa de habilidades para la vida y deporte con enfoque de "
     "género cambió las actitudes de hombres jóvenes sobre los roles de género, el control y la violencia contra las "
     "mujeres.", "Favorable"),
    ("E61", "T07", "Gurbuz Cuneo, A., Vaillant, J., Koussoubé, E., Pierotti, R. S., Falb, K. y Kabeya, R. (2023). "
     "Prevention, cessation, or harm reduction: Heterogeneous effects of an intimate partner violence prevention program "
     "in eastern Democratic Republic of the Congo. PLOS ONE, 18(3), e0282339.",
     "https://doi.org/10.1371/journal.pone.0282339", VP, "Experimental", "Ingreso bajo (República Democrática del Congo)",
     2023, EAC, "Sin efecto promedio del programa de grupos de hombres EMAP en la violencia de pareja; reducción entre los "
     "hombres más violentos al inicio.", "Mixto"),
    ("E62", "T06", "Harvey, S., Abramsky, T., Mshana, G., Hansen, C. H., Mtolela, G. J., Madaha, F. et al. (2021). A "
     "cluster randomised controlled trial to evaluate the impact of a gender transformative intervention on intimate "
     "partner violence against women in newly formed neighbourhood groups in Tanzania. BMJ Global Health, 6(7), e004555.",
     "https://doi.org/10.1136/bmjgh-2020-004555", VP, "Experimental", "Ingreso medio (Tanzania)", 2021, EAC,
     "Menos abuso emocional y menor aceptación de la violencia, sin reducción de la violencia física o sexual de pareja.",
     "Mixto"),
    ("E63", "T08", "Christia, F., Larreguy, H., Parker-Magyar, E. y Quintero, M. (2023). Empowering women facing "
     "gender-based violence amid COVID-19 through media campaigns. Nature Human Behaviour, 7(10), 1740–1752.",
     "https://doi.org/10.1038/s41562-023-01665-y", VP, "Experimental", "Ingreso medio (Egipto)", 2023, EAS,
     "Sin efecto creíble en las actitudes de género ni en la justificación de la violencia; aumentó el conocimiento y el "
     "uso declarado de los servicios disponibles.", "Mixto"),
    ("E64", "T08", "El-Khoury, J. R. y Shafer, A. (2016). Narrative exemplars and the celebrity spokesperson in Lebanese "
     "anti-domestic violence public service announcements. Journal of Health Communication, 21(8), 935–943.",
     "https://doi.org/10.1080/10810730.2016.1177146", VP, "Experimental",
     "Ingreso medio (Líbano, estudiantes universitarios)", 2016, "Experimento con estudiantes",
     "El efecto de los anuncios dependió de la experiencia personal con la violencia; la celebridad mejoró la valoración "
     "del anuncio, no las actitudes.", "Mixto"),
    ("E65", "T08", "Donati, D., Orozco-Olvera, V. y Rao, N. (2022). Using social media to change gender norms: An "
     "experiment within Facebook Messenger in India. Banco Mundial, Policy Research Working Paper 10199.",
     "https://doi.org/10.1596/1813-9450-10199", VR("EconBiz y Crossref (DOI)"), "Experimental", "Ingreso medio (India)",
     2022, EAS, "Un formato de entretenimiento redujo la aceptación social de la violencia contra las mujeres entre 0,16 y "
     "0,21 desviaciones estándar a la semana; el efecto se diluyó a los cuatro meses.", "Favorable"),
    ("E66", "T09", "Melendez-Torres, G. J., Orr, N., Farmer, C., Shaw, N., Chollet, A., Rizzo, A. J. et al. (2024). "
     "School-based interventions TO Prevent Dating and Relationship Violence and Gender-Based Violence: STOP-DRV-GBV "
     "systematic review. Public Health Research, 12(3).", "https://doi.org/10.3310/KTWR6997", VP, "Revisión sistemática",
     "Global", 2024, "No informada", "Evidencia más sólida de reducción de la violencia en el enamoramiento, con efectos "
     "de largo plazo en victimización y perpetración; alguna evidencia de países de ingreso alto para la violencia de "
     "género.", "Favorable"),
    ("E67", "T09", "Queiroz, M. N. R., de Oliveira, M. M., Schaedler, A. C. y de Souza, M. R. (2026). School-based sexual "
     "violence prevention: A systematic review. Child Abuse & Neglect, 180, 108259.",
     "https://doi.org/10.1016/j.chiabu.2026.108259", VP, "Revisión sistemática", "No informado en el resumen", 2026,
     "No informada", "En 18 ensayos aleatorios mejoran los conocimientos y las actitudes; la evidencia de cambios de "
     "conducta es menos frecuente y limitada.", "Mixto"),
    ("E68", "T09", "Fellmeth, G. L. T., Heffernan, C., Nurse, J., Habibula, S. y Sethi, D. (2013). Educational and "
     "skills-based interventions for preventing relationship and dating violence in adolescents and young adults: A "
     "systematic review. Campbell Systematic Reviews, 9(1).", "https://doi.org/10.4073/csr.2013.14", VC,
     "Revisión sistemática", "No informado en el resumen", 2013, "No informada",
     "Los programas mejoraron levemente el conocimiento sobre las relaciones, sin evidencia de reducción de la violencia.",
     "Mixto"),
    ("E69", "T09", "Saldívar Hernández, G. J., Rivera Guevara, B. E., Moreno Reyes, A. y Trejo Arizmendi, J. (2022). "
     "Evaluación de un programa de prevención e intervención en violencia en el noviazgo en adolescentes de Ciudad de "
     "México. Psicología Iberoamericana, 30(1).", "https://www.redalyc.org/journal/1339/133970714005/", VR("Redalyc"),
     "Cuasi experimental", "Ingreso medio (México)", 2022, "No informada (113 estudiantes de una escuela)",
     "Sin diferencias frente al grupo control en violencia en el noviazgo ni en sexismo; mejoras en autoestima y acoso "
     "escolar en el grupo experimental.", "Nulo"),
    ("E70", "T09", "Shinde, S., Weiss, H. A., Khandeparkar, P., Pereira, B., Sharma, A., Gupta, R. et al. (2020). A "
     "multicomponent secondary school health promotion intervention and adolescent health: An extension of the SEHER "
     "cluster randomised controlled trial in Bihar, India. PLOS Medicine, 17(2), e1003021.",
     "https://doi.org/10.1371/journal.pmed.1003021", VP, "Experimental", "Ingreso medio (India)", 2020,
     "Ensayo aleatorio por conglomerados (75 escuelas)", "Con un consejero no docente mejoró el clima escolar y se "
     "redujeron el acoso, la victimización y la perpetración de violencia; la versión conducida por docentes no tuvo "
     "efecto.", "Mixto"),
    ("E71", "T10", "Barros, S., Oliveira, C., Araújo, E., Moreira, D., Almeida, F. y Santos, A. (2022). Community "
     "intervention programs for sex offenders: A systematic review. Frontiers in Psychiatry, 13, 949899.",
     "https://doi.org/10.3389/fpsyt.2022.949899", VP, "Revisión sistemática", "No informado en el resumen", 2022,
     "No informada", "Los programas comunitarios, en su mayoría cognitivo conductuales, parecen reducir la reincidencia.",
     "Favorable"),
    ("E72", "T10", "Schmucker, M. y Lösel, F. (2008). Does sexual offender treatment work? A systematic review of outcome "
     "evaluations. Psicothema, 20(1), 10–19.", "https://pubmed.ncbi.nlm.nih.gov/18206060/", VR("PubMed"),
     "Revisión sistemática", "No informado en el resumen (búsqueda en cinco idiomas)", 2008, "No informada",
     "En 80 comparaciones, la reincidencia sexual fue de 11,1 % en los tratados frente a 17,5 % en los controles; los "
     "enfoques cognitivo conductuales y hormonales muestran efectos positivos.", "Favorable"),
    ("E73", "T10", "Kim, B., Benekos, P. J. y Merlo, A. V. (2016). Sex offender recidivism revisited: Review of recent "
     "meta-analyses on the effects of sex offender treatment. Trauma, Violence, & Abuse, 17(1), 105–117.",
     "https://doi.org/10.1177/1524838014566719", VP, "Revisión narrativa", "No informado en el resumen", 2016, "No aplica",
     "La revisión de metaanálisis califica los tratamientos como probados o prometedores, con diferencias según la edad "
     "y el tipo de intervención.", "Favorable"),
    ("E74", "T10", "Soldino, V. y Carbonell-Vayá, E. J. (2017). Effect of treatment on sex offenders' recidivism: A "
     "meta-analysis. Anales de Psicología, 33(3).", "https://doi.org/10.6018/analesps.33.3.267961",
     VR("LA Referencia y Crossref (DOI)"), "Revisión sistemática", "No informado en el resumen", 2017,
     "Los estudios de mejor calidad no muestran efecto significativo", "En 17 estudios, menor reincidencia sexual (OR "
     "0,69) y general en los tratados; sin efecto significativo al analizar solo los estudios de mejor calidad.", "Mixto"),
    ("E75", "T11", "Scher, L. S., Maynard, R. A. y Stagner, M. (2006). Interventions intended to reduce pregnancy-related "
     "outcomes among adolescents. Campbell Systematic Reviews, 2(1).", "https://doi.org/10.4073/csr.2006.12", VC,
     "Revisión sistemática", "Ingreso alto (Estados Unidos y otros países desarrollados)", 2006, "No informada",
     "En 31 estudios, sin evidencia consistente de que los programas cambien la actividad sexual o el riesgo de embarazo "
     "de los adolescentes.", "Nulo"),
    ("E76", "T11", "Librado-González, N., dos Santos-Monteiro, J. C., Ramírez-Calderón, A. J. y Nicolas-Santiago, M. "
     "(2024). Programas de prevención del embarazo adolescente en América Latina desde un enfoque interseccional: revisión "
     "sistemática. Revista de la Facultad de Ciencias Médicas (Quito), 49(3), 64–76.",
     "https://doi.org/10.29166/rfcmq.v49i3.6730", VR("la página de la revista"), "Revisión sistemática",
     "Ingreso medio (América Latina)", 2024, "No informada", "Ensayos aleatorios con intervenciones educativas y de "
     "autoeficacia que aumentan el uso de condón y reducen el embarazo y el inicio sexual; la revisión califica la "
     "evidencia de efectividad como limitada.", "Mixto"),
    ("E77", "T11; T12", "Castro-Benavides, M., Otzen, T., Palma-Millanao, K. y Manterola, C. (2023). Efectividad de las "
     "intervenciones educativas en prevención del embarazo adolescente en América del Sur: revisión sistemática. Revista "
     "de la Facultad de Ciencias Médicas (Quito), 48(2), 70–81.", "https://doi.org/10.29166/rfcmq.v48i2.5877",
     VR("la página de la revista"), "Revisión sistemática", "Ingreso medio (América del Sur, la mayoría en el Perú)", 2023,
     "No informada (16 estudios, en su mayoría cuasi experimentales)", "Las 16 intervenciones aumentaron el conocimiento "
     "sobre salud sexual y reproductiva y métodos anticonceptivos y mejoraron las actitudes de los adolescentes.",
     "Favorable"),
    ("E78", "T11", "Rokicki, S., Cohen, J., Salomon, J. A. y Fink, G. (2017). Impact of a text-messaging program on "
     "adolescent reproductive health: A cluster-randomized trial in Ghana. American Journal of Public Health, 107(2), "
     "298–305.", "https://doi.org/10.2105/AJPH.2016.303562", VP, "Experimental", "Ingreso medio (Ghana)", 2017, EAC,
     "Los mensajes de texto aumentaron el conocimiento entre 11 y 24 puntos porcentuales; sin cambios generales en salud "
     "reproductiva, pero menos embarazos autorreportados entre las estudiantes sexualmente activas.", "Mixto"),
    ("E79", "T12", "Lameiras-Fernández, M., Martínez-Román, R., Carrera-Fernández, M. V. y Rodríguez-Castro, Y. (2021). "
     "Sex education in the spotlight: What is working? Systematic review. International Journal of Environmental Research "
     "and Public Health, 18(5), 2555.", "https://doi.org/10.3390/ijerph18052555", VP, "Revisión sistemática", "Global",
     2021, "No informada", "En 20 revisiones, la efectividad de los programas escolares es discutida; las plataformas "
     "digitales y el aprendizaje combinado muestran más efectividad, con mayor riesgo de sesgo.", "Mixto"),
    ("E80", "T12", "Denford, S., Abraham, C., Campbell, R. y Busse, H. (2017). A comprehensive review of reviews of "
     "school-based interventions to improve sexual-health. Health Psychology Review, 11(1), 33–52.",
     "https://doi.org/10.1080/17437199.2016.1240625", VP, "Revisión sistemática", "Global", 2017, "No informada",
     "En 37 revisiones (224 estudios), los programas integrales, los de prevención del VIH y las clínicas escolares "
     "mejoran conocimientos, actitudes y conductas; los programas de solo abstinencia son inefectivos.", "Favorable"),
    ("E81", "T12", "Duflo, E., Dupas, P. y Kremer, M. (2015). Education, HIV, and early fertility: Experimental evidence "
     "from Kenya. American Economic Review, 105(9), 2757–2797.", "https://doi.org/10.1257/aer.20121607", VP,
     "Experimental", "Ingreso bajo (Kenia)", 2015, "Ensayo aleatorio (siete años)", "El currículo oficial de VIH, centrado "
     "en la abstinencia, no redujo el embarazo ni las ITS; los subsidios educativos sí redujeron la deserción, el embarazo "
     "y el matrimonio adolescente.", "Nulo"),
    ("E82", "T13", "Adida, C. L., Kosec, K., Kyle, J., Mo, C. H., Arriola, L. R., Matanock, A. M., Adeyanju, D. y Fisher, "
     "R. (2025). Bridging gender gaps in political participation: Experimental evidence on group-based trainings from "
     "Nigeria. IFPRI Discussion Paper.", "https://hdl.handle.net/10568/180330", VR("EconBiz y RePEc (IFPRI)"),
     "Experimental", "Ingreso medio (Nigeria)", 2025, "Ensayo aleatorio (300 comunidades)",
     "Las capacitaciones grupales a comités de mujeres aumentaron el nivel y la calidad de su participación política y la "
     "respuesta de los líderes locales a sus prioridades.", "Favorable"),
    ("E83", "T13", "Mvukiyehe, E. (2017). Can media interventions reduce gender gaps in political participation after "
     "civil war? Evidence from a field experiment in rural Liberia. Banco Mundial, Policy Research Working Paper 7942.",
     "https://doi.org/10.1596/1813-9450-7942", VR("EconBiz y Crossref (DOI)"), "Experimental", "Ingreso bajo (Liberia)",
     2017, EAS, "Grupos de escucha de radio sobre el proceso electoral aumentaron la participación política de las mujeres "
     "y redujeron las brechas de género.", "Favorable"),
    ("E84", "T14", "Castro, M. y da Cruz, B. (2024). Effects of a large-scale program for the construction of daycare and "
     "preschool centers on cognitive skills and female employment. Education Economics, 32(6), 786–811.",
     "https://doi.org/10.1080/09645292.2023.2254516", VR("EconBiz y Crossref (DOI)"), "Cuasi experimental",
     "Ingreso medio (Brasil)", 2024, "No informada", "El programa Proinfância aumentó la asistencia a la educación inicial "
     "y, en un punto porcentual, la participación laboral femenina.", "Favorable"),
    ("E85", "T14", "López Boo, F. y Hojman, A. (2019). Cost-effective public daycare in a low-income economy benefits "
     "children and mothers. Banco Interamericano de Desarrollo.", "https://doi.org/10.18235/0001849",
     VR("EconBiz y Crossref (DOI)"), "Experimental", "Ingreso medio (Nicaragua)", 2019,
     "Asignación aleatoria original con variables instrumentales", "El acceso a centros de cuidado infantil de medio "
     "tiempo aumentó en 14 puntos porcentuales la participación laboral de las madres y mejoró el desarrollo personal y "
     "social de los niños.", "Favorable"),
    ("E86", "T14", "Clark, S., Kabiru, C. W., Laszlo, S. y Muthuri, S. (2019). The impact of childcare on poor urban "
     "women's economic empowerment in Africa. Demography, 56(4), 1247–1272.", "https://doi.org/10.1007/s13524-019-00793-3",
     VP, "Experimental", "Ingreso medio (Kenia)", 2019, EAS, "Los vales de cuidado infantil subsidiado aumentaron en 8,5 "
     "puntos porcentuales el empleo de las madres en un asentamiento informal de Nairobi.", "Favorable"),
    ("E87", "T14", "Viitanen, T. K. (2007). Public versus private provision of daycare: An experimental evaluation. SSRN.",
     "https://doi.org/10.2139/ssrn.1012793", VR("EconBiz y Crossref (DOI)"), "Experimental",
     "Ingreso alto (país nórdico)", 2007, "No informada", "El vale para cuidado privado aumentó su uso, con efectos nulos "
     "o insignificantes en la participación laboral.", "Nulo"),
    ("E88", "T15", "Rains, E. y Huntington, H. (2026). Gatekeeping women's property rights: Evidence from a land titling "
     "program in urban Mozambique. The Journal of Development Studies.", "https://doi.org/10.1080/00220388.2026.2716048",
     VR("EconBiz y Crossref (DOI)"), "Descriptivo", "Ingreso bajo (Mozambique)", 2026, "No aplica",
     "Las mujeres rara vez figuraron en los documentos; las que figuran reportan más poder de decisión y derechos "
     "percibidos sobre la tierra, y las más vulnerables fueron las menos incluidas.", "Mixto"),
    ("E89", "T15", "Donald, A., Goldstein, M., Hartman, A., La Ferrara, E., O'Sullivan, M. y Stickler, M. (2020). What's "
     "mine is yours: Pilot evidence from a randomized impact evaluation on property rights and women's empowerment in "
     "Côte d'Ivoire. Banco Mundial.", "https://doi.org/10.1596/34081", VR("EconBiz y Crossref (DOI)"), "Experimental",
     "Ingreso medio (Côte d'Ivoire)", 2020, "Piloto de un ensayo aleatorio", "Destacar los beneficios de la propiedad de "
     "las mujeres llevó a los esposos a registrar tierra a nombre de sus esposas.", "Favorable"),
    ("E90", "T15", "Poulin, M. y Pierotti, R. S. (2025). Joint land titles and women's status in Uganda. Banco Mundial, "
     "Policy Research Working Paper 11241.", "https://doi.org/10.1596/1813-9450-11241", VR("EconBiz y Crossref (DOI)"),
     "Experimental", "Ingreso bajo (Uganda)", 2025, "Ensayo aleatorio con entrevistas",
     "La mayoría de los hombres aceptó titular la tierra junto con su esposa; la titulación conjunta aumentó el poder de "
     "negociación de las mujeres.", "Favorable"),
    ("E91", "T15", "O'Sullivan, M. (2017). Gender and property rights in Sub-Saharan Africa: A review of constraints and "
     "effective interventions. Banco Mundial, Policy Research Working Paper 8250.",
     "https://doi.org/10.1596/1813-9450-8250", VR("EconBiz y Crossref (DOI)"), "Revisión narrativa",
     "Ingreso bajo y medio (África subsahariana)", 2017, "No informada", "Evidencia prometedora sobre la titulación "
     "conjunta de tierras; falta evidencia en ámbitos urbanos.", "Favorable"),
    ("E92", "T15", "Botea, I., Goldstein, M., Houngbedji, K., Kondylis, F., O'Sullivan, M. y Selod, H. (2026). Right to "
     "stay? Long-run experimental evidence on land formalization and widows' tenure security in Benin. Journal of "
     "Development Economics, 179, 103675.", "https://doi.org/10.1016/j.jdeveco.2025.103675",
     VR("EconBiz y Crossref (DOI)"), "Experimental", "Ingreso bajo (Benín)", 2026, EAS,
     "Cuatro años después, las viudas en aldeas con formalización tenían más probabilidad de permanecer en su vivienda, "
     "sobre todo las que no tenían heredero varón.", "Favorable"),
    ("E93", "T15", "Boxho, C., Brudevold-Newman, A., Montalvao, J., Proença, M. y O'Sullivan, M. (2026). Land "
     "registration, agricultural technology adoption, and investment in Mozambique: Evidence from a field experiment with "
     "women farmers. Banco Mundial.", "https://doi.org/10.1596/45251", VR("EconBiz y Crossref (DOI)"), "Experimental",
     "Ingreso bajo (Mozambique)", 2026, EAS, "El registro de tierras elevó temporalmente la seguridad percibida de la "
     "tenencia, sin aumentar la inversión ni la productividad agrícola de las mujeres.", "Mixto"),
    ("E94", "T16", "Reyes-Retana, G., Pons, G., Siegmann, K., Afif, Z., Gomez-Garcia, M., Soto-Mota, P. y Castaneda "
     "Farill, C. E. (2023). Using behavioral science to increase women's participation in natural resource management in "
     "Mexico. Banco Mundial, Policy Research Working Paper 10419.", "https://doi.org/10.1596/1813-9450-10419",
     VR("EconBiz y Crossref (DOI)"), "Experimental", "Ingreso medio (México)", 2023, EAS,
     "La difusión con enfoque conductual aumentó las solicitudes de mujeres a proyectos forestales, sin efecto en la "
     "probabilidad de obtener el financiamiento.", "Mixto"),
    ("E95", "T17", "Croke, K., Goldstein, M. y Holla, A. (2023). The role of skills and gender norms in sector switches: "
     "Experimental evidence from a job training program in Nigeria. Journal of African Economies, 32(3), 237–268.",
     "https://doi.org/10.1093/jae/ejac007", VR("EconBiz y Crossref (DOI)"), "Experimental", "Ingreso medio (Nigeria)",
     2023, EAS, "La capacitación aumentó el paso a empleos en tecnologías de la información y servicios empresariales y "
     "ayudó a las mujeres a superar normas de género internalizadas.", "Favorable"),
    ("E96", "T17", "Frohnweiler, S. et al. (2024). Effects of skills training on employment and livelihood outcomes: A "
     "randomized controlled trial with young women in Ghana. RWI, Ruhr Economic Papers.",
     "https://doi.org/10.4419/96973272", VR("EconBiz y la página de RWI"), "Experimental", "Ingreso medio (Ghana)", 2024,
     EAS, "Sin efecto en empleo, horas ni ingresos; más trabajo en el oficio aprendido, alguna mejora en la calidad del "
     "empleo y mejores resultados en salud mental, matrimonio tardío y acceso a financiamiento.", "Mixto"),
    ("E97", "T17; T11", "Bandiera, O., Buehren, N., Burgess, R., Goldstein, M., Gulesci, S., Rasul, I. y Sulaiman, M. "
     "(2020). Women's empowerment in action: Evidence from a randomized control trial in Africa. American Economic "
     "Journal: Applied Economics, 12(1), 210–259.", "https://doi.org/10.1257/app.20170416", VR("EconBiz y Crossref (DOI)"),
     "Experimental", "Ingreso bajo (Uganda)", 2020, EAS, "Cuatro años después, las adolescentes con formación vocacional "
     "e información sobre salud sexual tenían más autoempleo y menos embarazo, matrimonio temprano y relaciones sexuales "
     "forzadas.", "Favorable"),
    ("E98", "T17", "Tripney, J., Hombrados, J., Newman, M., Hovish, K., Brown, C., Steinka-Fry, K. y Wilkey, E. (2013). "
     "Technical and vocational education and training (TVET) interventions to improve the employability and employment "
     "of young people in low- and middle-income countries: A systematic review. Campbell Systematic Reviews, 9(1).",
     "https://doi.org/10.4073/csr.2013.9", VC, "Revisión sistemática", "Ingreso bajo y medio", 2013, "No informada",
     "En 26 estudios, la formación técnica y vocacional tiene un efecto pequeño y positivo en el empleo de los jóvenes.",
     "Favorable"),
    ("E99", "T17", "Ñopo, H., Robles, M. y Saavedra, J. (2002). Una medición del impacto del Programa de Capacitación "
     "Laboral Juvenil PROJoven. Grupo de Análisis para el Desarrollo (GRADE).", "https://hdl.handle.net/20.500.12820/229",
     VR("LA Referencia (registro del repositorio de GRADE, que no respondió el 31-10-2026)"), "Cuasi experimental",
     "Perú", 2002, "No informada", "Con emparejamiento y grupo de control, PROJoven elevó en 18 % el ingreso por hora de "
     "los beneficiarios que consiguieron empleo.", "Favorable"),
    ("E100", "T17", "Galdo, J., Jaramillo, M. y Montalva, V. (2009). Pobreza e impactos heterogéneos de las políticas "
     "activas de empleo juvenil: el caso de PROJOVEN en el Perú. Grupo de Análisis para el Desarrollo (GRADE).",
     "https://hdl.handle.net/20.500.12820/248",
     VR("LA Referencia (registro del repositorio de GRADE, que no respondió el 31-10-2026)"), "Cuasi experimental",
     "Perú", 2009, "No informada", "Impactos heterogéneos, con diferencias fuertes entre varones y mujeres; los más pobres "
     "se benefician igual que los menos pobres y la calidad de la entidad capacitadora explica la heterogeneidad.",
     "Favorable"),
    ("E101", "T18", "Vu, N. T. H., van Velzen, R., Lensink, R. y Bulte, E. (2015). The impact of gender and business "
     "training for female microfinance clients in Vietnam. International Initiative for Impact Evaluation (3ie).",
     "https://doi.org/10.23846/ow31132", f"{FECHA}: título, autores, año y fuente comprobados en Crossref (DOI); "
     "resultado tomado del resumen. El resumen indica un grupo tratado y uno de comparación; la página de 3ie no abrió, "
     "por lo que no se precisó si la asignación fue aleatoria.", "Cuasi experimental", "Ingreso medio (Vietnam)", 2015,
     "No informada", "La capacitación en género y negocios mejoró los conocimientos y los márgenes de ganancia de los "
     "negocios de las mujeres; invitar a los esposos no tuvo efecto adicional significativo.", "Favorable"),
    ("E102", "T18", "Knox, S., Gok, A., Aydogdu, E., Carter, S., Levie, J. y Shaw, E. (2026). Empowering women's "
     "entrepreneurship: An evidence synthesis of policy and practice in developing countries. Entrepreneurship & Regional "
     "Development, 38(3-4), 410–434.", "https://doi.org/10.1080/08985626.2025.2503145", VR("EconBiz y Crossref (DOI)"),
     "Revisión narrativa", "Ingreso bajo y medio", 2026, "No aplica", "Síntesis realista: el impacto de las microfinanzas "
     "y la capacitación empresarial es mixto; el capital inicial rara vez basta sin mecanismos que conviertan recursos "
     "económicos, humanos y sociales.", "Mixto"),
    ("E103", "T18", "McKenzie, D. y Puerto, S. (2021). Growing markets through business training for female "
     "entrepreneurs: A market-level randomized experiment in Kenya. American Economic Journal: Applied Economics, 13(2), "
     "297–332.", "https://doi.org/10.1257/app.20180340", VR("EconBiz y Crossref (DOI)"), "Experimental",
     "Ingreso medio (Kenia)", 2021, EAS, "Tres años después, los negocios capacitados venden más y tienen mayores "
     "ganancias, sin perjuicio significativo para sus competidoras.", "Favorable"),
    ("E104", "T18", "Bakhtiar, M. M., Bastian, G. y Goldstein, M. (2021). Business training and mentoring: Experimental "
     "evidence from women-owned microenterprises in Ethiopia. Banco Mundial, Policy Research Working Paper 9552.",
     "https://doi.org/10.1596/1813-9450-9552", VR("EconBiz y Crossref (DOI)"), "Experimental", "Ingreso bajo (Etiopía)",
     2021, EAS, "La capacitación aumentó las ganancias y las ventas en 0,21 desviaciones estándar; la mentoría tuvo "
     "efectos débiles.", "Favorable"),
    ("E105", "T18", "Namayengo, F. M. M., Antonides, G. y van Ophem, J. (2025). Impact of microcredit on performance of "
     "women-owned non-farm microenterprises. Review of Development Economics, 29(3), 1940–1951.",
     "https://doi.org/10.1111/rode.13205", VR("EconBiz y Crossref (DOI)"), "Cuasi experimental", "Ingreso bajo (Uganda)",
     2025, "No informada", "El microcrédito aumentó el valor de los negocios y el número de trabajadores, sin efecto en "
     "las ganancias.", "Mixto"),
    ("E106", "T18", "de Mel, S., McKenzie, D. y Woodruff, C. (2014). Business training and female enterprise start-up, "
     "growth, and dynamics: Experimental evidence from Sri Lanka. Journal of Development Economics, 106, 199–210.",
     "https://doi.org/10.1016/j.jdeveco.2013.09.005", f"{FECHA}: título, autores, año y fuente comprobados en EconBiz y "
     "Crossref (DOI); resultado tomado del resumen publicado en Semantic Scholar.", "Experimental",
     "Ingreso medio (Sri Lanka)", 2014, EAS, "La capacitación sola no cambió las ganancias de los negocios existentes; "
     "con una donación las mejoró solo en el corto plazo; aceleró la creación de nuevos negocios.", "Mixto"),
    ("E107", "T18", "Karlan, D. y Valdivia, M. (2010). Formando microempresarias: impacto de la capacitación empresarial "
     "en las instituciones de microfinanzas y sus socias. Grupo de Análisis para el Desarrollo (GRADE).",
     "https://hdl.handle.net/20.500.12820/251",
     VR("LA Referencia (registro del repositorio de GRADE, que no respondió el 31-10-2026)"), "Experimental", "Perú", 2010,
     "Ensayo aleatorio (bancos comunales)", "La capacitación mejoró el conocimiento de prácticas de gestión; efectos nulos "
     "o poco robustos en ingresos, ganancias y empleo; aumentó la permanencia de las socias en la institución.", "Mixto"),
    ("E108", "T19", "Tang, X., Patterson, P., MacKenzie-Shalders, K., van Herwerden, L. A., Bishop, J., Rathbone, E. et al. "
     "(2021). Workplace programmes for supporting breast-feeding: A systematic review and meta-analysis. Public Health "
     "Nutrition, 24(6), 1501–1513.", "https://doi.org/10.1017/S1368980020004012", VP, "Revisión sistemática", "Global",
     2021, "Baja (sin ensayos aleatorios, según la revisión)", "Los programas en el centro de trabajo pueden promover la "
     "lactancia: en los estudios de un solo grupo, 76 % amamantaba a los seis meses.", "Favorable"),
    ("E109", "T19", "Souza, C. B., Venancio, S. I. y Silva, R. P. G. V. C. (2022). Breastfeeding support rooms, benefits "
     "and challenges for implantation: Cross-sectional study. Research, Society and Development, 11(17), "
     "e238111737973.", "https://doi.org/10.33448/rsd-v11i17.37973", VR("la página de la revista"), "Descriptivo",
     "Ingreso medio (Brasil)", 2022, "No aplica", "En 16 salas de apoyo a la lactancia certificadas, la mediana de "
     "usuarias fue cerca de cinco veces menor que la de empleadas en licencia de maternidad; el uso infrecuente es el "
     "principal desafío.", "Sin estimación"),
    ("E110", "T20", "Gaentzsch, A. (2020). Do conditional cash transfers (CCTs) raise educational attainment? An impact "
     "evaluation of Juntos in Peru. Development Policy Review, 38(6), 747–765.", "https://doi.org/10.1111/dpr.12468",
     VR("EconBiz y Crossref (DOI)"), "Cuasi experimental", "Perú", 2020, "No informada", "Juntos aumentó la matrícula y la "
     "culminación de la primaria entre los 12 y 18 años; el efecto en el paso a la secundaria fue débil y no hubo efecto "
     "en las pruebas de aprendizaje.", "Mixto"),
    ("E111", "T20", "Ham, A. (2010). The effect of conditional cash transfers on educational opportunities: Experimental "
     "evidence from Latin America. Centro de Estudios Distributivos, Laborales y Sociales (CEDLAS), documento de trabajo.",
     "http://hdl.handle.net/10419/127617", VR("EconBiz y EconStor"), "Experimental", "Ingreso medio (América Latina)",
     2010, "No informada", "Las transferencias condicionadas tuvieron un efecto redistributivo y un impacto inicial "
     "positivo en la igualdad de oportunidades educativas.", "Favorable"),
    ("E112", "T20", "Psaki, S., Haberland, N., Mensch, B., Woyczynski, L. y Chuang, E. (2022). Policies and interventions "
     "to remove gender-related barriers to girls' school participation and learning in low- and middle-income countries: "
     "A systematic review of the evidence. Campbell Systematic Reviews, 18(1), e1207.",
     "https://doi.org/10.1002/cl2.1207", VC, "Revisión sistemática", "Ingreso bajo y medio", 2022, "No informada",
     "Son efectivas las intervenciones que reducen el costo de la educación, dan alimentación o apoyo académico; para el "
     "embarazo adolescente, la violencia de género en la escuela y el entorno escolar insensible al género la evidencia es "
     "insuficiente.", "Favorable"),
    ("E113", "T21", "Niesten, H. M. L., Ferraz Di Ricco, L. y Sakhonchik, A. (2026). Making budgets work for women: "
     "Gender-responsive budgeting in 50 economies. Banco Mundial.", "https://doi.org/10.1596/44921",
     VR("EconBiz y Crossref (DOI)"), "Descriptivo", "Global (50 economías)", 2026, "No aplica",
     "Muchos sistemas de etiquetado presupuestal de género existen en el papel, pero no están integrados a los marcos de "
     "desempeño y rendición de cuentas.", "Sin estimación"),
    ("E114", "T23", "Srinivasan, V., San Sebastián, M., Rana, S., Bhatt, P., Armstrong, G., Deshpande, S. y Mathias, K. "
     "(2025). Effectiveness of a resilience, gender equity and mental health group intervention for young people living "
     "in informal urban communities in North India: A cluster randomized controlled trial. Global Health Action, 18(1), "
     "2455236.", "https://doi.org/10.1080/16549716.2025.2455236", VP, "Experimental", "Ingreso medio (India)", 2025, EAC,
     "Mejoró la resiliencia de hombres y mujeres jóvenes, sin efecto significativo en las actitudes de igualdad de género "
     "ni en la salud mental.", "Nulo"),
    ("E115", "T24", "Galvin, L., Verissimo, C. K., Ambikapathi, R., Gunaratna, N. S., Rudnicka, P., Sunseri, A. et al. "
     "(2023). Effects of engaging fathers and bundling nutrition and parenting interventions on household gender equality "
     "and women's empowerment in rural Tanzania: Results from EFFECTS, a five-arm cluster-randomized controlled trial. "
     "Social Science & Medicine, 324, 115869.", "https://doi.org/10.1016/j.socscimed.2023.115869", VP, "Experimental",
     "Ingreso medio (Tanzania)", 2023, EAC, "Involucrar a ambos miembros de la pareja aumentó las actitudes equitativas, el "
     "tiempo de los padres en las tareas domésticas y el poder de decisión de las madres.", "Favorable"),
]
EVIDENCIA += ACADEMICA

# Registros de la búsqueda académica leídos por su resumen (selección por título) y su decisión. Los que no figuran
# aquí se excluyeron en el cribado de títulos. Formato: registro (estrategia|clave del registro), tipos, decisión,
# evidencia, motivo.
INC = "Incluida"
EXC = "Excluida tras leer el resumen"
YA = "Ya incluida (búsqueda web)"
SIN_RESUMEN = "Resumen no disponible en las fuentes consultadas (base, Crossref, OpenAlex y Semantic Scholar); no se "
SIN_RESUMEN += "verificaron sus resultados."
SELECCION = [
    ("ACA-002|DOI:10.4073/csr.2016.2", "T01", YA, "E02", ""),
    ("ACA-005|PMID:32608505", "T02", YA, "E03", ""),
    ("ACA-039|DOI:10.4073/csr.2014.7", "T09", YA, "E15", ""),
    ("ACA-043|DOI:10.4073/csr.2017.8", "T10", YA, "E16", ""),
    ("ACA-062|DOI:10.23846/sr71094", "T14", YA, "E27", ""),
    ("ACA-065|DOI:10.4073/csr.2014.1", "T15", YA, "E24", ""),
    ("ACA-076|LAREF:PE_c7d72989ba6c7d2bb17b3ce66f4227fd", "T17", YA, "E31", ""),
    ("ACA-081|PMID:28394659", "T19", YA, "E35", ""),
    ("ACA-040|DOI:10.23846/pw2ie89", "T23", YA, "E14", "Informe de 3ie del mismo estudio."),
    ("ACA-085|DOI:10.4073/csr.2013.8", "T20", YA, "E36", "Versión Campbell de la misma revisión."),
    ("ACA-100|PMID:29617375", "T24", YA, "E12", ""),
    ("ACA-038|ERIC:EJ1132720", "T09", YA, "E15", "Versión publicada en revista de la misma revisión."),
    ("ACA-002|DOI:10.1002/cl2.1271", "T01", INC, "E40", ""),
    ("ACA-005|PMID:24482679", "T01", INC, "E41", ""),
    ("ACA-005|PMID:24439354", "T01", INC, "E42", ""),
    ("ACA-004|LAREF:PE_168faf5cc1c7e2c2ad0b6abd1d01e50d", "T01", EXC, "", "Resumen de congreso publicado en un suplemento (una página); no es un artículo completo."),
    ("ACA-008|LAREF:EC_d656b22706eb434d9bfda94f85ecce95", "T02", INC, "E43", ""),
    ("ACA-008|LAREF:MX_45384d5c789fcb6e92ded9133ff2cf80", "T02", EXC, "", "Intervención distinta: tratamiento de hombres que ejercen violencia de pareja; la revisión "
     "halló un solo estudio."),
    ("ACA-009|PMID:32954993", "T03", INC, "E44", ""),
    ("ACA-009|PMID:39394692", "T03", INC, "E45", ""),
    ("ACA-009|PMID:31658878", "T03", INC, "E46", ""),
    ("ACA-010|ECONBIZ:10014552740", "T03", EXC, "", SIN_RESUMEN),
    ("ACA-013|LAREF:PE_8042a78da4ef0e3c4b9298068f9973f9", "T03", EXC, "", "Diseño correlacional sin grupo de comparación."),
    ("ACA-013|LAREF:ES_ae909764a4ffa2b74a297ba8701b4830", "T03", INC, "E47", ""),
    ("ACA-010|ECONBIZ:10016098102", "T05", INC, "E48", ""),
    ("ACA-010|ECONBIZ:10015654409", "T05", EXC, "", "Versión previa del mismo ensayo (E48)."),
    ("ACA-014|PMID:36058604", "T04", INC, "E49", ""), ("ACA-014|PMID:28616887", "T04", INC, "E50", ""),
    ("ACA-016|ECONBIZ:10015643944", "T04", EXC, "", "Estudio primario en un país de ingreso alto (Noruega), según el título."),
    ("ACA-015|ERIC:EJ1425662", "T09", INC, "E51", ""),
    ("ACA-017|DOI:10.4073/csr.2015.10", "T09", INC, "E52", ""),
    ("ACA-017|DOI:10.4073/csr.2019.1", "T09", INC, "E53", ""),
    ("ACA-024|PMID:32334615", "T06", INC, "E54", ""),
    ("ACA-024|PMID:34770188", "T06", INC, "E55", ""),
    ("ACA-024|PMID:37329160", "T06", INC, "E56", ""),
    ("ACA-027|LAREF:PE_10711fe17cc28869a65e3b136ede024d", "T06", EXC, "", "Sistematización de una buena práctica del Programa Nacional Aurora, sin evaluación de "
     "efectos."),
    ("ACA-028|PMID:34282677", "T07", INC, "E57", ""),
    ("ACA-028|PMID:35140139", "T07", INC, "E58", ""),
    ("ACA-028|PMID:33355268", "T07", INC, "E59", ""),
    ("ACA-028|PMID:30730251", "T07", INC, "E60", ""),
    ("ACA-028|PMID:31227386", "T07", EXC, "", "Análisis secundario del mismo ensayo (E60)."),
    ("ACA-028|PMID:36888613", "T07", INC, "E61", ""), ("ACA-028|PMID:34301673", "T07", INC, "E62", ""),
    ("ACA-104|LAREF:MX_78691336ce1613383c013bf6b75ca369", "T07", EXC, "", "Sistematización cualitativa de un taller, sin diseño evaluativo."),
    ("ACA-032|PMID:37550411", "T08", INC, "E63", ""),
    ("ACA-032|PMID:27441946", "T08", INC, "E64", ""),
    ("ACA-033|ECONBIZ:10013545007", "T08", INC, "E65", ""),
    ("ACA-037|PMID:30264210", "T09", EXC, "", "Publicación derivada de la misma revisión (E53)."),
    ("ACA-037|PMID:34342255", "T09", EXC, "", "Población distinta: estudiantes universitarios de Estados Unidos."),
    ("ACA-037|PMID:38421001", "T09", INC, "E66", ""),
    ("ACA-037|PMID:42574787", "T09", INC, "E67", ""),
    ("ACA-039|DOI:10.4073/csr.2013.14", "T09", INC, "E68", ""),
    ("ACA-041|LAREF:MX_e6f9648fad42acf4d71b333cf844c3ec", "T09", INC, "E69", ""),
    ("ACA-041|LAREF:CL_8744f624320d310bdafa92a94aec9c89", "T09", EXC, "", "Diseño preexperimental de un solo grupo."),
    ("ACA-041|LAREF:BR_38ef66443bd02fcea5dd6f616508ee5c", "T09", EXC, "", "Medición antes y después sin grupo de comparación, en un país de ingreso alto (Portugal)."),
    ("ACA-096|PMID:32045409", "T09", INC, "E70", ""),
    ("ACA-042|PMID:36506440", "T10", INC, "E71", ""),
    ("ACA-042|PMID:18206060", "T10", INC, "E72", ""),
    ("ACA-042|PMID:25575803", "T10", INC, "E73", ""),
    ("ACA-045|LAREF:ES_849ad45af200ed597bc96f47250f50e2", "T10", INC, "E74", ""),
    ("ACA-047|DOI:10.4073/csr.2006.12", "T11", INC, "E75", ""),
    ("ACA-049|LAREF:EC_628d87fa97a57d4dff1eda29aafba465", "T11", INC, "E76", ""),
    ("ACA-049|LAREF:EC_6dc91c23704adb1067079215732cd417", "T11", INC, "E77", ""),
    ("ACA-050|PMID:27997236", "T11", INC, "E78", ""),
    ("ACA-049|LAREF:PE_bfc57ef0ff5b0d424c58727a309e45c7", "T12", EXC, "", "Análisis antes y después sin grupo de comparación."),
    ("ACA-050|PMID:33806507", "T12", INC, "E79", ""),
    ("ACA-050|PMID:27677440", "T12", INC, "E80", ""),
    ("ACA-050|PMID:26523067", "T12", INC, "E81", ""),
    ("ACA-050|PMID:32048888", "T12", EXC, "", "El resumen no informa la dirección de los efectos."),
    ("ACA-054|LAREF:PE_5fe809838c1f5e219eb3f85c40caedd6", "T12", EXC, "", "No se pudo verificar la fuente: la página de la revista no respondió el 31-10-2026."),
    ("ACA-055|ECONBIZ:10015656764", "T13", INC, "E82", ""), ("ACA-055|ECONBIZ:10012246504", "T13", INC, "E83", ""),
    ("ACA-055|ECONBIZ:10014316907", "T13", EXC, "", "Intervención distinta: campaña puerta a puerta de un candidato."),
    ("ACA-058|LAREF:CO_820450f1a88e9755a77aabe5413b4f35", "T13", EXC, "", "Evaluación cualitativa con 40 mujeres."),
    ("ACA-059|ECONBIZ:10014517577", "T14", EXC, "", SIN_RESUMEN), ("ACA-059|ECONBIZ:10014247851", "T14", EXC, "", SIN_RESUMEN),
    ("ACA-059|ECONBIZ:10013316846", "T14", INC, "E87", ""),
    ("ACA-059|ECONBIZ:10015532027", "T14", INC, "E84", ""),
    ("ACA-059|ECONBIZ:10012098043", "T14", INC, "E85", ""),
    ("ACA-060|PMID:31286428", "T14", INC, "E86", ""),
    ("ACA-062|DOI:10.23846/ow2223", "T14", EXC, "", "El resumen que devuelven Crossref y OpenAlex corresponde a otro documento; no se "
     "verificaron sus resultados."),
    ("ACA-064|ECONBIZ:10016096591", "T15", INC, "E88", ""),
    ("ACA-064|ECONBIZ:10012566631", "T15", INC, "E89", ""),
    ("ACA-064|ECONBIZ:10015556835", "T15", INC, "E90", ""),
    ("ACA-064|ECONBIZ:10012569701", "T15", INC, "E91", ""),
    ("ACA-064|ECONBIZ:10015560833", "T15", INC, "E92", ""),
    ("ACA-064|ECONBIZ:10016069783", "T15", INC, "E93", ""),
    ("ACA-066|DOI:10.23846/rps0009", "T15", EXC, "", "El resumen no informa resultados (réplica del estudio de Galiani y Schargrodsky)."),
    ("ACA-066|DOI:10.23846/b/ob/202203", "T15", EXC, "", "El resumen que devuelve Crossref corresponde a otro documento."),
    ("ACA-068|ECONBIZ:10013184271", "T16", EXC, "", SIN_RESUMEN), ("ACA-068|ECONBIZ:10014312636", "T16", INC, "E94", ""),
    ("ACA-068|ECONBIZ:10013209034", "T16", EXC, "", "Informe del mismo ensayo (E94)."),
    ("ACA-068|ECONBIZ:10003932223", "T16", EXC, "", SIN_RESUMEN),
    ("ACA-072|ECONBIZ:10014314261", "T17", INC, "E95", ""),
    ("ACA-072|ECONBIZ:10015673065", "T17", EXC, "", "Intervención distinta: habilidades socioemocionales, no formación vocacional."),
    ("ACA-072|ECONBIZ:10015051679", "T17", INC, "E96", ""),
    ("ACA-072|ECONBIZ:10012198381", "T17", INC, "E97", ""),
    ("ACA-074|DOI:10.4073/csr.2013.9", "T17", INC, "E98", ""),
    ("ACA-075|DOI:10.23846/b/ob/202107", "T17", EXC, "", "El resumen no informa resultados."),
    ("ACA-076|LAREF:PE_6b4e4fae65c9c0b4a548df63a27cdc4e", "T17", INC, "E99", ""),
    ("ACA-076|LAREF:MX_fa8df553a5a77b85e7be81652b99a4cd", "T17", EXC, "", "Panorama de evaluaciones sin estimación de efectos."),
    ("ACA-076|LAREF:PE_279b6a41057e7cb3c8783a348c89f64c", "T17", INC, "E100", ""),
    ("ACA-022|DOI:10.23846/ow31132", "T18", INC, "E101", ""),
    ("ACA-077|ECONBIZ:10015654216", "T18", INC, "E102", ""),
    ("ACA-077|ECONBIZ:10012570488", "T18", INC, "E103", ""),
    ("ACA-077|ECONBIZ:10012521772", "T18", INC, "E104", ""),
    ("ACA-077|ECONBIZ:10015481453", "T18", INC, "E105", ""),
    ("ACA-077|ECONBIZ:10012975105", "T18", INC, "E106", ""),
    ("ACA-077|ECONBIZ:10015073154", "T18", EXC, "", "Revisión conceptual sin estimación de efectos."),
    ("ACA-080|LAREF:EC_a5badbda3b83d4c497e4059cddb0a9f3", "T18", EXC, "", "Documento de repositorio institucional sin revista arbitrada identificada."),
    ("ACA-080|LAREF:PE_9d07ef25574fe0b43d09e942277c10a3", "T18", INC, "E107", ""),
    ("ACA-080|LAREF:MX_fd2e21ff958d36cb4a31194e1594faac", "T18", EXC, "", "Evaluación cualitativa."),
    ("ACA-081|PMID:33050979", "T19", INC, "E108", ""),
    ("ACA-081|PMID:32530352", "T19", EXC, "", "Revisión de percepciones sin estimación de efectos."),
    ("ACA-082|LAREF:BR_5743f64f858851725935e8d4312fb42d", "T19", INC, "E109", ""),
    ("ACA-084|ECONBIZ:10012951863", "T20", EXC, "", "El resumen no informa la dirección de los efectos; pendiente de lectura del texto "
     "completo (metaanálisis de 94 estudios)."),
    ("ACA-084|ECONBIZ:10012428645", "T20", INC, "E110", ""), ("ACA-084|ECONBIZ:10008702795", "T20", INC, "E111", ""),
    ("ACA-086|DOI:10.23846/ow4.ie69", "T20", EXC, "", "El resumen que devuelve Crossref corresponde a otro documento."),
    ("ACA-087|LAREF:PE_4819da2a9a98162215fe1fc77faea4d1", "T20", EXC, "", "El resumen no informa resultados; síntesis de las evaluaciones de Juntos pendiente de "
     "lectura."),
    ("ACA-089|DOI:10.1002/cl2.1207", "T20", INC, "E112", ""),
    ("ACA-088|ECONBIZ:10014344806", "T21", EXC, "", SIN_RESUMEN), ("ACA-088|ECONBIZ:10015655934", "T21", INC, "E113", ""),
    ("ACA-095|ERIC:ED567431", "T23", EXC, "", "Resumen del diseño de la evaluación, sin resultados."),
    ("ACA-096|PMID:34923962", "T23", EXC, "", "Ensayo sin grupo sin intervención: compara variantes del programa."),
    ("ACA-096|PMID:39898764", "T23", INC, "E114", ""),
    ("ACA-099|LAREF:MX_c08ace5207414b90d60b3fe1450f49a7", "T23", EXC, "", "El resumen no describe método ni resultados."),
    ("ACA-099|LAREF:ES_2d163806a308ce0723622ea687de50c9", "T23", EXC, "", "Medición antes y después sin grupo de comparación, en un país de ingreso alto."),
    ("ACA-100|PMID:37023660", "T24", INC, "E115", ""),
]
# Ampliación del 2026-10-31 (ronda 2): PubMed y EconBiz hasta el registro 30, Banco Mundial, OpenAlex, portales (J-PAL,
# What Works, SVRI) y evaluaciones peruanas. Extracción orientada a la decisión: se extraen y verifican las fuentes que
# pueden cambiar la solidez o la dirección del tipo (todas en los tipos con evidencia limitada o de dirección favorable;
# revisiones sistemáticas en los de evidencia moderada) y toda evaluación realizada en el Perú. Las demás fuentes
# pertinentes se registran como identificadas, sin extracción, porque una fuente más no cambia la clasificación del tipo.
IDE = "Identificada, no extraída"
ACADEMICA2 = [
    ("E116", "T02", "Mahapatro, M., Roy, S., Nayar, P., Jadhav, A., Panchkaran, S., Srivastava, D. y Prasad, S. (2025). "
     "Behavioral intervention for mental health of pregnant women experiencing domestic violence: A randomized "
     "controlled trial. Frontiers in Public Health, 13, 1541169.", "https://doi.org/10.3389/fpubh.2025.1541169", VP,
     "Experimental", "Ingreso medio (India)", 2025, EAS, "Un paquete de intervención conductual con yoga en el control "
     "prenatal redujo la depresión, el estrés postraumático y la violencia doméstica.", "Favorable"),
    ("E117", "T02", "Arroyo, K., Lundahl, B., Butters, R., Vanderloo, M. y Wood, D. S. (2017). Short-term interventions "
     "for survivors of intimate partner violence: A systematic review and meta-analysis. Trauma, Violence, & Abuse, "
     "18(2), 155–171.", "https://doi.org/10.1177/1524838015602736", VP, "Revisión sistemática",
     "No informado en el resumen", 2017, "No informada", "En 21 estudios, las intervenciones breves tuvieron efectos "
     "grandes (g = 1,02); las cognitivo conductuales adaptadas a sobrevivientes, los mayores.", "Favorable"),
    ("E118", "T02", "Micklitz, H. M., Glass, C. M., Bengel, J. y Sander, L. B. (2024). Efficacy of psychosocial "
     "interventions for survivors of intimate partner violence: A systematic review and meta-analysis. Trauma, "
     "Violence, & Abuse, 25(2), 1000–1017.", "https://doi.org/10.1177/15248380231169481", VP, "Revisión sistemática",
     "No informado en el resumen", 2024, "Baja (según la revisión)", "En 40 ensayos, las intervenciones psicosociales "
     "redujeron levemente la depresión (DME −0,15) y el estrés postraumático, sin efecto en la revictimización; los "
     "efectos son modestos y no se mantienen.", "Mixto"),
    ("E119", "T06", "Noble, E., Ward, L., French, S. y Falb, K. (2019). State of the evidence: A systematic review of "
     "approaches to reduce gender-based violence and support the empowerment of adolescent girls in humanitarian "
     "settings. Trauma, Violence, & Abuse, 20(3), 428–434.", "https://doi.org/10.1177/1524838017699601", VP,
     "Revisión sistemática", "Ingreso bajo y medio (contextos humanitarios)", 2019,
     "Baja (evaluaciones antes y después, según la revisión)", "Pocos programas evaluados para adolescentes en crisis "
     "humanitarias, con resultados prometedores y sin evaluaciones rigurosas.", "Mixto"),
    ("E120", "T07; T06", "Sabri, B., Sellke, R., Smudde, M., Bourey, C. y Murray, S. M. (2023). Gender-based violence "
     "interventions in low- and middle-income countries: A systematic review of interventions at structural, community, "
     "interpersonal, individual, and multiple levels. Trauma, Violence, & Abuse, 24(5), 3170–3186.",
     "https://doi.org/10.1177/15248380221126181", VP, "Revisión sistemática", "Ingreso bajo y medio", 2023, "No informada",
     "La mayoría de las intervenciones, en varios niveles, tuvo resultados positivos en la victimización; hay menos "
     "evidencia para reducir la perpetración.", "Favorable"),
    ("E121", "T07", "Dworkin, S. L., Treves-Kagan, S. y Lippman, S. A. (2013). Gender-transformative interventions to "
     "reduce HIV risks and violence with heterosexually-active men: A review of the global evidence. AIDS and Behavior, "
     "17(9), 2845–2863.", "https://doi.org/10.1007/s10461-013-0565-2", VP, "Revisión sistemática", "Global", 2013,
     "No informada", "En 15 estudios, las intervenciones transformadoras de género con hombres pueden prevenir la "
     "violencia de pareja y modificar actitudes inequitativas.", "Favorable"),
    ("E122", "T09", "Erk, M. A. y Sanberk, İ. (2026). School-based prevention programmes for sexual abuse: A "
     "meta-analysis study. Journal of Prevention, 47(2), 217–237.", "https://doi.org/10.1007/s10935-025-00885-4", VP,
     "Revisión sistemática", "No informado en el resumen", 2026, "No informada", "Los programas escolares de prevención "
     "del abuso sexual mejoran conocimientos, actitudes y conductas de protección; son más eficaces con varias sesiones "
     "y en niños más pequeños.", "Favorable"),
    ("E123", "T09", "Piolanti, A., Jouriles, E. N. y Foran, H. M. (2022). Assessment of psychosocial programs to prevent "
     "sexual violence during adolescence: A systematic review and meta-analysis. JAMA Network Open, 5(11), e2240895.",
     "https://doi.org/10.1001/jamanetworkopen.2022.40895", VP, "Revisión sistemática", "No informado en el resumen", 2022,
     "Riesgo de sesgo en varios ensayos, según la revisión", "Los programas se asociaron con menos violencia sexual en "
     "adolescentes, sobre todo en la escuela y con adolescentes mayores.", "Favorable"),
    ("E124", "T11", "Zapata, L. B., Pazol, K., Dehlendorf, C., Curtis, K. M., Malcolm, N. M., Rosmarin, R. B. y "
     "Frederiksen, B. N. (2018). Contraceptive counseling in clinical settings: An updated systematic review. American "
     "Journal of Preventive Medicine, 55(5), 677–690.", "https://doi.org/10.1016/j.amepre.2018.07.006", VP,
     "Revisión sistemática", "No informado en el resumen", 2018, "No informada", "16 de 23 estudios en adultos y 6 de 9 "
     "en adolescentes hallaron un efecto positivo de la consejería anticonceptiva en al menos un resultado.",
     "Favorable"),
    ("E125", "T12", "Marseille, E., Mirzazadeh, A., Biggs, M. A., Miller, A. P., Horvath, H., Lightfoot, M. et al. "
     "(2018). Effectiveness of school-based teen pregnancy prevention programs in the USA: A systematic review and "
     "meta-analysis. Prevention Science, 19(4), 468–489.", "https://doi.org/10.1007/s11121-017-0861-6", VP,
     "Revisión sistemática", "Ingreso alto (Estados Unidos y Canadá)", 2018, "Baja o muy baja (según la revisión)",
     "Sin diferencia consistente en las tasas de embarazo; 24 de 30 comparaciones no fueron significativas.", "Nulo"),
    ("E126", "T14", "Nandi, A., Agarwal, P., Chandrashekar, A., Maloney, S., Richardson, R., Thakur, L. y Harper, S. "
     "(2024). Access to affordable daycare and women's mental health in Rajasthan, India: Evidence from a "
     "cluster-randomised social intervention. Journal of Global Health, 14, 04063.",
     "https://doi.org/10.7189/jogh.14.04063", VP, "Experimental", "Ingreso medio (India)", 2024, EAC, "Las guarderías "
     "comunitarias asequibles tuvieron alta utilización y mostraron potencial para mejorar la salud mental de las "
     "madres; el estudio no mide el empleo como resultado principal.", "Favorable"),
    ("E127", "T14", "Halim, D., Johnson, H. C. y Perova, E. (2019). Preschool availability and female labor force "
     "participation: Evidence from Indonesia. Banco Mundial, Policy Research Working Paper 8915.",
     "https://doi.org/10.1596/1813-9450-8915", VR("la API del Banco Mundial y Crossref (DOI)"), "Cuasi experimental",
     "Ingreso medio (Indonesia)", 2019, "No informada", "Una escuela inicial pública adicional por cada 1000 niños "
     "aumentó entre 11 % y 16 % la participación laboral de las madres.", "Favorable"),
    ("E128", "T14", "Abdurazzakova, D., Niu, C. y Purevjav, A.-O. (2025). Working for yourself or for your kids? Childcare "
     "expansion policy in Uzbekistan. Banco Mundial, Policy Research Working Paper 11038.",
     "https://doi.org/10.1596/1813-9450-11038", VR("la API del Banco Mundial y Crossref (DOI)"), "Cuasi experimental",
     "Ingreso medio (Uzbekistán)", 2025, "No informada", "La expansión de la educación inicial (2018 a 2022) aumentó en "
     "12 % la oferta laboral femenina; el efecto se explica más por el costo del cuidado que por el tiempo.",
     "Favorable"),
    ("E129", "T14", "Robayo-Abril, M. y Rude, B. (2023). Preparatory school years and maternal employment in Romania. "
     "Banco Mundial, Policy Research Working Paper 10638.", "https://doi.org/10.1596/1813-9450-10638",
     VR("la API del Banco Mundial y Crossref (DOI)"), "Cuasi experimental", "Ingreso medio (Rumania, al momento de la "
     "reforma)", 2023, "No informada", "El año preparatorio escolar universal aumentó el empleo de las madres de niños "
     "de seis años.", "Favorable"),
    ("E131", "T19", "Nardi, A. L., Frankenberg, A. D. V., Franzosi, O. S. y Santo, L. C. E. (2020). Impacto dos aspectos "
     "institucionais no aleitamento materno em mulheres trabalhadoras: uma revisão sistemática. Ciência & Saúde "
     "Coletiva, 25(4), 1445–1462.", "https://doi.org/10.1590/1413-81232020254.20382018", VP, "Revisión sistemática",
     "No informado en el resumen", 2020, "No informada", "Las salas de apoyo a la lactancia, la reducción de jornada y el "
     "tiempo para extraer leche se asocian positivamente con la lactancia de las trabajadoras.", "Favorable"),
    ("E132", "T19", "Abdulwadud, O. A. y Snow, M. E. (2012). Interventions in the workplace to support breastfeeding for "
     "women in employment. Cochrane Database of Systematic Reviews, CD006177.pub3.",
     "https://doi.org/10.1002/14651858.CD006177.pub3", VP, "Revisión sistemática", "Global", 2012,
     "No aplica (sin ensayos)", "No se identificaron ensayos aleatorios ni cuasi aleatorios sobre intervenciones en el "
     "lugar de trabajo para apoyar la lactancia.", "Sin estimación"),
    ("E133", "T19", "Kim, J. H., Shin, J. C. y Donovan, S. M. (2019). Effectiveness of workplace lactation interventions "
     "on breastfeeding outcomes in the United States: An updated systematic review. Journal of Human Lactation, 35(1), "
     "100–113.", "https://doi.org/10.1177/0890334418765464", VP, "Revisión sistemática", "Ingreso alto (Estados Unidos)",
     2019, "No informada", "Las intervenciones de lactancia en el lugar de trabajo aumentaron el inicio, la duración y la "
     "exclusividad de la lactancia.", "Favorable"),
    ("E134", "T19", "van Dellen, S. A., Wisse, B. y Mobach, M. P. (2022). Effects of lactation room quality on working "
     "mothers' feelings and thoughts related to breastfeeding and work: A randomized controlled experiment. "
     "International Breastfeeding Journal, 17, 57.", "https://doi.org/10.1186/s13006-022-00499-0", VP, "Experimental",
     "No informado en el resumen", 2022, "Experimento de campo", "Usar un lactario de alta calidad redujo el estrés de "
     "las madres frente a uno de baja calidad.", "Favorable"),
    ("E135", "T19", "Núñez Pineda, C. M., Franco Soto, J. V. et al. (2017). Impacto del lactario domiciliario en madres "
     "trabajadoras. Hospital Central de San Cristóbal, 2015. Archivos Venezolanos de Puericultura y Pediatría.",
     "https://openalex.org/W2781891875", VR("OpenAlex (registro de la revista)"), "Cuasi experimental",
     "Ingreso medio (Venezuela)", 2017, "No informada", "Las madres trabajadoras que crearon un lactario domiciliario "
     "abandonaron menos la lactancia exclusiva (28 % frente a 60 %).", "Favorable"),
    ("E136", "T04", "Diez-Canseco, F., Toyama, M., Hidalgo-Padilla, L. y Bird, V. J. (2022). Systematic review of "
     "policies and interventions to prevent sexual harassment in the workplace in order to prevent depression. "
     "International Journal of Environmental Research and Public Health, 19(20), 13278.",
     "https://doi.org/10.3390/ijerph192013278", VP, "Revisión sistemática", "No informado en el resumen", 2022,
     "No informada", "Evidencia limitada sobre la efectividad de las políticas y capacitaciones contra el hostigamiento "
     "sexual laboral, centrada en conocimientos y actitudes.", "Mixto"),
    ("E137", "T04", "Dalsbø, T. K., Greve, R. A., Jørgensen, I. L. y Fønhus, M. S. (2026). Education and training "
     "interventions for healthcare workers to prevent sexual harassment. Cochrane Database of Systematic Reviews, "
     "CD016096.pub2.", "https://doi.org/10.1002/14651858.CD016096.pub2", VP, "Revisión sistemática",
     "Ingreso alto (Taiwán)", 2026, "No informada", "Dos ensayos en Taiwán compararon dos tipos de educación; ningún "
     "estudio comparó la capacitación con la ausencia de intervención.", "Sin estimación"),
    ("E138", "T04", "Chang, T.-S., Tzeng, Y.-L., Teng, Y.-K. y Chen, C.-H. (2025). Effectiveness of an e-book in enhancing "
     "knowledge, coping behaviors and preventive strategies for sexual harassment prevention among new nurses: A "
     "randomized controlled trial. Nurse Education in Practice, 82, 104198.", "https://doi.org/10.1016/j.nepr.2024.104198",
     VP, "Experimental", "Ingreso alto (Taiwán)", 2025, EAS, "Un libro electrónico interactivo mejoró los conocimientos "
     "y las estrategias de prevención del hostigamiento en enfermeras nuevas.", "Favorable"),
    ("E139", "T10", "Rocha, I. C. O. y Valença, A. M. (2023). The efficacy of CBT based interventions to sexual offenders: "
     "A systematic review of the last decade literature. International Journal of Law and Psychiatry, 87, 101856.",
     "https://doi.org/10.1016/j.ijlp.2022.101856", VP, "Revisión sistemática", "No informado en el resumen", 2023,
     "No informada", "En ocho estudios, las intervenciones cognitivo conductuales son prometedoras para reducir la "
     "reincidencia.", "Favorable"),
    ("E140", "T10", "Lösel, F., Link, E., Schmucker, M., Bender, D., Breuer, M., Carl, L. et al. (2020). On the "
     "effectiveness of sexual offender treatment in prisons: A comparison of two different evaluation designs in "
     "routine practice. Sexual Abuse, 32(4), 452–475.", "https://doi.org/10.1177/1079063219871576", VP,
     "Cuasi experimental", "Ingreso alto (Alemania)", 2020, "No informada", "Resultados mayormente no significativos; "
     "efecto en la reincidencia general con un método y tendencias opuestas en la sexual según el método.", "Mixto"),
    ("E141", "T10", "Aebi, M., Krause, C., Barra, S., Vogt, G., Vertone, L., Manetsch, M. et al. (2022). What kind of "
     "therapy works with juveniles who have sexually offended? A randomized-controlled trial of two versions of a "
     "specialized cognitive behavioral outpatient treatment program. Sexual Abuse, 34(8), 973–1002.",
     "https://doi.org/10.1177/10790632211070804", VP, "Experimental", "Ingreso alto (Suiza)", 2022, EAS,
     "La terapia centrada en la conducta sexual tuvo menos reincidencia sexual que la de habilidades generales.",
     "Favorable"),
    ("E142", "T10", "Sousa, M., Andrade, J., de Castro-Rodrigues, A. y Gonçalves, R. A. (2023). The effectiveness of "
     "psychological treatment in adult male convicted for sexual offenses against children: A systematic review. "
     "Trauma, Violence, & Abuse, 24(3), 1867–1881.", "https://doi.org/10.1177/15248380221082080", VP,
     "Revisión sistemática", "No informado en el resumen", 2023, "No informada", "El diseño y la calidad de los estudios "
     "limitan la acumulación de evidencia sobre la efectividad del tratamiento.", "Sin estimación"),
    ("E143", "T10", "Redondo Illescas, S. (2006). ¿Sirve el tratamiento para rehabilitar a los delincuentes sexuales? "
     "Revista Española de Investigación Criminológica, 4, 1–22.", "https://doi.org/10.46381/reic.v4i0.26", VC,
     "Cuasi experimental", "Ingreso alto (España)", 2006, "No informada", "El programa cognitivo conductual de la "
     "prisión de Brians mostró un efecto rehabilitador frente al grupo control, condicionado por factores de riesgo "
     "estáticos.", "Favorable"),
    ("E144", "T13", "Pang, X., Zeng, J. y Rozelle, S. (2013). Does women's knowledge of voting rights affect their voting "
     "behaviour in village elections? Evidence from a randomized controlled trial in China. The China Quarterly, 213, "
     "39–59.", "https://doi.org/10.1017/s0305741012001531", VR("EconBiz y Crossref (DOI)"), "Experimental",
     "Ingreso medio (China)", 2013, EAS, "La capacitación a mujeres sobre sus derechos electorales aumentó su "
     "conocimiento y su votación; capacitar solo a los líderes no tuvo efecto.", "Favorable"),
    ("E145", "T21", "Pulikkamath, A. y Sunny, A. (2025). A decade of subnational gender budgeting and its effects in "
     "India: A difference-in-differences analysis. International Public Management Journal, 29(3), 284–300.",
     "https://doi.org/10.1080/10967494.2025.2536267", VR("EconBiz y Crossref (DOI)"), "Cuasi experimental",
     "Ingreso medio (India)", 2025, "No informada", "Los estados que aplican presupuestos con enfoque de género superan "
     "a los demás en indicadores de desarrollo y empoderamiento de las mujeres.", "Favorable"),
    ("E146", "T24", "Querejeta, M., Olivieri, C., Tomio, A., Castaneda, J. L. y Muñoz Boudet, A. M. (2023). Sharing "
     "parental leave between mothers and fathers: Experimental evidence from a messaging intervention in Uruguay. Banco "
     "Mundial, Policy Research Working Paper 10340.", "https://doi.org/10.1596/1813-9450-10340",
     VR("la API del Banco Mundial y Crossref (DOI)"), "Experimental", "Ingreso alto (Uruguay)", 2023, EAS,
     "Los mensajes aumentaron el conocimiento de la licencia parental y cambiaron normas entre las mujeres, con impacto "
     "limitado en el uso efectivo de la licencia por los padres.", "Mixto"),
    ("E147", "T11", "Gil Gonzales, V. R. (2026). Efectos de programa estructurado de planificación familiar en "
     "conocimiento sobre prevención del embarazo adolescente en Cusco. Vive, Revista de Investigación en Salud, 9(26), "
     "1–14.", "https://doi.org/10.33996/revistavive.v9i26.500", VR("OpenAlex y Crossref (DOI)"), "Cuasi experimental",
     "Perú", 2026, "No informada (260 estudiantes, grupo control no equivalente)", "Un programa de ocho sesiones en "
     "colegios públicos del Cusco aumentó el conocimiento sobre prevención del embarazo (d = 1,02 en mujeres y 0,68 en "
     "hombres).", "Favorable"),
    ("E148", "T11", "Ramos Taype, Z. Z. y Oquendo Olivera, S. K. (2021). Efectos del programa de consejería de métodos de "
     "planificación familiar en puérperas adolescentes del C.M.I. San José, Villa El Salvador, Lima. Alpha Centauri, "
     "2(4), 26–35.", "https://doi.org/10.47422/ac.v2i4.63", VR("OpenAlex y Crossref (DOI)"), "Cuasi experimental", "Perú",
     2021, "No informada (60 puérperas)", "La consejería en planificación familiar aumentó el conocimiento de puérperas "
     "adolescentes frente al grupo control.", "Favorable"),
    ("E149", "T23", "Solís-Ceballos, C., Rodríguez-López, M. y Rebollo-Catalán, Á. (2026). Formación inicial del "
     "profesorado en coeducación e igualdad de género: una revisión sistemática de experiencias de éxito. Teoría de la "
     "Educación. Revista Interuniversitaria, 38(1), 61–84.", "https://doi.org/10.14201/teri.32579",
     VR("OpenAlex y Crossref (DOI)"), "Revisión sistemática", "No informado en el resumen", 2026,
     "Baja (debilidades metodológicas, según la revisión)", "Las experiencias de formación docente en coeducación tienen "
     "evaluaciones con debilidades metodológicas y evidencia de logro muy frágil.", "Sin estimación"),
]
# Fuentes halladas en la búsqueda web de evaluaciones peruanas y en los portales (SVRI y What Works)
PORTALES_EV = [
    ("E130", "T14", "Boyd, C. y Rentería, J. M. (2016). Economía del cuidado, desigualdades de género y participación en "
     "el mercado laboral: el caso de Cuna Más. Investigaciones, Consorcio de Investigación Económica y Social (CIES).",
     "https://ideas.repec.org/a/bbj/invcie/618.html", VR("RePEc (IDEAS)"), "Cuasi experimental", "Perú", 2016,
     "No informada", "Sin efecto del Servicio de Cuidado Diurno en el empleo de las madres con emparejamiento por puntaje "
     "de propensión; con balance por entropía, aumento de 14 %; efectos negativos en el empleo informal.", "Mixto"),
    ("E150", "T01", "Calderon-Anyosa, R., Koski, A., Carcamo, C., Garcia, P. J., Nandi, A. y Kaufman, J. (2025). Impact "
     "of Peru's women's emergency centres on the reporting of physical, psychological and sexual intimate partner "
     "violence. Journal of Epidemiology and Community Health, 79(8), 646–652.", "https://doi.org/10.1136/jech-2024-222140",
     VC, "Cuasi experimental", "Perú", 2025, "No informada", "Con diferencias en diferencias escalonadas (ENDES 2004 a "
     "2016), los CEM aumentaron en 3 puntos porcentuales el reporte de violencia de pareja, sobre todo psicológica; sin "
     "efectos concluyentes en la física y la sexual.", "Mixto"),
    ("E151", "T01; T02", "Kirk, L., Terry, S., Lokuge, K. y Watterson, J. L. (2017). Effectiveness of secondary and "
     "tertiary prevention for violence against women in low and low-middle income countries: A systematic review. BMC "
     "Public Health, 17, 622.", "https://doi.org/10.1186/s12889-017-4502-6", VP, "Revisión sistemática",
     "Ingreso bajo y medio (ingreso bajo y medio bajo)", 2017, "Baja (evidencia débil, según la revisión)",
     "En 22 estudios la evidencia es débil; hay indicios a favor de intervenciones sobre el alcohol, de la psicoterapia "
     "para sobrevivientes de violencia sexual y de los centros de crisis para el acceso a servicios.", "Mixto"),
]
EVIDENCIA += ACADEMICA2 + PORTALES_EV
ORIGEN_PORTALES = {"E130": "Búsqueda web de evaluaciones peruanas (WEB-64)", "E150": "Búsqueda web de evaluaciones "
                   "peruanas (WEB-65)", "E151": "Búsqueda en portales: SVRI (WEB-73)"}
SELECCION += [
    ("ACA-005|PMID:40552233", "T02", INC, "E116", ""),
    ("ACA-005|PMID:26335794", "T02", INC, "E117", ""),
    ("ACA-005|PMID:37148270", "T02", INC, "E118", ""),
    ("ACA-024|PMID:29334024", "T06", INC, "E119", ""),
    ("ACA-028|PMID:36226579", "T07; T06", INC, "E120", ""),
    ("ACA-028|PMID:23934267", "T07", INC, "E121", ""),
    ("ACA-037|PMID:41460577", "T09", INC, "E122", ""),
    ("ACA-037|PMID:36346627", "T09", INC, "E123", ""),
    ("ACA-046|PMID:30342631", "T11", INC, "E124", ""),
    ("ACA-050|PMID:29374797", "T12", INC, "E125", ""),
    ("ACA-060|PMID:39513282", "T14", INC, "E126", ""),
    ("ACA-118|BM:31195258", "T14", INC, "E127", ""),
    ("ACA-118|BM:34446022", "T14", INC, "E128", ""),
    ("ACA-118|BM:34212724", "T14", INC, "E129", ""),
    ("ACA-081|PMID:32267445", "T19", INC, "E131", ""),
    ("ACA-081|PMID:23076920", "T19", INC, "E132", ""),
    ("ACA-081|PMID:29928834", "T19", INC, "E133", ""),
    ("ACA-081|PMID:35945573", "T19", INC, "E134", ""),
    ("ACA-147|OPENALEX:W2781891875", "T19", INC, "E135", ""),
    ("ACA-014|PMID:36293858", "T04", INC, "E136", ""),
    ("ACA-014|PMID:42861686", "T04", INC, "E137", ""),
    ("ACA-014|PMID:39612767", "T04", INC, "E138", ""),
    ("ACA-042|PMID:36599224", "T10", INC, "E139", ""),
    ("ACA-042|PMID:31451086", "T10", INC, "E140", ""),
    ("ACA-042|PMID:35230203", "T10", INC, "E141", ""),
    ("ACA-042|PMID:35411837", "T10", INC, "E142", ""),
    ("ACA-138|OPENALEX:W157072455", "T10", INC, "E143", ""),
    ("ACA-055|ECONBIZ:10014159172", "T13", INC, "E144", ""),
    ("ACA-088|ECONBIZ:10016079729", "T21", INC, "E145", ""),
    ("ACA-123|BM:34015996", "T24", INC, "E146", ""),
    ("ACA-139|OPENALEX:W7169867796", "T11", INC, "E147", ""),
    ("ACA-139|OPENALEX:W4212989561", "T11", INC, "E148", ""),
    ("ACA-151|OPENALEX:W7117897845", "T23", INC, "E149", ""),
    ("ACA-005|PMID:20601730", "T02", EXC, "", "El resumen no informa resultados; estudio en Hong Kong (ingreso alto)."),
    ("ACA-046|PMID:30342637", "T11", EXC, "", "El resumen no informa la dirección de los efectos."),
    ("ACA-059|ECONBIZ:10014500537", "T14", EXC, "", SIN_RESUMEN),
    ("ACA-059|ECONBIZ:10012841571", "T14", EXC, "", SIN_RESUMEN),
    ("ACA-117|BM:16974758", "T13", EXC, "", "Intervención distinta: participación obligatoria de mujeres en consejos de desarrollo, no educación electoral ni asistencia técnica."),
    ("ACA-120|BM:27678214", "T13", EXC, "", "Intervención distinta: programa de medios de vida con grupos de mujeres."),
    ("ACA-131|OPENALEX:W7140321821", "T03", EXC, "", "Revisión sin estimación de efectos."),
    ("ACA-131|OPENALEX:W7220450246", "T03", EXC, "", "Revisión bibliográfica cualitativa sin estimación de efectos."),
    ("ACA-131|OPENALEX:W4386798084", "T03", EXC, "", "Diseño no experimental (encuesta a 25 magistrados)."),
    ("ACA-138|OPENALEX:W2244331379", "T10", EXC, "", SIN_RESUMEN),
    ("ACA-138|OPENALEX:W68348789", "T10", EXC, "", "No se pudo verificar la revista de publicación."),
    ("ACA-142|OPENALEX:W7117580117", "T14", EXC, "", "Diseño no experimental descriptivo correlacional."),
    ("ACA-142|OPENALEX:W7221739864", "T14", EXC, "", "Tesis no arbitrada."),
    ("ACA-142|OPENALEX:W4407871708", "T14", EXC, "", "No se pudo verificar la autoría: Crossref no la registra y la editorial bloqueó la consulta."),
    ("ACA-142|OPENALEX:W2924111035", "T14", EXC, "", "Resultado distinto: desarrollo infantil, no empleo de las madres."),
    ("ACA-147|OPENALEX:W3206814975", "T19", EXC, "", "Intervención distinta: promoción de la salud laboral, no lactancia."),
    ("ACA-151|OPENALEX:W4406397277", "T23", EXC, "", "Revisión sin estimación de efectos."),
    ("ACA-152|OPENALEX:W4411215010", "T24", EXC, "", "Resultado distinto: lactancia, no distribución del trabajo doméstico y de cuidado."),
    ("ACA-129|OPENALEX:W7154490168", "T01", EXC, "", "Revisión sin estimación de efectos."),
    ("ACA-014|PMID:42894178", "T04", EXC, "", "Población distinta: universidades; el servicio se dirige al hostigamiento laboral."),
    ("ACA-010|ECONBIZ:10012059191", "T03", EXC, "", SIN_RESUMEN),
    ("ACA-138|OPENALEX:W2895581372", "T10", YA, "E16", "Resumen en español de la misma revisión."),
    ("ACA-107|BM:25470471", "T03", YA, "E04", "Versión de documento de trabajo del mismo estudio."),
    ("ACA-107|BM:32990600", "T03", YA, "E01", "Versión de documento de trabajo del mismo estudio."),
    ("ACA-118|BM:33568630", "T14", YA, "E23", "Versión de documento de trabajo de la misma revisión."),
    ("ACA-024|PMID:37861113", "T06", YA, "E08", ""),
    ("ACA-028|PMID:37861113", "T07", YA, "E08", ""),
    ("ACA-024|PMID:32302308", "T06", IDE, "", ""),
    ("ACA-024|PMID:32257154", "T06", IDE, "", ""),
    ("ACA-024|PMID:29972287", "T06", IDE, "", ""),
    ("ACA-024|PMID:29609568", "T06", IDE, "", ""),
    ("ACA-024|PMID:29574438", "T06", IDE, "", ""),
    ("ACA-134|OPENALEX:W4205600850", "T06", IDE, "", ""),
    ("ACA-028|PMID:32467354", "T07", IDE, "", ""),
    ("ACA-028|PMID:32810146", "T07", IDE, "", ""),
    ("ACA-111|BM:34447966", "T07", IDE, "", ""),
    ("ACA-135|OPENALEX:W4412872645", "T07", IDE, "", ""),
    ("ACA-037|PMID:37569038", "T09", IDE, "", ""),
    ("ACA-037|PMID:33446045", "T09", IDE, "", ""),
    ("ACA-113|BM:40065480", "T09", IDE, "", ""),
    ("ACA-132|OPENALEX:W4404247036", "T09", IDE, "", ""),
    ("ACA-137|OPENALEX:W4417406123", "T09", IDE, "", ""),
    ("ACA-137|OPENALEX:W3004161625", "T09", IDE, "", ""),
    ("ACA-137|OPENALEX:W4313186686", "T09", IDE, "", ""),
    ("ACA-046|PMID:35727080", "T11", IDE, "", ""),
    ("ACA-046|PMID:29372086", "T11", IDE, "", ""),
    ("ACA-115|BM:33427864", "T11", IDE, "", ""),
    ("ACA-139|OPENALEX:W7104516604", "T11", IDE, "", ""),
    ("ACA-139|OPENALEX:W7118709519", "T11", IDE, "", ""),
    ("ACA-050|PMID:31660918", "T12", IDE, "", ""),
    ("ACA-050|PMID:35724585", "T12", IDE, "", ""),
    ("ACA-050|PMID:26346692", "T12", IDE, "", ""),
    ("ACA-113|BM:30990271", "T12", IDE, "", ""),
    ("ACA-116|BM:34280293", "T12", IDE, "", ""),
    ("ACA-140|OPENALEX:W4318762617", "T12", IDE, "", ""),
    ("ACA-072|ECONBIZ:10014263099", "T17", IDE, "", ""),
    ("ACA-121|BM:18030787", "T17", IDE, "", ""),
    ("ACA-121|BM:16530040", "T17", IDE, "", ""),
    ("ACA-121|BM:33725383", "T17", IDE, "", ""),
    ("ACA-121|BM:26286952", "T17", IDE, "", ""),
    ("ACA-115|BM:19338849", "T17", IDE, "", ""),
    ("ACA-110|BM:30867947", "T17", IDE, "", ""),
    ("ACA-118|BM:16704314", "T17", IDE, "", ""),
    ("ACA-121|BM:20346367", "T17", IDE, "", ""),
    ("ACA-077|ECONBIZ:10014337080", "T18", IDE, "", ""),
    ("ACA-077|ECONBIZ:10014358711", "T18", IDE, "", ""),
    ("ACA-077|ECONBIZ:10012566984", "T18", IDE, "", ""),
    ("ACA-122|BM:34432241", "T18", IDE, "", ""),
    ("ACA-122|BM:40024555", "T18", IDE, "", ""),
    ("ACA-122|BM:34168976", "T18", IDE, "", ""),
    ("ACA-122|BM:40010530", "T18", IDE, "", ""),
    ("ACA-146|OPENALEX:W7202271106", "T18", IDE, "", ""),
    ("ACA-119|BM:32908035", "T15", IDE, "", ""),
    ("ACA-119|BM:33129225", "T15", IDE, "", ""),
    ("ACA-119|BM:14833489", "T15", IDE, "", ""),
    ("ACA-064|ECONBIZ:10012973143", "T15", IDE, "", ""),
    ("ACA-064|ECONBIZ:10012970863", "T15", IDE, "", ""),
    ("ACA-064|ECONBIZ:10012702407", "T15", IDE, "", ""),
    ("ACA-064|ECONBIZ:10011637919", "T15", IDE, "", ""),
    ("ACA-105|BM:6030503", "T01", IDE, "", ""),
    ("ACA-009|PMID:31254283", "T01", IDE, "", ""),
    ("ACA-032|PMID:27002859", "T08", IDE, "", ""),
]
# Portales y búsqueda web de la ronda 2: cadena, buscador o portal, resultados revisados, tipos
JPAL = "Portal J-PAL (povertyactionlab.org/evaluations)"
G = {"G1": "T01; T02; T03; T05", "G2": "T06; T07; T08; T09", "G3": "T04", "G5": "T11; T12", "G6": "T14; T19; T24",
     "G7": "T15; T16; T17; T18; T20", "G8": "T13; T21; T22; T23", "PE": "T01; T06; T08; T11; T14; T17; T18; T20"}
BUSQUEDAS += [
    ("intimate partner violence", JPAL, 27, G["G1"]), ("violence against women police", JPAL, 27, G["G1"]),
    ("gender norms violence", JPAL, 27, G["G2"]), ("dating violence school", JPAL, 30, G["G2"]),
    ("sexual harassment", JPAL, 27, G["G3"]), ("adolescent pregnancy", JPAL, 27, G["G5"]),
    ("sexual education", JPAL, 30, G["G5"]), ("childcare", JPAL, 28, G["G6"]), ("breastfeeding", JPAL, 29, G["G6"]),
    ("fathers childcare", JPAL, 32, G["G6"]), ("vocational training women", JPAL, 27, G["G7"]),
    ("business training women", JPAL, 29, G["G7"]), ("land titling", JPAL, 27, G["G7"]),
    ("conditional cash transfers schooling", JPAL, 30, G["G7"]), ("women political participation", JPAL, 30, G["G8"]),
    ("gender attitudes school", JPAL, 28, G["G8"]), ("gender budgeting", JPAL, 30, G["G8"]),
    ("Peru gender", JPAL, 33, G["PE"]),
    ("evaluación de impacto Cuna Más cuidado diurno empleo de las madres Perú", "Buscador web general (evaluaciones "
     "peruanas)", 10, "T14"),
    ("evaluación de impacto Centros Emergencia Mujer Perú violencia", "Buscador web general (evaluaciones peruanas)", 10,
     "T01; T03"),
    ("MEF evaluación independiente programa presupuestal 1002 violencia contra la mujer evaluación de diseño y "
     "ejecución", "Buscador web general (evaluaciones peruanas)", 20, "T01; T03; T06"),
    ("evaluación de impacto Jóvenes Productivos Impulsa Perú mujeres empleo", "Buscador web general (evaluaciones "
     "peruanas)", 10, "T17"),
    ("Wawa Wasi evaluación de impacto participación laboral madres Perú", "Buscador web general (evaluaciones peruanas)",
     10, "T14"),
    ("evaluación programa Hombres por el Cambio Perú violencia resultados", "Buscador web general (evaluaciones "
     "peruanas)", 10, "T07"),
    ("evaluación de impacto educación sexual integral Perú escuelas resultados", "Buscador web general (evaluaciones "
     "peruanas)", 10, "T12"),
    ("mef.gob.pe eval_indep evaluaciones independientes presupuesto por resultados violencia familiar Cuna Más Juntos",
     "Buscador web general (evaluaciones peruanas)", 20, "T01; T14; T20"),
    ("What Works to Prevent Violence against Women and Girls global programme results synthesis interventions effective "
     "design elements Jewkes", "Buscador web general (What Works)", 10, "T06; T07; T08"),
    ("SVRI evidence review prevention violence against women low and middle income countries interventions "
     "effectiveness", "Buscador web general (SVRI)", 10, "T01; T02; T06"),
]
NO_EXTRAIDA = "Identificada en un portal; no se extrajo porque una fuente más no cambia la clasificación del tipo."
CRIBADO2 = [
    ("Sukhtankar, S., Kruks-Wisner, G. y Mangla, A. (2022). Policing in patriarchy: An experimental evaluation of reforms "
     "to improve police responsiveness to women in India. Science, 377(6602), 191–198.", "T03",
     "https://doi.org/10.1126/science.abm7387", NO_EXTRAIDA, "Portal J-PAL"),
    ("Chong, A., Gonzalez-Navarro, M., Karlan, D. y Valdivia, M. (2020). Do information technologies improve teenagers' "
     "sexual education? Evidence from a randomized evaluation in Colombia. Banco Mundial.", "T12",
     "https://doi.org/10.1596/36715", NO_EXTRAIDA, "Portal J-PAL y Banco Mundial"),
    ("Banerjee, A., La Ferrara, E. y Orozco-Olvera, V. (2019). The entertaining way to behavioral change: Fighting HIV "
     "with MTV. Banco Mundial, Policy Research Working Paper 8998.", "T11; T08", "https://doi.org/10.1596/1813-9450-8998",
     NO_EXTRAIDA, "Portal J-PAL"),
    ("Fulu, E., Kerr-Wilson, A. y Lang, J. (2014). What works to prevent violence against women and girls? Evidence "
     "review of interventions to prevent violence against women and girls. DFID.", "T06; T07",
     "https://assets.publishing.service.gov.uk/media/5a7e41f340f0b6230268a2f1/evidence-review-interventions-F.pdf",
     "Revisión rápida narrativa identificada en la búsqueda de What Works; no cambia la clasificación de los tipos.",
     "Búsqueda en portales: What Works"),
    ("J-PAL. The impact of a text-based campaign on intimate partner violence in Peru.", "T08",
     "https://www.povertyactionlab.org/evaluation/impact-text-based-campaign-intimate-partner-violence-peru",
     "Evaluación en curso; sin resultados publicados al 31-10-2026.", "Portal J-PAL"),
    ("J-PAL. Training local leaders to prevent gender-based violence in Peru.", "T06",
     "https://www.povertyactionlab.org/evaluation/training-local-leaders-prevent-gender-based-violence-peru",
     "Evaluación en curso; sin resultados publicados al 31-10-2026.", "Portal J-PAL"),
    ("J-PAL. Mass media experiments to reduce violence against women in Tanzania.", "T08",
     "https://www.povertyactionlab.org/evaluation/mass-media-experiments-reduce-violence-against-women-tanzania",
     "Evaluación en curso; sin resultados publicados al 31-10-2026.", "Portal J-PAL"),
    ("J-PAL. The impact of a gender-sensitive training for police officers on gender-based violence in India.", "T05",
     "https://www.povertyactionlab.org/evaluation/impact-gender-sensitive-training-police-officers-gender-based-violence-"
     "india", "Ficha de J-PAL sin resultados publicados al 31-10-2026.", "Portal J-PAL"),
    ("J-PAL. Increasing women's local political participation through top-down and bottom-up training in Indonesia.",
     "T13", "https://www.povertyactionlab.org/evaluation/increasing-womens-local-political-participation-through-top-"
     "down-and-bottom-training", "Evaluación en curso; sin resultados publicados al 31-10-2026.", "Portal J-PAL"),
    ("J-PAL. Empowering female leaders and voters in Rajasthan, India.", "T13",
     "https://www.povertyactionlab.org/evaluation/empowering-female-leaders-and-voters-rajasthan-india",
     "Intervención distinta: cuotas y campaña general de información al votante.", "Portal J-PAL"),
    ("Sharma, G. (2026). Tackling sexual harassment: Short and long-run experimental evidence from India. SSRN.", "T04",
     "https://doi.org/10.2139/ssrn.6460764", "Población distinta: estudiantes universitarios; el servicio se dirige al "
     "hostigamiento laboral.", "Portal J-PAL"),
    ("J-PAL. A safer monitoring tool to help workers report harassment in Bangladesh (NBER Working Paper 31011).", "T04",
     "https://www.povertyactionlab.org/evaluation/safer-monitoring-tool-help-workers-report-harassment-bangladesh",
     "Intervención distinta: método de medición del hostigamiento en encuestas, no "
     "un servicio de prevención o atención.", "Portal J-PAL"),
    ("J-PAL. The impact of psychosocial therapy and vocational skills training on intimate partner violence in urban "
     "Liberia.", "T02",
     "https://www.povertyactionlab.org/evaluation/impact-psychosocial-therapy-and-vocational-skills-training-intimate-"
     "partner-violence", "No se identificó el documento de resultados; la ficha de J-PAL no enlaza la publicación.",
     "Portal J-PAL"),
    ("J-PAL. Edutainment campaigns to address social norms and perceptions about child marriage in Indonesia.", "T08",
     "https://www.povertyactionlab.org/evaluation/edutainment-campaigns-address-social-norms-and-perceptions-about-child-"
     "marriage", "Resultado distinto: percepciones sobre el matrimonio infantil.", "Portal J-PAL"),
    ("MEF. Informe final de la evaluación de procesos de los Centros Emergencia Mujer (EPROC CEM 22).", "T01",
     "https://mef.gob.pe/contenidos/archivos-descarga/Informe_final_EPROC_CEM_22.pdf", "Evaluación de procesos sin "
     "estimación de efectos; pertinente para el análisis de implementación (REV_DOC y REP_ENT), no para REV_EVI.",
     "Búsqueda web de evaluaciones peruanas"),
    ("MEF (2016). Informe final de la evaluación de diseño y ejecución de presupuesto de Cuna Más.", "T14",
     "https://www.mef.gob.pe/contenidos/presu_publ/ppr/eval_indep/2016_cuna_mas.pdf", "Evaluación de diseño y ejecución "
     "sin estimación de efectos en el empleo de las madres.", "Búsqueda web de evaluaciones peruanas"),
]
# Búsqueda en el Development Evidence Portal de 3ie (31-10-2026), estrategias DEP-15 a DEP-22 ejecutadas mediante la API
# GraphQL del portal (src/analisis_pnig/rev_evi_busqueda_3ie_dep.py). Mismo cribado y misma regla de extracción orientada
# a la decisión que en la ampliación de la búsqueda académica. Los registros repetidos entre estrategias se deciden en
# su primera aparición.
VD = lambda donde: f"{FECHA}: título, autores y año comprobados en el registro del 3ie DEP y en {donde}; resultado tomado del resumen."
ACADEMICA3 = [
    ("E152", "T05", "Romero, O., Peñaranda, C., Paredes, M. y Armbrister, A. (2017). Capacitación de funcionarias de "
     "Comisarías de Familia en Medellín, Colombia, sobre servicios amigables para víctimas de violencia íntima de pareja: "
     "resultados de una evaluación de impacto. Banco Interamericano de Desarrollo.",
     "https://publications.iadb.org/handle/11319/8238",
     f"{FECHA}: título, autores y año comprobados en el registro del 3ie DEP (DEP:7631, con enlace a la publicación del "
     "BID); resultado tomado de la Plataforma de Evidencias del BID (plataformadeevidencias.iadb.org).",
     "Experimental", "Ingreso medio (Colombia)", 2017, EAS, "Capacitación de 20 horas al personal de las Comisarías de "
     "Familia (2014): el reporte formal de violencia contra las mujeres aumentó 7 % y el de violencia de pareja 8 %, y "
     "mejoró la confidencialidad percibida (77 % frente a 69 %); sin efecto en la incidencia de la violencia de pareja "
     "ni en la calidad percibida del servicio.", "Mixto"),
    ("E153", "T17", "Galdo, J. y Chong, A. (2012). Does the quality of public-sponsored training programs matter? "
     "Evidence from bidding processes data. Labour Economics, 19(6), 970–986.",
     "https://doi.org/10.1016/j.labeco.2012.08.001", VC, "Cuasi experimental", "Perú", 2012, "No informada",
     "En ProJoven, con puntaje de propensión generalizado, los beneficiarios de cursos de alta calidad obtienen mayores "
     "ingresos y mejor calidad del empleo que los de cursos de baja calidad y que los no participantes; los retornos son "
     "particularmente robustos para las mujeres; el atributo que más importa es el gasto por capacitado.", "Favorable"),
    ("E154", "T21", "Ghosh, A. (2026). Fiscal pathways to gender equality: Evidence from gender-responsive budgeting in "
     "emerging Asia. International Economics and Economic Policy, 23(4).",
     "https://doi.org/10.1007/s10368-026-00783-y", VC, "Cuasi experimental", "Ingreso medio (15 economías de Asia "
     "emergente)", 2026, "No informada", "Con efectos fijos, variables instrumentales y un control sintético para la "
     "reforma de la India de 2005 (1990 a 2023), el presupuesto con enfoque de género mejora el índice de desarrollo de "
     "género y reduce la desigualdad, solo cuando se cumplen umbrales de calidad institucional, capacidad fiscal y "
     "educación femenina.", "Favorable"),
    ("E155", "T13", "Cheema, A., Khan, S., Liaqat, A. y Mohmand, S. K. (2023). Canvassing the gatekeepers: A field "
     "experiment to increase women voters' turnout in Pakistan. American Political Science Review, 117(1), 1–21.",
     "https://doi.org/10.1017/S0003055422000375", VC, "Experimental", "Ingreso medio (Pakistán)", 2023, EAS,
     "Una campaña no partidaria de promoción del voto dirigida solo a las mujeres no cambió su participación en una "
     "elección nacional; cuando se visitó también a los hombres del hogar para que apoyaran la participación de las "
     "mujeres, la participación femenina aumentó sustancialmente, con más discusión política y apoyo práctico el día "
     "de la elección.", "Mixto"),
    ("E156", "T13", "Ichino, N. y Nathan, N. L. (2017). Political party organization and women's empowerment: A field "
     "experiment in Ghana. Innovations for Poverty Action.",
     "https://www.poverty-action.org/publication/political-party-organization-and-womens-empowerment-field-experiment-ghana",
     VD("la página de Innovations for Poverty Action"), "Experimental", "Ingreso medio (Ghana)", 2017, EAS,
     "Reuniones comunitarias con la Comisión Nacional de Educación Cívica antes de las elecciones de 2016, en 22 pares "
     "de comunidades rurales, para fomentar la participación de las mujeres en la política partidaria local: sin cambios "
     "en la participación de las mujeres ni en las normas de género; la implementación fue incompleta por la resistencia "
     "de líderes partidarios locales.", "Nulo"),
    ("E157", "T13", "Roza, V., Rodríguez-Tejedo, I., Monje Silva, A., Cruz-Aguayo, Y. y Vega, G. (2014). Women's voter "
     "mobilization campaign in Guatemala: A field experiment. Banco Interamericano de Desarrollo.",
     "https://doi.org/10.18235/0012804", VC, "Experimental", "Ingreso medio (Guatemala)", 2014, EAS,
     "Campaña pública de sensibilización sobre la participación cívica y política de las mujeres en 40 municipios, en "
     "las elecciones generales de 2011: sin impacto medible en la participación política de las mujeres.", "Nulo"),
    ("E158", "T13", "McClendon, G., Sperber, E. y Kaaba, O. (2025). Reducing gender gaps in political participation with "
     "efficacy promotion: Evidence from a civic education experiment in Zambia. British Journal of Political Science.",
     "https://doi.org/10.1017/S000712342510121X", VC, "Experimental", "Ingreso medio (Zambia)", 2025, EAS,
     "Un curso de educación cívica con mensajes de promoción de la eficacia política, comparado con el mismo curso solo "
     "informativo, aumentó sustancialmente el interés y la participación política de las mujeres jóvenes y redujo las "
     "brechas de género en un amplio conjunto de resultados conductuales y actitudinales.", "Favorable"),
    ("E159", "T15; T13; T17; T18", "Dekker, M. (2013). Promoting gender equality and female empowerment: A systematic "
     "review of the evidence on property rights, labour markets, political participation and violence against women. "
     "ASC Working Paper 111. African Studies Centre, Leiden.", "http://hdl.handle.net/1887/21859",
     f"{FECHA}: título, autora, año y serie comprobados en el repositorio de la Universidad de Leiden; texto completo "
     "descargado (pp. 25, 53, 54 y 72 a 74).", "Revisión sistemática", "Ingreso bajo y medio (varios países)", 2013,
     "Baja (confianza baja según el 3ie DEP)", "57 estudios. Los programas de titulación de tierras muestran efectos "
     "positivos en la seguridad de la tenencia, la inversión y la productividad; las reformas de herencia en la India "
     "dan resultados mixtos. Las reservas de cargos para mujeres en la India aumentan su participación política, con "
     "resultados mixtos en las políticas. La mayoría de las intervenciones tuvo algún efecto positivo en los resultados, "
     "pero rara vez se midió el impacto en el empoderamiento.", "Mixto"),
    ("E160", "T08", "Evans, W. D., Larson, E. A., McLarnon, C. J., Hauer, M., Marian, M., Agha, S. et al. (2025). "
     "Entertainment media and gender norm transformation interventions for young women and girls in sub-Saharan Africa: "
     "A systematic review. Behavioral Sciences, 15(11), 1596.", "https://doi.org/10.3390/bs15111596", VC,
     "Revisión sistemática", "Ingreso bajo y medio (África subsahariana)", 2025, "Baja (confianza baja según el 3ie DEP)",
     "35 artículos sobre 24 intervenciones en 15 países, con diseños observacionales, cuasi experimentales y aleatorios: "
     "evidencia de efectividad variable, de emergente a demostrada; la revisión concluye que los medios de "
     "entretenimiento son un enfoque efectivo para cambiar normas, actitudes y conductas de género entre adolescentes y "
     "jóvenes, y pide investigación más rigurosa.", "Favorable"),
]
EVIDENCIA += ACADEMICA3
YD = lambda e, nota="": f"Ya incluida ({e}{', ' + nota if nota else ''})"
SELECCION_DEP = [
    # DEP-15: atención, salud mental, justicia y policía (T01, T02, T03, T05)
    ("DEP-15|DEP:11662", "T03", YD("E05"), "E05", ""),
    ("DEP-15|DEP:10211", "T06", EXC, "", "Protocolo de ensayo sin resultados; los resultados (DEP:35120) ya se revisaron en la búsqueda académica"),
    ("DEP-15|DEP:11688", "T02", IDE, "", "Ensayo aleatorio con refugiadas en Tanzania: intervención psicológica integrada para violencia de pareja y malestar psicológico"),
    ("DEP-15|DEP:10415", "T06", EXC, "", "Población y resultado distintos (violencia contra niños)"),
    ("DEP-15|DEP:10996", "T06", IDE, "", "Ensayo aleatorio en grupos de microfinanzas de Tanzania (MAISHA)"),
    ("DEP-15|DEP:11664", "T03", YD("E01", "versión de 2018 del documento de trabajo"), "E01", ""),
    ("DEP-15|DEP:11713", "T06; T07", IDE, "", "Ensayo aleatorio Indashyikirwa en Ruanda: currículo de parejas y activismo comunitario"),
    ("DEP-15|DEP:10411", "T06", IDE, "", "Cuasi experimental en Etiopía: intervención comunitaria de prevención"),
    ("DEP-15|DEP:7631", "T05", INC, "E152", ""),
    ("DEP-15|DEP:10562", "T03", EXC, "", "Resultado distinto (opiniones políticas de las víctimas)"),
    ("DEP-15|DEP:10469", "T07", EXC, "", "Protocolo de revisión sin resultados"),
    ("DEP-15|DEP:34100", "T05", EXC, "", "Mismo estudio que DEP:7631 (registro duplicado en el portal; incluido como E152)"),
    ("DEP-15|DEP:6004", "T07; T06", IDE, "", "Ensayo aleatorio en Costa de Marfil: grupos de diálogo de género añadidos a grupos de ahorro"),
    ("DEP-15|DEP:12942", "T06", IDE, "", "Ensayo aleatorio por conglomerados SAFE en Bangladesh"),
    ("DEP-15|DEP:7192", "T01", IDE, "", "Ensayo aleatorio en clínicas públicas de Ciudad de México: intervención de enfermería para violencia de pareja"),
    ("DEP-15|DEP:10574", "T01; T02", IDE, "", "Ensayo en Egipto: intervención de enfermería basada en el empoderamiento para gestantes"),
    ("DEP-15|DEP:24936", "T06; T07", IDE, "", "Revisión sistemática narrativa en contextos humanitarios"),
    ("DEP-15|DEP:10337", "T02", IDE, "", "Ensayo aleatorio en Irán: habilidades de resolución de problemas para gestantes"),
    ("DEP-15|DEP:10455", "T06; T07", IDE, "", "Revisión sistemática de intervenciones sobre normas de género en países de ingreso bajo y medio bajo"),
    ("DEP-15|DEP:24634", "T01", IDE, "", "Ensayo aleatorio en Turquía: capacitación de enfermeras y parteras sobre violencia de pareja"),
    ("DEP-15|DEP:10462", "T06", EXC, "", "Protocolo de revisión sin resultados"),
    ("DEP-15|DEP:24048", "T01; T06", IDE, "", "Cuasi experimental en Tanzania: búsqueda de atención de sobrevivientes de violencia sexual"),
    ("DEP-15|DEP:24986", "T06", IDE, "", "Evaluación de impacto integrada en Filipinas"),
    ("DEP-15|DEP:26477", "T06; T07", IDE, "", "Evaluación cuantitativa de Zindagii Shoista en Tayikistán"),
    ("DEP-15|DEP:32938", "T01", YD("E150"), "E150", ""),
    ("DEP-15|DEP:10273", "T08", IDE, "", "Evaluación cuasi experimental de la campaña We Can en Bangladesh (Oxfam)"),
    ("DEP-15|DEP:10552", "T02", IDE, "", "Ensayo aleatorio en Colombia: intervención psicosocial con apoyo de trabajadores comunitarios"),
    # DEP-16: prevención comunitaria, hombres, comunicación y escuela (T06 a T09)
    ("DEP-16|DEP:11711", "T08", YD("E13", "versión previa del documento de trabajo"), "E13", ""),
    ("DEP-16|DEP:10265", "T08", IDE, "", "Ensayo aleatorio en Túnez: serie radial sobre violencia contra las mujeres"),
    ("DEP-16|DEP:10251", "T06; T07", IDE, "", "Ensayo piloto por conglomerados en Tanzania con hombres y comunidades"),
    ("DEP-16|DEP:6992", "T09", IDE, "", "Taller educativo sobre violencia en el noviazgo en universitarios de México"),
    ("DEP-16|DEP:10416", "T09", IDE, "", "Cuasi experimental del programa escolar Amor... pero del bueno en México"),
    ("DEP-16|DEP:11717", "T06; T07", IDE, "", "Serie de tiempo interrumpida en Nepal: intervención familiar Sammanit Jeevan con actividades generadoras de ingreso"),
    ("DEP-16|DEP:39969", "T08", IDE, "", "Evaluación experimental de la serie C'est la vie! en Senegal"),
    ("DEP-16|DEP:24446", "T06", IDE, "", "Evaluación de impacto de la iniciativa Ciudades Seguras de Madhya Pradesh"),
    ("DEP-16|DEP:35317", "T07; T09", IDE, "", "Revisión sistemática de intervenciones con adolescentes y jóvenes varones en países de ingreso bajo y medio"),
    ("DEP-16|DEP:10475", "T09; T06", IDE, "", "Revisión de revisiones sobre prevención de la violencia contra adolescentes en países de ingreso bajo y medio"),
    ("DEP-16|DEP:10293", "T08", IDE, "", "Experimentos con medios masivos en Tanzania"),
    ("DEP-16|DEP:11714", "T07", IDE, "", "Subanálisis por matrimonio infantil del ensayo de Costa de Marfil"),
    ("DEP-16|DEP:31439", "T06", IDE, "", "Revisión sistemática de intervenciones conjuntas sobre violencia de pareja y violencia contra niños"),
    ("DEP-16|DEP:34728", "T06; T09", IDE, "", "Análisis secundario cuasi experimental de Rethinking Power en Haití"),
    ("DEP-16|DEP:10178", "T08", YD("E13"), "E13", ""),
    ("DEP-16|DEP:31588", "T06; T08", IDE, "", "Revisión sistemática en la Liga Árabe (ámbito mixto, incluye ingreso alto)"),
    ("DEP-16|DEP:35075", "T08", INC, "E160", ""),
    ("DEP-16|DEP:35151", "T09; T06", IDE, "", "Revisión sistemática What Works sobre violencia contra adolescentes en países de ingreso bajo y medio"),
    ("DEP-16|DEP:10217", "T07; T09", IDE, "", "Programa Parivartan con deportistas escolares en Bombay"),
    ("DEP-16|DEP:30095", "T06", IDE, "", "Evaluación de Change Starts at Home en Nepal con difusión de normas"),
    ("DEP-16|DEP:6014", "T07", IDE, "", "Ensayo piloto por conglomerados con hombres en Costa de Marfil"),
    ("DEP-16|DEP:21989", "T06", IDE, "", "Estudio de costo-efectividad en seis países de ingreso bajo y medio"),
    ("DEP-16|DEP:30399", "T06; T07", IDE, "", "Revisión sistemática y metaanálisis de 26 estudios (2008 a 2022): razón de riesgo 0,85; las intervenciones que incluyen a los hombres son más efectivas"),
    # DEP-17: hostigamiento sexual laboral (T04)
    ("DEP-17|DEP:7451", "T04", EXC, "", "Intervención distinta (programa de empleo rural; resultado violencia de pareja)"),
    ("DEP-17|DEP:14408", "T04", EXC, "", "Resultado distinto (abuso verbal en fábricas, no hostigamiento sexual)"),
    ("DEP-17|DEP:10448", "T07", IDE, "", "Revisión de 58 evaluaciones de programas con hombres y niños (2009)"),
    ("DEP-17|DEP:33133", "T04", EXC, "", "Ámbito distinto (acoso en el transporte público, no el laboral)"),
    ("DEP-17|DEP:10109", "T04", EXC, "", "Tema distinto (matrimonio infantil)"),
    ("DEP-17|DEP:23578", "T04", EXC, "", "Tema distinto (matrimonio infantil)"),
    ("DEP-17|DEP:10277", "T04", EXC, "", "Resultado agregado de cumplimiento laboral en Jordania; no mide el hostigamiento sexual"),
    ("DEP-17|DEP:10268", "T04", EXC, "", "Resultado agregado de cumplimiento laboral en Camboya; no mide el hostigamiento sexual"),
    ("DEP-17|DEP:33078", "T06; T18", IDE, "", "Revisión sistemática de microfinanzas y violencia contra las mujeres (2015 a 2023)"),
    ("DEP-17|DEP:32716", "T09", IDE, "", "Ensayo aleatorio por conglomerados de autodefensa con empoderamiento en Sudáfrica"),
    ("DEP-17|DEP:31238", "T04", EXC, "", "Pregunta distinta (empoderamiento económico y desarrollo humano)"),
    ("DEP-17|DEP:7728", "T04", EXC, "", "Intervención distinta (registro de empresas, sin enfoque de género)"),
    ("DEP-17|DEP:22738", "T17; T18", IDE, "", "Revisión rápida de 57 intervenciones de capacidades económicas de adolescentes"),
    # DEP-18: tratamiento de agresores sexuales (T10): sin resultados en el portal
    # DEP-19: salud sexual y reproductiva y educación sexual integral (T11, T12)
    ("DEP-19|DEP:24912", "T12", IDE, "", "Ensayo aleatorio por conglomerados de educación sexual por internet en China"),
    ("DEP-19|DEP:7181", "T11", IDE, "", "Evaluación de impacto de servicios de planificación familiar amigables para jóvenes en Uganda"),
    ("DEP-19|DEP:15239", "T11", EXC, "", "Población y resultado distintos (riesgo de VIH en adultos)"),
    ("DEP-19|DEP:9432", "T11", EXC, "", "Población distinta (atención preconcepcional)"),
    ("DEP-19|DEP:9125", "T11", IDE, "", "Revisión sobre fecundidad adolescente en países de ingreso bajo y medio (2012)"),
    ("DEP-19|DEP:13191", "T20; T11", IDE, "", "Ensayo de transferencias en Malawi y Zambia: matrimonio temprano y fecundidad"),
    ("DEP-19|DEP:5313", "T12", IDE, "", "Ensayo del programa TeenSTAR en Chile"),
    ("DEP-19|DEP:18805", "T11", EXC, "", "Resultado distinto (asistencia escolar)"),
    ("DEP-19|DEP:14134", "T12", IDE, "", "Ensayo de intervención educativa sobre abstinencia en Ghana"),
    ("DEP-19|DEP:24698", "T11", IDE, "", "Cuasi experimental del programa de parteras en Indonesia"),
    ("DEP-19|DEP:24757", "T12; T11", IDE, "", "Ensayo en Zambia: educación sexual integral vinculada a servicios"),
    ("DEP-19|DEP:22935", "T12", IDE, "", "Panel de 17 países latinoamericanos: la legislación sobre educación sexual escolar se asocia a menor fecundidad adolescente"),
    ("DEP-19|DEP:12970", "T12", EXC, "", "Intervención distinta (habilidades de negociación)"),
    ("DEP-19|DEP:24442", "T11", EXC, "", "Protocolo sin resultados"),
    ("DEP-19|DEP:7800", "T12", IDE, "", "Experimento en Kenia: información sobre riesgo relativo de VIH"),
    ("DEP-19|DEP:10825", "T11", IDE, "", "Experimento natural de clínicas amigables para adolescentes en Sudáfrica"),
    ("DEP-19|DEP:10738", "T12", IDE, "", "Ensayo en Camerún: información de riesgo en escolares"),
    ("DEP-19|DEP:4462", "T11", EXC, "", "No evalúa una intervención"),
    ("DEP-19|DEP:24923", "T20; T11", IDE, "", "Revisión sistemática de transferencias y anticoncepción en países de ingreso bajo y medio"),
    ("DEP-19|DEP:24394", "T12", IDE, "", "Evaluación de impacto del Join-In-Circuit en escuelas de Zambia"),
    ("DEP-19|DEP:24818", "T12", EXC, "", "Población distinta (candidatos a consejeros escolares)"),
    ("DEP-19|DEP:16072", "T11", EXC, "", "Tema distinto (razón de sexos al nacer)"),
    ("DEP-19|DEP:24878", "T20", IDE, "", "Ensayo de transferencia con sesiones para mujeres jóvenes en Sudáfrica"),
    ("DEP-19|DEP:21998", "T12", IDE, "", "Ensayo por conglomerados del programa COMPAS en Colombia"),
    ("DEP-19|DEP:6941", "T12", EXC, "", "Resultado distinto (conocimientos sobre ITS en varones jóvenes)"),
    # DEP-20: cuidado, lactarios y corresponsabilidad (T14, T19, T24)
    ("DEP-20|DEP:4558", "T14", IDE, "", "Evaluación del Programa de Estancias Infantiles de México"),
    ("DEP-20|DEP:23462", "T14", IDE, "", "Experimento en Uganda: subsidio de cuidado infantil frente a transferencia"),
    ("DEP-20|DEP:34865", "T14", IDE, "", "Diferencias en diferencias del cuidado público en China"),
    ("DEP-20|DEP:12955", "T14", IDE, "", "Ensayo por conglomerados en la India sobre oportunidades económicas de las mujeres (versión de 2020; E126 es el artículo sobre salud mental del mismo ensayo)"),
    ("DEP-20|DEP:7242", "T14", IDE, "", "Cuasi experimental del ICDS en la India rural"),
    ("DEP-20|DEP:36544", "T14", EXC, "", "Intervención distinta (reforma tributaria)"),
    ("DEP-20|DEP:4819", "T14", EXC, "", "Tema distinto (nutrición)"),
    ("DEP-20|DEP:14121", "T24", EXC, "", "Intervención distinta (reducción de la semana laboral)"),
    ("DEP-20|DEP:35009", "T24", EXC, "", "Intervención distinta (reducción de la semana laboral)"),
    ("DEP-20|DEP:4783", "T14", IDE, "", "Programa de Hogares Comunitarios de Guatemala (sin resumen en el portal)"),
    ("DEP-20|DEP:24086", "T14", EXC, "", "Tema distinto (agricultura sensible a la nutrición)"),
    ("DEP-20|DEP:7510", "T14", EXC, "", "Intervención distinta (liberalización comercial)"),
    ("DEP-20|DEP:4524", "T14", IDE, "", "Evaluación de Hogares Comunitarios en Colombia con efectos en la oferta laboral femenina"),
    ("DEP-20|DEP:6455", "T14", EXC, "", "Resultado distinto (interacción entre padres e hijos)"),
    ("DEP-20|DEP:6921", "T17", IDE, "", "Ensayo de Jóvenes en Acción en Colombia: las mujeres ofrecidas capacitación ganan 19,6 % más"),
    ("DEP-20|DEP:35326", "T14", EXC, "", "Intervención distinta (política de dos hijos)"),
    ("DEP-20|DEP:12508", "T14", EXC, "", "Tema distinto (nutrición)"),
    ("DEP-20|DEP:7120", "T14", EXC, "", "Tema distinto (nutrición)"),
    ("DEP-20|DEP:7648", "T18", YD("E33", "uno de los seis ensayos de microcrédito que resume"), "E33", ""),
    ("DEP-20|DEP:4551", "T14", IDE, "", "Expansión de la educación preescolar en Argentina y empleo materno"),
    ("DEP-20|DEP:9749", "T14", EXC, "", "Resultado distinto (salud y desarrollo infantil)"),
    ("DEP-20|DEP:19158", "T14", EXC, "", "Tema distinto (alimentación infantil)"),
    ("DEP-20|DEP:33068", "T14", EXC, "", "Resultado distinto (oferta laboral de personas mayores)"),
    ("DEP-20|DEP:26135", "T14", EXC, "", "Tesis sobre determinantes generales de disparidades, sin intervención de cuidado"),
    ("DEP-20|DEP:7955", "T15", IDE, "", "Reforma del sistema de herencia matrilineal en Ghana"),
    ("DEP-20|DEP:4520", "T14", IDE, "", "Ensayo de preescolar en Mozambique con efectos en los cuidadores"),
    ("DEP-20|DEP:6561", "T14", EXC, "", "Intervención distinta (pensión de vejez)"),
    # DEP-21: titulación, recursos naturales, formación, emprendimiento y transferencias (T15 a T18, T20)
    ("DEP-21|DEP:15605", "T18", IDE, "", "Ensayo en Sri Lanka: capacitación en metas con participación de los esposos"),
    ("DEP-21|DEP:9190", "T18; T17", IDE, "", "Metarregresión de 37 evaluaciones de programas de emprendimiento: el financiamiento parece más efectivo para las mujeres"),
    ("DEP-21|DEP:18719", "T18", IDE, "", "Programa multicomponente para mujeres en Mongolia Interior"),
    ("DEP-21|DEP:7673", "T17", IDE, "", "Ensayo de capacitación laboral en República Dominicana: mejora para las mujeres"),
    ("DEP-21|DEP:8923", "T17", EXC, "", "Población no específica (aprendices, sin enfoque de género)"),
    ("DEP-21|DEP:10443", "T15; T13; T17; T18", INC, "E159", ""),
    ("DEP-21|DEP:32814", "T17", IDE, "", "Institutos de formación vocacional de Punjab: aumento del empleo femenino"),
    ("DEP-21|DEP:21614", "T18", IDE, "", "Panel de 20 años de participantes de microcrédito en Bangladesh"),
    ("DEP-21|DEP:14892", "T20", IDE, "", "Progresa y dinámicas del hogar"),
    ("DEP-21|DEP:19314", "T20", EXC, "", "Resultado distinto (salud preventiva)"),
    ("DEP-21|DEP:8112", "T20", IDE, "", "Ensayo HPTN 068: transferencia condicionada a la asistencia escolar en Sudáfrica"),
    ("DEP-21|DEP:5956", "T18", YD("E107", "versión de 2008 del documento de trabajo"), "E107", ""),
    ("DEP-21|DEP:7529", "T17", EXC, "", "Población no específica (jóvenes, sin enfoque de género)"),
    ("DEP-21|DEP:9095", "T18", EXC, "", "Resultado distinto (prevención del VIH)"),
    ("DEP-21|DEP:20369", "T17", EXC, "", "Población no específica (migración laboral)"),
    ("DEP-21|DEP:20919", "T20; T17", IDE, "", "Experimento en Nicaragua: transferencia con capacitación o donación productiva"),
    ("DEP-21|DEP:6882", "T20; T11", IDE, "", "Oportunidades: embarazo y anticoncepción en mujeres jóvenes rurales"),
    ("DEP-21|DEP:9133", "T18", EXC, "", "Tema distinto (nutrición materna)"),
    ("DEP-21|DEP:6457", "T17", INC, "E153", ""),
    ("DEP-21|DEP:9883", "T17", YD("E97", "versión de 2012 del documento de trabajo"), "E97", ""),
    ("DEP-21|DEP:7700", "T17", IDE, "", "Seguimiento de largo plazo de Jóvenes en Acción en Colombia"),
    ("DEP-21|DEP:10745", "T18", IDE, "", "Experimento en Uganda: préstamos, donaciones y capacitación según el sexo"),
    ("DEP-21|DEP:32981", "T17", IDE, "", "Loterías de admisión a formación en Brasil: las mujeres obtienen mejores resultados"),
    ("DEP-21|DEP:9920", "T18", EXC, "", "Población no específica (medios de vida rurales)"),
    # DEP-22: participación política, transversalización, estadísticas y educación para la igualdad (T13, T21 a T23)
    ("DEP-22|DEP:10491", "T13", EXC, "", "Estudio en curso sin resultados en el registro"),
    ("DEP-22|DEP:10561", "T13", EXC, "", "Tesis cuyo experimento principal está publicado (DEP:17826, incluido como E155)"),
    ("DEP-22|DEP:39197", "T21", INC, "E154", ""),
    ("DEP-22|DEP:17826", "T13", INC, "E155", ""),
    ("DEP-22|DEP:7117", "T07; T11", IDE, "", "Talleres comunitarios de equidad de género y planificación familiar en Guatemala"),
    ("DEP-22|DEP:10310", "T13", INC, "E156", ""),
    ("DEP-22|DEP:11636", "T13", EXC, "", "Resultado distinto (desarrollo infantil)"),
    ("DEP-22|DEP:10190", "T13", EXC, "", "Pregunta distinta (efectos del liderazgo femenino sobre el empleo público)"),
    ("DEP-22|DEP:10577", "T13", INC, "E157", ""),
    ("DEP-22|DEP:11447", "T16", IDE, "", "Extensión agraria para mujeres agricultoras en Etiopía"),
    ("DEP-22|DEP:10541", "T13; T23", IDE, "", "Revisión rigurosa de ODI sobre medidas antidiscriminación (confianza baja según el 3ie DEP); el texto completo no fue accesible desde el repositorio"),
    ("DEP-22|DEP:10529", "T13", EXC, "", "Población no específica (sin enfoque de género)"),
    ("DEP-22|DEP:23005", "T17", IDE, "", "Programa Plan-It Girls de empleabilidad en la India"),
    ("DEP-22|DEP:6000", "T13", EXC, "", "Sin enfoque de género"),
    ("DEP-22|DEP:16523", "T23", IDE, "", "Experimento de educación en actitudes de género en secundaria en El Salvador"),
    ("DEP-22|DEP:10270", "T13", IDE, "", "Reserva de cargos para mujeres en Maharashtra (instrumento distinto del servicio)"),
    ("DEP-22|DEP:37679", "T13", INC, "E158", ""),
    ("DEP-22|DEP:14778", "T21", EXC, "", "No evalúa una intervención de transversalización (análisis de transferencias fiscales)"),
    ("DEP-22|DEP:13122", "T18", IDE, "", "Grupos de autoayuda en la India"),
    ("DEP-22|DEP:13131", "T13", EXC, "", "Estudio descriptivo sin intervención"),
    ("DEP-22|DEP:13136", "T13", IDE, "", "Capítulo de tesis sobre la brecha de participación política en la India (información); sin resultados en el registro"),
    ("DEP-22|DEP:13162", "T13", IDE, "", "Capítulo de tesis sobre grupos de mujeres y participación política en la India; sin resultados en el registro"),
    ("DEP-22|DEP:10581", "T06; T18", IDE, "", "Ensayo de empoderamiento social y económico en Afganistán"),
    ("DEP-22|DEP:17477", "T13", IDE, "", "Reservas para mujeres en consejos de aldea del sur de la India (instrumento distinto del servicio)"),
]
# Búsqueda manual en Scopus (estrategias MAN-23 a MAN-30, ejecutadas desde el navegador el 2026-10-31 e importadas con
# rev_evi_importar_manual.py). Registros `M0xx|SCOPUS:eid`. La regla de extracción es la misma que en el DEP: se extraen las
# fuentes que pueden cambiar la clasificación del tipo (todas en los tipos con evidencia limitada, revisiones en los tipos
# con evidencia moderada, direcciones no favorables en los tipos sólidos favorables, siempre el Perú); el resto queda
# «Identificada, no extraída» con su resultado anotado en el motivo. Los duplicados de búsquedas anteriores no se listan.
VS = f"{FECHA}: título, autores, año y fuente comprobados en el registro de Scopus y en Crossref; resultado tomado del resumen."
AIR = "No informado en el resumen"
ACADEMICA4 = [
    ("E161", "T10", "Olver, M. E. y Stockdale, K. C. (2025). Sexual offense treatment programming and recidivism reduction: "
     "A meta-meta-analysis of program outcomes and sources of effect size heterogeneity. Current Sexual Health Reports, "
     "17(1), 13.", "https://doi.org/10.1007/s11930-025-00409-5", VS, "Revisión sistemática", AIR, 2025, "No informada",
     "Metametaanálisis de 21 metaanálisis: los programas de tratamiento para personas que cometieron delitos sexuales "
     "reducen la reincidencia sexual en cerca de un tercio frente a los grupos de comparación (OR 0,67; IC 95 % 0,61 a "
     "0,73), con efectos mayores en programas para adolescentes, comunitarios y de alto riesgo o mayor dosis; los diseños "
     "rigurosos muestran efectos menores.", "Favorable"),
    ("E162", "T10", "Schmucker, M. y Lösel, F. (2015). The effects of sexual offender treatment on recidivism: An "
     "international meta-analysis of sound quality evaluations. Journal of Experimental Criminology, 11(4), 597–630.",
     "https://doi.org/10.1007/s11292-015-9241-z", VS, "Revisión sistemática", AIR, 2015,
     "Moderada (solo comparaciones con grupos equivalentes y reincidencia oficial)",
     "29 comparaciones (4939 tratados y 5448 no tratados): el tratamiento psicosocial, sobre todo cognitivo conductual, "
     "reduce la reincidencia sexual (OR 1,41; p < 0,01), con un efecto menor que en el metaanálisis de 2005.", "Favorable"),
    ("E163", "T10", "Lösel, F. y Schmucker, M. (2005). The effectiveness of treatment for sexual offenders: A comprehensive "
     "meta-analysis. Journal of Experimental Criminology, 1(1), 117–146.", "https://doi.org/10.1007/s11292-004-6466-7", VS,
     "Revisión sistemática", AIR, 2005, "No informada",
     "69 estudios con 80 comparaciones (N = 22 181): los tratados muestran 6 puntos porcentuales (37 %) menos de "
     "reincidencia sexual que los controles; entre los programas psicológicos, los cognitivo conductuales tienen el efecto "
     "más robusto y los no conductuales no muestran efecto.", "Favorable"),
    ("E164", "T10", "Hanson, R. K., Bourgon, G., Helmus, L. y Hodgson, S. (2009). The principles of effective correctional "
     "treatment also apply to sexual offenders: A meta-analysis. Criminal Justice and Behavior, 36(9), 865–891.",
     "https://doi.org/10.1177/0093854809338545", VS, "Revisión sistemática", AIR, 2009,
     "No informada (estudios que cumplen criterios básicos de calidad)",
     "23 estudios: reincidencia sexual de 10,9 % en tratados frente a 19,2 % en comparación, y reincidencia general de "
     "31,8 % frente a 48,3 %; los programas que siguen los principios de riesgo, necesidad y responsividad logran las "
     "mayores reducciones.", "Favorable"),
    ("E165", "T10", "Gannon, T. A., Olver, M. E., Mallion, J. S. y James, M. (2019). Does specialized psychological treatment "
     "for offending reduce recidivism? A meta-analysis examining staff and program variables as predictors of treatment "
     "effectiveness. Clinical Psychology Review, 73, 101752.", "https://doi.org/10.1016/j.cpr.2019.101752", VS,
     "Revisión sistemática", AIR, 2019, "No informada",
     "70 estudios (55 604 personas): los programas especializados para delitos sexuales reducen la reincidencia específica "
     "en 32,6 % y los de violencia doméstica en 36,0 %; la efectividad mejora con participación de psicólogos "
     "colegiados y supervisión clínica del personal.", "Favorable"),
    ("E166", "T10", "Kettrey, H. H. y Lipsey, M. W. (2018). The effects of specialized treatment on the recidivism of "
     "juvenile sex offenders: A systematic review and meta-analysis. Journal of Experimental Criminology, 14(3), 361–387.",
     "https://doi.org/10.1007/s11292-018-9329-3", VS, "Revisión sistemática", AIR, 2018,
     "Baja (ocho estudios, siete cuasiexperimentales)",
     "En adolescentes, el tratamiento especializado no reduce de forma significativa la reincidencia sexual (OR 0,74; "
     "IC 95 % 0,40 a 1,36) y sí la reincidencia general (OR 0,58); la evidencia creíble es escasa.", "Nulo"),
    ("E167", "T10", "Grønnerød, C., Grønnerød, J. S. y Grøndahl, P. (2015). Psychological treatment of sexual offenders "
     "against children: A meta-analytic review of treatment outcome studies. Trauma, Violence, & Abuse, 16(3), 280–290.",
     "https://doi.org/10.1177/1524838014526043", VS, "Revisión sistemática", AIR, 2015,
     "Baja (solo 9 de 14 estudios con calidad aceptable según los criterios CODC)",
     "14 estudios (1421 tratados y 1509 controles): el tamaño del efecto del tratamiento psicológico en agresores sexuales "
     "de niñas y niños es r = 0,03 en los estudios de calidad aceptable; la investigación disponible no permite "
     "establecer un efecto.", "Nulo"),
    ("E169", "T01", "Tirado-Muñoz, J., Gilchrist, G., Farré, M., Hegarty, K. y Torrens, M. (2014). The efficacy of cognitive "
     "behavioural therapy and advocacy interventions for women who have experienced intimate partner violence: A "
     "systematic review and meta-analysis. Annals of Medicine, 46(8), 567–586.",
     "https://doi.org/10.3109/07853890.2014.941918", VS, "Revisión sistemática", "No informado en el resumen", 2014,
     "No informada (pocos estudios e intervenciones heterogéneas)",
     "12 ensayos aleatorios (2666 participantes): el acompañamiento (advocacy) y la terapia cognitivo conductual reducen "
     "la violencia física y psicológica de pareja, pero no la sexual ni la violencia en conjunto.", "Mixto"),
    ("E170", "T23", "Visser, M., Jansen van Rensburg, M., Garforth, L. y Tefera, N. (2021). A large-scale community "
     "intervention to change gender perceptions in rural Ethiopia. Journal of Community and Applied Social Psychology, "
     "31(5), 571–589.", "https://doi.org/10.1002/casp.2540", VS, "Cuasi experimental", "Ingreso bajo (Etiopía)", 2021,
     "No informada", "Intervención comunitaria multicomponente en 123 escuelas primarias, evaluada en 30 escuelas (15 de "
     "proyecto y 15 de control) durante tres años: mejoraron las actitudes de género de docentes y familias y el apoyo a "
     "la educación de las niñas, y aumentaron la autoestima y las aspiraciones educativas de las niñas.", "Favorable"),
    # Ampliación de MAN-25 a 40 registros (aprobada por el usuario el 2026-11-01).
    ("E173", "T04", "Roehling, M. V., Wu, D., Choi, M. G. y Dulebohn, J. H. (2022). The effects of sexual harassment training "
     "on proximal and transfer training outcomes: A meta-analytic investigation. Personnel Psychology, 75(1), 3–31.",
     "https://doi.org/10.1111/peps.12492", VS, "Revisión sistemática", AIR, 2022, "No informada",
     "Primer metaanálisis de la capacitación contra el hostigamiento sexual en el trabajo: efecto grande en el conocimiento "
     "declarativo (δ = 1,06), moderado en las actitudes (δ = 0,41) y en las habilidades (δ = 0,39) y pequeño en la "
     "transferencia al puesto (δ = 0,14); el alcance de la capacitación, el método instructivo y el sexo de los participantes "
     "moderan los efectos, y el medio, la duración y el entorno no.", "Favorable"),
]
EVIDENCIA += ACADEMICA4
SELECCION_MANUAL = [
    # MAN-23: atención, salud mental, justicia y policía (T01, T02, T03, T05)
    ("MAN-23|SCOPUS:2-s2.0-85149042701", "T03", IDE, "", "Metaanálisis de 12 evaluaciones de cámaras corporales policiales: sin efecto global en resultados judiciales, con efectos grandes en los estudios de violencia doméstica (cuasiexperimentales); la intervención no corresponde a un servicio de la PNIG"),
    ("MAN-23|SCOPUS:2-s2.0-85194153841", "T07", IDE, "", "Metaanálisis actualizado de 59 evaluaciones de programas para hombres que ejercen violencia de pareja: efectos pequeños sobre la reincidencia, menores en los experimentos; T07 ya tiene evidencia sólida mixta"),
    ("MAN-23|SCOPUS:2-s2.0-85101699042", "T01", IDE, "", "Revisión de 13 estudios de centros integrados (Family Justice Centers) en Estados Unidos: solo evaluaciones de proceso o sin grupo de comparación; alta satisfacción, sin evidencia de resultados en las sobrevivientes"),
    ("MAN-23|SCOPUS:2-s2.0-105010602845", "T02", IDE, "", "Metaanálisis de ocho estudios: las intervenciones psicosociales y combinadas reducen la depresión en sobrevivientes de violencia de pareja (DM −1,22); T02 ya tiene evidencia sólida"),
    ("MAN-23|SCOPUS:2-s2.0-85198070936", "T02", IDE, "", "Metaanálisis de 28 estudios: las intervenciones psicoterapéuticas reducen el estrés postraumático (DME −0,59) y la depresión (DME −0,48) en mujeres expuestas a violencia, sin efecto en la autoestima; T02 ya tiene evidencia sólida"),
    ("MAN-23|SCOPUS:2-s2.0-85125262254", "T02", IDE, "", "Metaanálisis de 25 ensayos aleatorios (4683 participantes): resultados prometedores en ansiedad y bienestar de sobrevivientes adultas; T02 ya tiene evidencia sólida"),
    ("MAN-23|SCOPUS:2-s2.0-105021868922", "T18", IDE, "", "Metaanálisis de 49 ensayos en 31 países de ingreso bajo y medio: las intervenciones (transferencias, capacitación empresarial, enfoques transformadores de género) reducen la violencia económica con un efecto pequeño (g −0,076); T18 ya tiene evidencia sólida mixta"),
    ("MAN-23|SCOPUS:2-s2.0-105031747582", "T02", IDE, "", "Metaanálisis de 13 estudios: las intervenciones psicosociales reducen los síntomas depresivos (−0,26) y de ansiedad (−0,23) en mujeres con experiencia de violencia de pareja; T02 ya tiene evidencia sólida"),
    ("MAN-23|SCOPUS:2-s2.0-105042026572", "T02", IDE, "", "Revisión de diez estudios y metaanálisis de cinco sobre EMDR en sobrevivientes de violencia de pareja: reducciones pre y post de estrés postraumático, ansiedad y depresión; sin grupo de comparación en el metaanálisis"),
    ("MAN-23|SCOPUS:2-s2.0-84919382570", "T01", INC, "E169", ""),
    # MAN-24: prevención comunitaria, hombres, comunicación y escuela (T06, T07, T08, T09)
    ("MAN-24|SCOPUS:2-s2.0-85074553850", "T06; T09", IDE, "", "Revisión de 18 estudios en países de ingreso bajo y medio con jóvenes afectados por el VIH: las intervenciones de salud sexual y empoderamiento reducen poco la exposición a violencia de género y no la perpetración; las escuelas seguras no muestran efecto; calidad de la evidencia baja; T06 y T09 ya tienen evidencia sólida mixta"),
    ("MAN-24|SCOPUS:2-s2.0-85152352076", "T06", IDE, "", "Metaanálisis de ocho evaluaciones en la India: las intervenciones comunitarias reducen la violencia física y psicológica de pareja y la aprobación de la violencia; T06 ya tiene evidencia sólida"),
    ("MAN-24|SCOPUS:2-s2.0-85074998267", "T07", IDE, "", "Nueve ensayos (1014 hombres que consumen sustancias): el metaanálisis no encuentra diferencias en consumo ni en perpetración de violencia frente a la atención habitual; T07 ya tiene evidencia sólida mixta"),
    ("MAN-24|SCOPUS:2-s2.0-85128989556", "T09", IDE, "", "Metaanálisis de 13 ensayos aleatorios: los programas de prevención reducen la victimización (g 0,23) y la perpetración (g 0,24) de violencia psicológica en el noviazgo adolescente; T09 ya tiene evidencia sólida"),
    ("MAN-24|SCOPUS:2-s2.0-84907264834", "T09", IDE, "", "Metaanálisis de ocho programas escolares de prevención de la violencia en el noviazgo: efecto global positivo pequeño (r 0,11), con un cuarto de los estudios en dirección negativa; T09 ya tiene evidencia sólida mixta"),
    # MAN-25: hostigamiento sexual laboral (T04)
    ("MAN-25|SCOPUS:2-s2.0-105051786430", "T09", EXC, "", "Resultado distinto (violencia escolar general, no violencia de género)"),
    # MAN-26: tratamiento de agresores sexuales (T10): tipo con evidencia limitada, se extraen las revisiones principales
    ("MAN-26|SCOPUS:2-s2.0-105002863193", "T10", IDE, "", "Revisión narrativa de metaanálisis para peritos forenses: la mayoría de los metaanálisis respalda un beneficio sustancial del tratamiento, incluso en personas de alto riesgo"),
    ("MAN-26|SCOPUS:2-s2.0-85150895994", "T10", IDE, "", "Actualización del metaanálisis de Schmucker y Lösel (37 muestras, 30 394 personas): efecto pequeño (OR 1,54), mayor en riesgo medio y alto y en tratamientos especializados"),
    ("MAN-26|SCOPUS:2-s2.0-85046700805", "T10", INC, "E166", ""),
    ("MAN-26|SCOPUS:2-s2.0-84916983043", "T10", IDE, "", "Metaanálisis de diez estudios con adolescentes varones (N = 644): efecto favorable (r 0,37), mayor con medidas de autoinforme que con reincidencia real"),
    ("MAN-26|SCOPUS:2-s2.0-84953638571", "T10", INC, "E162", ""),
    ("MAN-26|SCOPUS:2-s2.0-24944518153", "T10", INC, "E163", ""),
    ("MAN-26|SCOPUS:2-s2.0-84931331761", "T10", INC, "E167", ""),
    ("MAN-26|SCOPUS:2-s2.0-0029129345", "T10", IDE, "", "Metaanálisis de 12 estudios (N = 1313): reincidencia de 19 % en tratados frente a 27 % en no tratados (r 0,12); mayor efecto de los tratamientos cognitivo conductuales y hormonales"),
    ("MAN-26|SCOPUS:2-s2.0-33845486676", "T10", IDE, "", "Metaanálisis de nueve estudios con adolescentes (N = 2986): efecto significativo del tratamiento en la reincidencia sexual (0,43), con cautela por el manejo de abandonos y seguimientos desiguales"),
    ("MAN-26|SCOPUS:2-s2.0-85078440946", "T10", IDE, "", "Metaanálisis de 25 estudios de programas cognitivo conductuales con adultos: efectos significativos en reincidencia sexual y violenta, heterogéneos y mayores en los programas de los años noventa"),
    ("MAN-26|SCOPUS:2-s2.0-105007648086", "T10", INC, "E161", ""),
    ("MAN-26|SCOPUS:2-s2.0-85025656480", "T10", YD("E74", "el DOI que consigna Scopus, analesps.33.2.267961, no resuelve; el de la revista es analesps.33.3.267961"), "E74", ""),
    ("MAN-26|SCOPUS:2-s2.0-70249090442", "T10", INC, "E164", ""),
    ("MAN-26|SCOPUS:2-s2.0-85071492663", "T10", INC, "E165", ""),
    ("MAN-26|SCOPUS:2-s2.0-105037716827", "T10", IDE, "", "Revisión rápida de nueve cuasiexperimentos en prisiones de Canadá (3568 hombres): razón de riesgos de 0,40 para la reincidencia sexual"),
    # MAN-27: salud sexual y reproductiva y educación sexual integral (T11, T12)
    ("MAN-27|SCOPUS:2-s2.0-0041047669", "T11", IDE, "", "Metaanálisis de 32 estudios de prevención del embarazo adolescente: sin efecto en la actividad sexual, aumento del uso de anticonceptivos y reducción menor de los embarazos; T11 ya tiene evidencia sólida mixta"),
    ("MAN-27|SCOPUS:2-s2.0-85066080188", "T12", IDE, "", "Revisión de 11 evaluaciones de intervenciones sobre el entorno escolar: tres ensayos sugieren que retrasan el inicio sexual (OR 0,5), con resultados mixtos en lo demás; T12 ya tiene evidencia sólida mixta"),
    ("MAN-27|SCOPUS:2-s2.0-85204485080", "T12", IDE, "", "Metaanálisis de 47 estudios (2000 a 2020): los programas escolares de educación sexual mejoran actitudes, conocimientos, autoeficacia e intención de uso del condón y la abstinencia, con alta heterogeneidad; T12 ya tiene evidencia sólida"),
    ("MAN-27|SCOPUS:2-s2.0-85134412941", "T11", IDE, "", "Metaanálisis de 52 evaluaciones financiadas por el gobierno de Estados Unidos: efectos pequeños, de dirección mixta y no significativos en embarazo y conducta sexual; T11 ya tiene evidencia sólida mixta"),
    # MAN-28: cuidado, lactarios y corresponsabilidad (T14, T19, T24)
    ("MAN-28|SCOPUS:2-s2.0-105000987632", "T19", IDE, "", "Metaanálisis de 13 estudios (4564 participantes, 3582 madres empleadas): la telelactancia en tiempo real aumenta la lactancia exclusiva a seis meses (RR 1,31), sobre todo los servicios proactivos; intervención distinta de los lactarios, T19 ya tiene evidencia sólida favorable"),
    ("MAN-28|SCOPUS:2-s2.0-105008658746", "T24", EXC, "", "Sin intervención evaluada (revisión de determinantes de la división del trabajo doméstico en América del Norte)"),
    ("MAN-28|SCOPUS:2-s2.0-77955167986", "T24", IDE, "", "Metaanálisis de 16 estudios de programas de educación para padres residentes: efecto global de d 0,26, sin seguimiento de largo plazo; T24 ya tiene evidencia sólida favorable"),
    ("MAN-28|SCOPUS:2-s2.0-105038353378", "T14", EXC, "", "Resultado distinto (bienestar y autoeficacia de las educadoras de cuidado infantil, no acceso ni calidad del cuidado)"),
    # MAN-30: participación política, transversalización, estadísticas y educación para la igualdad (T13, T21, T22, T23)
    ("MAN-30|SCOPUS:2-s2.0-85116943465", "T13", IDE, "", "Síntesis narrativa de más de 400 estudios sobre estrategias para aumentar la participación cívica y política de las mujeres en países en desarrollo (cuotas, redes sociales, programas focalizados); sin estimación de efectos"),
    ("MAN-30|SCOPUS:2-s2.0-105024829906", "T23", IDE, "", "Revisión narrativa de 21 programas de sensibilización de género para adolescentes: la mayoría sin diseño estructurado ni evaluación de efectividad"),
    ("MAN-30|SCOPUS:2-s2.0-84976484372", "T23", EXC, "", "Sin intervención evaluada (factores que forman las actitudes de género en la adolescencia temprana)"),
    ("MAN-30|SCOPUS:2-s2.0-85106272380", "T23", INC, "E170", ""),
    # MAN-25, puestos 31 a 40 (profundidad ampliada a 40 el 2026-11-01); Dalsbø y colegas (2026), en el puesto 39, ya estaba
    # incluida como E137 desde PubMed y figura como duplicado.
    ("MAN-25|SCOPUS:2-s2.0-105048470014", "T06", IDE, "", "Revisión de alcance de 55 currículos de 42 programas de prevención de la violencia de pareja en 23 países de ingreso bajo y medio (2001 a 2024): describe contenidos, metodologías e implementación sin estimar efectos; T06 y T09 ya tienen evidencia sólida mixta"),
    ("MAN-25|SCOPUS:2-s2.0-85192591246", "T04", IDE, "", "Revisión sistemática de 10 estudios de prevalencia de acoso en obstetricia y ginecología de Estados Unidos y Canadá y de 12 estudios de intervenciones en especialidades quirúrgicas (2906 participantes), sin metaanálisis; el resumen no informa el efecto de las intervenciones"),
    ("MAN-25|SCOPUS:2-s2.0-85121741407", "T04", INC, "E173", ""),
    ("MAN-30|SCOPUS:2-s2.0-85148677910", "T24", IDE, "", "Revisión de 44 artículos sobre la implementación de intervenciones con padres en países de ingreso bajo y medio: describe diseños y canales, no efectos"),
]


# Búsqueda manual en SciELO (MAN-01 a MAN-08, ejecutada el 2026-10-31 desde el navegador; importada con rev_evi_importar_manual.py).
# Misma regla de cribado que en Scopus. Las versiones repetidas del mismo artículo y los duplicados de búsquedas anteriores no se listan.
VSS = f"{FECHA}: título, autores, año y fuente comprobados en el registro RIS exportado de SciELO; resultado tomado del resumen."
ACADEMICA5 = [
    ("E171", "T11", "Sánchez, S., Atencio, G., Duy, N., Grande, M., Flores, M., Chiappe, M., Nalvarte, R., Sánchez, J. y Holmes, "
     "K. K. (2003). Comparación de la educación por pares y por profesionales de la salud para mejorar el conocimiento, "
     "percepción y la conducta sexual de riesgo en adolescentes. Revista Peruana de Medicina Experimental y Salud Pública, "
     "20(4), 206–210.", "http://www.scielo.org.pe/scielo.php?script=sci_arttext&pid=S1726-46342003000400006",
     VSS + " Sin DOI registrado en Crossref.", "Experimental", "Perú", 2003,
     "No informada (ensayo simple ciego con 206 adolescentes; retorno a los tres meses de 84 % y 70 %)",
     "206 adolescentes en riesgo, usuarios de los consultorios de planificación familiar de dos hospitales de Lima, "
     "asignados al azar a educación por pares o por profesionales de salud: a los tres meses, la educación por pares logra "
     "una mayor mejora del conocimiento (p = 0,047; p = 0,07 al ajustar por el puntaje inicial) y mejor retorno al "
     "seguimiento; las mejoras en percepción del riesgo y conducta sexual no difieren de forma significativa.", "Mixto"),
    ("E172", "T12", "Lazo Martínez, B. A., Sulca Carbajo, K. Y. y Vega Mamani, S. Z. (2026). Efectividad del programa "
     "«Educando para decidir mejor» en la prevención de conductas sexuales riesgosas: estudio preexperimental en escolares "
     "peruanos. Revista InveCom, 6(3).", "https://doi.org/10.5281/zenodo.17211713",
     VSS + " DOI comprobado en DataCite (depósito en Zenodo con año 2025; se cita el año de la revista).", "Descriptivo",
     "Perú", 2026, "Baja (diseño preexperimental de un solo grupo, 35 estudiantes seleccionados por conveniencia)",
     "35 estudiantes de quinto de secundaria de un colegio de Andahuaylas, medidos antes y después del programa: cambios "
     "significativos (p < 0,05) en la percepción del riesgo de las relaciones sexuales tempranas y del consumo de alcohol "
     "y drogas, menor intención de inicio sexual temprano y mayor uso del condón entre quienes ya tenían actividad sexual.",
     "Favorable"),
]
EVIDENCIA += ACADEMICA5
SELECCION_SCIELO = [
    # MAN-01: atención, salud mental, justicia y policía (T01, T02, T03, T05)
    ("MAN-01|SCIELO:S1516-36872025000103003-psi", "T02", IDE, "", "Revisión sistemática de alcance de 17 estudios sobre psicoterapia con mujeres que viven violencia de pareja: la mayoría son comparaciones pre-post sin grupo de control y la revisión no estima efectos; T02 ya tiene evidencia sólida"),
    ("MAN-01|SCIELO:S2178-86502022000100501-rve", "T05", IDE, "", "Revisión sistemática de métodos mixtos de nueve artículos (2015 a 2020): el escaso conocimiento del personal de atención primaria es una barrera y la capacitación una potencialidad; no evalúa efectos de la capacitación"),
    ("MAN-01|SCIELO:S0211-57352026000100127-esp", "T02", IDE, "", "Estudio pre-post sin grupo de control con 39 mujeres con trastorno mental grave en España: mejora el autoconcepto físico y disminuyen los síntomas depresivos; ingreso alto y sin comparación"),
    ("MAN-01|SCIELO:S2216-09732018000101988-rve", "T01", EXC, "", "Estudio cualitativo con 11 mujeres atendidas en una comisaría de la mujer de Brasil sobre el impacto de la violencia en la autoestima; no evalúa una intervención"),
    ("MAN-01|SCIELO:S1132-05592016000100001-esp", "T03", IDE, "", "Validación predictiva del protocolo policial de valoración del riesgo (VPR, sistema VioGén) en 407 denunciantes en España: AUC 0,71 y sensibilidad 85 %; mide la capacidad predictiva del instrumento, no el efecto de la respuesta policial"),
    ("MAN-01|SCIELO:S0213-91112023000100180-spa", "T05", IDE, "", "Evaluación cualitativa de la formación de 32 docentes para aplicar una intervención de prevención de la violencia de género en Terrassa: la formación asegura conocimientos mínimos pero no logra cambios profundos de actitudes"),
    ("MAN-01|SCIELO:S0104-026X2025000200206-scl", "T22", IDE, "", "Análisis documental del sistema de información de la Policía Civil de Río de Janeiro sobre violencia de pareja: concluye que falta el enfoque interseccional para producir estadísticas precisas; sin medición de efectos"),
    ("MAN-01|SCIELO:S1657-92672018000300052-col", "T02", IDE, "", "Protocolo de terapia cognitivo conductual de 13 sesiones evaluado antes y después en 11 mujeres en Brasil, sin grupo de control: reduce depresión, ansiedad y estrés y aumenta la satisfacción con la vida; los síntomas de estrés postraumático no cambian"),
    ("MAN-01|SCIELO:S1132-05592021000100005-esp", "T03", IDE, "", "Validación del algoritmo policial VPR5.0-H para predecir homicidios de pareja en España (2159 registros): sensibilidad 84 % y especificidad 60 %; no mide efectos sobre la violencia"),
    ("MAN-01|SCIELO:S2739-00632026000302026-ven", "T03", IDE, "", "Análisis normativo, doctrinario y jurisprudencial del principio de intervención inmediata de la Ley N.° 30364, con dos entrevistas a operadores: concluye que la respuesta rápida previene nuevos episodios, pero sin medición de efectos"),
    ("MAN-01|SCIELO:S1132-05592010000200007-esp", "T07", IDE, "", "Descripción del Programa Galicia de reeducación de agresores de pareja (2005 a 2010) con revisión narrativa de la eficacia de las intervenciones; declara resultados satisfactorios sin datos ni grupo de comparación"),
    # MAN-02: prevención en comunidad, hombres, comunicación y escuela (T06, T07, T08, T09)
    ("MAN-02|SCIELO:S0185-33252010000200006-mex", "T09", IDE, "", "Programa breve de sensibilización sobre violencia en el noviazgo (conferencias interactivas de dos horas y media en escuelas públicas de México, 2250 participantes), evaluado con cuestionarios antes (311) y después (184) en muestras distintas y sin grupo de control: mejora el reconocimiento de la violencia, de sus causas y de las formas sanas de expresar afecto; T09 ya tiene evidencia sólida mixta"),
    ("MAN-02|SCIELO:S1132-05592010000200006-esp", "T07", IDE, "", "Descripción de un programa psicosocial con perspectiva de género para hombres condenados por violencia de género en España, con escalas de ideología de género y actitudes antes y después; sin grupo de control ni datos de reincidencia"),
    ("MAN-02|SCIELO:S0034-98872024000100128-chl", "T01", IDE, "", "Estudio descriptivo transversal del estado de implementación de las Unidades Clínico Forense Hospitalarias para víctimas de violencia sexual en Chile: 12 de 24 servicios de salud las tienen y ninguna cuenta con recursos humanos completos; no mide efectos"),
    # MAN-03: hostigamiento sexual laboral (T04)
    ("MAN-03|SCIELO:S0718-50062015000600009-chl", "T04", IDE, "", "Estudio de caso descriptivo transversal con 901 estudiantes de una escuela del Instituto Politécnico Nacional de México: desconocimiento de los procedimientos de denuncia, falta de protocolos y escasa capacitación de los directivos; no evalúa efectos de una intervención"),
    # MAN-04: tratamiento de agresores sexuales (T10)
    ("MAN-04|SCIELO:S1889-18612017000100041-esp", "T10", IDE, "", "Comparación intrasujetos de 153 agresores sexuales tratados en prisiones españolas frente a un grupo de control: mejora el cambio terapéutico global (escala PASSO) y la mayoría de las variables específicas; no mide reincidencia y es de ingreso alto"),
    # MAN-05: salud sexual y reproductiva y educación sexual integral (T11, T12)
    ("MAN-05|SCIELO:S0874-02832013000200015-prt", "T12", IDE, "", "Revisión sistemática de 13 estudios primarios (2008 a 2010) sobre educación sexual en adolescentes: aumenta el conocimiento y en algunos estudios mejora las actitudes ante conductas de riesgo; sin estimación de efectos; T12 ya tiene evidencia sólida"),
    ("MAN-05|SCIELO:S1409-42582022000200329-cri", "T12", IDE, "", "Revisión de 16 programas de educación sexual (12 de España) de 1992 a 2016: mayor efectividad declarada de los enfoques integrales; sin estimación de efectos"),
    ("MAN-05|SCIELO:S1020-49892000000200006-spa", "T12", IDE, "", "Cuasiexperimento con 945 escolares de séptimo grado en Jamaica (1995 a 1997): sin efecto en el inicio de la actividad sexual; efecto positivo de corto plazo en el uso de anticonceptivos en la primera relación (p = 0,08) y en conocimientos y actitudes; T12 ya tiene evidencia sólida mixta"),
    ("MAN-05|SCIELO:S2739-00632026000302009-ven", "T12", INC, "E172", ""),
    ("MAN-05|SCIELO:S0036-36342019000500009-mex", "T12", IDE, "", "Cuasiexperimento con diferencias en diferencias en dos comunidades rurales de México: reduce 22,1 % las respuestas erróneas sobre la fase fértil del ciclo menstrual frente al control (p < 0,001); resultado de conocimiento, T12 ya tiene evidencia sólida"),
    ("MAN-05|SCIELO:S1726-46342003000400006-per", "T11", INC, "E171", ""),
    ("MAN-05|SCIELO:S1684-18242014000500005-cub", "T12", IDE, "", "Intervención educativa pre-post sin grupo de control en una secundaria rural de Cuba: mejora conocimientos y actitudes sobre infecciones de transmisión sexual; sin datos numéricos en el resumen"),
    ("MAN-05|SCIELO:S1794-99982007000200002-psi", "T12", IDE, "", "Programa de educación afectivo sexual para padres y madres de adolescentes en un centro de secundaria de España, con evaluación procesual sin grupo de control: aumentan la satisfacción corporal, la autoestima y las actitudes igualitarias"),
    ("MAN-05|SCIELO:S2007-07052020000200121-mex", "T12", IDE, "", "Adaptación y pilotaje de aceptabilidad de la intervención «Usando Condón» con 56 adolescentes de Puebla: comparación test-retest de la autoeficacia sin grupo de control; estudio de factibilidad"),
    ("MAN-05|SCIELO:S2304-51322017000100008-per", "T11", IDE, "", "Evaluación rápida de la oferta de servicios de salud sexual y reproductiva en el Perú ante el zika (2016): presupuesto de planificación familiar desviado a otros usos, 83 % de establecimientos con desabastecimiento de anticonceptivos inyectables y horarios restringidos para adolescentes; diagnóstico de la oferta, no mide efectos"),
    ("MAN-05|SCIELO:S0122-06672012000200003-col", "T12", IDE, "", "Estudio con 13 adolescentes en instituciones de semilibertad y abandono de Caracas, sin grupo de control: informar sobre anticoncepción mejora el nivel de conocimiento"),
    # MAN-06: cuidado, lactarios y corresponsabilidad (T14, T19, T24)
    ("MAN-06|SCIELO:S2594-066X2025000200206-mex", "T24", IDE, "", "Análisis cualitativo de los alcances y limitaciones de la norma mexicana NMX-R-025-SCFI-2015 de igualdad y no discriminación como mecanismo para la corresponsabilidad del sector privado en los cuidados; sin medición de efectos"),
    ("MAN-06|SCIELO:S2250-81392023000200093-arg", "T14", IDE, "", "Revisión sistemática de estudios empíricos sobre el vínculo entre trabajo remunerado y cuidado (percepciones, servicios de cuidado infantil, trayectorias y representaciones sociales); no evalúa intervenciones"),
    ("MAN-06|SCIELO:S0004-06222004000400002-ven", "T19", IDE, "", "Cuatro encuestas nacionales en Chile (1993 a 2002, unos 10 000 niños menores de 18 meses): la lactancia exclusiva a los seis meses sube de 16 % a 43,1 % durante el programa nacional de fomento y el trabajo materno fuera del hogar se asocia inversamente con la lactancia; serie sin grupo de control y no evalúa lactarios"),
    ("MAN-06|SCIELO:S2393-66062022000201201-ury", "T19", IDE, "", "Intervención educativa virtual con 40 gestantes de un centro de salud de Lima (2021): aumenta el conocimiento y 83,3 % del grupo experimental inicia la lactancia de forma temprana; no evalúa lactarios ni el retorno al trabajo"),
    # MAN-07: autonomía económica (T15 a T18, T20)
    ("MAN-07|SCIELO:S1415-98482014000300453-scl", "T18", IDE, "", "Panel administrativo de 20 628 clientes de microcrédito productivo de un banco comercial en Brasil, sin grupo de control: ser mujer se asocia con mayor ingreso y las ventas de las mujeres crecen cerca de 10 % por cada renovación del crédito"),
    # MAN-08: participación política, transversalización, estadísticas y educación para la igualdad (T13, T21, T22, T23)
    ("MAN-08|SCIELO:S0185-19182023000100119-mex", "T13", IDE, "", "Revisión sistemática de 13 estudios cualitativos sobre trayectorias políticas de mujeres: identifica obstáculos (estereotipos, prácticas machistas en los partidos) y facilitadores; no evalúa intervenciones"),
    ("MAN-08|SCIELO:S2183-54622025000100001-prt", "T08", IDE, "", "Estudio cualitativo (entrevistas, grupo focal y observación) de la emisora Rádio Mulher en Bafatá, Guinea-Bisáu: la radio se percibe como plataforma de empoderamiento local; sin medición de efectos"),
    ("MAN-08|SCIELO:S2071-081X2022000200011-bol", "T21", IDE, "", "Revisión descriptiva de fuentes oficiales sobre el gasto etiquetado para la igualdad de género en México (2008 a 2020): el monto crece cada año pero no alcanzó el 1 % del presupuesto federal hasta 2018; no mide efectos"),
    ("MAN-08|SCIELO:S2219-71682024000400346-per", "T13", IDE, "", "Estudio correlacional con 383 mujeres de la provincia de San Martín: la violencia de género (β = −0,505) y el empoderamiento (β = 0,387) predicen la participación política; no evalúa intervenciones"),
    ("MAN-08|SCIELO:S0185-19182013000100009-mex", "T13", IDE, "", "Análisis de elecciones locales de 12 estados de México (1998 a 2012): la descentralización de la selección de candidaturas no afecta las nominaciones de mayoría relativa y la centralización favorece la nominación de mujeres por representación proporcional; estudio observacional sin intervención de política pública"),
    ("MAN-08|SCIELO:S0102-69922023000100125-scl", "T13", IDE, "", "Estudio de caso con fuentes secundarias sobre 20 años de leyes de paridad en la legislatura de Córdoba, Argentina: la paridad vertical en las listas se diluye en el sistema mixto con distritos uninominales y persiste la subrepresentación en cargos jerárquicos; sin contrafactual"),
    ("MAN-08|SCIELO:S0185-16162019000100073-mex", "T13", IDE, "", "Análisis comparado de la representación descriptiva y sustantiva de las mujeres en los congresos subnacionales de México (2018): el acceso crece con las reformas político electorales de 15 años, con sesgos persistentes en el liderazgo parlamentario; descriptivo"),
    ("MAN-08|SCIELO:S1665-80272016000100013-mex", "T13", IDE, "", "Examen del cumplimiento de la paridad en las elecciones locales de 2015 en Tabasco, Distrito Federal y Estado de México: la paridad se logró en el Congreso de Tabasco y no en las otras dos entidades; descriptivo"),
    ("MAN-08|SCIELO:S1853-19702020000100041-arg", "T13", IDE, "", "Estudio de caso de las leyes de cupo (1993) y paridad (2002) en la legislatura de Río Negro, Argentina: analiza el proceso legislativo y la composición de la cámara; sin contrafactual"),
    ("MAN-08|SCIELO:S0104-83332016000200702-scl", "T21", IDE, "", "Ensayo sobre las políticas de igualdad de género en la ciencia de la Unión Europea y Alemania: pese a la institucionalización, persiste la subrepresentación de las mujeres; sin evaluación de efectos"),
    ("MAN-08|SCIELO:S0120-386X2020000200008-col", "T20", IDE, "", "Revisión sistemática de 34 evaluaciones aleatorizadas de programas contra la pobreza en la obra de Esther Duflo (8,8 % sobre participación política de mujeres y 5,9 % sobre transferencias directas), con síntesis cualitativa sin estimaciones"),
]

# Búsqueda manual en LILACS (MAN-09 a MAN-14, ejecutada el 2026-10-31 desde el portal de la BVS con las cadenas corregidas que
# constan en estrategias_manuales.csv). Ninguna fuente nueva: las seleccionadas se identifican sin extraer o se excluyen.
SELECCION_LILACS = [
    # MAN-09: atención, salud mental, justicia y policía (T01, T02, T03, T05)
    ("MAN-09|LILACS:biblio-1620248", "T09", IDE, "", "Revisión sistemática de 53 estudios (2010 a 2021) de prevención de la violencia en el noviazgo, en su mayoría de Estados Unidos y en escuelas: evidencia a favor de las intervenciones grupales multicomponente frente a las breves psicoeducativas; mapeo metodológico sin estimación de efectos; T09 ya tiene evidencia sólida mixta"),
    ("MAN-09|LILACS:biblio-1141543", "T02", IDE, "", "Resumen de evidencia de la revista Evidencia (el registro no trae resumen) sobre terapias psicológicas que podrían reducir la depresión y la ansiedad en víctimas de violencia de pareja; T02 ya tiene evidencia sólida"),
    ("MAN-09|LILACS:biblio-1013338", "T07", IDE, "", "Observación evaluativa del Núcleo de Atención al Hombre Autor de Violencia Doméstica de la Defensoría Pública de Pará, Brasil, con testimonios de atendidos y cónyuges: los hombres declaran cambios de conducta y las parejas expresan insatisfacción con los resultados; sin medición de efectos"),
    ("MAN-09|LILACS:biblio-846324", "T02", IDE, "", "Cuasiexperimento con grupo de control (107 mujeres maltratadas por su pareja, España) y seguimiento a tres y seis meses: disminuyen el estrés postraumático, la depresión y la ansiedad y aumentan la autoestima y el apoyo social; ingreso alto y T02 ya tiene evidencia sólida"),
    # MAN-10: prevención en comunidad, hombres, comunicación y escuela (T06, T07, T08, T09)
    ("MAN-10|LILACS:biblio-1435663", "T09", IDE, "", "Revisión sistemática de ocho ensayos escolares (la mayoría de Estados Unidos): la mayoría de las intervenciones reduce la victimización y la perpetración de violencia de pareja entre adolescentes; sin estimación conjunta; T09 ya tiene evidencia sólida mixta"),
    # MAN-11: hostigamiento sexual laboral (T04)
    ("MAN-11|LILACS:lil-714911", "T04", IDE, "", "Análisis de los relatos de un encuentro de mujeres policías de Venezuela (2009) sobre acoso sexual, afrontamiento y respuesta institucional; descriptivo, sin evaluación de las medidas adoptadas"),
    ("MAN-11|LILACS:lil-651165", "T04", EXC, "", "Estudio transversal con 460 trabajadores del sector servicios de Madrid sobre la asociación entre agresión laboral, satisfacción y conductas contraproducentes; no evalúa una intervención"),
    # MAN-12: tratamiento de agresores sexuales (T10)
    ("MAN-12|LILACS:biblio-1359362", "T10", IDE, "", "Revisión narrativa sobre la terapia hormonal (castración química) de la pedofilia: no está probada la cura ni la prevención primaria o secundaria; sin estimación de efectos"),
    ("MAN-12|LILACS:biblio-1418990", "T04", IDE, "", "Análisis de contenido de los 29 protocolos de acoso sexual de los servicios de salud de Chile: calidad heterogénea, con deficiencias en el enfoque preventivo, la claridad del proceso y las medidas reparatorias; no mide efectos"),
    ("MAN-12|LILACS:biblio-1249050", "T10", EXC, "", "Comparación de perfiles de 30 agresores sexuales psicópatas y no psicópatas en prisión en Brasil; no evalúa tratamiento"),
    ("MAN-12|LILACS:biblio-1057237", "T10", IDE, "", "Revisión sistemática de 56 artículos sobre instrumentos de distorsión cognitiva en agresores sexuales de niñas, niños y adolescentes: la distorsión cognitiva se usa como medida de resultado del tratamiento; no estima efectos"),
    ("MAN-12|LILACS:biblio-990975", "T10", IDE, "", "Estudio cualitativo con egresados chilenos de un programa para adolescentes que cometieron agresiones sexuales: importancia de la especialización de las técnicas y del terapeuta; sin medición de efectos"),
    ("MAN-12|LILACS:biblio-830418", "T10", IDE, "", "Revisión descriptiva de la validez en América Latina de los instrumentos de riesgo de reincidencia sexual en adolescentes (J-SOAP-II, ERASOR 2.0, JSORRAT-II): escasa investigación psicométrica regional; no evalúa tratamiento"),
    ("MAN-12|LILACS:lil-699186", "T05", IDE, "", "Estudio de casos múltiples con tres psicólogas capacitadas en terapia cognitivo conductual grupal para víctimas de violencia sexual en Brasil: impacto percibido en el desempeño y transferencia del aprendizaje; sin grupo de comparación"),
    # MAN-13: salud sexual y reproductiva y educación sexual integral (T11, T12)
    ("MAN-13|LILACS:biblio-1717464", "T11", IDE, "", "Revisión bibliográfica de publicaciones brasileñas sobre educación sexual en la consulta ginecológica de adolescentes: declara mayor adhesión a los métodos de barrera y a las pruebas de ITS; sin estimación de efectos"),
    ("MAN-13|LILACS:biblio-1612033", "T11", IDE, "", "Análisis documental de políticas y estrategias de salud sexual y reproductiva en Medellín, Antioquia y Colombia (2012 a 2022) con indicadores del DANE: concluye un impacto positivo sobre el embarazo adolescente, sin un diseño que permita atribuirlo"),
    ("MAN-13|LILACS:biblio-1618657", "T09", IDE, "", "Cuasiexperimento con 57 estudiantes de sexto grado en Colombia: talleres sobre derechos sexuales, autoconocimiento y límites mejoran las conductas de autoprotección frente al abuso sexual (Test de prudencia) respecto del grupo de comparación (p < 0,001); T09 ya tiene evidencia sólida mixta"),
    # MAN-14: cuidado, lactarios y corresponsabilidad (T14, T19, T24)
    ("MAN-14|LILACS:biblio-1614406", "T19", IDE, "", "Evaluación del grado de implantación de 13 salas de apoyo a la lactancia certificadas por el Ministerio de Salud en Paraná, Brasil: puntaje medio 8,53 de 10 (consolidado), con debilidades en la capacitación de tutores y el monitoreo; evalúa la implementación, no los efectos"),
    ("MAN-14|LILACS:biblio-978096", "T19", IDE, "", "Estudio cualitativo con 15 funcionarias de un hospital de Curanilahue, Chile: la extensión del posnatal (Ley 20.545) y el acceso a una sala cuna cercana permitieron mantener la lactancia al volver al trabajo; sin medición de efectos"),
    ("MAN-14|LILACS:biblio-959550", "T19", IDE, "", "Cohorte retrospectiva de 938 lactantes de un centro de salud de Santiago (2009 a 2013), antes y después de la extensión del posnatal de 12 a 21 semanas (Ley 20.545): la lactancia exclusiva a los seis meses sube de 33 % a 42 % (p = 0,004) y a los 12 meses de 10 % a 16 % (HR 0,85 para el abandono); T19 ya tiene evidencia sólida favorable"),
    ("MAN-14|LILACS:lil-558787", "T19", IDE, "", "Estudio de intervención no aleatorizado con 200 madres trabajadoras en Brasil: las participantes del programa de incentivo a la lactancia mantienen más la lactancia exclusiva (p < 0,0001) y no poder amamantar durante la jornada multiplica por 4,98 la probabilidad de destete antes del cuarto mes; T19 ya tiene evidencia sólida favorable"),
]

# Búsqueda manual en JSTOR (MAN-43 a MAN-46, ejecutada el 2026-11-01 desde el navegador con el acceso institucional de la UNMSM;
# sustituye a EconLit para los grupos económicos). Ninguna fuente nueva cambia la clasificación de su tipo: los estudios
# primarios de ingreso alto no entran a la matriz y los demás recaen en tipos ya sólidos o moderados por otras fuentes.
SELECCION_JSTOR = [
    # MAN-43: atención, salud mental, justicia y policía (T01, T02, T03, T05)
    ("MAN-43|JSTOR:26802791", "T03", IDE, "", "Efectos fijos y planes de acción afirmativa impuestos judicialmente en la policía de Estados Unidos (fines de los setenta a inicios de los noventa): más mujeres policías aumentan la denuncia de violencia doméstica y reducen el homicidio de pareja; ingreso alto, no entra a la matriz"),
    # MAN-44: cuidado, lactarios y corresponsabilidad (T14, T19, T24)
    ("MAN-44|JSTOR:10.15609/annaeconstat2009.129.0103", "T14", IDE, "", "Diferencias en diferencias por municipio en Valonia, Bélgica: la expansión de plazas de cuidado infantil de 0 a 3 años aumenta el empleo materno, menos de lo esperado por desplazamiento del cuidado informal; ingreso alto"),
    ("MAN-44|JSTOR:26727057", "T14", IDE, "", "Modelo estructural con contrafactuales en Estados Unidos: ampliar Head Start tiene efectos grandes en habilidades cognitivas y los subsidios de cuidado infantil efectos nulos; ingreso alto"),
    ("MAN-44|JSTOR:48695747", "T18", IDE, "", "Cuasiexperimento en Croacia: pequeñas subvenciones de contrapartida para mujeres emprendedoras (usadas en cuidado infantil y asesoría) aumentan el crédito bancario, la facturación y el empleo de las más experimentadas; ingreso alto"),
    ("MAN-44|JSTOR:48601827", "T18", YD("E34"), "E34", ""),
    ("MAN-44|JSTOR:26605583", "T14", IDE, "", "Modelo estructural con datos de Australia: los créditos tributarios por cuidado infantil aumentan más las horas trabajadas que los subsidios de precio, pero benefician sobre todo a mujeres con más ingresos y educación; ingreso alto"),
    ("MAN-44|JSTOR:43738181", "T24", IDE, "", "Experimento natural de la reforma de licencia parental de Alemania (2007): baja el empleo materno en los 12 meses tras el parto y sube al vencer la transferencia; los dos meses reservados al padre no cambian de forma significativa su tiempo de cuidado; ingreso alto, no entra a la matriz"),
    ("MAN-44|JSTOR:27254426", "T19", IDE, "", "Normas hospitalarias de atención posparto en Estados Unidos: aumentan el inicio de la lactancia en 3,8 puntos y la lactancia a los 3 y 6 meses en 7 %, y desplazan tiempo de trabajo formal hacia el cuidado en los tres primeros meses; no evalúa lactarios y es de ingreso alto"),
    ("MAN-44|JSTOR:28036354", "T14", IDE, "", "Revisión narrativa sobre la ampliación de los servicios de educación y cuidado infantil en Canadá: la calidad y la atención a familias de bajos ingresos condicionan sus efectos; sin estimación propia"),
    ("MAN-44|JSTOR:27031111", "T14", IDE, "", "Encuesta en barrios de bajos ingresos de Daca, Bangladés: la necesidad de cuidado infantil y la seguridad en el espacio público y el trabajo se asocian con menor participación laboral femenina; correlacional, no evalúa una intervención"),
    ("MAN-44|JSTOR:26818747", "T14", IDE, "", "Revisión narrativa de las políticas de cuidado infantil de Estados Unidos y sus efectos en el desarrollo infantil y el empleo materno, con tres propuestas de reforma; T14 ya tiene evidencia sólida"),
    # MAN-45: autonomía económica (T15 a T18, T20)
    ("MAN-45|JSTOR:43189512", "T18", IDE, "", "Ensayo aleatorio de microcrédito grupal en 52 barrios de Hyderabad, India: aumentan la inversión y las ganancias de negocios existentes, sin cambios en consumo, salud, educación ni empoderamiento de las mujeres; es uno de los seis ensayos que resume E33"),
    ("MAN-45|JSTOR:27290718", "T20", IDE, "", "Familias en Acción, Colombia: la transferencia condicionada mejora la salud de los miembros del hogar no focalizados, más en mujeres y adultos mayores en el corto plazo; T20 ya tiene evidencia sólida"),
    ("MAN-45|JSTOR:26598334", "T18", IDE, "", "Emparejamiento y diferencias en diferencias con beneficiarias de grupos de autoayuda en Karnataka, India (2010 a 2015): aumenta el ingreso de los hogares; T18 ya tiene evidencia sólida"),
    ("MAN-45|JSTOR:26832113", "T18", IDE, "", "Encuesta a 2598 hogares de Bangladés con cuatro grupos de control para la autoselección: los prestatarios de ASA tienen menor incidencia y brecha de pobreza; T18 ya tiene evidencia sólida"),
    ("MAN-45|JSTOR:27004172", "T18", IDE, "", "Análisis de economía política feminista del microcrédito en Timor Oriental: el endeudamiento se suma al precio de la novia y agrava la violencia de género; cualitativo, sin medición de efectos"),
    ("MAN-45|JSTOR:43189514", "T18", IDE, "", "Ensayo aleatorio en Mongolia rural: el microcrédito grupal para mujeres aumenta el emprendimiento femenino y el consumo de alimentos, sin efecto en horas de trabajo ni ingreso; el crédito individual no tiene efectos; T18 ya tiene evidencia sólida"),
    ("MAN-45|JSTOR:27346168", "T18", IDE, "", "Estimadores corregidos por sesgo y emparejamiento con datos de Bangladés: el acceso al crédito tiene un efecto limitado en el poder de decisión de las mujeres y la educación uno mucho mayor; coincide con los ensayos aleatorios"),
    ("MAN-45|JSTOR:26643866", "T18", IDE, "", "Modelo de ecuaciones estructurales con 353 integrantes de grupos de autoayuda en Odisha, India: mayor participación se asocia con mayor empoderamiento económico, social y político; correlacional"),
    ("MAN-45|JSTOR:10.1086/674968", "T20", IDE, "", "Modelo colectivo del hogar con la asignación aleatoria de Progresa en México rural: transferir a las mujeres cambia la estructura del consumo de alimentos; estudio de mecanismos, T20 ya tiene evidencia sólida"),
    ("MAN-45|JSTOR:26538639", "T18", IDE, "", "Emparejamiento por puntaje de propensión con hogares del programa de microfinanzas público BRDB en Bangladés: mejores ingresos, gasto y empoderamiento de las mujeres, sin efecto en capital humano; pobreza de 24 % frente a 35 % en el grupo de control"),
    ("MAN-45|JSTOR:26699166", "T18", IDE, "", "Encuesta a 2615 hogares de la Misión Nacional de Medios de Vida Rurales en Madhya Pradesh, India: beneficia a los pobres, pero los más pobres no logran pagar y quedan rezagados; sin grupo de control"),
    ("MAN-45|JSTOR:26372560", "T17", IDE, "", "Ensayo aleatorio del Youth Opportunities Program en el norte de Uganda: las subvenciones para capacitación vocacional y negocio aumentan los activos (57 %), las horas (17 %) y los ingresos (38 %) a cuatro años, con efectos relativos mayores en mujeres; T17 ya tiene evidencia sólida"),
    ("MAN-45|JSTOR:27031081", "T20", IDE, "", "Puntaje de propensión con 2500 hogares del Programa de Becas de Educación Primaria de Bangladés: la transferencia por medio de las madres mejora los resultados educativos y las empodera como madres, no como mujeres; T20 ya tiene evidencia sólida"),
    ("MAN-45|JSTOR:44004619", "T20", IDE, "", "Análisis de la implementación del derecho a prestaciones de maternidad mediante la transferencia condicionada IGMSY en India: la condicionalidad excluye a las mujeres más vulnerables; sin medición de efectos"),
    ("MAN-45|JSTOR:41679557", "T20", IDE, "", "Experimento en Malawi rural con incentivos por mantener el estado serológico: sin efecto en VIH ni en conducta sexual; tras recibir el dinero, los hombres aumentan y las mujeres reducen el sexo de riesgo; T20 ya tiene evidencia sólida mixta"),
    ("MAN-45|JSTOR:24243705", "T15", IDE, "", "Proyecto de reasentamiento y reforma de tierras en el sur de Malawi: mejora la seguridad alimentaria de largo plazo; la formalización mejora la seguridad de la tierra de jefes y jefas de hogar, pero el reasentamiento la debilita para las mujeres de hogares con jefe varón; T15 ya tiene evidencia moderada mixta y un estudio primario más no la cambia"),
    ("MAN-45|JSTOR:23723977", "T18", IDE, "", "Diseño transversal con muestreo estratificado de clientas del microcrédito AIM en Malasia urbana: la participación se asocia con mayor empoderamiento; sin contrafactual creíble"),
    ("MAN-45|JSTOR:23528575", "T20", IDE, "", "Evaluación de funcionamiento de Janani Suraksha Yojana en ocho distritos de siete estados rezagados de India: el incentivo se paga en la mayoría de los partos institucionales, pero una fracción importante de mujeres, más desfavorecidas, sigue dando a luz en casa; descriptivo"),
    # MAN-46: participación política, transversalización, estadísticas y educación para la igualdad (T13, T21, T22, T23)
    ("MAN-46|JSTOR:27108637", "T21", IDE, "", "Análisis de los programas de financiamiento de la Unión Europea 2014 a 2020: el presupuesto con enfoque de género no se aplica de forma sistemática y la igualdad queda como objetivo horizontal sin metas ni fondos asignados; descriptivo"),
    ("MAN-46|JSTOR:48601830", "T13", IDE, "", "Revisión narrativa de estudios microempíricos sobre reformas legales contra la desigualdad de género (herencia, propiedad, cuotas): algunas reformas reducen las brechas y las intervenciones temporales pueden tener efectos duraderos; T13 ya tiene evidencia moderada mixta"),
    ("MAN-46|JSTOR:27004171", "T21", IDE, "", "Análisis de economía política del presupuesto con enfoque de género en los Balcanes Occidentales: las iniciativas resultan «gestos vacíos» frente a las reformas fiscales de las instituciones financieras internacionales; descriptivo y de dirección desfavorable, sin estimación de efectos"),
    ("MAN-46|JSTOR:45023161", "T13", IDE, "", "Regresión discontinua con datos de distritos electorales de India: la elección de una mujer como legisladora aumenta su recandidatura pero reduce la entrada de nuevas candidatas, sobre todo en estados con más sesgo de género; T13 ya tiene evidencia moderada mixta"),
    ("MAN-46|JSTOR:10.7758/rsf.2019.5.5.09", "T24", EXC, "", "Describe el acceso a licencias y cuidado infantil provistos por empleadores en Estados Unidos según sexo; no evalúa una intervención"),
    ("MAN-46|JSTOR:26746471", "T13", EXC, "", "Modelo teórico de candidatura ciudadana sobre cuotas de género y calidad de los políticos; no hay evaluación empírica"),
    ("MAN-46|JSTOR:26585555", "T21", IDE, "", "Análisis temático de 32 entrevistas y tres informes sobre el Análisis de Género Plus (ACS+) en Canadá: problemas para operacionalizar la interseccionalidad; cualitativo, sin medición de efectos"),
    ("MAN-46|JSTOR:26291683", "T21", IDE, "", "Estudio cualitativo de cómo las analistas de género aplican el ACS+ en la burocracia federal de Canadá; sin medición de efectos"),
    ("MAN-46|JSTOR:4417355", "T21", EXC, "", "Ensayo sobre transversalización del enfoque de género y reducción de la pobreza en Asia del Sur; no evalúa una intervención"),
    ("MAN-46|JSTOR:23602096", "T22", IDE, "", "Encuesta nacional de uso del tiempo de Australia (1992): registrar las actividades simultáneas mejora la medición del trabajo no remunerado, en especial el de las mujeres; metodológico, no mide efectos de producir estadísticas"),
    ("MAN-46|JSTOR:4404710", "T22", IDE, "", "Inventario de las estadísticas de India en educación, trabajo, salud y participación política para indicadores de desarrollo sensibles al género (1996); descriptivo"),
]

ids_academica = {e[0] for e in ACADEMICA + ACADEMICA2 + ACADEMICA3 + ACADEMICA4 + ACADEMICA5}
assert {s[3] for s in SELECCION + SELECCION_DEP + SELECCION_MANUAL + SELECCION_SCIELO + SELECCION_LILACS + SELECCION_JSTOR if s[2] == INC} == ids_academica, "cada fuente académica debe venir de un registro"
ORIGEN = {e[0]: "Búsqueda web" for e in EVIDENCIA}
for reg, _, decision, ev_id, _ in SELECCION:
    if decision == INC:
        ORIGEN[ev_id] = f"Búsqueda académica ({reg.replace('|', ', ')})"
    elif decision == YA:
        ORIGEN[ev_id] = ORIGEN[ev_id] + (f"; recuperada en la búsqueda académica ({reg.replace('|', ', ')})" if ORIGEN[ev_id].startswith(
            "Búsqueda web") else f"; {reg.replace('|', ', ')}")
CRIBADO = [(*c, "Búsqueda web") for c in CRIBADO]
CRIBADO += CRIBADO2
ORIGEN.update(ORIGEN_PORTALES)
for reg, _, decision, ev_id, _ in SELECCION_DEP:
    if decision == INC:
        ORIGEN[ev_id] = f"Búsqueda en el 3ie DEP ({reg.replace('|', ', ')})"
    elif decision.startswith("Ya incluida"):
        ORIGEN[ev_id] += f"; recuperada en el 3ie DEP ({reg.replace('|', ', ')})"
BASE_MANUAL = {"SCOPUS": "Scopus", "SCIELO": "SciELO", "LILACS": "LILACS", "WOS": "Web of Science", "ECONLIT": "EconLit",
               "JSTOR": "JSTOR"}
for reg, _, decision, ev_id, _ in SELECCION_MANUAL + SELECCION_SCIELO + SELECCION_LILACS + SELECCION_JSTOR:
    base = BASE_MANUAL[reg.split('|')[1].split(':')[0]]
    if decision == INC:
        ORIGEN[ev_id] = f"Búsqueda en {base} ({reg.replace('|', ', ')})"
    elif decision.startswith("Ya incluida"):
        ORIGEN[ev_id] += f"; recuperada en {base} ({reg.replace('|', ', ')})"
SELECCION += SELECCION_DEP + SELECCION_MANUAL + SELECCION_SCIELO + SELECCION_LILACS + SELECCION_JSTOR

# Autores completos (APA 7) de las referencias que decían «et al.»: los obtiene src/fuentes/completar_autores_evidencia.py
# de Crossref, PubMed, OpenAlex o doi.org, con cada respuesta registrada por su hash (ADR-010). Se aplican aquí para no
# tocar los literales del registro verificado el 2026-10-31; la verificación anota la fuente de los autores.
AUTORES = DEST / "autores_completados.csv"
if AUTORES.exists():
    completos = {r["id"]: r for r in csv.DictReader(open(AUTORES, encoding="utf-8")) if r["aceptado"] == "Sí"}
    nuevas = []
    for e in EVIDENCIA:
        r = completos.get(e[0])
        if r:
            assert e[2].startswith(r["tramo_actual"]), f"{e[0]}: la referencia ya no empieza por {r['tramo_actual']}"
            e = (e[0], e[1], r["tramo_nuevo"] + e[2][len(r["tramo_actual"]):], e[3],
                 f"{e[4].rstrip('.')}. Autores completos tomados de {r['fuente']} el 2026-11-08.", *e[5:])
        nuevas.append(e)
    EVIDENCIA = nuevas

DEST.mkdir(parents=True, exist_ok=True)
for nombre, cols, filas in [
    ("tipologia.csv", ["tipo", "nombre", "servicios", "nodo_modelo", "poblacion", "intervencion", "resultado_esperado",
                       "pertinencia", "cadena_es", "cadena_en"], TIPOS),
    ("busquedas.csv", ["id", "fecha", "cadena", "buscador", "resultados_revisados", "tipos"],
     [(f"WEB-{i:02d}", FECHA, *b) for i, b in enumerate(BUSQUEDAS, 1)]),
    ("cribado.csv", ["referencia", "tipos", "enlace", "motivo", "origen"], CRIBADO),
    ("evidencia.csv", ["id", "tipos", "referencia", "enlace", "verificacion", "diseno", "ambito", "anio", "calidad",
                       "resultado_principal", "direccion", "origen"], [(*e, ORIGEN[e[0]]) for e in EVIDENCIA]),
    ("seleccion_academica.csv", ["registro", "tipos", "decision", "evidencia", "motivo"], SELECCION),
]:
    with open(DEST / nombre, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f, quoting=csv.QUOTE_ALL)
        w.writerow(cols)
        w.writerows(filas)
print(len(TIPOS), len(BUSQUEDAS), len(CRIBADO), len(EVIDENCIA))
