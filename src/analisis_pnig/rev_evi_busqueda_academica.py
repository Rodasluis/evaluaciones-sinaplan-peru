"""REV_EVI: búsqueda en bases académicas con vocabulario controlado (nota metodológica, numeral 5.1 y Anexo 6).

Ejecuta, por tipo de intervención, las estrategias de búsqueda definidas abajo en siete bases con interfaz pública:
  - PubMed (MEDLINE), con descriptores MeSH, mediante E-utilities;
  - ERIC (educación), con descriptores del tesauro de ERIC, mediante su API;
  - EconBiz (economía; incluye RePEc y EconStor), con descriptores del Tesauro de Economía STW, mediante su API;
  - Campbell Systematic Reviews (ISSN 1891-1803) y publicaciones de 3ie (prefijo DOI 10.23846), mediante la API de
    Crossref;
  - LA Referencia (repositorios de América Latina, incluido ALICIA del Perú), con términos libres en español, mediante
    su API (VuFind). Las tildes se omiten en las cadenas porque el buscador no las admite dentro de paréntesis;
  - Banco Mundial, Documents & Reports (serie Policy Research Working Papers), con términos libres en inglés, mediante
    su API. El cortafuegos del sitio bloquea las cadenas con comillas u operadores, por lo que se usan palabras sueltas,
    que la API combina como conjunción;
  - OpenAlex (indexa SciELO, Redalyc y repositorios latinoamericanos), con términos libres en español y filtro de idioma,
    mediante su API. Si existe la variable de entorno OPENALEX_API_KEY, la clave se envía en la consulta y nunca se
    escribe en los registros; sin clave, la API admite un cupo diario reducido.
SciELO, LILACS (descriptores DeCS), el Development Evidence Portal de 3ie, Scopus, Web of Science y EconLit no admiten
consultas automatizadas o requieren suscripción: sus estrategias están en rev_evi_insumos/busqueda_manual/ y se ejecutan a
mano.

Escribe en la carpeta de destino (por defecto evaluaciones/pnig_2026/04_analisis/rev_evi_insumos):
  - estrategias_academicas.csv: una fila por estrategia, con la cadena exacta, los descriptores, los términos libres, los
    filtros, la fecha, el total de resultados, los registros revisados y la URL de la consulta;
  - descriptores.csv: cada descriptor usado, con su identificador en el tesauro y la verificación hecha;
  - resultados_academicos.csv: los registros revisados de cada estrategia (los primeros por relevancia), con su clave
    (PMID, ERIC, EconBiz, DOI o LA Referencia) y la marca de duplicado.
Las respuestas crudas se guardan en evaluaciones/pnig_2026/03_datos/raw/rev_evi_busqueda/<fecha>/ (no se versionan) y se
registran con su hash en docs/fuentes/descargas_rev06-busqueda_<fecha>.tsv. Repetir la búsqueda en otra fecha puede dar
totales distintos, porque las bases se actualizan; el registro versionado es el del 2026-10-31. Se revisan los primeros
30 registros por relevancia en PubMed y EconBiz, 15 en ERIC, LA Referencia, Banco Mundial y OpenAlex, y 10 en Campbell
y 3ie (en Campbell, 3ie y LA Referencia el aporte cae después del registro 5).

Uso (desde la raíz del repositorio):
    python src/analisis_pnig/rev_evi_busqueda_academica.py [carpeta_destino]
"""
import csv
import hashlib
import json
import os
import re
import sys
import time
import urllib.parse
import urllib.request
from datetime import date
from pathlib import Path
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

DEST = Path(sys.argv[1] if len(sys.argv) > 1 else f"{EVAL}/04_analisis/rev_evi_insumos")
FECHA = date.today().isoformat()
CRUDO = Path(f"{EVAL}/03_datos/raw/rev_evi_busqueda/{FECHA}")
REGISTRO = Path(f"docs/fuentes/descargas_rev06-busqueda_{FECHA}.tsv")
UA = {"User-Agent": "Mozilla/5.0 (compatible; evaluacion-pnig-rev_evi)"}
OA_CLAVE = os.environ.get("OPENALEX_API_KEY", "")

# Filtros de diseño y de contexto reutilizados en las cadenas
DIS_PM = ('(systematic[sb] OR "Randomized Controlled Trial"[pt] OR "Evaluation Study"[pt] OR "Program Evaluation"[mh] '
          'OR "Controlled Before-After Studies"[mh] OR "Interrupted Time Series Analysis"[mh])')
GEO_PM = ('("Developing Countries"[mh] OR "South America"[mh] OR "Central America"[mh] OR "Mexico"[mh] OR "Africa"[mh] '
          'OR "Asia"[mh] OR "low- and middle-income"[tiab] OR "low and middle income"[tiab] OR LMIC*[tiab] OR systematic[sb])')
DIS_EB = ('(subject:"Impact assessment" OR subject:"Evaluation" OR subject:"Randomized controlled trial" OR '
          'subject:"Experiment" OR "impact evaluation" OR randomized OR randomised OR "systematic review" OR "meta-analysis")')
GEO_EB = '(subject:"Developing countries" OR subject:"Latin America" OR subject:"Peru" OR "middle-income")'
DIS_ER = ('(subject:"Program Effectiveness" OR subject:"Program Evaluation" OR subject:"Randomized Controlled Trials" OR '
          'subject:"Meta Analysis")')
GEO_ER = '(subject:"Developing Nations" OR subject:"Meta Analysis")'
VIOL_PM = ('("Intimate Partner Violence"[mh] OR "Domestic Violence"[mh] OR "Battered Women"[mh] OR '
           '"Gender-Based Violence"[mh])')
DISENO_LR = "(impacto OR evaluacion OR efectos OR experimental)"
DIS_OA = '(evaluación OR impacto OR efecto OR efectos OR efectividad OR eficacia OR "revisión sistemática")'

# Estrategias: tipos, base, cadena (para Crossref y LA Referencia, términos libres)
ESTRATEGIAS = [
    # T01 Atención integral a víctimas
    ("T01", "PubMed", f'{VIOL_PM} AND ("Crisis Intervention"[mh] OR "Crime Victims"[mh] OR advocacy[tiab] OR '
     f'"one-stop"[tiab] OR "crisis centre*"[tiab] OR "crisis center*"[tiab] OR "justice center*"[tiab]) AND {DIS_PM} AND '
     f'{GEO_PM}'),
    ("T01", "Campbell", "intimate partner violence survivors advocacy support services"),
    ("T01", "3ie", "intimate partner violence survivors support services"),
    ("T01", "LA Referencia", f'("Centro Emergencia Mujer" OR "centros de atencion" OR "atencion integral") AND violencia AND '
     f'mujer AND {DISENO_LR}'),
    # T02 Salud mental para víctimas
    ("T02", "PubMed", f'{VIOL_PM} AND ("Psychotherapy"[mh] OR "Mental Health Services"[mh] OR "Counseling"[mh]) AND '
     f'{DIS_PM} AND {GEO_PM}'),
    ("T02", "Campbell", "psychological interventions women intimate partner violence mental health"),
    ("T02", "3ie", "psychological intervention violence against women mental health"),
    ("T02", "LA Referencia", f'(psicologica OR "salud mental" OR terapia) AND "violencia de pareja" AND {DISENO_LR}'),
    # T03 Respuesta policial y judicial especializada
    ("T03", "PubMed", f'{VIOL_PM} AND ("Police"[mh] OR "Law Enforcement"[mh] OR court*[tiab] OR "protection order*"[tiab] '
     f'OR "restraining order*"[tiab]) AND {DIS_PM} AND {GEO_PM}'),
    ("T03", "EconBiz", '(subject:"Domestic violence" OR subject:"Gender-based violence") AND (subject:"Police" OR police OR '
     f'court OR courts OR "protection orders") AND {DIS_EB}'),
    ("T03", "Campbell", "police courts protection orders domestic violence"),
    ("T03", "3ie", "police stations courts violence against women"),
    ("T03", "LA Referencia", f'(comisaria OR fiscalia OR juzgado OR "medidas de proteccion") AND violencia AND mujer AND '
     f'{DISENO_LR}'),
    # T04 Hostigamiento sexual laboral
    ("T04", "PubMed", f'"Sexual Harassment"[mh] AND ("Workplace"[mh] OR "Inservice Training"[mh] OR training[tiab] OR '
     f'policy[tiab] OR policies[tiab]) AND {DIS_PM}'),
    ("T04", "ERIC", f'subject:"Sexual Harassment" AND (subject:"Prevention" OR training OR policy) AND {DIS_ER}'),
    ("T04", "EconBiz", 'subject:"Sexual harassment" AND (training OR prevention OR policy OR intervention)'),
    ("T04", "Campbell", "sexual harassment workplace prevention training"),
    ("T04", "3ie", "sexual harassment workplace"),
    ("T04", "LA Referencia", f'"hostigamiento sexual" AND (laboral OR trabajo) AND (prevencion OR capacitacion) AND '
     f'{DISENO_LR}'),
    # T05 Capacitación de operadores de justicia
    ("T05", "PubMed", f'("Police"[mh] OR "Law Enforcement"[mh] OR police[tiab] OR judge*[tiab] OR prosecutor*[tiab]) AND '
     f'("Inservice Training"[mh] OR training[tiab]) AND ({VIOL_PM} OR "Sex Offenses"[mh]) AND {DIS_PM}'),
    ("T05", "Campbell", "police training response domestic violence"),
    ("T05", "3ie", "police training gender violence"),
    ("T05", "LA Referencia", f'capacitacion AND (policia OR policias OR "operadores de justicia") AND violencia AND '
     f'{DISENO_LR}'),
    # T06 Prevención comunitaria y cambio de normas
    ("T06", "PubMed", f'{VIOL_PM} AND ("Community Participation"[mh] OR "Community Networks"[mh] OR "Social Norms"[mh] OR '
     f'"community mobili*"[tiab]) AND {DIS_PM} AND {GEO_PM}'),
    ("T06", "Campbell", "community mobilisation prevention violence against women"),
    ("T06", "3ie", "community mobilization intimate partner violence prevention"),
    ("T06", "LA Referencia", f'prevencion AND violencia AND (comunitaria OR comunidad) AND (mujer OR genero) AND '
     f'{DISENO_LR}'),
    # T07 Trabajo con hombres y masculinidades
    ("T07", "PubMed", '("Masculinity"[mh] OR "gender transformative"[tiab] OR "engaging men"[tiab] OR "men and boys"[tiab]) '
     f'AND ({VIOL_PM} OR "Gender Role"[mh] OR "Gender Equity"[mh]) AND {DIS_PM} AND {GEO_PM}'),
    ("T07", "Campbell", "engaging men boys gender transformative violence prevention"),
    ("T07", "3ie", "engaging men gender norms violence"),
    ("T07", "LA Referencia", f'(masculinidades OR "nuevas masculinidades" OR "trabajo con hombres") AND (programa OR '
     f'intervencion) AND {DISENO_LR}'),
    # T08 Información y comunicación
    ("T08", "PubMed", '("Mass Media"[mh] OR "Health Communication"[mh] OR "Social Marketing"[mh] OR campaign*[tiab] OR '
     f'edutainment[tiab] OR "entertainment-education"[tiab]) AND ({VIOL_PM} OR "Sexism"[mh] OR "Gender Equity"[mh]) AND '
     f'{DIS_PM} AND {GEO_PM}'),
    ("T08", "EconBiz", '(subject:"Communication media" OR "mass media" OR edutainment OR television OR "soap opera") AND '
     '(subject:"Gender-based violence" OR subject:"Domestic violence" OR subject:"Gender role" OR '
     f'subject:"Gender equality" OR subject:"Social norm") AND {DIS_EB}'),
    ("T08", "Campbell", "mass media campaigns violence against women attitudes"),
    ("T08", "3ie", "mass media edutainment gender norms violence"),
    ("T08", "LA Referencia", f'(campana OR "medios de comunicacion" OR comunicacional) AND violencia AND (mujer OR genero) '
     f'AND {DISENO_LR}'),
    # T09 Prevención de la violencia en la escuela
    ("T09", "PubMed", '("Schools"[mh] OR "School Health Services"[mh] OR school*[ti]) AND ("Intimate Partner Violence"[mh] '
     f'OR "dating violence"[tiab] OR "Sex Offenses"[mh] OR "Gender-Based Violence"[mh]) AND {DIS_PM} AND {GEO_PM}'),
    ("T09", "ERIC", '(subject:"Violence" OR subject:"Sexual Abuse" OR subject:"Sexual Harassment" OR "dating violence") AND '
     f'subject:"Prevention" AND {DIS_ER} AND {GEO_ER}'),
    ("T09", "Campbell", "school-based programs dating violence sexual violence prevention"),
    ("T09", "3ie", "school-based violence prevention adolescents"),
    ("T09", "LA Referencia", f'(escuela OR escolar) AND violencia AND prevencion AND programa AND {DISENO_LR}'),
    # T10 Tratamiento de agresores sexuales
    ("T10", "PubMed", '("Sex Offenses"[mh] OR "sex offender*"[tiab] OR "sexual offender*"[tiab]) AND ("Recidivism"[mh] OR '
     f'recidivism[tiab]) AND (treatment[tiab] OR "Psychotherapy"[mh]) AND {DIS_PM}'),
    ("T10", "Campbell", "sex offender treatment recidivism"),
    ("T10", "3ie", "offender treatment recidivism"),
    ("T10", "LA Referencia", '("agresores sexuales" OR "delincuentes sexuales") AND tratamiento AND reincidencia'),
    # T11 Servicios e información en salud sexual y reproductiva
    ("T11", "PubMed", '("Contraception"[mh] OR "Contraception, Postcoital"[mh] OR "Pregnancy in Adolescence"[mh] OR '
     '"Pregnancy, Unplanned"[mh]) AND ("Reproductive Health Services"[mh] OR "Adolescent Health Services"[mh] OR '
     f'"Counseling"[mh]) AND {DIS_PM} AND {GEO_PM}'),
    ("T11", "Campbell", "adolescent pregnancy prevention contraception"),
    ("T11", "3ie", "adolescent sexual reproductive health services contraception"),
    ("T11", "LA Referencia", f'("embarazo adolescente" OR anticoncepcion OR "planificacion familiar") AND (intervencion OR '
     f'programa) AND {DISENO_LR}'),
    # T12 Educación sexual integral
    ("T12", "PubMed", f'"Sex Education"[mh] AND ("Schools"[mh] OR school*[tiab]) AND {DIS_PM} AND {GEO_PM}'),
    ("T12", "ERIC", f'subject:"Sex Education" AND {DIS_ER} AND {GEO_ER}'),
    ("T12", "Campbell", "school-based sexuality education"),
    ("T12", "3ie", "comprehensive sexuality education school"),
    ("T12", "LA Referencia", f'"educacion sexual" AND (escuela OR escolar OR estudiantes) AND {DISENO_LR}'),
    # T13 Participación política de las mujeres
    ("T13", "EconBiz", '(subject:"Political participation" OR subject:"Election" OR subject:"Voting behaviour") AND '
     f'subject:"Women" AND {DIS_EB}'),
    ("T13", "Campbell", "women political participation representation"),
    ("T13", "3ie", "women political participation voting"),
    ("T13", "LA Referencia", f'"participacion politica" AND mujeres AND (capacitacion OR candidatas OR cuotas) AND '
     f'{DISENO_LR}'),
    # T14 Servicios de cuidado
    ("T14", "EconBiz", '(subject:"Child care" OR childcare OR daycare) AND (subject:"Mothers" OR subject:"Women") AND '
     f'(subject:"Employment" OR subject:"Labour supply") AND {DIS_EB}'),
    ("T14", "PubMed", '("Child Day Care Centers"[mh] OR "Child Care"[mh]) AND ("Women, Working"[mh] OR "Employment"[mh] OR '
     f'"Mothers"[mh]) AND {DIS_PM} AND {GEO_PM}'),
    ("T14", "Campbell", "childcare maternal employment"),
    ("T14", "3ie", "childcare women labour market"),
    ("T14", "LA Referencia", f'("cuidado infantil" OR guarderias OR "Wawa Wasi" OR "Cuna Mas") AND (empleo OR trabajo) AND '
     f'(madres OR mujeres) AND {DISENO_LR}'),
    # T15 Formalización de la propiedad y de derechos sobre recursos
    ("T15", "EconBiz", '(subject:"Land tenure" OR titling OR "property rights") AND (subject:"Women" OR subject:"Gender" '
     f'OR women) AND {DIS_EB}'),
    ("T15", "Campbell", "land titling property rights"),
    ("T15", "3ie", "land titling property rights women"),
    ("T15", "LA Referencia", f'(titulacion OR "derechos de propiedad" OR "derechos de agua") AND (mujeres OR genero) AND '
     f'{DISENO_LR}'),
    # T16 Participación de las mujeres en la gestión de recursos naturales
    ("T16", "EconBiz", '(subject:"Forestry" OR subject:"Fisheries" OR subject:"Natural resources" OR '
     'subject:"Water management") AND (subject:"Women" OR subject:"Gender") AND (subject:"Participation" OR participation)'),
    ("T16", "Campbell", "women participation natural resource management forest"),
    ("T16", "3ie", "women participation natural resource management"),
    ("T16", "LA Referencia", f'participacion AND mujeres AND ("gestion forestal" OR pesca OR "recursos naturales" OR '
     f'"gestion del agua")'),
    # T17 Formación laboral, certificación y carreras no tradicionales
    ("T17", "EconBiz", '(subject:"Vocational training" OR subject:"Further training" OR subject:"Labour market policy") AND '
     f'(subject:"Women" OR subject:"Gender") AND {DIS_EB}'),
    ("T17", "ERIC", '(subject:"Vocational Education" OR subject:"Job Training") AND (subject:"Females" OR '
     f'subject:"Womens Education") AND {DIS_ER} AND {GEO_ER}'),
    ("T17", "Campbell", "vocational training women employment"),
    ("T17", "3ie", "vocational training women employment"),
    ("T17", "LA Referencia", f'("capacitacion laboral" OR "formacion tecnica" OR "certificacion de competencias" OR '
     f'ProJoven OR "Jovenes Productivos" OR "Impulsa Peru") AND {DISENO_LR}'),
    # T18 Desarrollo empresarial y financiamiento
    ("T18", "EconBiz", '(subject:"Microfinance" OR subject:"Women entrepreneurs" OR subject:"Business start-up" OR '
     f'"business training") AND subject:"Women" AND {DIS_EB}'),
    ("T18", "Campbell", "business training microfinance women entrepreneurs"),
    ("T18", "3ie", "business training women entrepreneurs"),
    ("T18", "LA Referencia", f'(emprendimiento OR microcredito OR "capacitacion empresarial") AND mujeres AND {DISENO_LR}'),
    # T19 Lactarios institucionales
    ("T19", "PubMed", '"Breast Feeding"[mh] AND ("Workplace"[mh] OR "Women, Working"[mh] OR workplace[tiab] OR '
     f'"lactation room*"[tiab] OR employer*[tiab]) AND {DIS_PM}'),
    ("T19", "LA Referencia", f'(lactario OR lactarios OR "lactancia materna") AND (trabajo OR laboral OR trabajadoras) AND '
     f'{DISENO_LR}'),
    ("T19", "3ie", "breastfeeding working mothers workplace"),
    # T20 Transferencias condicionadas y culminación de la secundaria
    ("T20", "EconBiz", '("conditional cash transfer" OR "conditional cash transfers" OR subject:"Social security benefits") '
     f'AND (subject:"School attendance" OR subject:"Educational attainment" OR subject:"Education") AND {DIS_EB} AND '
     f'{GEO_EB}'),
    ("T20", "Campbell", "cash transfers schooling"),
    ("T20", "3ie", "conditional cash transfers education"),
    ("T20", "LA Referencia", f'("programa Juntos" OR "transferencias condicionadas") AND (escolar OR educacion OR '
     f'secundaria) AND {DISENO_LR}'),
    # T21 Transversalización del enfoque de género
    ("T21", "EconBiz", '(subject:"Gender mainstreaming" OR "gender budgeting" OR "gender-responsive budgeting") AND '
     '(subject:"Evaluation" OR subject:"Impact assessment" OR effects OR outcomes OR impact)'),
    ("T21", "Campbell", "gender mainstreaming"),
    ("T21", "3ie", "gender mainstreaming gender responsive budgeting"),
    ("T21", "LA Referencia", '("transversalizacion del enfoque de genero" OR "presupuesto con enfoque de genero" OR '
     '"presupuestos sensibles al genero") AND (evaluacion OR efectos OR resultados)'),
    # T22 Estadísticas y estudios de género
    ("T22", "EconBiz", 'subject:"Time use" AND (subject:"Social statistics" OR survey OR surveys) AND (policy OR policies)'),
    ("T22", "3ie", "gender data statistics policy use"),
    ("T22", "LA Referencia", '("encuesta de uso del tiempo" OR "estadisticas de genero") AND ("politicas publicas" OR '
     'politica)'),
    # T23 Educación para la igualdad en la escuela
    ("T23", "ERIC", '(subject:"Sex Fairness" OR subject:"Sex Stereotypes" OR subject:"Gender Bias" OR subject:"Sex Role") '
     f'AND (subject:"Attitude Change" OR subject:"Program Effectiveness") AND (subject:"Randomized Controlled Trials" OR '
     f'{GEO_ER})'),
    ("T23", "PubMed", '("Gender Role"[mh] OR "Gender Equity"[mh] OR "Sexism"[mh]) AND ("Schools"[mh] OR school*[tiab]) AND '
     f'{DIS_PM} AND {GEO_PM}'),
    ("T23", "Campbell", "school-based gender attitudes adolescents"),
    ("T23", "3ie", "gender attitudes school adolescents"),
    ("T23", "LA Referencia", f'("igualdad de genero" OR "estereotipos de genero") AND (escuela OR docentes OR estudiantes) '
     f'AND (intervencion OR programa) AND {DISENO_LR}'),
    # T24 Corresponsabilidad familiar en el cuidado
    ("T24", "PubMed", '("Fathers"[mh] OR "Paternal Behavior"[mh] OR "Spouses"[mh]) AND ("Child Care"[mh] OR '
     '"Child Rearing"[mh] OR household[tiab] OR "domestic work"[tiab] OR "unpaid care"[tiab]) AND ("Gender Role"[mh] OR '
     f'"Gender Equity"[mh] OR "gender transformative"[tiab]) AND {DIS_PM} AND {GEO_PM}'),
    ("T24", "EconBiz", '(subject:"Unpaid work" OR subject:"Care work" OR subject:"Household production" OR '
     'subject:"Division of labour") AND (subject:"Fathers" OR subject:"Gender role" OR men) AND (intervention OR '
     'randomized OR experiment OR program)'),
    ("T24", "Campbell", "fathers involvement childcare household intervention"),
    ("T24", "3ie", "men unpaid care work household chores"),
    ("T24", "LA Referencia", '(corresponsabilidad OR "trabajo domestico" OR "tareas domesticas") AND (padres OR varones OR '
     'hombres) AND (intervencion OR programa OR taller)'),
]
# Banco Mundial (palabras sueltas en inglés) y OpenAlex (español), añadidas el 2026-10-31 a pedido del usuario
BANCO_MUNDIAL = {
    "T01": "violence women survivors services", "T02": "violence women mental health intervention",
    "T03": "police violence women", "T04": "sexual harassment workplace", "T05": "police training violence women",
    "T06": "intimate partner violence prevention community", "T07": "men gender norms violence",
    "T08": "media gender norms violence", "T09": "school violence prevention", "T10": "sexual offenders recidivism",
    "T11": "adolescent pregnancy contraception", "T12": "sexuality education adolescents",
    "T13": "women political participation", "T14": "childcare female employment", "T15": "land titling women",
    "T16": "women natural resource management", "T17": "vocational training women",
    "T18": "business training women entrepreneurs", "T19": "breastfeeding working mothers",
    "T20": "conditional cash transfers schooling", "T21": "gender budgeting", "T22": "time use survey gender policy",
    "T23": "gender attitudes school", "T24": "fathers childcare household",
}
OPENALEX = {
    "T01": '("violencia contra la mujer" OR "violencia de pareja" OR "violencia de género") AND ("Centro Emergencia Mujer" '
           f'OR "atención integral" OR "centros de atención" OR "servicios de atención") AND {DIS_OA}',
    "T02": '("violencia de pareja" OR "violencia contra la mujer") AND (psicoterapia OR "intervención psicológica" OR '
           f'"salud mental") AND {DIS_OA}',
    "T03": '("violencia contra la mujer" OR "violencia familiar" OR "violencia de género") AND (comisaría OR policía OR '
           f'"medidas de protección" OR fiscalía OR juzgado) AND {DIS_OA}',
    "T04": '("hostigamiento sexual" OR "acoso sexual") AND (laboral OR trabajo) AND (prevención OR capacitación OR '
           f'intervención) AND {DIS_OA}',
    "T05": '(capacitación OR formación) AND (policías OR "operadores de justicia" OR jueces OR fiscales) AND ("violencia '
           f'contra la mujer" OR "violencia de género" OR "violencia familiar") AND {DIS_OA}',
    "T06": '("violencia de pareja" OR "violencia contra la mujer") AND (comunitaria OR comunidad OR "normas sociales") AND '
           f'prevención AND {DIS_OA}',
    "T07": '(masculinidades OR "trabajo con hombres" OR "grupos de hombres") AND (violencia OR igualdad) AND (programa OR '
           f'intervención) AND {DIS_OA}',
    "T08": '(campaña OR "medios de comunicación" OR eduentretenimiento OR telenovela) AND ("violencia contra la mujer" OR '
           f'"violencia de género") AND {DIS_OA}',
    "T09": '(escuela OR escolar OR colegio) AND ("violencia en el noviazgo" OR "violencia escolar" OR "acoso escolar") AND '
           f'(programa OR intervención) AND {DIS_OA}',
    "T10": '("agresores sexuales" OR "delincuentes sexuales") AND tratamiento AND reincidencia',
    "T11": '("embarazo adolescente" OR anticoncepción OR "planificación familiar") AND (intervención OR programa OR '
           f'consejería) AND {DIS_OA}',
    "T12": f'"educación sexual" AND (escuela OR escolar OR adolescentes) AND {DIS_OA}',
    "T13": '("participación política" OR candidatas OR "representación política") AND mujeres AND (capacitación OR '
           f'programa OR formación) AND {DIS_OA}',
    "T14": '("cuidado infantil" OR guarderías OR "Cuna Más" OR "Wawa Wasi" OR "centros infantiles") AND (empleo OR '
           '"participación laboral" OR trabajo) AND (madres OR mujeres)',
    "T15": f'(titulación OR "derechos de propiedad" OR "tenencia de la tierra") AND (mujeres OR género) AND {DIS_OA}',
    "T16": 'participación AND mujeres AND ("gestión forestal" OR pesca OR "recursos naturales" OR "gestión del agua") AND '
           f'{DIS_OA}',
    "T17": '("capacitación laboral" OR "formación técnica" OR "formación profesional" OR "certificación de competencias") '
           f'AND (mujeres OR jóvenes) AND {DIS_OA}',
    "T18": f'(microcrédito OR microfinanzas OR "capacitación empresarial" OR emprendimiento) AND mujeres AND {DIS_OA}',
    "T19": f'(lactario OR lactarios OR "lactancia materna") AND (trabajo OR laboral OR trabajadoras) AND {DIS_OA}',
    "T20": '("transferencias condicionadas" OR "programa Juntos" OR Oportunidades OR "Bolsa Família") AND (escolar OR '
           f'educación OR secundaria) AND {DIS_OA}',
    "T21": '("presupuesto con enfoque de género" OR "presupuestos sensibles al género" OR "transversalización del enfoque de '
           f'género" OR "presupuesto de género") AND {DIS_OA}',
    "T22": '("encuesta de uso del tiempo" OR "estadísticas de género") AND ("políticas públicas" OR política)',
    "T23": '("igualdad de género" OR "estereotipos de género" OR coeducación) AND (escuela OR docentes OR estudiantes) AND '
           f'(programa OR intervención) AND {DIS_OA}',
    "T24": '(corresponsabilidad OR "trabajo doméstico" OR "tareas domésticas" OR paternidad) AND (hombres OR padres OR '
           f'varones) AND (programa OR intervención) AND {DIS_OA}',
}
ESTRATEGIAS += [(t, "Banco Mundial", c) for t, c in BANCO_MUNDIAL.items()]
ESTRATEGIAS += [(t, "OpenAlex", c) for t, c in OPENALEX.items()]
REVISADOS = {"PubMed": 30, "ERIC": 15, "EconBiz": 30, "LA Referencia": 15, "Campbell": 10, "3ie": 10, "Banco Mundial": 15,
             "OpenAlex": 15}
INTERFAZ = {
    "PubMed": "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi (orden: relevancia)",
    "ERIC": "https://api.ies.ed.gov/eric/ (solo arbitrados: peerreviewed:T)",
    "EconBiz": "https://api.econbiz.de/v1/search (orden: relevancia)",
    "Campbell": "https://api.crossref.org/works (filtro issn:1891-1803, artículos; orden: relevancia)",
    "3ie": "https://api.crossref.org/works (filtro prefix:10.23846; orden: relevancia)",
    "LA Referencia": "https://buscador.lareferencia.info/api/v1/search (artículos e informes; orden: relevancia)",
    "Banco Mundial": "https://search.worldbank.org/api/v3/wds (Policy Research Working Papers; orden: relevancia)",
    "OpenAlex": "https://api.openalex.org/works (orden: relevancia)",
}
FILTROS = {
    "PubMed": "Diseño: revisión sistemática, ensayo aleatorio, estudio de evaluación o evaluación de programa; contexto "
              "(cuando la cadena lo incluye): países en desarrollo, América Latina, África, Asia, ingreso bajo y medio, o "
              "revisión sistemática. Sin límite de año.",
    "ERIC": "Solo arbitrados; diseño mediante descriptores de efectividad, evaluación, ensayo o metaanálisis. Sin límite "
            "de año.",
    "EconBiz": "Diseño mediante descriptores STW o términos de evaluación de impacto; contexto (cuando la cadena lo "
               "incluye): países en desarrollo, América Latina o Perú. Sin límite de año.",
    "Campbell": "Revista Campbell Systematic Reviews (revisiones, protocolos y mapas). Sin límite de año.",
    "3ie": "Publicaciones de 3ie con DOI (revisiones, evaluaciones de impacto, mapas y documentos de trabajo). Sin límite "
           "de año.",
    "LA Referencia": "Formato artículo o informe (se excluyen tesis). Sin límite de año.",
    "Banco Mundial": "Tipo de documento Policy Research Working Paper. Palabras combinadas como conjunción. Sin límite de "
                     "año.",
    "OpenAlex": "Idioma español; tipos artículo, revisión o informe; búsqueda en título, resumen y texto completo indexado. "
                "Sin límite de año.",
}
TESAURO = {"PubMed": "MeSH (NLM)", "ERIC": "Tesauro de ERIC", "EconBiz": "Tesauro de Economía STW (ZBW)"}


def pedir(url, nombre, url_registro=None):
    """Descarga una URL, guarda la respuesta cruda y la registra con su hash (url_registro: la URL sin credenciales)."""
    for intento in range(4):
        try:
            datos = urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=60).read()
            break
        except Exception:
            time.sleep(3 * (intento + 1))
    else:
        raise RuntimeError(f"sin respuesta: {url_registro or url}")
    ruta = CRUDO / nombre
    ruta.write_bytes(datos)
    descargas.append((url_registro or url, ruta.as_posix(), hashlib.sha256(datos).hexdigest(), FECHA,
                      "Respuesta de la API (metadatos)"))
    time.sleep(0.4)
    return datos


def pubmed(cadena, n, eid):
    q = urllib.parse.urlencode({"db": "pubmed", "term": cadena, "retmode": "json", "retmax": n, "sort": "relevance"})
    url = f"https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi?{q}"
    r = json.loads(pedir(url, f"{eid}_esearch.json"))["esearchresult"]
    ids = r["idlist"]
    regs = []
    if ids:
        s = json.loads(pedir("https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esummary.fcgi?db=pubmed&retmode=json&id="
                             + ",".join(ids), f"{eid}_esummary.json"))["result"]
        pedir("https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=pubmed&rettype=abstract&retmode=xml&id="
              + ",".join(ids), f"{eid}_efetch.xml")
        for pmid in ids:
            d = s[pmid]
            doi = next((a["value"] for a in d.get("articleids", []) if a["idtype"] == "doi"), "")
            regs.append({"clave": f"PMID:{pmid}", "doi": doi, "titulo": d.get("title", ""),
                         "anio": d.get("pubdate", "")[:4], "fuente": d.get("fulljournalname", d.get("source", "")),
                         "tipo_documento": "; ".join(d.get("pubtype", [])),
                         "enlace": f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/"})
    return int(r["count"]), url, regs


def eric(cadena, n, eid):
    q = urllib.parse.urlencode({"search": f"({cadena}) AND peerreviewed:T", "format": "json", "rows": n,
                                "fields": "id,title,author,publicationdateyear,source,publicationtype,url,description"})
    url = f"https://api.ies.ed.gov/eric/?{q}"
    r = json.loads(pedir(url, f"{eid}.json"))["response"]
    regs = [{"clave": f"ERIC:{d['id']}", "doi": "", "titulo": d.get("title", ""), "anio": str(d.get("publicationdateyear", "")),
             "fuente": d.get("source", ""), "tipo_documento": "; ".join(d.get("publicationtype", [])),
             "enlace": f"https://eric.ed.gov/?id={d['id']}"} for d in r["docs"]]
    return r["numFound"], url, regs


def econbiz(cadena, n, eid):
    q = urllib.parse.urlencode({"q": cadena, "size": n})
    url = f"https://api.econbiz.de/v1/search?{q}"
    r = json.loads(pedir(url, f"{eid}.json"))["hits"]
    regs = []
    for d in r["hits"]:
        urls = d.get("identifier_url", [])
        doi = next((u.split("doi.org/", 1)[1] for u in urls if "doi.org/" in u), "")
        regs.append({"clave": f"ECONBIZ:{d['id']}", "doi": doi, "titulo": " ".join(d.get("title", [])) if isinstance(
            d.get("title"), list) else d.get("title", ""), "anio": (d.get("date") or [""])[0][:4],
            "fuente": "; ".join(d.get("isPartOf") or d.get("publisher") or []), "tipo_documento": d.get("type", ""),
            "enlace": f"https://www.econbiz.de/Record/{d['id']}"})
    return r["total"], url, regs


def crossref(cadena, n, eid, filtro):
    q = urllib.parse.urlencode({"query.bibliographic": cadena, "filter": filtro, "rows": n,
                                "select": "DOI,title,issued,type,container-title,author,abstract"})
    url = f"https://api.crossref.org/works?{q}"
    m = json.loads(pedir(url, f"{eid}.json"))["message"]
    regs = [{"clave": f"DOI:{d['DOI'].lower()}", "doi": d["DOI"].lower(), "titulo": " ".join(d.get("title", [""])),
             "anio": str(d.get("issued", {}).get("date-parts", [[""]])[0][0] or ""),
             "fuente": " ".join(d.get("container-title", [])) or ("3ie" if filtro.startswith("prefix") else ""),
             "tipo_documento": d.get("type", ""), "enlace": f"https://doi.org/{d['DOI']}"} for d in m["items"]]
    return m["total-results"], url, regs


def lareferencia(cadena, n, eid):
    params = [("lookfor", cadena), ("type", "AllFields"), ("limit", n), ("filter[]", '~format:"article"'),
              ("filter[]", '~format:"report"')] + [("field[]", f) for f in (
                  "id", "title", "formats", "publicationDates", "urls", "primaryAuthors", "summary", "containerTitle")]
    url = f"https://buscador.lareferencia.info/api/v1/search?{urllib.parse.urlencode(params)}"
    r = json.loads(pedir(url, f"{eid}.json"))
    regs = []
    for d in r.get("records", []):
        enlace = (d.get("urls") or [{}])[0].get("url", "")
        doi = enlace.split("doi.org/", 1)[1].lower() if "doi.org/" in enlace else ""
        regs.append({"clave": f"LAREF:{d['id']}", "doi": doi, "titulo": d.get("title", ""),
                     "anio": (d.get("publicationDates") or [""])[0][:4], "fuente": d.get("containerTitle", ""),
                     "tipo_documento": "; ".join(d.get("formats", [])), "enlace": enlace})
    return r.get("resultCount", 0), url, regs


def bancomundial(cadena, n, eid):
    q = urllib.parse.urlencode({"format": "json", "qterm": cadena, "rows": n, "docty_exact": "Policy Research Working Paper",
                                "fl": "display_title,docdt,docty,authors,url,repnb,abstracts"})
    url = f"https://search.worldbank.org/api/v3/wds?{q}"
    d = json.loads(pedir(url, f"{eid}.json").decode("utf-8-sig"))
    time.sleep(1.5)
    regs = [{"clave": f"BM:{v['id']}", "doi": "", "titulo": v.get("display_title", ""), "anio": v.get("docdt", "")[:4],
             "fuente": "Banco Mundial, Policy Research Working Paper", "tipo_documento": v.get("docty", ""),
             "enlace": v.get("url", "")} for k, v in d.get("documents", {}).items() if k != "facets"]
    return d.get("total", 0), url, regs


def openalex(cadena, n, eid):
    q = urllib.parse.urlencode({"search": cadena, "filter": "language:es,type:article|review|report", "per_page": n,
                                "select": "id,doi,display_name,publication_year,type,primary_location,authorships,"
                                          "abstract_inverted_index"})
    url = f"https://api.openalex.org/works?{q}"
    con_clave = url + (f"&api_key={OA_CLAVE}" if OA_CLAVE else "")
    d = json.loads(pedir(con_clave, f"{eid}.json", url_registro=url))
    regs = []
    for w in d.get("results", []):
        doi = (w.get("doi") or "").replace("https://doi.org/", "").lower()
        fuente = ((w.get("primary_location") or {}).get("source") or {}).get("display_name") or ""
        regs.append({"clave": f"OPENALEX:{w['id'].rsplit('/', 1)[1]}", "doi": doi, "titulo": w.get("display_name") or "",
                     "anio": str(w.get("publication_year") or ""), "fuente": fuente, "tipo_documento": w.get("type", ""),
                     "enlace": f"https://doi.org/{doi}" if doi else w["id"]})
    return d.get("meta", {}).get("count", 0), url, regs


BUSCAR = {"PubMed": pubmed, "Banco Mundial": bancomundial, "OpenAlex": openalex, "ERIC": eric, "EconBiz": econbiz, "LA Referencia": lareferencia,
          "Campbell": lambda c, n, e: crossref(c, n, e, "issn:1891-1803,type:journal-article"),
          "3ie": lambda c, n, e: crossref(c, n, e, "prefix:10.23846")}


def componentes(base, cadena):
    """Separa descriptores controlados, términos libres y filtros de una cadena."""
    if base == "PubMed":
        desc = re.findall(r'"([^"]+)"\[mh\]', cadena)
        libres = re.findall(r'("[^"]+"|[\w*\'-]+)\[(?:tiab|ti)\]', cadena)
        filtros = re.findall(r'("[^"]+"|\w+)\[(?:pt|sb)\]', cadena)
    elif base in ("ERIC", "EconBiz"):
        desc = re.findall(r'subject:"([^"]+)"', cadena)
        resto = re.sub(r'subject:"[^"]+"', " ", cadena)
        libres = [t for t in re.findall(r'"[^"]+"|[\w*-]+', resto) if t not in ("AND", "OR", "NOT")]
        filtros = ["peerreviewed:T"] if base == "ERIC" else []
    else:
        desc, filtros = [], []
        libres = [t for t in re.findall(r'"[^"]+"|[\w*-]+', cadena) if t not in ("AND", "OR", "NOT")]
    unico = lambda xs: list(dict.fromkeys(x.strip('"') for x in xs))
    return unico(desc), unico(libres), unico(filtros)


def verificar(tesauro, termino):
    """Identificador del descriptor en su tesauro (MeSH y STW) o uso en registros de ERIC."""
    qt = urllib.parse.quote
    if tesauro.startswith("MeSH"):
        d = json.loads(pedir(f"https://id.nlm.nih.gov/mesh/lookup/descriptor?label={qt(termino)}&match=exact",
                             f"mesh_{re.sub(r'[^A-Za-z0-9]+', '_', termino)}.json"))
        return (d[0]["resource"].rsplit("/", 1)[1], d[0]["resource"], "Descriptor exacto en la API de MeSH") if d else (
            "", "", "No encontrado")
    if tesauro.startswith("Tesauro de Economía"):
        d = json.loads(pedir(f"https://zbw.eu/beta/skosmos/rest/v1/stw/search?query={qt(termino)}&lang=en",
                             f"stw_{re.sub(r'[^A-Za-z0-9]+', '_', termino)}.json"))
        r = d.get("results", [])
        if not r:
            return "", "", "No encontrado"
        etiqueta = r[0].get("prefLabel", "")
        nota = "Término preferente en STW" if etiqueta == termino else f"Término no preferente; preferente: {etiqueta}"
        return r[0]["uri"].rsplit("/", 1)[1], r[0]["uri"], nota
    s = urllib.parse.quote(f'subject:"{termino}"')
    d = json.loads(pedir(f"https://api.ies.ed.gov/eric/?search={s}&format=json&rows=0",
                         f"eric_{re.sub(r'[^A-Za-z0-9]+', '_', termino)}.json"))
    n = d["response"]["numFound"]
    return "", f"https://eric.ed.gov/?ti={qt(termino)}", f"Descriptor asignado en {n} registros de ERIC" if n else "Sin uso"


if __name__ == "__main__":
    CRUDO.mkdir(parents=True, exist_ok=True)
    DEST.mkdir(parents=True, exist_ok=True)
    descargas = []
    estrategias, resultados, usados = [], [], {}
    vistos = {}
    for i, (tipo, base, cadena) in enumerate(ESTRATEGIAS, 1):
        eid = f"ACA-{i:03d}"
        n = REVISADOS[base]
        total, url, regs = BUSCAR[base](cadena, n, eid)
        desc, libres, filtros = componentes(base, cadena)
        for dsc in desc:
            usados.setdefault((TESAURO[base], dsc), set()).add(eid)
        estrategias.append({"id": eid, "tipo": tipo, "base": base, "interfaz": INTERFAZ[base], "cadena": cadena,
                            "tesauro": TESAURO.get(base, "Sin tesauro (términos libres)"),
                            "descriptores": "; ".join(desc), "terminos_libres": "; ".join(libres),
                            "filtros": "; ".join(filtros + [FILTROS[base]]), "fecha": FECHA, "total_resultados": total,
                            "revisados": len(regs), "url_consulta": url})
        for rango, r in enumerate(regs, 1):
            norm = re.sub(r"[^a-z0-9]", "", r["titulo"].lower())[:80]
            llave = r["doi"].lower() or norm
            dup = vistos.get(llave) or vistos.get(norm, "")
            vistos.setdefault(llave, r["clave"])
            vistos.setdefault(norm, r["clave"])
            resultados.append({"estrategia": eid, "tipo": tipo, "base": base, "rango": rango, **r,
                               "duplicado_de": dup if dup != r["clave"] else ""})
        print(eid, tipo, base, total, len(regs), flush=True)
    descriptores = []
    for (tes, dsc), eids in sorted(usados.items()):
        ident, uri, nota = verificar(tes, dsc)
        descriptores.append({"tesauro": tes, "descriptor": dsc, "identificador": ident, "uri": uri,
                             "verificacion": f"{FECHA}: {nota}", "estrategias": ", ".join(sorted(eids))})
    for nombre, filas in (("estrategias_academicas.csv", estrategias), ("resultados_academicos.csv", resultados),
                          ("descriptores.csv", descriptores)):
        with open(DEST / nombre, "w", encoding="utf-8", newline="") as f:
            w = csv.DictWriter(f, fieldnames=list(filas[0]), quoting=csv.QUOTE_ALL)
            w.writeheader()
            w.writerows(filas)
    with open(REGISTRO, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f, delimiter="\t", lineterminator="\n")
        w.writerow(["url_origen", "destino", "sha256", "fecha_descarga", "verificacion_identidad"])
        w.writerows(descargas)
    sin = [d for d in descriptores if not d["identificador"] and not d["verificacion"].endswith("registros de ERIC")]
    print(len(estrategias), "estrategias;", len(resultados), "registros;", len(descriptores), "descriptores;",
          "sin verificar:", [(d["tesauro"], d["descriptor"]) for d in sin])
