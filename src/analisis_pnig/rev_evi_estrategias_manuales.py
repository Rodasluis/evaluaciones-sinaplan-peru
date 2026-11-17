"""REV_EVI: estrategias de búsqueda que se ejecutan a mano (nota metodológica, Anexo 6).

Las bases de esta lista no admiten consultas automatizadas o requieren suscripción:
  - SciELO (search.scielo.org) y LILACS (Biblioteca Virtual en Salud, pesquisa.bvsalud.org), que bloquean las consultas
    automatizadas; LILACS se busca con descriptores DeCS;
  - el Development Evidence Portal de 3ie (developmentevidence.3ieimpact.org), que carga sus resultados con JavaScript;
    su página consume una API GraphQL pública, por lo que sus estrategias (DEP-15 a DEP-22) se ejecutan con
    src/analisis_pnig/rev_evi_busqueda_3ie_dep.py, que completa aquí las columnas de ejecución;
  - Scopus, Web of Science y EconLit, que requieren suscripción institucional. En la evaluación de la PNIG (2026) se usó
    Scopus con el acceso remoto de la UNMSM; Web of Science y EconLit no tuvieron acceso y se dejan registradas como
    opción para otro usuario con suscripción;
  - JSTOR (MAN-43 a MAN-46), que sustituye a EconLit para los grupos económicos cuando hay acceso institucional: búsqueda
    avanzada con la población y la intervención en el campo Abstract y el diseño en All fields, disciplina Economics;
    las cadenas son versiones compactas de las de los otros grupos porque la interfaz limita el número de términos.
Para reducir el número de búsquedas, los 24 tipos de intervención se agrupan en ocho grupos temáticos. Cada estrategia se
registra con su cadena exacta y los filtros que deben aplicarse en la interfaz; quien la ejecuta completa la fecha, el
total de resultados y el archivo exportado (ver busqueda_manual/protocolo_busqueda_manual.md).

Descriptores DeCS: son los equivalentes en español de los descriptores MeSH verificados en descriptores.csv (DeCS es la
traducción trilingüe de MeSH que mantiene BIREME). No se pudieron verificar en línea porque el portal de DeCS bloquea las
consultas automatizadas; se comprobaron por su conteo en LILACS el 2026-10-31 (busqueda_manual/lilacs_descriptores_verificados.tsv).
Las cadenas de LILACS (LILACS_CADENAS) tienen la estructura población AND intervención AND diseño: las generadas al inicio
juntaban población e intervención en un solo bloque OR y se corrigieron antes de ejecutarlas (la original consta en la
columna observaciones). Códigos JEL verificados en la guía de la American Economic Association
(aeaweb.org/econlit/jelCodes.php) el 2026-10-31.

Uso (desde la raíz del repositorio):
    python src/analisis_pnig/rev_evi_estrategias_manuales.py [carpeta_destino]
La carpeta de destino por defecto es evaluaciones/pnig_2026/04_analisis/rev_evi_insumos/busqueda_manual. Si el archivo
de estrategias ya existe, se conservan las columnas que completa quien ejecuta la búsqueda.
"""
import csv
import sys
from pathlib import Path
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

DEST = Path(sys.argv[1] if len(sys.argv) > 1 else f"{EVAL}/04_analisis/rev_evi_insumos/busqueda_manual")

# Grupo: tipos, términos en inglés (intervención y resultado), términos en español, descriptores DeCS, códigos JEL
GRUPOS = [
    ("G1", "Atención, salud mental, justicia y policía ante la violencia contra las mujeres", "T01; T02; T03; T05",
     '("intimate partner violence" OR "domestic violence" OR "violence against women" OR "gender-based violence") AND '
     '(advocacy OR "support services" OR "crisis cent*" OR "justice cent*" OR psychotherapy OR "mental health" OR police '
     'OR "protection order*" OR court* OR training)',
     '("violencia contra la mujer" OR "violencia de pareja" OR "violencia familiar" OR "violencia de género") AND '
     '("atención integral" OR "Centro Emergencia Mujer" OR psicoterapia OR "salud mental" OR policía OR comisaría OR '
     '"medidas de protección" OR fiscalía OR capacitación)',
     ["Violencia de Pareja", "Violencia Doméstica", "Mujeres Maltratadas", "Violencia de Género", "Intervención en la "
      "Crisis (Psiquiatría)", "Psicoterapia", "Servicios de Salud Mental", "Policía", "Aplicación de la Ley"],
     ["J12", "J16", "K42"]),
    ("G2", "Prevención de la violencia: comunidad, hombres, comunicación y escuela", "T06; T07; T08; T09",
     '("intimate partner violence" OR "violence against women" OR "gender-based violence" OR "dating violence") AND '
     '(prevention) AND (community OR "social norms" OR men OR masculinit* OR "mass media" OR campaign* OR edutainment OR '
     'school*)',
     '("violencia contra la mujer" OR "violencia de pareja" OR "violencia de género" OR "violencia en el noviazgo") AND '
     'prevención AND (comunidad OR "normas sociales" OR hombres OR masculinidades OR campaña OR "medios de comunicación" '
     'OR escuela)',
     ["Violencia de Pareja", "Violencia de Género", "Participación de la Comunidad", "Normas Sociales", "Masculinidad",
      "Medios de Comunicación de Masas", "Comunicación en Salud", "Instituciones Académicas"],
     ["J12", "J16", "D91"]),
    ("G3", "Hostigamiento sexual laboral", "T04",
     '("sexual harassment") AND (workplace OR employ* OR work) AND (prevention OR training OR policy OR intervention)',
     '("hostigamiento sexual" OR "acoso sexual") AND (laboral OR trabajo) AND (prevención OR capacitación OR política)',
     ["Acoso Sexual", "Lugar de Trabajo"], ["J16", "J71", "J81"]),
    ("G4", "Tratamiento de agresores sexuales", "T10",
     '("sex offender*" OR "sexual offender*" OR "sexual offense*") AND (treatment OR therapy) AND recidivism',
     '("agresores sexuales" OR "delincuentes sexuales" OR "delitos sexuales") AND tratamiento AND reincidencia',
     ["Delitos Sexuales", "Reincidencia", "Criminales"], ["K42"]),
    ("G5", "Salud sexual y reproductiva y educación sexual integral", "T11; T12",
     '("adolescent pregnancy" OR "teen* pregnancy" OR contracepti* OR "family planning" OR "sexuality education" OR '
     '"sex education") AND (adolescen* OR youth OR school*)',
     '("embarazo adolescente" OR anticoncepción OR "planificación familiar" OR "educación sexual") AND (adolescentes OR '
     'escuela OR escolares)',
     ["Embarazo en Adolescencia", "Anticoncepción", "Embarazo no Planeado", "Servicios de Salud Reproductiva",
      "Servicios de Salud para Adolescentes", "Educación Sexual"], ["J13", "I12", "I21"]),
    ("G6", "Cuidado, lactarios y corresponsabilidad", "T14; T19; T24",
     '(childcare OR "child care" OR daycare OR "day care" OR preschool OR breastfeeding OR "lactation room*" OR fathers OR '
     '"unpaid care" OR housework) AND (employment OR "labor supply" OR "labour supply" OR work* OR "division of labo*")',
     '("cuidado infantil" OR guarderías OR "Cuna Más" OR "Wawa Wasi" OR "lactancia materna" OR lactario OR '
     'corresponsabilidad OR "trabajo doméstico") AND (empleo OR trabajo OR "participación laboral" OR padres)',
     ["Cuidado del Niño", "Guarderías Infantiles", "Lactancia Materna", "Mujeres Trabajadoras", "Padre", "Conducta "
      "Paterna", "Rol de Género"], ["J13", "J22", "J16", "D13"]),
    ("G7", "Autonomía económica: formación, emprendimiento, propiedad, recursos naturales y transferencias",
     "T15; T16; T17; T18; T20",
     '(women OR gender OR female) AND ("vocational training" OR "job training" OR "business training" OR microcredit OR '
     'microfinance OR "land titl*" OR "property rights" OR "natural resource management" OR forest* OR fisher* OR '
     '"conditional cash transfer*")',
     '(mujeres OR género) AND ("capacitación laboral" OR "formación técnica" OR "capacitación empresarial" OR '
     'microcrédito OR microfinanzas OR titulación OR "derechos de propiedad" OR "gestión forestal" OR pesca OR '
     '"transferencias condicionadas" OR Juntos)',
     ["Capacitación Profesional", "Empleo", "Equidad de Género"],
     ["J16", "J24", "L26", "G21", "Q15", "Q22", "Q23", "I38", "I25"]),
    ("G8", "Gestión pública y participación: participación política, transversalización, estadísticas y educación para la "
     "igualdad", "T13; T21; T22; T23",
     '("political participation" OR "women candidates" OR "voter education" OR "gender budgeting" OR "gender-responsive '
     'budgeting" OR "gender mainstreaming" OR "time use survey*" OR "gender statistics" OR "gender attitudes") AND (women '
     'OR gender OR school*)',
     '("participación política" OR candidatas OR "presupuesto con enfoque de género" OR "presupuestos sensibles al género" '
     'OR "transversalización del enfoque de género" OR "encuesta de uso del tiempo" OR "estadísticas de género" OR '
     '"igualdad de género") AND (mujeres OR escuela OR docentes)',
     ["Equidad de Género", "Sexismo", "Rol de Género", "Política"], ["J16", "D72", "H61", "H72", "C83"]),
]
DIS_EN = ('("randomi*ed" OR "impact evaluation" OR "quasi-experiment*" OR "difference-in-difference*" OR "regression '
          'discontinuity" OR "propensity score" OR "systematic review" OR "meta-analys*" OR evaluation)')
DIS_ES = '(evaluación OR impacto OR efecto OR efectividad OR "ensayo aleatorio" OR cuasiexperimental OR "revisión sistemática")'
CTX_EN = ('("developing countr*" OR "low- and middle-income" OR "low and middle income" OR "Latin America" OR Peru OR '
          '"systematic review" OR "meta-analys*")')
LILACS_DIS = ["Evaluación de Programas y Proyectos de Salud", "Ensayos Clínicos Controlados Aleatorios como Asunto",
              "Revisión Sistemática", "Resultado del Tratamiento"]
GRUPOS_SALUD = {"G1", "G2", "G3", "G4", "G5", "G6"}
GRUPOS_ECONOMIA = {"G1", "G6", "G7", "G8"}
LILACS_DISENO = (' OR '.join(f'mh:"{d}"' for d in LILACS_DIS)) + " OR " + DIS_ES[1:-1]
LILACS_CADENAS = {
    "G1": '(mh:"Violencia de Pareja" OR mh:"Violencia Doméstica" OR mh:"Mujeres Maltratadas" OR mh:"Violencia de Género") AND '
          '(mh:"Intervención en la Crisis (Psiquiatría)" OR mh:"Psicoterapia" OR mh:"Servicios de Salud Mental" OR mh:"Policía" '
          'OR mh:"Aplicación de la Ley")',
    "G2": '(mh:"Violencia de Pareja" OR mh:"Violencia de Género") AND (mh:"Participación de la Comunidad" OR mh:"Normas Sociales" '
          'OR mh:"Masculinidad" OR mh:"Medios de Comunicación de Masas" OR mh:"Comunicación en Salud" OR mh:"Instituciones Académicas")',
    "G3": 'mh:"Acoso Sexual" AND (mh:"Lugar de Trabajo" OR trabajo OR laboral)',
    "G4": 'mh:"Delitos Sexuales" AND (mh:"Reincidencia" OR tratamiento OR terapia)',
    "G5": '(mh:"Embarazo en Adolescencia" OR mh:"Anticoncepción" OR mh:"Embarazo no Planeado" OR mh:"Servicios de Salud Reproductiva" '
          'OR mh:"Servicios de Salud para Adolescentes" OR mh:"Educación Sexual") AND (mh:"Adolescente" OR adolescentes OR escolares)',
    "G6": '(mh:"Cuidado del Niño" OR mh:"Guarderías Infantiles" OR mh:"Lactancia Materna" OR mh:"Padre" OR mh:"Conducta Paterna" '
          'OR mh:"Rol de Género") AND (mh:"Mujeres Trabajadoras" OR mh:"Empleo" OR mh:"Trabajo" OR empleo OR "participación laboral")',
}
# JSTOR rechaza las consultas de más de unos doce términos («We could not complete your search because it is too long»,
# comprobado el 2026-11-01: nueve términos en Abstract más tres en All fields funcionan; más de eso no). Tampoco admite
# comodines dentro de frases ni en medio de palabra. Por eso cada grupo lleva una versión compacta de sus términos en
# inglés (hasta nueve en la fila Abstract) y un bloque de diseño de tres términos en All fields.
JSTOR_CADENAS = {
    "G1": '("intimate partner violence" OR "domestic violence" OR "violence against women") AND (advocacy OR '
          '"support services" OR psychotherapy OR "mental health" OR police OR training)',
    "G6": '(childcare OR "child care" OR daycare OR breastfeeding OR housework OR fathers) AND (employment OR '
          '"labor supply" OR work*)',
    "G7": '(women OR female) AND ("vocational training" OR "business training" OR microcredit OR microfinance OR '
          '"land titling" OR "property rights" OR "conditional cash transfer")',
    "G8": '("political participation" OR "women candidates" OR "gender budgeting" OR "gender mainstreaming" OR '
          '"time use survey" OR "gender statistics" OR "gender attitudes") AND (women OR gender)',
}
DIS_JSTOR = '(randomized OR "impact evaluation" OR evaluation)'

BASES = {
    "SciELO": ("https://search.scielo.org/", "Todas las colecciones; tipo de literatura: artículo de investigación y "
               "artículo de revisión; sin límite de año. Exportar en CSV o RIS."),
    "LILACS": ("https://pesquisa.bvsalud.org/portal/", "Base de datos: LILACS; sin límite de año. Exportar en RIS o CSV. "
               "Confirmar cada descriptor en DeCS (decs.bvsalud.org) antes de ejecutar."),
    "3ie DEP": ("https://developmentevidence.3ieimpact.org/ (API GraphQL en api.developmentevidence.3ieimpact.org/graphql, "
                "ejecutada por rev_evi_busqueda_3ie_dep.py)", "Tipo de producto: evaluaciones de impacto (ier) y revisiones "
                "sistemáticas (srr); sin límite de año; orden por relevancia; primeros 30 registros."),
    "Scopus": ("https://www.scopus.com/ (requiere suscripción)", "Tipos de documento: artículo y revisión (LIMIT-TO "
               "DOCTYPE ar y re); sin límite de año. Exportar en CSV o RIS con resumen."),
    "Web of Science": ("https://www.webofscience.com/ (requiere suscripción)", "Colección principal; tipos de documento: "
                       "Article y Review; sin límite de año. Exportar en texto delimitado o RIS con resumen."),
    "EconLit": ("EconLit en EBSCOhost o ProQuest (requiere suscripción)", "Tipos de documento: artículo de revista y "
                "documento de trabajo; sin límite de año. El campo de clasificación JEL se llama CC en EBSCOhost y "
                "cc() en ProQuest [POR VERIFICAR según la interfaz disponible]. Exportar en RIS."),
    "JSTOR": ("https://www.jstor.org/action/showAdvancedSearch (acceso institucional)", "Búsqueda avanzada con dos filas unidas "
              "por AND: la primera en el campo Abstract y la segunda en All fields; tipo de ítem: Articles; disciplina: "
              "Economics; acceso: Everything; sin límite de año; orden por relevancia. Exportar en RIS."),
}


def cadena(base, g):
    _, _, _, en, es, decs, jel = g
    if base == "SciELO":
        return f"({es}) AND {DIS_ES}"
    if base == "LILACS":
        return f'{LILACS_CADENAS[g[0]]} AND ({LILACS_DISENO}) AND db:("LILACS")'
    if base == "3ie DEP":
        return en
    if base == "Scopus":
        return f"TITLE-ABS-KEY({en}) AND TITLE-ABS-KEY{DIS_EN} AND TITLE-ABS-KEY{CTX_EN}"
    if base == "Web of Science":
        return f"TS=({en}) AND TS={DIS_EN} AND TS={CTX_EN}"
    if base == "EconLit":
        return f"({' OR '.join(jel)}) [campo JEL] AND ({en}) AND {DIS_EN}"
    if base == "JSTOR":
        return f"[Abstract] {JSTOR_CADENAS[g[0]]} AND [All fields] {DIS_JSTOR}"


COLS = ["id", "grupo", "tipos", "base", "interfaz", "cadena", "descriptores", "filtros", "fecha_ejecucion",
        "total_resultados", "archivo_exportado", "ejecutado_por", "observaciones"]
MANUALES = ["fecha_ejecucion", "total_resultados", "archivo_exportado", "ejecutado_por", "observaciones"]

if __name__ == "__main__":
    DEST.mkdir(parents=True, exist_ok=True)
    ruta = DEST / "estrategias_manuales.csv"
    previas = {}
    if ruta.exists():
        previas = {r["id"]: r for r in csv.DictReader(open(ruta, encoding="utf-8"))}
    filas, i = [], 0
    for base, (interfaz, filtros) in BASES.items():
        for g in GRUPOS:
            if base == "LILACS" and g[0] not in GRUPOS_SALUD or base in ("EconLit", "JSTOR") and g[0] not in GRUPOS_ECONOMIA:
                continue
            i += 1
            eid = ("DEP" if base == "3ie DEP" else "MAN") + f"-{i:02d}"  # el número corre entre todas las bases
            desc = ("DeCS: " + "; ".join(g[5] + LILACS_DIS)) if base == "LILACS" else (
                "JEL: " + "; ".join(g[6])) if base == "EconLit" else ""
            fila = {"id": eid, "grupo": f"{g[0]}: {g[1]}", "tipos": g[2], "base": base, "interfaz": interfaz,
                    "cadena": cadena(base, g), "descriptores": desc, "filtros": filtros}
            fila.update({c: previas.get(eid, {}).get(c, "") for c in MANUALES})
            filas.append(fila)
    with open(ruta, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=COLS, quoting=csv.QUOTE_ALL)
        w.writeheader()
        w.writerows(filas)
    print(len(filas), "estrategias manuales:", {b: sum(f["base"] == b for f in filas) for b in BASES})
