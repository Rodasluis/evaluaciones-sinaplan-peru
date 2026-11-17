"""REV_EVI: matriz de revisión de evidencia y plausibilidad causal de los servicios (nota metodológica v15, numeral 5.1 y
Anexo 6).

Lee los insumos de 04_analisis/rev_evi_insumos/ (rev_evi_insumos.py y rev_evi_busqueda_academica.py) y servicios.csv, y genera:
  - 04_analisis/REV_EVI_matriz_evidencia.xlsx: matriz por tipo de intervención y causa, matriz de evidencia en el
    formato del Anexo 6, plausibilidad por servicio, flujo de la búsqueda por tipo, estrategias académicas con sus
    descriptores, registros recuperados y seleccionados, tipología, búsquedas web, cribado, referencias y criterios;
  - 04_analisis/REV_EVI_plausibilidad_por_servicio.csv, que lee rev_ali_coherencia_alineamiento.py para la columna de
    plausibilidad de REV_ALI.1.

Criterios de inclusión (Anexo 6, ajustado el 2026-10-31): sin límite de año; las evaluaciones primarias (experimentales y
cuasi experimentales) deben ser del Perú o de países de ingreso bajo o medio; las revisiones y los estudios descriptivos se
admiten por su alcance. La solidez se calcula también con el criterio anterior (2010 a 2026, Perú o países de ingreso
medio), solo como sensibilidad.

Solidez (Anexo 6), por tipo de intervención y con las fuentes incluidas:
  - sólida: al menos una revisión sistemática con resultado favorable cuya evidencia no sea de calidad baja, o al menos
    dos evaluaciones con diseño explícito, todas con resultado favorable;
  - moderada: evaluaciones aisladas con diseño explícito, o revisiones sistemáticas con resultados mixtos o de calidad
    baja;
  - limitada: solo estudios descriptivos, revisiones narrativas o evidencia de contextos poco comparables (países de
    ingreso alto, o revisiones cuyo resumen no informa el ámbito);
  - sin evidencia identificada.

Uso (desde la raíz del repositorio, después de rev_evi_busqueda_academica.py y rev_evi_insumos.py):
    uv run --no-project --with openpyxl --with pandas python src/analisis_pnig/rev_evi_evidencia.py
"""
import os

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

RAIZ = f"{EVAL}/04_analisis"
INSUMOS = f"{RAIZ}/rev_evi_insumos"
SALIDA = f"{RAIZ}/REV_EVI_matriz_evidencia.xlsx"
PLAUSIBILIDAD = f"{RAIZ}/REV_EVI_plausibilidad_por_servicio.csv"
ANIO_MIN_ANTERIOR = 2010
DISENOS = {"Revisión sistemática", "Revisión narrativa", "Experimental", "Cuasi experimental", "Descriptivo"}
PRIMARIOS = {"Experimental", "Cuasi experimental"}
DIRECCIONES = {"Favorable", "Mixto", "Nulo", "Desfavorable", "Sin estimación"}
SOLIDEZ = ["Sólida", "Moderada", "Limitada", "Sin evidencia identificada"]

leer = lambda n: pd.read_csv(f"{INSUMOS}/{n}", dtype=str).fillna("")
tip, bus, cri, ev = leer("tipologia.csv"), leer("busquedas.csv"), leer("cribado.csv"), leer("evidencia.csv")
est, desc, res, sel = (leer("estrategias_academicas.csv"), leer("descriptores.csv"), leer("resultados_academicos.csv"),
                       leer("seleccion_academica.csv"))
serv = pd.read_csv("src/formulario_fie03/datos/servicios.csv", dtype=str)
man = leer("busqueda_manual/estrategias_manuales.csv")
dep = leer("busqueda_manual/resultados_dep.csv")  # registros del 3ie DEP (estrategias DEP-15 a DEP-22, por grupo de tipos)
# Exportaciones de la búsqueda manual (rev_evi_importar_manual.py); vacío hasta que se importe la primera
manual = (leer("busqueda_manual/resultados_manual.csv")
          if os.path.exists(f"{INSUMOS}/busqueda_manual/resultados_manual.csv") else dep.iloc[0:0].copy())

# Controles
lista = lambda s: [x.strip() for x in s.split(";") if x.strip()]
asignados = [c for s in tip.servicios for c in lista(s)]
assert sorted(asignados) == sorted(serv.codigo_corto), "la tipología no cubre los 52 servicios exactamente una vez"
assert tip.tipo.is_unique and ev.id.is_unique
for df, col in ((ev, "tipos"), (bus, "tipos"), (cri, "tipos"), (est, "tipo"), (sel, "tipos")):
    fuera = {t for s in df[col] for t in lista(s)} - set(tip.tipo)
    assert not fuera, f"tipos inexistentes en {col}: {fuera}"
assert set(ev.diseno) <= DISENOS, set(ev.diseno) - DISENOS
assert set(ev.direccion) <= DIRECCIONES, set(ev.direccion) - DIRECCIONES
assert (ev.enlace != "").all() and (ev.verificacion != "").all(), "toda fuente necesita enlace y verificación"
res["registro"] = res.estrategia + "|" + res.clave
dep["registro"] = dep.estrategia + "|" + dep.clave
manual["registro"] = manual.estrategia + "|" + manual.clave
todos = pd.concat([res] + [d.rename(columns={"tipos": "tipo"})[res.columns] for d in (dep, manual)], ignore_index=True)
assert set(sel.registro) <= set(todos.registro), set(sel.registro) - set(todos.registro)
assert set(dep.estrategia) <= set(man[man.base == "3ie DEP"].id), "registro del DEP sin estrategia"
assert set(manual.estrategia) <= set(man[man.base != "3ie DEP"].id), "registro manual sin estrategia"
assert set(sel.evidencia[sel.evidencia != ""]) <= set(ev.id), "selección con evidencia inexistente"
assert (desc.verificacion.str.contains("Descriptor|Término|registros de ERIC")).all(), "descriptor sin verificar"
ev["anio"] = ev.anio.astype(int)


def incluida(r, anterior=False):
    """Criterio del Anexo 6; con anterior=True, el de la v15 previa (2010 a 2026, Perú o ingreso medio)."""
    if anterior and r.anio < ANIO_MIN_ANTERIOR:
        return False
    if r.diseno in PRIMARIOS:
        # «Ingreso bajo y medio»: estudios multipaís que incluyen países de ingreso medio
        return r.ambito.startswith(("Perú", "Ingreso medio", "Ingreso bajo y medio")) or (
            not anterior and r.ambito.startswith("Ingreso bajo ("))
    return True


ev["incluida"] = ev.apply(incluida, axis=1)
ev["incluida_anterior"] = ev.apply(lambda r: incluida(r, True), axis=1)
ev["comparable"] = ~ev.ambito.str.startswith(("Ingreso alto", "No informado"))
larga = ev.assign(tipo=ev.tipos.map(lista)).explode("tipo")


from nucleo.evidencia import plausibilidad, solidez as _solidez  # noqa: E402,F401  (regla del núcleo)


def solidez(g):
    """Regla del Anexo 6 sobre las fuentes incluidas de un tipo (ver encabezado y nucleo.evidencia)."""
    return _solidez(g, PRIMARIOS)


filas = []
for t in tip.itertuples():
    g = larga[larga.tipo == t.tipo]
    nivel, direccion = solidez(g[g.incluida])
    nivel_ant, _ = solidez(g[g.incluida_anterior])
    filas.append({
        "Tipo": t.tipo, "Tipo de intervención": t.nombre, "Servicios": t.servicios, "Causa o efecto que atiende": t.nodo_modelo,
        "Resultado esperado": t.resultado_esperado,
        "Fuentes incluidas": ", ".join(g[g.incluida].id) or None,
        "Fuentes excluidas por el criterio": ", ".join(g[~g.incluida].id) or None,
        "Solidez": nivel, "Dirección": direccion or None, "Pertinencia del resultado": t.pertinencia,
        "Plausibilidad": plausibilidad(nivel, direccion, t.pertinencia),
        "Solidez con el criterio anterior": nivel_ant,
        "Cambia con el criterio anterior": "Sí" if nivel_ant != nivel else "No",
        "Búsquedas web": ", ".join(bus[bus.tipos.map(lambda s: t.tipo in lista(s))].id),
        "Estrategias académicas": ", ".join(est[est.tipo == t.tipo].id),
        "Estrategias 3ie DEP": ", ".join(man[(man.base == "3ie DEP") & man.tipos.map(lambda s: t.tipo in lista(s))].id),
    })
por_tipo = pd.DataFrame(filas)

# Flujo de la búsqueda académica por tipo (registros recuperados, revisados, seleccionados por título e incluidos)
flujo = []
for t in tip.itertuples():
    e = est[est.tipo == t.tipo]
    r = res[res.tipo == t.tipo]
    s = sel[sel.registro.isin(r.registro)]
    flujo.append({
        "Tipo": t.tipo, "Tipo de intervención": t.nombre, "Estrategias": len(e),
        "Bases": ", ".join(dict.fromkeys(e.base)),
        "Registros recuperados": int(e.total_resultados.astype(int).sum()),
        "Registros revisados (primeros por relevancia)": len(r),
        "Duplicados": int((r.duplicado_de != "").sum()),
        "Excluidos en el cribado de títulos": int(((r.duplicado_de == "") & ~r.registro.isin(sel.registro)).sum()),
        "Seleccionados por título": len(s),
        "Incluidos": int((s.decision == "Incluida").sum()),
        "Identificados, no extraídos": int((s.decision == "Identificada, no extraída").sum()),
        "Ya incluidos por la búsqueda web": int((s.decision.str.startswith("Ya incluida")).sum()),
        "Excluidos tras leer el resumen": int((s.decision.str.startswith("Excluida")).sum()),
    })
flujo = pd.DataFrame(flujo)
total = flujo.drop(columns=["Tipo", "Tipo de intervención", "Bases"]).sum()
flujo = pd.concat([flujo, pd.DataFrame([{"Tipo": "Total", **total.to_dict()}])], ignore_index=True)

# Flujo por estrategia de la búsqueda en el 3ie DEP y de las exportaciones manuales (cada estrategia cubre un grupo de
# tipos). Los registros ya revisados en una estrategia anterior del mismo bloque se cuentan como repetidos y no como
# excluidos en el cribado. Entran las estrategias ejecutadas (con fecha) aunque no hayan devuelto registros.
def flujo_estrategias(df, bases, duplicados):
    filas = []
    for e in man[man.base.isin(bases) & ((man.fecha_ejecucion != "") | man.id.isin(df.estrategia))].itertuples():
        r = df[df.estrategia == e.id]
        repetido = r.clave.isin(df[df.estrategia < e.id].clave) & (r.duplicado_de == "")
        s = sel[sel.registro.isin(r.registro)]
        filas.append({
            "Estrategia": e.id, "Base": e.base, "Grupo": e.grupo, "Tipos": e.tipos,
            "Registros recuperados": int(e.total_resultados or 0),
            "Registros revisados (primeros por relevancia)": len(r),
            duplicados: int((r.duplicado_de != "").sum()),
            "Repetidos de una estrategia anterior": int(repetido.sum()),
            "Excluidos en el cribado de títulos": int(((r.duplicado_de == "") & ~repetido & ~r.registro.isin(sel.registro)).sum()),
            "Seleccionados por título": len(s),
            "Incluidos": int((s.decision == "Incluida").sum()),
            "Identificados, no extraídos": int((s.decision == "Identificada, no extraída").sum()),
            "Ya incluidos": int(s.decision.str.startswith("Ya incluida").sum()),
            "Excluidos tras leer el resumen": int(s.decision.str.startswith("Excluida").sum()),
        })
    filas = pd.DataFrame(filas)
    if filas.empty:
        return filas
    total = filas.drop(columns=["Estrategia", "Base", "Grupo", "Tipos"]).sum()
    return pd.concat([filas, pd.DataFrame([{"Estrategia": "Total", **total.to_dict()}])], ignore_index=True)


flujo_dep = flujo_estrategias(dep, ["3ie DEP"], "Duplicados de la búsqueda académica")
flujo_manual = flujo_estrategias(manual, set(man.base) - {"3ie DEP"}, "Duplicados de la búsqueda académica o del 3ie DEP")
# El libro solo lista las estrategias ejecutadas; las de bases sin acceso (Web of Science y EconLit en la PNIG) quedan en
# estrategias_manuales.csv como opción para otro usuario.
man_ejecutadas = man[man.fecha_ejecucion != ""]
estado_manual = (f"Importadas {manual.estrategia.nunique()} estrategias ({', '.join(sorted(set(manual.base)))}) con "
                 f"{len(manual)} registros." if len(manual) else "Ninguna exportación importada todavía.")

# Matriz de evidencia con el formato del Anexo 6 (una fila por tipo y fuente)
matriz = larga.merge(tip[["tipo", "nombre", "nodo_modelo"]], on="tipo").merge(
    por_tipo[["Tipo", "Solidez"]], left_on="tipo", right_on="Tipo")
matriz = matriz.assign(
    fuente=lambda d: d.id + ": " + d.referencia.str.extract(r"^(.*?\(\d{4}\))")[0].fillna(d.referencia.str[:60]),
    criterio=lambda d: d.incluida.map({True: "Incluida", False: "Excluida (país)"}))
matriz["orden"] = matriz.id.str[1:].astype(int)
matriz = matriz.sort_values(["tipo", "incluida", "orden"], ascending=[True, False, True])[[
    "tipo", "nombre", "nodo_modelo", "fuente", "diseno", "ambito", "calidad", "resultado_principal", "direccion",
    "criterio", "Solidez"]].rename(columns={
        "tipo": "Tipo", "nombre": "Tipo de servicio", "nodo_modelo": "Causa que atiende", "fuente": "Fuente",
        "diseno": "Diseño del estudio", "ambito": "Ámbito", "calidad": "Calidad según la fuente",
        "resultado_principal": "Resultado principal", "direccion": "Dirección", "criterio": "Criterio de inclusión",
        "Solidez": "Solidez del tipo"})

# Plausibilidad por servicio (para REV_ALI.1)
por_servicio = (tip.assign(codigo=tip.servicios.map(lista)).explode("codigo")[["codigo", "tipo"]]
                .merge(por_tipo, left_on="tipo", right_on="Tipo")
                .merge(serv[["codigo_corto", "servicio"]], left_on="codigo", right_on="codigo_corto"))
clave = lambda c: tuple(int(i) for i in c.split("."))
por_servicio = por_servicio.sort_values("codigo", key=lambda s: s.map(clave))
por_servicio[["codigo", "tipo", "Tipo de intervención", "Solidez", "Dirección", "Plausibilidad",
              "Solidez con el criterio anterior"]].rename(columns={"codigo": "Código", "tipo": "Tipo"}).to_csv(
    PLAUSIBILIDAD, index=False, encoding="utf-8")

referencias = ev.assign(orden=ev.id.str[1:].astype(int)).sort_values("orden")[[
    "id", "referencia", "enlace", "verificacion", "diseno", "ambito", "anio", "origen", "incluida",
    "incluida_anterior"]].rename(columns={
        "id": "Id", "referencia": "Referencia", "enlace": "Enlace", "verificacion": "Verificación", "diseno": "Diseño",
        "ambito": "Ámbito", "anio": "Año", "origen": "Origen", "incluida": "Incluida (criterio de la nota)",
        "incluida_anterior": "Incluida (criterio anterior)"})
referencias[["Incluida (criterio de la nota)", "Incluida (criterio anterior)"]] = referencias[[
    "Incluida (criterio de la nota)", "Incluida (criterio anterior)"]].replace({True: "Sí", False: "No"})

seleccion = sel.merge(todos[["registro", "base", "clave", "titulo", "anio", "fuente", "enlace"]], on="registro")[[
    "registro", "tipos", "base", "clave", "titulo", "anio", "fuente", "enlace", "decision", "evidencia", "motivo"]]

n_web = (ev.origen == "Búsqueda web").sum() + ev.origen.str.startswith("Búsqueda web;").sum()
n_acad = ev.origen.str.startswith("Búsqueda académica").sum()
n_port = ev.origen.str.contains("evaluaciones peruanas|portales").sum()
n_dep = ev.origen.str.startswith("Búsqueda en el 3ie DEP").sum()
n_dep_ya = int((sel.registro.str.startswith(("DEP-", "MAN-")) & sel.decision.str.startswith("Ya incluida")).sum())
recuperadas = ev.origen.str.contains("recuperada en la búsqueda académica").sum()
bases = ", ".join(f"{b} ({n})" for b, n in est.base.value_counts().items())
criterios = pd.DataFrame([
    ("Protocolo", "Revisión rápida de evidencia por tipo de intervención y por la causa o el efecto del modelo que "
     "atiende (Anexo 6). Dos componentes: una búsqueda en bases académicas con vocabulario controlado, replicable por su "
     "cadena exacta, y una búsqueda complementaria en la web y en repositorios de organismos especializados (Campbell, "
     "Cochrane, 3ie, J-PAL, MEF), con revisión de referencias de las revisiones incluidas.", "Nota metodológica, Anexo 6"),
    ("Tipología", f"{len(tip)} tipos de intervención para los 52 servicios; la causa o el efecto de cada tipo es el del "
     "Anexo 5 de la PNIG, como lo clasifica REV_ALI.1.", "Equipo evaluador"),
    ("Búsqueda académica", f"{len(est)} estrategias ejecutadas el {est.fecha.iloc[0]} en {bases}. PubMed con "
     "descriptores MeSH, ERIC con su tesauro y EconBiz con el Tesauro de Economía STW; Campbell y 3ie mediante la API de "
     "Crossref (ISSN 1891-1803 y prefijo DOI 10.23846); LA Referencia y OpenAlex con términos libres en español; Banco "
     "Mundial (Policy Research Working Papers) con términos libres en inglés. Cada estrategia "
     "registra su cadena exacta, descriptores, términos libres, filtros, fecha, total de resultados y URL de la consulta "
     f"(hoja «Estrategias académicas»); los {len(desc)} descriptores se verificaron en su tesauro (hoja «Descriptores»). "
     "Sin límite de año. OpenAlex se consultó sin clave, dentro del cupo diario gratuito; el script usa la variable "
     "OPENALEX_API_KEY si existe y no la escribe en los registros.", "Equipo evaluador"),
    ("Búsqueda en el 3ie DEP", "Las ocho estrategias DEP-15 a DEP-22 se ejecutaron el 31-10-2026 mediante la API GraphQL "
     "pública del Development Evidence Portal (consulta advancedSearch, src/analisis_pnig/rev_evi_busqueda_3ie_dep.py), con "
     "la cadena exacta registrada, orden por relevancia, filtro a evaluaciones de impacto y revisiones sistemáticas y los "
     "primeros 30 registros por estrategia (hoja «Flujo 3ie DEP»). El cribado y la regla de extracción son los de la "
     "búsqueda académica. Aportó {n_dep} fuentes nuevas; {n_dep_ya} registros correspondían a fuentes ya incluidas. No "
     "devolvió registros sobre tratamiento de agresores sexuales (T10) ni sobre hostigamiento sexual laboral (T04).",
     "Equipo evaluador"),
    ("Búsqueda manual", "Bases consultadas desde el navegador, con las cadenas y los filtros registrados en "
     "rev_evi_insumos/busqueda_manual/estrategias_manuales.csv y el protocolo de esa carpeta: SciELO (MAN-01 a MAN-08, cadena en "
     "español, artículos y artículos de revisión, orden por relevancia), LILACS (MAN-09 a MAN-14, descriptores DeCS comprobados "
     "por su conteo en la base, estructura población AND intervención AND diseño), Scopus (MAN-23 a MAN-30, acceso remoto "
     "institucional de un integrante del equipo, 30 registros por estrategia y 40 en MAN-25) y JSTOR (MAN-43 a MAN-46, disciplina "
     "Economics, para los grupos económicos). Las exportaciones se incorporan con src/analisis_pnig/rev_evi_importar_manual.py "
     f"y siguen el mismo cribado (hojas «Resultados manuales» y «Flujo manual»). {estado_manual}", "Equipo evaluador"),
    ("Cribado", "En dos etapas. Primero, títulos de los primeros registros por relevancia de cada estrategia (30 en "
     "PubMed y EconBiz; 15 en ERIC, LA Referencia, Banco Mundial y OpenAlex; 10 en Campbell y 3ie), sin contar "
     "duplicados. La profundidad se fijó con el aporte por tramo: en PubMed y EconBiz los registros 11 a 15 aportaron "
     "tantas fuentes como los primeros cinco; en Campbell, 3ie y LA Referencia el aporte cae después del registro 5. "
     "Después, lectura del resumen de los seleccionados, con decisión y motivo (hoja «Selección académica»). La hoja "
     "«Flujo por tipo» resume los registros recuperados, revisados, seleccionados e incluidos.", "Equipo evaluador"),
    ("Extracción", "En la ampliación del 31-10-2026 la extracción se orienta a la decisión: se extraen y verifican las "
     "fuentes que pueden cambiar la solidez o la dirección de un tipo (todas en los tipos con evidencia limitada o de "
     "dirección favorable, y las revisiones sistemáticas en los de evidencia moderada) y toda evaluación realizada en el "
     "Perú. Las demás fuentes pertinentes quedan como «Identificada, no extraída», porque una fuente más no cambia la "
     "clasificación de su tipo.", "Equipo evaluador"),
    ("Búsqueda web y portales", f"{len(bus)} búsquedas registradas el 31-10-2026: en un buscador web general (algunas "
     "restringidas a Campbell, 3ie y J-PAL), en el portal de evaluaciones de J-PAL, en evaluaciones peruanas (MEF, "
     "MIDIS, CIES, GRADE) y en What Works y SVRI. La búsqueda web inicial aportó "
     f"{n_web} fuentes; la búsqueda académica recuperó {recuperadas} de ellas y aportó {n_acad} nuevas; la búsqueda de "
     f"evaluaciones peruanas y de portales aportó {n_port}.", "Equipo evaluador"),
    ("Verificación", "Cada fuente se verificó antes de registrarse: título, autores, año y fuente en Crossref, PubMed o "
     "la página de la revista o del repositorio; el resultado se toma del resumen. Lo que no se pudo verificar se dejó "
     "fuera con su motivo.", "Nota metodológica, Anexo 6"),
    ("Inclusión", "Sin límite de año. Las evaluaciones primarias (experimentales y cuasi experimentales) deben ser del "
     "Perú o de países de ingreso bajo o medio; las revisiones y los estudios descriptivos se admiten por su alcance. "
     "Se excluyen opiniones, tesis, resúmenes de congreso y documentos sin método declarado.",
     "Nota metodológica, Anexo 6 (ajuste aprobado el 31-10-2026)"),
    ("Criterio anterior", f"{ANIO_MIN_ANTERIOR} a 2026, Perú o países de ingreso medio. Se informa como sensibilidad: "
     "excluye evaluaciones peruanas previas a 2010 (Field, 2007; Ñopo et al., 2008; Perova y Vakis, 2009) y ensayos en "
     "países de ingreso bajo.", "Equipo evaluador"),
    ("Comparabilidad", "La evidencia de países de ingreso alto y la de revisiones cuyo resumen no informa el ámbito se "
     "tratan como de contexto poco comparable: cuentan para la solidez limitada, no para la moderada o la sólida.",
     "Equipo evaluador"),
    ("Solidez", "Sólida: al menos una revisión sistemática favorable cuya evidencia no sea de calidad baja, o dos "
     "evaluaciones con diseño explícito, todas favorables. Moderada: evaluaciones aisladas, o revisiones con resultados "
     "mixtos o de calidad baja. Limitada: solo estudios descriptivos, revisiones narrativas o contextos poco comparables. "
     "Sin evidencia identificada.", "Nota metodológica, Anexo 6; precisión del equipo evaluador"),
    ("Plausibilidad", "Solidez y dirección de la evidencia del tipo de intervención al que pertenece el servicio. "
     "«Evidencia parcial» indica que la evidencia mide un resultado distinto del que persigue el servicio o cubre solo "
     "una parte de su población (por ejemplo, lactancia en lugar de permanencia laboral, o cuidado infantil y no de "
     "personas adultas mayores). Dirección mixta: resultados favorables junto con resultados nulos o no concluyentes.",
     "Equipo evaluador"),
    ("Limitaciones", "Revisión rápida: cribado por un solo revisor; se revisan los primeros registros por relevancia de "
     "cada estrategia y no el total recuperado; los resultados se toman del resumen y no del texto completo; en la "
     "ampliación, la extracción se limita a las fuentes que pueden cambiar la clasificación; "
     + "el texto completo de una revisión del 3ie DEP (ODI, 2016) no fue accesible. Repetir la búsqueda en otra fecha "
     "puede dar totales distintos.", "Equipo evaluador"),
], columns=["Elemento", "Criterio", "Plano"])

hojas = [("Matriz por tipo", por_tipo), ("Matriz de evidencia", matriz),
         ("Plausibilidad por servicio", por_servicio[["codigo", "servicio", "tipo", "Tipo de intervención", "Solidez",
                                                      "Dirección", "Plausibilidad", "Solidez con el criterio anterior"]]
          .rename(columns={"codigo": "Código", "servicio": "Servicio", "tipo": "Tipo"})),
         ("Flujo por tipo", flujo), ("Estrategias académicas", est), ("Descriptores", desc),
         ("Selección académica", seleccion), ("Resultados académicos", res.drop(columns="registro")),
         ("Tipología", tip), ("Búsquedas web y portales", bus), ("Cribado web y portales", cri),
         ("Estrategias manuales", man_ejecutadas), ("Resultados 3ie DEP", dep), ("Flujo 3ie DEP", flujo_dep),
         *([("Resultados manuales", manual), ("Flujo manual", flujo_manual)] if len(manual) else []),
         ("Referencias", referencias),
         ("Criterios", criterios)]
ANCHOS = {80: {"Resultado principal", "Referencia", "Criterio", "resultado_principal", "cadena", "titulo"},
          40: {"Servicio", "Fuente", "Verificación", "Tipo de intervención", "Tipo de servicio", "Causa que atiende",
               "Causa o efecto que atiende", "Plausibilidad", "motivo", "referencia", "Enlace", "nombre", "intervencion",
               "descriptores", "terminos_libres", "filtros", "interfaz", "url_consulta", "verificacion", "fuente",
               "enlace", "Origen", "Bases", "tipo_documento", "uri"}}
with pd.ExcelWriter(SALIDA, engine="openpyxl") as xw:
    for nombre, df in hojas:
        df.to_excel(xw, sheet_name=nombre, index=False)
        ws = xw.sheets[nombre]
        ws.freeze_panes = "B2"
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F3864")
            c.alignment = Alignment(wrap_text=True, vertical="top")
        for col in ws.columns:
            titulo = str(col[0].value)
            ancho = next((a for a, nombres in ANCHOS.items() if titulo in nombres), 16)
            ws.column_dimensions[col[0].column_letter].width = ancho
            if ancho >= 40:
                for c in col[1:]:
                    c.alignment = Alignment(wrap_text=True, vertical="top")
print(por_tipo[["Tipo", "Solidez", "Dirección", "Solidez con el criterio anterior"]].to_string(index=False))
print(por_tipo.Solidez.value_counts().reindex(SOLIDEZ, fill_value=0).to_dict())
print(flujo.iloc[-1].to_dict())
