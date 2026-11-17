"""REV_EVI: importación de las exportaciones de la búsqueda manual (SciELO, LILACS, Scopus, JSTOR, Web of Science y EconLit).

Estas bases no admiten consultas automatizadas o requieren suscripción (busqueda_manual/protocolo_busqueda_manual.md):
cada estrategia MAN-01 a MAN-14 y MAN-23 a MAN-46 se ejecuta a mano en la interfaz de la base y se exporta en RIS, CSV o texto
delimitado. Este script incorpora esas exportaciones al mismo flujo que la búsqueda académica y la del 3ie DEP:

  - lee evaluaciones/pnig_2026/03_datos/raw/rev_evi_busqueda_manual/<base>/<id>.<ext>, con <base> en scielo, lilacs,
    scopus, jstor, wos o econlit y <id> el de la estrategia (por ejemplo scopus/MAN-23.csv); un archivo por estrategia;
  - reconoce el formato por el contenido: CSV de Scopus, texto delimitado por tabulaciones de Web of Science (etiquetas
    PT, TI, PY, DI, UT), RIS (EBSCOhost, SciELO, BVS, Scopus o Web of Science) y, como respaldo, cualquier CSV con
    columnas de título, autores, año, fuente, DOI, enlace y resumen (SciELO y BVS);
  - conserva los primeros 30 registros en el orden de la exportación (orden por relevancia, según el protocolo). Si junto
    al archivo hay <id>.orden.txt (un título por línea, copiado de la lista en pantalla ordenada por relevancia), los
    registros se reordenan según esa lista y los que no figuran en ella quedan fuera: Scopus no exporta en el orden
    de relevancia que muestra, así que se exporta el conjunto completo y se recorta con la lista en pantalla;
  - en SciELO (COLAPSAR_REPETIDOS) descarta las versiones repetidas del mismo artículo que el buscador lista por cada
    colección o idioma (mismo PID con otro sufijo, mismo DOI o mismo título), antes de contar los 30;
  - incorpora las bases en el orden en que se consultaron (ORDEN_EJECUCION), de modo que un registro repetido entre
    bases cuente como duplicado de la consultada primero y las decisiones ya tomadas no cambien;
  - marca los duplicados frente a la búsqueda académica y el 3ie DEP (DOI o título normalizado); un registro repetido
    por otra estrategia manual no se marca y se decide en su primera aparición, igual que en las otras búsquedas;
  - escribe busqueda_manual/resultados_manual.csv y busqueda_manual/resumenes_manual.json, completa archivo_exportado en
    estrategias_manuales.csv y registra cada archivo con su hash en docs/fuentes/descargas_rev06-manual.tsv.

Los archivos exportados son datos no confiables: se leen como texto y nunca se ejecutan. Las columnas fecha_ejecucion,
total_resultados y ejecutado_por las completa a mano quien ejecuta la búsqueda; el script avisa cuando faltan. El script
no usa ni guarda credenciales: la cuenta de acceso a cada base queda fuera del repositorio.

Uso (desde la raíz del repositorio):
    python src/analisis_pnig/rev_evi_importar_manual.py
    python src/analisis_pnig/rev_evi_importar_manual.py --raiz <carpeta> --salida <carpeta> --registro <tsv>   (pruebas)
Después se añaden las decisiones del cribado en rev_evi_insumos.py (SELECCION_MANUAL) y se corre rev_evi_evidencia.py.
"""
import argparse
import csv
import hashlib
import io
import json
import re
import sys
import unicodedata
from datetime import date, datetime
from pathlib import Path
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

INSUMOS = Path(f"{EVAL}/04_analisis/rev_evi_insumos")
MANUAL = INSUMOS / "busqueda_manual"
RAIZ = Path(f"{EVAL}/03_datos/raw/rev_evi_busqueda_manual")
REGISTRO = Path("docs/fuentes/descargas_rev06-manual.tsv")
REVISADOS = 30
COLAPSAR_REPETIDOS = {"scielo"}  # bases cuya lista repite el mismo artículo por colección o idioma
ORDEN_EJECUCION = ("scopus", "scielo", "lilacs", "wos", "econlit", "jstor")  # orden en que se consultaron las bases
# Estrategias con una profundidad distinta de REVISADOS, aprobada por el usuario (MAN-25: 40, el 2026-11-01, para alcanzar dos
# revisiones sobre hostigamiento sexual laboral que quedaban en los puestos 34 y 39).
PROFUNDIDAD = {"MAN-25": 40}
BASES = {"scielo": "SciELO", "lilacs": "LILACS", "scopus": "Scopus", "wos": "Web of Science", "econlit": "EconLit",
         "jstor": "JSTOR"}
PREFIJO = {"scielo": "SCIELO", "lilacs": "LILACS", "scopus": "SCOPUS", "wos": "WOS", "econlit": "ECONLIT", "jstor": "JSTOR"}
EXTENSIONES = {".csv", ".tsv", ".txt", ".ris"}
COLUMNAS = {  # nombres equivalentes en los CSV de las interfaces (en minúsculas y sin acentos)
    "titulo": ["title", "article title", "titulo", "document title"],
    "autores": ["authors", "author full names", "autores", "autor", "author"],
    "anio": ["year", "publication year", "anio", "ano", "date", "fecha"],
    "fuente": ["source title", "revista", "journal", "publication title", "fonte", "source"],
    "doi": ["doi"],
    "enlace": ["link", "url", "enlace", "permalink"],
    "resumen": ["abstract", "resumen", "resumo"],
    "tipo_documento": ["document type", "tipo de documento", "tipo", "type"],
    "identificador": ["eid", "ut (unique wos id)", "ut", "id", "accession number", "pid", "identifier", "an"],
    "palabras": ["author keywords", "index keywords", "keywords", "palabras clave", "descriptores"],
}
WOS = {"titulo": ["TI"], "autores": ["AU"], "anio": ["PY"], "fuente": ["SO"], "doi": ["DI"], "resumen": ["AB"],
       "tipo_documento": ["DT"], "identificador": ["UT"], "palabras": ["DE", "ID"], "enlace": []}
RIS = {"titulo": ["TI", "T1"], "autores": ["AU", "A1"], "anio": ["PY", "Y1"], "fuente": ["JO", "JF", "T2", "J2"],
       "doi": ["DO"], "enlace": ["UR"], "resumen": ["AB", "N2"], "tipo_documento": ["TY"],
       "identificador": ["ID", "AN", "UT", "M1"], "palabras": ["KW"]}
PATRONES_ID = [r"eid=(2-s2\.0-\d+)", r"pid=(S\d{4}-\d{3}[\dX]\d{10})", r"[?&]AN=([^&\s]+)", r"\b((?:lil|biblio|mdl)-\d+)\b",
               r"\b(WOS:\d{15})\b", r"jstor\.org/stable/(10\.\d{4,9}/[^\s?#]+|\d+)"]  # JSTOR: número estable o DOI de la URL estable

normalizar = lambda s: unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode().lower().strip()
norm_titulo = lambda s: re.sub(r"[^a-z0-9]", "", normalizar(s))[:80]
limpiar_doi = lambda d: re.sub(r"^(https?://(dx\.)?doi\.org/|doi:\s*)", "", d.strip(), flags=re.I)


def leer_texto(ruta):
    datos = ruta.read_bytes()
    for codificacion in ("utf-8-sig", "utf-16", "cp1252"):
        try:
            return datos.decode(codificacion), datos
        except UnicodeDecodeError:
            continue
    return datos.decode("latin-1"), datos


def leer_ris(texto):
    """Registros RIS como diccionarios etiqueta -> lista de valores; las líneas sin etiqueta continúan el campo previo."""
    registros, actual, ultimo = [], {}, None
    for linea in texto.splitlines():
        m = re.match(r"^([A-Z][A-Z0-9])  -\s?(.*)$", linea)
        if not m:
            if ultimo and linea.strip():
                actual[ultimo][-1] = (actual[ultimo][-1] + " " + linea.strip()).strip()
            continue
        etiqueta, valor = m.group(1), m.group(2).strip()
        if etiqueta == "ER":
            if actual:
                registros.append(actual)
            actual, ultimo = {}, None
            continue
        actual.setdefault(etiqueta, []).append(valor)
        ultimo = etiqueta
    if actual:
        registros.append(actual)
    return registros


def leer_tabla(texto):
    """CSV o texto delimitado (coma, punto y coma o tabulación) como lista de diccionarios."""
    primera = texto.splitlines()[0] if texto.strip() else ""
    delimitador = "\t" if primera.count("\t") >= 3 else max(",;", key=primera.count)
    filas = list(csv.DictReader(io.StringIO(texto), delimiter=delimitador))
    return [{(k or "").strip(): (v or "").strip() for k, v in f.items() if k is not None} for f in filas]


def campo(fila, nombres, mapa, unir="; "):
    """Primer valor no vacío de las columnas equivalentes; para palabras clave, todas unidas."""
    valores = []
    for nombre in mapa.get(nombres, []):
        v = fila.get(nombre, "")
        v = unir.join(v) if isinstance(v, list) else v
        if v:
            valores.append(v)
    if nombres == "palabras" and valores:
        return unir.join(valores)
    return valores[0] if valores else ""


def normalizar_tabla(filas):
    """Columnas de un CSV en minúsculas y sin acentos, para buscarlas con COLUMNAS."""
    return [{normalizar(k): v for k, v in f.items()} for f in filas]


def formato_de(texto, ruta):
    if re.search(r"^TY  - ", texto, flags=re.M):
        return "RIS"
    primera = texto.splitlines()[0] if texto.strip() else ""
    etiquetas = [c.strip().strip('"') for c in primera.split("\t")]
    if primera.count("\t") >= 3 and {"TI", "PY", "UT"} <= set(etiquetas):
        return "Web of Science (texto delimitado)"
    if "EID" in [c.strip().strip('"') for c in re.split(r"[,;\t]", primera)]:
        return "Scopus (CSV)"
    if ruta.suffix.lower() in {".csv", ".tsv", ".txt"}:
        return "CSV"
    raise ValueError(f"formato no reconocido: {ruta}")


def registros_de(texto, ruta):
    """Lista de registros con los campos comunes, en el orden de la exportación."""
    formato = formato_de(texto, ruta)
    if formato == "RIS":
        crudos, mapa = leer_ris(texto), RIS
    elif formato.startswith("Web of Science"):
        crudos, mapa = leer_tabla(texto), WOS
    else:
        crudos, mapa = normalizar_tabla(leer_tabla(texto)), COLUMNAS
    salida = []
    for r in crudos:
        reg = {c: campo(r, c, mapa) for c in COLUMNAS}
        if not reg["titulo"]:
            continue
        # JSTOR exporta el título en TI y el subtítulo en T1: se unen para que el título quede completo.
        if formato == "RIS" and r.get("TI") and r.get("T1") and r["T1"][0] != r["TI"][0]:
            reg["titulo"] = f"{r['TI'][0].rstrip(':?. ')}{'?' if r['TI'][0].rstrip().endswith('?') else ':'} {r['T1'][0]}"
        reg["anio"] = (re.search(r"\d{4}", reg["anio"]) or [""])[0] if reg["anio"] else ""
        reg["doi"] = limpiar_doi(reg["doi"])
        salida.append(reg)
    return formato, salida


def clave_de(reg, base):
    """Identificador estable del registro: el de la base; si falta, el DOI o un hash del título."""
    prefijo = PREFIJO[base]
    ident = reg["identificador"].strip()
    if not ident:
        for patron in PATRONES_ID:
            m = re.search(patron, reg["enlace"], flags=re.I)
            if m:
                ident = m.group(1)
                break
    if ident:
        ident = re.sub(rf"^{prefijo}:", "", ident, flags=re.I)
        return f"{prefijo}:{ident}"
    if reg["doi"]:
        return f"{prefijo}:doi:{reg['doi']}"
    return f"{prefijo}:h:{hashlib.sha1(norm_titulo(reg['titulo']).encode()).hexdigest()[:10]}"


def enlace_de(reg, base, clave):
    if reg["enlace"]:
        return reg["enlace"].split(";")[0].strip()
    if base == "wos" and clave.startswith("WOS:"):
        return f"https://www.webofscience.com/wos/woscc/full-record/{clave}"
    if base == "scopus" and clave.startswith("SCOPUS:2-s2.0-"):
        return f"https://www.scopus.com/record/display.uri?eid={clave[7:]}&origin=resultslist"
    return f"https://doi.org/{reg['doi']}" if reg["doi"] else ""


def reordenar(regs, orden):
    """Ordena los registros según la lista de títulos en pantalla (uno por línea); los que no figuran quedan fuera."""
    titulos = [t.strip() for t in orden.read_text(encoding="utf-8").splitlines() if t.strip()]
    completo = lambda s: re.sub(r"[^a-z0-9]", "", normalizar(s))
    pares = [(completo(r["titulo"]), r) for r in regs]
    ordenados, faltantes = [], []
    for t in titulos:
        nt = completo(t)
        # Coincidencia exacta o, si la exportación trae el título bilingüe («título | título traducido»), por prefijo
        r = next((r for k, r in pares if k == nt), None) or next((r for k, r in pares if k.startswith(nt)), None)
        (ordenados if r else faltantes).append(r or t)
    return ordenados, faltantes


def main(raiz, salida, registro, estrategias_csv):
    estrategias = list(csv.DictReader(open(estrategias_csv, encoding="utf-8")))
    por_id = {e["id"]: e for e in estrategias}
    vistos = {}
    for archivo in ("resultados_academicos.csv", "busqueda_manual/resultados_dep.csv"):
        ruta = INSUMOS / archivo
        if ruta.exists():
            for a in csv.DictReader(open(ruta, encoding="utf-8")):
                vistos.setdefault(a["doi"].lower() or norm_titulo(a["titulo"]), a["clave"])
                vistos.setdefault(norm_titulo(a["titulo"]), a["clave"])
    # Las bases se incorporan en el orden en que se ejecutaron (ORDEN_EJECUCION), de modo que un registro repetido
    # entre bases cuente como duplicado de la base consultada primero y las decisiones ya tomadas no cambien.
    archivos = sorted((p for p in raiz.glob("*/*") if p.is_file() and p.suffix.lower() in EXTENSIONES
                       and ".orden." not in p.name.lower()),
                      key=lambda p: (ORDEN_EJECUCION.index(p.parent.name.lower()) if p.parent.name.lower() in ORDEN_EJECUCION
                                     else len(ORDEN_EJECUCION), p.stem.upper()))
    if not archivos:
        print(f"sin exportaciones en {raiz.as_posix()}; nada que importar")
        return 0
    resultados, resumenes, descargas, avisos = [], {}, [], []
    for ruta in archivos:
        base, ident = ruta.parent.name.lower(), ruta.stem.upper()
        if base not in BASES:
            avisos.append(f"{ruta.as_posix()}: carpeta desconocida (usar {', '.join(BASES)})")
            continue
        e = por_id.get(ident)
        if e is None or e["base"] != BASES[base]:
            avisos.append(f"{ruta.as_posix()}: la estrategia {ident} no existe o no es de {BASES[base]}")
            continue
        if sum(1 for p in archivos if p.parent == ruta.parent and p.stem.upper() == ident) > 1:
            raise SystemExit(f"más de un archivo para {ident} en {ruta.parent.as_posix()}: dejar solo uno")
        texto, datos = leer_texto(ruta)
        formato, regs = registros_de(texto, ruta)
        if not regs:
            avisos.append(f"{ruta.as_posix()}: sin registros legibles")
            continue
        fecha = e["fecha_ejecucion"] or datetime.fromtimestamp(ruta.stat().st_mtime).date().isoformat()
        # SciELO lista el mismo artículo una vez por colección o idioma (PID distinguido solo por el sufijo de colección):
        # se conserva la primera aparición y las repetidas no cuentan entre los 30. En las demás bases cada registro es único.
        if base in COLAPSAR_REPETIDOS:
            unicos, llaves_archivo, repetidos = [], set(), 0
            for r in regs:
                pid = re.sub(r"^(S\d{4}-\d{3}[\dX]\d{12})-[a-z]{3}$", r"\1", r["identificador"].strip(), flags=re.I).lower()
                llaves = {k for k in (r["doi"].lower(), norm_titulo(r["titulo"]), pid) if k}
                if llaves & llaves_archivo:
                    repetidos += 1
                    continue
                llaves_archivo |= llaves
                unicos.append(r)
            if repetidos:
                formato += f", {repetidos} versiones repetidas del mismo artículo descartadas"
            regs = unicos
        orden = ruta.with_name(f"{ruta.stem}.orden.txt")
        if orden.exists():
            exportados = len(regs)
            regs, faltantes = reordenar(regs, orden)
            if faltantes:
                avisos.append(f"{ident}: {len(faltantes)} títulos de la lista en pantalla no están en la exportación: "
                              f"{'; '.join(t[:60] for t in faltantes[:3])}")
            formato += f", {exportados} registros reordenados con la lista en pantalla"
            datos_orden = orden.read_bytes()
            descargas.append((f"{e['interfaz'].split(' (')[0]} (orden de la lista en pantalla, {ident})", orden.as_posix(),
                              hashlib.sha256(datos_orden).hexdigest(), fecha,
                              "Títulos copiados de la lista de resultados ordenada por relevancia"))
        revisados = PROFUNDIDAD.get(ident, REVISADOS)
        if not orden.exists() and len(regs) > revisados:
            avisos.append(f"{ident}: la exportación trae {len(regs)} registros; se conservan los primeros {revisados}")
        regs = regs[:revisados]
        for rango, r in enumerate(regs, 1):
            clave = clave_de(r, base)
            norm = norm_titulo(r["titulo"])
            llave = r["doi"].lower() or norm
            dup = vistos.get(llave) or vistos.get(norm, "")
            vistos.setdefault(llave, clave)
            vistos.setdefault(norm, clave)
            resultados.append({
                "estrategia": ident, "tipos": e["tipos"], "base": BASES[base], "rango": rango, "clave": clave,
                "doi": r["doi"], "titulo": r["titulo"], "anio": r["anio"], "fuente": r["fuente"],
                "tipo_documento": r["tipo_documento"], "enlace": enlace_de(r, base, clave),
                "duplicado_de": dup if dup != clave else "", "pais": "", "ingreso": "", "formato": formato})
            resumenes[clave] = {"titulo": r["titulo"], "autores": r["autores"], "resumen": r["resumen"],
                                "palabras": r["palabras"]}
        for columna in ("fecha_ejecucion", "total_resultados", "ejecutado_por"):
            if not e[columna]:
                avisos.append(f"{ident}: falta {columna} en estrategias_manuales.csv")
        e["archivo_exportado"] = e["archivo_exportado"] or ruta.as_posix()
        descargas.append((f"{e['interfaz'].split(' (')[0]} (exportación manual, {ident})", ruta.as_posix(),
                          hashlib.sha256(datos).hexdigest(), fecha, f"Exportación de la interfaz ({formato}, metadatos)"))
        print(ident, BASES[base], formato, e["total_resultados"] or "?", len(regs), flush=True)
    if not resultados:
        print("ninguna exportación válida")
        return 1
    salida.mkdir(parents=True, exist_ok=True)
    with open(salida / "resultados_manual.csv", "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(resultados[0]), quoting=csv.QUOTE_ALL)
        w.writeheader()
        w.writerows(resultados)
    (salida / "resumenes_manual.json").write_text(json.dumps(resumenes, ensure_ascii=False, indent=1), encoding="utf-8")
    with open(salida / "estrategias_manuales.csv", "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(estrategias[0]), quoting=csv.QUOTE_ALL)
        w.writeheader()
        w.writerows(estrategias)
    registro.parent.mkdir(parents=True, exist_ok=True)
    with open(registro, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f, delimiter="\t", lineterminator="\n")
        w.writerow(["url_origen", "destino", "sha256", "fecha_descarga", "verificacion_identidad"])
        w.writerows(descargas)
    print(len(resultados), "registros;", sum(1 for r in resultados if r["duplicado_de"]), "duplicados;",
          len(descargas), "archivos registrados")
    for a in avisos:
        print("AVISO:", a)
    return 0


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    p.add_argument("--raiz", type=Path, default=RAIZ, help="carpeta con las exportaciones (<base>/<id>.<ext>)")
    p.add_argument("--salida", type=Path, default=MANUAL, help="carpeta de resultados_manual.csv y estrategias_manuales.csv")
    p.add_argument("--registro", type=Path, default=REGISTRO, help="TSV de registro de archivos con su hash")
    p.add_argument("--estrategias", type=Path, default=MANUAL / "estrategias_manuales.csv")
    a = p.parse_args()
    sys.exit(main(a.raiz, a.salida, a.registro, a.estrategias))
