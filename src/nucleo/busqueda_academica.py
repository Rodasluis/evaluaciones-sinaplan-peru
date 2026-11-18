"""Búsqueda académica replicable para la revisión de evidencia (Fase 4, ADR-017), generalizada de
src/analisis_pnig/rev_evi_busqueda_academica.py.

Ejecuta las estrategias de evaluaciones/<id>/02_insumos/rev_evi_estrategias.yaml en bases con interfaz pública:
PubMed (MeSH), ERIC (tesauro de ERIC), EconBiz (tesauro STW), Campbell Systematic Reviews y 3ie (Crossref),
LA Referencia, Banco Mundial (Policy Research Working Papers) y OpenAlex. Cada estrategia lleva la cadena exacta, y se
registran la fecha, el total de resultados, los registros revisados (los primeros por relevancia) y la URL de la
consulta. SciELO, LILACS, Scopus, Web of Science y EconLit no admiten consultas automatizadas o requieren suscripción:
sus estrategias se ejecutan a mano y se registran en rev_evi_busquedas.csv con el mismo formato.

rev_evi_estrategias.yaml:
    revisados: {PubMed: 30}          # opcional: registros revisados por base (por defecto, los del piloto)
    estrategias:
      - {id: ACA-001, tipos: T01, base: PubMed, cadena: '"Domestic Violence"[mh] AND …'}

Escribe en 02_insumos/: rev_evi_busquedas.csv (una fila por estrategia; conserva las columnas seleccionados e incluidos
ya llenadas), rev_evi_resultados.csv (registros revisados con su clave y la marca de duplicado, para el cribado) y
rev_evi_descriptores.csv (descriptores controlados usados). Las respuestas crudas van a
03_datos/raw/rev_evi_busqueda/<fecha>/ (no se versionan) y se registran con su hash en
docs/fuentes/descargas_<id>_rev-evi-busqueda_<fecha>.tsv (CLAUDE.md, regla 7). Si existe la variable OPENALEX_API_KEY,
la clave se envía en la consulta y nunca se escribe en los registros.

Uso (desde la raíz del repositorio; requiere red):
    uv run python src/nucleo/busqueda_academica.py <id>
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

import yaml

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from nucleo.config import RAIZ_REPO, cargar_desde  # noqa: E402

UA = {"User-Agent": "Mozilla/5.0 (compatible; evaluaciones-sinaplan, busqueda academica)"}
REVISADOS = {"PubMed": 30, "ERIC": 15, "EconBiz": 30, "LA Referencia": 15, "Campbell": 10, "3ie": 10,
             "Banco Mundial": 15, "OpenAlex": 15}
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
TESAURO = {"PubMed": "MeSH (NLM)", "ERIC": "Tesauro de ERIC", "EconBiz": "Tesauro de Economía STW (ZBW)"}
BUSQUEDAS = ["id", "base", "tipos", "cadena", "filtros", "fecha", "total", "revisados", "seleccionados", "incluidos",
             "url", "interfaz", "tesauro", "descriptores", "terminos_libres"]
RESULTADOS = ["estrategia", "tipos", "base", "rango", "clave", "doi", "titulo", "anio", "fuente", "tipo_documento",
              "enlace", "duplicado_de"]


class Cliente:
    """Descargas con reintentos; guarda cada respuesta cruda y su hash."""

    def __init__(self, crudo, fecha):
        self.crudo = Path(crudo)
        self.fecha = fecha
        self.descargas = []

    def pedir(self, url, nombre, url_registro=None):
        for intento in range(4):
            try:
                datos = urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=60).read()
                break
            except Exception:
                time.sleep(3 * (intento + 1))
        else:
            raise RuntimeError(f"sin respuesta: {url_registro or url}")
        self.crudo.mkdir(parents=True, exist_ok=True)
        ruta = self.crudo / nombre
        ruta.write_bytes(datos)
        self.descargas.append((url_registro or url, ruta.as_posix(), hashlib.sha256(datos).hexdigest(), self.fecha,
                               "Respuesta de la API (metadatos)"))
        time.sleep(0.4)
        return datos


def pubmed(c, cadena, n, eid):
    q = urllib.parse.urlencode({"db": "pubmed", "term": cadena, "retmode": "json", "retmax": n, "sort": "relevance"})
    url = f"https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi?{q}"
    r = json.loads(c.pedir(url, f"{eid}_esearch.json"))["esearchresult"]
    regs = []
    if r["idlist"]:
        s = json.loads(c.pedir("https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esummary.fcgi?db=pubmed&retmode=json&id="
                               + ",".join(r["idlist"]), f"{eid}_esummary.json"))["result"]
        for pmid in r["idlist"]:
            d = s[pmid]
            doi = next((a["value"] for a in d.get("articleids", []) if a["idtype"] == "doi"), "")
            regs.append({"clave": f"PMID:{pmid}", "doi": doi, "titulo": d.get("title", ""),
                         "anio": d.get("pubdate", "")[:4], "fuente": d.get("fulljournalname", d.get("source", "")),
                         "tipo_documento": "; ".join(d.get("pubtype", [])),
                         "enlace": f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/"})
    return int(r["count"]), url, regs


def eric(c, cadena, n, eid):
    q = urllib.parse.urlencode({"search": f"({cadena}) AND peerreviewed:T", "format": "json", "rows": n,
                                "fields": "id,title,author,publicationdateyear,source,publicationtype,url"})
    url = f"https://api.ies.ed.gov/eric/?{q}"
    r = json.loads(c.pedir(url, f"{eid}.json"))["response"]
    regs = [{"clave": f"ERIC:{d['id']}", "doi": "", "titulo": d.get("title", ""),
             "anio": str(d.get("publicationdateyear", "")), "fuente": d.get("source", ""),
             "tipo_documento": "; ".join(d.get("publicationtype", [])), "enlace": f"https://eric.ed.gov/?id={d['id']}"}
            for d in r["docs"]]
    return r["numFound"], url, regs


def econbiz(c, cadena, n, eid):
    url = f"https://api.econbiz.de/v1/search?{urllib.parse.urlencode({'q': cadena, 'size': n})}"
    r = json.loads(c.pedir(url, f"{eid}.json"))["hits"]
    regs = []
    for d in r["hits"]:
        doi = next((u.split("doi.org/", 1)[1] for u in d.get("identifier_url", []) if "doi.org/" in u), "")
        titulo = d.get("title", "")
        regs.append({"clave": f"ECONBIZ:{d['id']}", "doi": doi,
                     "titulo": " ".join(titulo) if isinstance(titulo, list) else titulo,
                     "anio": (d.get("date") or [""])[0][:4],
                     "fuente": "; ".join(d.get("isPartOf") or d.get("publisher") or []),
                     "tipo_documento": d.get("type", ""), "enlace": f"https://www.econbiz.de/Record/{d['id']}"})
    return r["total"], url, regs


def crossref(c, cadena, n, eid, filtro):
    q = urllib.parse.urlencode({"query.bibliographic": cadena, "filter": filtro, "rows": n,
                                "select": "DOI,title,issued,type,container-title"})
    url = f"https://api.crossref.org/works?{q}"
    m = json.loads(c.pedir(url, f"{eid}.json"))["message"]
    regs = [{"clave": f"DOI:{d['DOI'].lower()}", "doi": d["DOI"].lower(), "titulo": " ".join(d.get("title", [""])),
             "anio": str(d.get("issued", {}).get("date-parts", [[""]])[0][0] or ""),
             "fuente": " ".join(d.get("container-title", [])) or ("3ie" if filtro.startswith("prefix") else ""),
             "tipo_documento": d.get("type", ""), "enlace": f"https://doi.org/{d['DOI']}"} for d in m["items"]]
    return m["total-results"], url, regs


def lareferencia(c, cadena, n, eid):
    params = [("lookfor", cadena), ("type", "AllFields"), ("limit", n), ("filter[]", '~format:"article"'),
              ("filter[]", '~format:"report"')] + [("field[]", f) for f in (
                  "id", "title", "formats", "publicationDates", "urls", "containerTitle")]
    url = f"https://buscador.lareferencia.info/api/v1/search?{urllib.parse.urlencode(params)}"
    r = json.loads(c.pedir(url, f"{eid}.json"))
    regs = []
    for d in r.get("records", []):
        enlace = (d.get("urls") or [{}])[0].get("url", "")
        regs.append({"clave": f"LAREF:{d['id']}", "doi": enlace.split("doi.org/", 1)[1].lower()
                     if "doi.org/" in enlace else "", "titulo": d.get("title", ""),
                     "anio": (d.get("publicationDates") or [""])[0][:4], "fuente": d.get("containerTitle", ""),
                     "tipo_documento": "; ".join(d.get("formats", [])), "enlace": enlace})
    return r.get("resultCount", 0), url, regs


def bancomundial(c, cadena, n, eid):
    q = urllib.parse.urlencode({"format": "json", "qterm": cadena, "rows": n,
                                "docty_exact": "Policy Research Working Paper", "fl": "display_title,docdt,docty,url"})
    url = f"https://search.worldbank.org/api/v3/wds?{q}"
    d = json.loads(c.pedir(url, f"{eid}.json").decode("utf-8-sig"))
    regs = [{"clave": f"BM:{v['id']}", "doi": "", "titulo": v.get("display_title", ""), "anio": v.get("docdt", "")[:4],
             "fuente": "Banco Mundial, Policy Research Working Paper", "tipo_documento": v.get("docty", ""),
             "enlace": v.get("url", "")} for k, v in d.get("documents", {}).items() if k != "facets"]
    return d.get("total", 0), url, regs


def openalex(c, cadena, n, eid):
    q = urllib.parse.urlencode({"search": cadena, "filter": "language:es,type:article|review|report", "per_page": n,
                                "select": "id,doi,display_name,publication_year,type,primary_location"})
    url = f"https://api.openalex.org/works?{q}"
    clave = os.environ.get("OPENALEX_API_KEY", "")
    d = json.loads(c.pedir(url + (f"&api_key={clave}" if clave else ""), f"{eid}.json", url_registro=url))
    regs = []
    for w in d.get("results", []):
        doi = (w.get("doi") or "").replace("https://doi.org/", "").lower()
        fuente = ((w.get("primary_location") or {}).get("source") or {}).get("display_name") or ""
        regs.append({"clave": f"OPENALEX:{w['id'].rsplit('/', 1)[1]}", "doi": doi, "titulo": w.get("display_name") or "",
                     "anio": str(w.get("publication_year") or ""), "fuente": fuente,
                     "tipo_documento": w.get("type", ""), "enlace": f"https://doi.org/{doi}" if doi else w["id"]})
    return d.get("meta", {}).get("count", 0), url, regs


BUSCAR = {"PubMed": pubmed, "ERIC": eric, "EconBiz": econbiz, "LA Referencia": lareferencia,
          "Banco Mundial": bancomundial, "OpenAlex": openalex,
          "Campbell": lambda c, x, n, e: crossref(c, x, n, e, "issn:1891-1803,type:journal-article"),
          "3ie": lambda c, x, n, e: crossref(c, x, n, e, "prefix:10.23846")}


def componentes(base, cadena):
    """Descriptores controlados, términos libres y filtros de una cadena."""
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
    unico = lambda xs: list(dict.fromkeys(x.strip('"') for x in xs))  # noqa: E731
    return unico(desc), unico(libres), unico(filtros)


def leer_estrategias(cfg):
    archivo = cfg.abs("insumos", "rev_evi_estrategias.yaml")
    with open(archivo, encoding="utf-8") as f:
        datos = yaml.safe_load(f) or {}
    estrategias = datos.get("estrategias") or []
    problemas = [f"estrategia {i}: base «{e.get('base')}» sin interfaz pública ({', '.join(BUSCAR)})"
                 for i, e in enumerate(estrategias, 1) if e.get("base") not in BUSCAR]
    problemas += [f"estrategia {i}: falta {c}" for i, e in enumerate(estrategias, 1) for c in ("tipos", "cadena")
                  if not e.get(c)]
    if problemas:
        raise ValueError("rev_evi_estrategias.yaml:\n  - " + "\n  - ".join(problemas))
    return {**REVISADOS, **(datos.get("revisados") or {})}, estrategias


def _escribir(ruta, columnas, filas):
    with open(ruta, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=columnas, quoting=csv.QUOTE_ALL, extrasaction="ignore")
        w.writeheader()
        w.writerows(filas)


def buscar(cfg, cliente=None, fecha=None):
    """Ejecuta las estrategias y escribe los registros; devuelve (estrategias, registros)."""
    fecha = fecha or date.today().isoformat()
    revisados, estrategias = leer_estrategias(cfg)
    cliente = cliente or Cliente(cfg.abs("datos_raw", "rev_evi_busqueda", fecha), fecha)
    previas = {}
    archivo = cfg.abs("insumos", "rev_evi_busquedas.csv")
    if archivo.exists():
        with open(archivo, encoding="utf-8-sig") as f:
            previas = {r["id"]: r for r in csv.DictReader(f)}
    filas, resultados, vistos = [], [], {}
    for i, e in enumerate(estrategias, 1):
        eid = e.get("id") or f"ACA-{i:03d}"
        base, cadena = e["base"], e["cadena"]
        total, url, regs = BUSCAR[base](cliente, cadena, int(revisados[base]), eid)
        desc, libres, filtros = componentes(base, cadena)
        previa = previas.get(eid, {})
        filas.append({"id": eid, "base": base, "tipos": e["tipos"], "cadena": cadena, "filtros": "; ".join(filtros),
                      "fecha": fecha, "total": total, "revisados": len(regs),
                      "seleccionados": previa.get("seleccionados", ""), "incluidos": previa.get("incluidos", ""),
                      "url": url, "interfaz": INTERFAZ[base], "tesauro": TESAURO.get(base, "Sin tesauro"),
                      "descriptores": "; ".join(desc), "terminos_libres": "; ".join(libres)})
        for rango, r in enumerate(regs, 1):
            norm = re.sub(r"[^a-z0-9]", "", r["titulo"].lower())[:80]
            llave = r["doi"].lower() or norm
            dup = vistos.get(llave) or vistos.get(norm, "")
            vistos.setdefault(llave, r["clave"])
            vistos.setdefault(norm, r["clave"])
            resultados.append({"estrategia": eid, "tipos": e["tipos"], "base": base, "rango": rango, **r,
                               "duplicado_de": dup if dup != r["clave"] else ""})
    _escribir(archivo, BUSQUEDAS, filas)
    _escribir(cfg.abs("insumos", "rev_evi_resultados.csv"), RESULTADOS, resultados)
    descriptores = sorted({(TESAURO[f["base"]], d) for f in filas for d in base_lista(f["descriptores"])})
    _escribir(cfg.abs("insumos", "rev_evi_descriptores.csv"), ["tesauro", "descriptor"],
              [{"tesauro": t, "descriptor": d} for t, d in descriptores])
    registro = cfg.repo / "docs" / "fuentes" / f"descargas_{cfg.id}_rev-evi-busqueda_{fecha}.tsv"
    if cliente.descargas:
        registro.parent.mkdir(parents=True, exist_ok=True)
        with open(registro, "w", encoding="utf-8", newline="") as f:
            w = csv.writer(f, delimiter="\t", lineterminator="\n")
            w.writerow(["url_origen", "destino", "sha256", "fecha_descarga", "verificacion_identidad"])
            w.writerows(cliente.descargas)
    return filas, resultados


def base_lista(texto):
    return [x.strip() for x in texto.split(";") if x.strip()]


def main(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    if not argv:
        print(__doc__)
        return 2
    cfg = cargar_desde(RAIZ_REPO, argv[0])
    try:
        filas, resultados = buscar(cfg)
    except (ValueError, FileNotFoundError) as e:
        print(e)
        return 1
    print(f"{len(filas)} estrategias y {len(resultados)} registros revisados en 02_insumos/ "
          f"({sum(1 for r in resultados if r['duplicado_de'])} duplicados).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
