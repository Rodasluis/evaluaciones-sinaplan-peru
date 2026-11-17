"""REV_EVI: búsqueda en el Development Evidence Portal (DEP) de 3ie mediante su API GraphQL.

El portal (developmentevidence.3ieimpact.org) carga sus resultados con JavaScript, por lo que en un inicio se dejó para
búsqueda manual (estrategias DEP-15 a DEP-22 de busqueda_manual/estrategias_manuales.csv). Su página consume una API GraphQL
pública, sin autenticación para las búsquedas (api.developmentevidence.3ieimpact.org/graphql), que acepta la misma
sintaxis de la búsqueda avanzada de la interfaz: operadores AND y OR en mayúsculas, frases entre comillas rectas y
truncamiento con *. Este script ejecuta esas ocho estrategias con la cadena exacta registrada, sin límite de año, con
orden por relevancia y con el filtro de tipo de producto a evaluaciones de impacto (ier) y revisiones sistemáticas (srr).

Para cada estrategia:
  - pide las dos primeras páginas de 25 registros (la API devuelve 25 por llamada) y conserva los primeros 30, igual
    que en PubMed y EconBiz;
  - guarda cada respuesta en evaluaciones/pnig_2026/03_datos/raw/rev_evi_busqueda/<fecha>/dep/ y la registra con su
    hash en docs/fuentes/descargas_rev06-dep_<fecha>.tsv;
  - marca los duplicados frente a los registros de la búsqueda académica (resultados_academicos.csv, por DOI o
    título normalizado) y entre las propias estrategias del DEP;
  - escribe busqueda_manual/resultados_dep.csv (una fila por registro, con país, nivel de ingreso y diseño según
    el portal) y busqueda_manual/resumenes_dep.json (resumen y hallazgos principales de cada registro, para el cribado);
  - completa en estrategias_manuales.csv las columnas fecha_ejecucion, total_resultados, archivo_exportado,
    ejecutado_por y observaciones de DEP-15 a DEP-22.

Uso (desde la raíz del repositorio):
    python src/analisis_pnig/rev_evi_busqueda_3ie_dep.py
Repetir la búsqueda en otra fecha puede dar totales distintos.
"""
import csv
import hashlib
import json
import re
import time
import urllib.request
from datetime import date
from pathlib import Path
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

INSUMOS = Path(f"{EVAL}/04_analisis/rev_evi_insumos")
MANUAL = INSUMOS / "busqueda_manual"
FECHA = date.today().isoformat()
CRUDO = Path(f"{EVAL}/03_datos/raw/rev_evi_busqueda/{FECHA}/dep")
REGISTRO = Path(f"docs/fuentes/descargas_rev06-dep_{FECHA}.tsv")
API = "https://api.developmentevidence.3ieimpact.org/graphql"
PORTAL = "https://developmentevidence.3ieimpact.org"
UA = {"User-Agent": "evaluacion-pnig-rev_evi (MIMP)", "Content-Type": "application/json"}
REVISADOS = 30
PAGINA = 25
PRODUCTOS = {"ier": ("impact-evaluation-repository", "Evaluación de impacto"),
             "srr": ("systematic-review-repository", "Revisión sistemática")}

CONSULTA = """query A($data: AdvancedSearchInput!) { advancedSearch(data: $data) { total_count
  search_result { id product_type title year_of_publication doi open_access publication_type journal
    evaluation_design evaluation_method review_type language sector_name keywords
    continent { continent countries { country income_level } }
    authors { author } abstract main_findings headline_findings
    based_on_the_above_assessments_of_the_methods_how_would_you_rate_the_reliability_of_the_review } } }"""

descargas = []


def pedir(variables, nombre):
    """Envía la consulta GraphQL, guarda la respuesta cruda y la registra con su hash."""
    cuerpo = json.dumps({"query": CONSULTA, "variables": variables}).encode()
    for intento in range(4):
        try:
            datos = urllib.request.urlopen(urllib.request.Request(API, data=cuerpo, headers=UA), timeout=90).read()
            break
        except Exception:
            time.sleep(3 * (intento + 1))
    else:
        raise RuntimeError(f"sin respuesta: {nombre}")
    ruta = CRUDO / nombre
    ruta.write_bytes(datos)
    descargas.append((f"{API} (POST advancedSearch, {nombre[:-5]})", ruta.as_posix(),
                      hashlib.sha256(datos).hexdigest(), FECHA, "Respuesta de la API (metadatos)"))
    time.sleep(0.5)
    return json.loads(datos.decode("utf-8"))


def enlace(r):
    tipo = PRODUCTOS[r["product_type"]][0]
    slug = re.sub(r"""["~/!@#$%^&*_+=`{}\\:;'<>,.?\- \t\r\n]+""", "-", r["title"]).lower()
    return f"{PORTAL}/search-result-details/{tipo}/{slug}/{r['id']}"


def doi(r):
    d = (r.get("doi") or "").strip()
    d = re.sub(r"^https?://(dx\.)?doi\.org/", "", d, flags=re.I)
    return "" if d.lower() in {"", "no doi", "none", "n/a"} else d


def paises(r):
    pares = [(c["country"], c.get("income_level") or "") for k in (r.get("continent") or []) for c in (k.get("countries") or [])]
    return ("; ".join(dict.fromkeys(p for p, _ in pares if p)), "; ".join(dict.fromkeys(i for _, i in pares if i)))


if __name__ == "__main__":
    CRUDO.mkdir(parents=True, exist_ok=True)
    estrategias = list(csv.DictReader(open(MANUAL / "estrategias_manuales.csv", encoding="utf-8")))
    academicos = list(csv.DictReader(open(INSUMOS / "resultados_academicos.csv", encoding="utf-8")))
    vistos = {}
    for a in academicos:
        norm = re.sub(r"[^a-z0-9]", "", a["titulo"].lower())[:80]
        vistos.setdefault(a["doi"].lower() or norm, a["clave"])
        vistos.setdefault(norm, a["clave"])
    resultados, resumenes = [], {}
    for e in estrategias:
        if e["base"] != "3ie DEP":
            continue
        regs, total = [], 0
        for desde in range(0, REVISADOS, PAGINA):
            d = pedir({"data": {"query": e["cadena"], "from": desde, "size": PAGINA, "sort_by": "relevance",
                                "filters": {"product_type": list(PRODUCTOS)}}}, f"{e['id']}_p{desde // PAGINA + 1}.json")
            assert "errors" not in d, d.get("errors")
            r = d["data"]["advancedSearch"]
            total = r["total_count"]
            regs += r["search_result"]
        regs = regs[:REVISADOS]
        for rango, r in enumerate(regs, 1):
            clave = f"DEP:{r['id']}"
            norm = re.sub(r"[^a-z0-9]", "", r["title"].lower())[:80]
            llave = doi(r).lower() or norm
            # Duplicado de un registro académico (por DOI o título normalizado). Igual que en la búsqueda académica, el
            # mismo registro recuperado por otra estrategia del DEP no se marca: cada estrategia revisa su propia lista y
            # el registro se decide en su primera aparición.
            dup = vistos.get(llave) or vistos.get(norm, "")
            vistos.setdefault(llave, clave)
            vistos.setdefault(norm, clave)
            pais, ingreso = paises(r)
            resultados.append({
                "estrategia": e["id"], "tipos": e["tipos"], "base": "3ie DEP", "rango": rango, "clave": clave,
                "doi": doi(r), "titulo": r["title"].strip(), "anio": r.get("year_of_publication") or "",
                "fuente": r.get("journal") or r.get("publication_type") or "", "tipo_documento": PRODUCTOS[r["product_type"]][1],
                "enlace": enlace(r), "duplicado_de": dup if dup != clave else "", "pais": pais, "ingreso": ingreso,
                "diseno_portal": r.get("evaluation_design") or r.get("review_type") or "",
                "metodo_portal": r.get("evaluation_method") or "",
                "confianza_portal": r.get("based_on_the_above_assessments_of_the_methods_how_would_you_rate_the_reliability_of_the_review") or ""})
            resumenes[clave] = {"titulo": r["title"].strip(), "autores": "; ".join(a["author"] for a in (r.get("authors") or []) if a.get("author")),
                                "resumen": r.get("abstract") or "", "hallazgos": r.get("main_findings") or "",
                                "titulares": r.get("headline_findings") or ""}
        e.update({"fecha_ejecucion": FECHA, "total_resultados": str(total),
                  "archivo_exportado": f"{CRUDO.as_posix()}/{e['id']}_p1.json y _p2.json",
                  "ejecutado_por": "OMEP, mediante la API GraphQL del portal (src/analisis_pnig/rev_evi_busqueda_3ie_dep.py)",
                  "observaciones": "Cadena ejecutada sin cambios en advancedSearch, orden por relevancia, filtro "
                                   "product_type = ier y srr; conservados los primeros 30 registros."})
        print(e["id"], e["tipos"], total, len(regs), flush=True)
    with open(MANUAL / "resultados_dep.csv", "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(resultados[0]), quoting=csv.QUOTE_ALL)
        w.writeheader()
        w.writerows(resultados)
    (MANUAL / "resumenes_dep.json").write_text(json.dumps(resumenes, ensure_ascii=False, indent=1), encoding="utf-8")
    with open(MANUAL / "estrategias_manuales.csv", "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(estrategias[0]), quoting=csv.QUOTE_ALL)
        w.writeheader()
        w.writerows(estrategias)
    with open(REGISTRO, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f, delimiter="\t", lineterminator="\n")
        w.writerow(["url_origen", "destino", "sha256", "fecha_descarga", "verificacion_identidad"])
        w.writerows(descargas)
    print(len(resultados), "registros;", sum(1 for r in resultados if r["duplicado_de"]), "duplicados")
