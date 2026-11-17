"""Completa la lista de autores de las referencias de la revisión de evidencia que solo dicen «et al.» (APA 7 pide
hasta veinte autores en la lista de referencias; ADR-010).

Para cada referencia con «et al.» del registro de evidencia (04_analisis/rev_evi_insumos/evidencia.csv), obtiene los
autores de una fuente bibliográfica oficial, sin claves ni datos personales en la consulta:
  - Crossref (api.crossref.org/works/{doi}) cuando el enlace trae el DOI;
  - PubMed (E-utilities, esummary) cuando el enlace es de PubMed;
  - OpenAlex (api.openalex.org/works/{id}) cuando el enlace es de OpenAlex;
  - si no hay identificador, búsqueda bibliográfica en Crossref por el título, aceptada solo si el título coincide
    (similitud de 0,92 o más) y el primer autor y el año son los del registro.
Cada respuesta se guarda en 03_datos/raw/rev_evi_autores/<fecha>/ (no versionada) y se registra con su SHA-256 en
docs/fuentes/descargas_autores-evidencia_<fecha>.tsv. La propuesta queda en 04_analisis/rev_evi_insumos/
autores_completados.csv; el modo «aplicar» reemplaza el tramo de autores en rev_evi_insumos.py, que genera el
registro, y agrega a la verificación la fuente de los autores.

Formato APA 7 en español: «Apellido, I. I.», separados por comas y con «y» antes del último; con más de veinte
autores, los diecinueve primeros, puntos suspensivos y el último.

Uso (desde la raíz del repositorio):
    python src/fuentes/completar_autores_evidencia.py consultar
    python src/fuentes/completar_autores_evidencia.py aplicar
"""
import csv
import difflib
import hashlib
import json
import re
import sys
import time
import unicodedata
from datetime import date
from pathlib import Path

import requests
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

EVIDENCIA = Path(f"{EVAL}/04_analisis/rev_evi_insumos/evidencia.csv")
PROPUESTA = Path(f"{EVAL}/04_analisis/rev_evi_insumos/autores_completados.csv")
INSUMOS = Path("src/analisis_pnig/rev_evi_insumos.py")
HOY = date.today().isoformat()
RAW = Path(f"{EVAL}/03_datos/raw/rev_evi_autores/{HOY}")
REGISTRO = Path(f"docs/fuentes/descargas_autores-evidencia_{HOY}.tsv")
CABECERA = {"User-Agent": "evaluacion-pnig-mimp/1.0 (consulta bibliografica)"}
# DOI de la versión citada cuando el enlace del registro no lo trae; identificados con una búsqueda bibliográfica en
# Crossref por título (mismo título, primer autor y año) el 2026-11-08
DOI_VERSION = {"E12": "10.1371/journal.pone.0192756", "E28": "10.1016/j.worlddev.2018.10.004"}


def _plano(t):
    return unicodedata.normalize("NFKD", t).encode("ascii", "ignore").decode().lower()


def _get(url, nombre, registro):
    r = requests.get(url, headers=CABECERA, timeout=60)
    r.raise_for_status()
    RAW.mkdir(parents=True, exist_ok=True)
    ruta = RAW / nombre
    ruta.write_bytes(r.content)
    registro.append({"fecha": HOY, "url": url, "archivo": ruta.as_posix(),
                     "sha256": hashlib.sha256(r.content).hexdigest(), "bytes": len(r.content)})
    time.sleep(1)
    return r.json()


def _iniciales(nombres):
    """«Ruti G.» -> «R. G.»; «G.J.» -> «G. J.»; «Jean-Pierre» -> «J.-P.»."""
    salida = []
    for p in re.split(r"\s+", nombres.strip()):
        for trozo in [x for x in p.split(".") if x]:
            sub = [s for s in trozo.split("-") if s and s[0].isalpha()]
            if sub:
                salida.append("-".join(f"{s[0].upper()}." for s in sub))
    return " ".join(salida)


def apa(autores):
    """autores: lista de (apellido, iniciales) -> «Apellido, I., Apellido, I. y Apellido, I.»"""
    items = [f"{a}, {i}".strip().rstrip(",") for a, i in autores]
    if len(items) > 20:
        return ", ".join(items[:19]) + ", … " + items[-1]
    return items[0] if len(items) == 1 else ", ".join(items[:-1]) + " y " + items[-1]


def autores_crossref(msg):
    return [(a.get("family", "").strip(), _iniciales(a.get("given", ""))) for a in msg.get("author", [])
            if a.get("family")]


def consultar():
    filas = list(csv.DictReader(EVIDENCIA.open(encoding="utf-8")))
    objetivo = [f for f in filas if "et al" in f["referencia"]]
    registro, propuesta = [], []
    for f in objetivo:
        ref, enlace, anio = f["referencia"], f["enlace"], f["anio"]
        m = re.match(r"^(.+?)\s*\((\d{4})\)\.\s*(.+?)[.?!]\s", ref)
        tramo, titulo = m.group(1), m.group(3)
        primero = _plano(tramo.split(",")[0])
        autores, fuente = None, ""
        doi = re.search(r"(10\.\d{4,9}/[^\s?#]+?)(?:/(?:information|pdf|abstract|full))?(?:[/?#].*)?$", enlace)
        try:
            if "pubmed.ncbi.nlm.nih.gov" in enlace:
                pmid = re.search(r"/(\d+)/?$", enlace).group(1)
                js = _get(f"https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esummary.fcgi?db=pubmed&id={pmid}"
                          f"&retmode=json", f"{f['id']}_pubmed_{pmid}.json", registro)
                res = js["result"][pmid]
                autores = [(re.sub(r"\s+[A-Z]+$", "", a["name"]), " ".join(f"{c}." for c in a["name"].split()[-1]))
                           for a in res["authors"] if a.get("authtype") == "Author"]
                fuente = f"PubMed (PMID {pmid})"
            elif "openalex.org" in enlace:
                wid = re.search(r"(W\d+)", enlace).group(1)
                js = _get(f"https://api.openalex.org/works/{wid}", f"{f['id']}_openalex_{wid}.json", registro)
                autores = []
                for a in js["authorships"]:
                    # nombres hispanos: dos nombres y dos apellidos (con cuatro palabras o más) o un nombre y el resto
                    nombre = (a.get("raw_author_name") or a["author"]["display_name"]).split()
                    k = 2 if len(nombre) >= 4 else 1
                    familia = " ".join(x + "." if len(x) == 1 else x for x in nombre[k:])
                    autores.append((familia, _iniciales(" ".join(nombre[:k]))))
                fuente = f"OpenAlex ({wid})"
            elif doi or f["id"] in DOI_VERSION:
                d = DOI_VERSION.get(f["id"]) or doi.group(1).rstrip(".")
                try:
                    js = _get(f"https://api.crossref.org/works/{d}", f"{f['id']}_crossref.json", registro)
                    autores = autores_crossref(js["message"])
                    fuente = f"Crossref (DOI {d})"
                except requests.HTTPError:
                    # DOI de otra agencia (DataCite): metadatos por negociación de contenido en doi.org
                    CABECERA["Accept"] = "application/vnd.citationstyles.csl+json"
                    try:
                        js = _get(f"https://doi.org/{d}", f"{f['id']}_doi-csl.json", registro)
                    finally:
                        CABECERA.pop("Accept")
                    autores = autores_crossref(js)
                    fuente = f"doi.org (DOI {d}, DataCite)"
            else:
                q = requests.utils.quote(titulo)
                js = _get(f"https://api.crossref.org/works?query.bibliographic={q}&rows=3",
                          f"{f['id']}_crossref_busqueda.json", registro)
                for it in js["message"]["items"]:
                    t = (it.get("title") or [""])[0]
                    sim = difflib.SequenceMatcher(None, _plano(t), _plano(titulo)).ratio()
                    au = autores_crossref(it)
                    an = str((it.get("issued", {}).get("date-parts") or [[None]])[0][0])
                    if sim >= 0.92 and au and _plano(au[0][0]) == primero and an == anio:
                        autores, fuente = au, f"Crossref (DOI {it['DOI']}, por título)"
                        break
        except (requests.RequestException, KeyError, AttributeError) as e:
            fuente = f"error: {e}"
        ok = bool(autores) and _plano(autores[0][0]).split()[-1] in primero
        if ok:
            # los autores que el registro ya nombra (verificados el 2026-10-31) se conservan tal cual; de la fuente
            # bibliográfica solo se agregan los que faltan
            previos = re.findall(r"\s*([^,]+?), ((?:[A-ZÁÉÍÓÚÑ]\.[\s\-]?)+)(?:,|$)",
                                 re.sub(r"\s*et al\.?$", "", tramo).strip() + ",")
            previos = [(s.strip(), i.strip()) for s, i in previos]
            ok = len(autores) > len(previos)
            autores = previos + autores[len(previos):]
        propuesta.append({"id": f["id"], "tramo_actual": tramo, "tramo_nuevo": apa(autores) if ok else "",
                          "n_autores": len(autores or []), "fuente": fuente, "aceptado": "Sí" if ok else "No"})
        print(f["id"], "Sí" if ok else "No", len(autores or []), fuente)
    with PROPUESTA.open("w", encoding="utf-8", newline="") as h:
        w = csv.DictWriter(h, fieldnames=list(propuesta[0]))
        w.writeheader()
        w.writerows(propuesta)
    nuevo = not REGISTRO.exists()
    with REGISTRO.open("a", encoding="utf-8", newline="") as h:
        w = csv.DictWriter(h, fieldnames=["fecha", "url", "archivo", "sha256", "bytes"], delimiter="\t")
        if nuevo:
            w.writeheader()
        w.writerows(registro)


def aplicar():
    texto = INSUMOS.read_text(encoding="utf-8")
    hechos = 0
    for p in csv.DictReader(PROPUESTA.open(encoding="utf-8")):
        if p["aceptado"] != "Sí":
            continue
        viejo = p["tramo_actual"]
        assert "et al" in viejo, f"{p['id']}: el tramo no termina en «et al.»"
        n = texto.count(viejo)
        assert n >= 1, f"{p['id']}: el tramo no está en {INSUMOS}"
        texto = texto.replace(viejo, p["tramo_nuevo"])
        hechos += 1
    INSUMOS.write_text(texto, encoding="utf-8")
    print("tramos de autores reemplazados:", hechos)


if __name__ == "__main__":
    assert Path("CLAUDE.md").exists(), "correr desde la raíz del repositorio"
    {"consultar": consultar, "aplicar": aplicar}[sys.argv[1]]()
