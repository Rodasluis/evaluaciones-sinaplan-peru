"""Inventario de los documentos de docs/fuentes con su hash, páginas y origen, para el registro de fuentes.

El origen se toma de los registros de descargas y reorganización (docs/fuentes/*.tsv): URL de descarga cuando el
documento se obtuvo de un portal oficial, o la ruta original cuando lo entregó el usuario. La columna ocr marca los
escaneos a los que se agregó capa de texto (registro ocr_*.tsv); su hash es el del PDF con OCR.
El acceso y el enlace público salen del registro de enlaces (enlaces_publicos_*.tsv, de verificar_enlaces_publicos.py)
para los documentos entregados que están publicados; para los descargados, el enlace es la URL de descarga. El origen
no cambia: dice cómo llegó el archivo al repositorio.

Uso (desde la raíz del repositorio): python src/fuentes/inventario_fuentes.py
"""
import csv
import hashlib
import re
import subprocess
from pathlib import Path

BASE = Path("docs/fuentes")
SALIDA = BASE / "inventario_fuentes.csv"
EXT = {".pdf", ".docx", ".doc", ".xlsx", ".xls", ".pptx", ".html"}

origen, ocr, enlaces = {}, set(), {}
for log in sorted(BASE.glob("*.tsv")):
    with open(log, encoding="utf-8") as f:
        filas_log = list(csv.reader(f, delimiter="\t"))
    if filas_log and filas_log[0][0] == "archivo":  # registro de OCR: archivo, hashes, páginas
        ocr.update(fila[0] for fila in filas_log[1:] if fila)
        continue
    if filas_log and filas_log[0][0] == "documento":  # registro de enlaces públicos
        cab = filas_log[0]
        enlaces.update((fila[0], dict(zip(cab, fila))) for fila in filas_log[1:] if fila)
        continue
    for fila in filas_log:
        if len(fila) < 2 or fila[0] in ("origen", "url_origen"):
            continue
        src, dst = fila[0], fila[1].replace("\\", "/")
        tipo = "descarga" if src.startswith("http") else "entregado por el usuario"
        origen[dst] = (tipo, src)


def paginas(p):
    if p.suffix.lower() != ".pdf":
        return ""
    try:
        out = subprocess.run(["pdfinfo", str(p)], capture_output=True, text=True, encoding="latin-1").stdout
        m = re.search(r"Pages:\s+(\d+)", out)
        return m.group(1) if m else ""
    except OSError:
        return ""


filas = []
for p in sorted(BASE.rglob("*")):
    if not p.is_file() or p.suffix.lower() not in EXT:
        continue
    rel = p.as_posix()
    m = re.match(r"(\d{4}|\d{7,})_([^_]+)_(.+)\.\w+$", p.name)
    anio, organismo, titulo = (m.group(1), m.group(2), m.group(3)) if m else ("", "", p.stem)
    tipo, src = origen.get(rel, ("entregado por el usuario (Fase 0, sin registro de ruta)", ""))
    e = enlaces.get(rel)
    if e:
        acceso = (e["acceso"], e["enlace_archivo"], e["pagina_publicacion"],
                  f"{e['verificacion']} ({e['fecha_verificacion']})")
    elif tipo == "descarga":
        acceso = ("público", src, "", "Descargado del enlace (registro de descargas)")
    else:
        acceso = ("sin enlace público identificado", "", "", "")
    filas.append({
        "ruta": rel, "carpeta": p.parent.relative_to(BASE).as_posix(), "anio": anio if len(anio) == 4 else "",
        "organismo": organismo, "titulo_corto": titulo, "formato": p.suffix.lower().lstrip("."),
        "tamano_kb": round(p.stat().st_size / 1024), "paginas": paginas(p),
        "sha256": hashlib.sha256(p.read_bytes()).hexdigest(), "origen": tipo, "fuente_original": src,
        "ocr": "sí" if rel in ocr else "",
        **dict(zip(("acceso", "enlace_publico", "pagina_publicacion", "verificacion_enlace"), acceso)),
    })
assert not set(enlaces) - {f["ruta"] for f in filas}, "hay enlaces de documentos que no están en docs/fuentes"
with open(SALIDA, "w", newline="", encoding="utf-8-sig") as f:
    w = csv.DictWriter(f, fieldnames=list(filas[0]), lineterminator="\n")   # LF: el CSV se versiona sin conversión
    w.writeheader()
    w.writerows(filas)
print("OK", len(filas), "documentos")
