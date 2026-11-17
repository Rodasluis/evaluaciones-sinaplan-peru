"""Descarga los documentos públicos del corpus desde su enlace oficial y verifica su SHA-256 (versión pública del
repositorio: los PDF no se versionan, se obtienen con este script).

Lee docs/fuentes/inventario_publico.csv (ruta, enlace, sha256, nota) y guarda cada documento en su ruta. Un documento
que ya existe con el hash del inventario no se vuelve a descargar. Los enlaces caídos o los archivos con otro hash se
informan al final, sin detener la descarga. Dos casos tienen otro hash por razones conocidas y se informan como tales:
los documentos a los que el corpus aplicó OCR (la copia del corpus tiene capa de texto) y los que tienen publicada otra
versión (columna nota).

Uso (desde la raíz del repositorio):
    uv run python src/fuentes/descargar_inventario.py                 # todo el inventario (unos 900 MB)
    uv run python src/fuentes/descargar_inventario.py --carpeta ceplan  # solo una carpeta de docs/fuentes
    uv run python src/fuentes/descargar_inventario.py --muestra 5     # prueba con cinco documentos
"""
import argparse
import csv
import hashlib
import sys
import time
import urllib.request
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
INVENTARIO = RAIZ / "docs/fuentes/inventario_publico.csv"
AGENTE = "Mozilla/5.0 (compatible; evaluaciones-sinaplan; descarga de fuentes publicas)"


def bajar(url, intentos=2):
    for k in range(intentos):
        try:
            req = urllib.request.Request(url.replace("*", "%2A"), headers={"User-Agent": AGENTE})
            with urllib.request.urlopen(req, timeout=180) as r:
                return r.read()
        except Exception:   # noqa: BLE001  (se reintenta y luego se informa)
            if k == intentos - 1:
                raise
            time.sleep(3)


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    p.add_argument("--carpeta", help="solo los documentos de esta carpeta de docs/fuentes (p. ej. ceplan)")
    p.add_argument("--muestra", type=int, help="solo los primeros N documentos")
    a = p.parse_args(argv)
    with open(INVENTARIO, encoding="utf-8") as f:
        filas = list(csv.DictReader(f))
    if a.carpeta:
        filas = [x for x in filas if x["ruta"].startswith(f"docs/fuentes/{a.carpeta.strip('/')}/")]
    if a.muestra:
        filas = filas[:a.muestra]
    ok, ya, conocidos, fallas = 0, 0, [], []
    for x in filas:
        destino = RAIZ / x["ruta"]
        if destino.exists() and hashlib.sha256(destino.read_bytes()).hexdigest() == x["sha256"]:
            ya += 1
            continue
        try:
            contenido = bajar(x["enlace"])
        except Exception as e:   # noqa: BLE001
            fallas.append((x["ruta"], f"enlace sin respuesta ({type(e).__name__})"))
            continue
        destino.parent.mkdir(parents=True, exist_ok=True)
        destino.write_bytes(contenido)
        h = hashlib.sha256(contenido).hexdigest()
        if h == x["sha256"]:
            ok += 1
        elif x.get("nota"):
            conocidos.append((x["ruta"], x["nota"]))
        else:
            fallas.append((x["ruta"], "el archivo publicado tiene otro SHA-256: puede haberse actualizado"))
        print("descargado", x["ruta"], flush=True)
    print(f"\n{len(filas)} documentos: {ok} descargados con el hash del inventario, {ya} ya estaban, "
          f"{len(conocidos)} con otro hash por una razón conocida, {len(fallas)} con problemas.")
    for ruta, nota in conocidos:
        print("  conocido:", ruta, "|", nota)
    for ruta, motivo in fallas:
        print("  revisar:", ruta, "|", motivo)
    return 1 if fallas else 0


if __name__ == "__main__":
    sys.exit(main())
