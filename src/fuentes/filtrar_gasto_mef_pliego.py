"""Extrae del conjunto «Presupuesto y Ejecución de Gasto» del MEF todo el detalle de un pliego del Gobierno Nacional: mes,
unidad ejecutora, meta, fuente, rubro y clasificador de gasto hasta la específica de detalle (REV_IND.3).

Sirve para seguir las modificaciones presupuestarias entre metas y órganos durante el año (por ejemplo, las metas de los
CEM en AURORA o las de la transversalización del enfoque de género en la UE 001). Usa los mismos ZIP anuales que
filtrar_gasto_mef.py; las líneas se preseleccionan por el texto «"<pliego>"» y luego se validan con el lector CSV.

Uso (desde la raíz del repositorio; <dir_zip> fuera del repositorio: los ZIP 2019–2025 están en
<carpeta administrativa>/datos_abiertos_mef_gasto; ruta en config/entorno.local.yaml, clave carpeta_administrativa):
    uv run --no-project --with pandas --with requests python src/fuentes/filtrar_gasto_mef_pliego.py <dir_zip> <pliego> 2019 ...
"""
import csv
import io
import sys
import zipfile
from datetime import date
from pathlib import Path

from filtrar_gasto_mef import NOMBRES, URL, bajar, sha
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

DESTINO = Path(f"{EVAL}/03_datos/raw/datos_abiertos_mef")
LOG = Path(f"docs/fuentes/descargas_datos-abiertos-mef_{date.today().isoformat()}.tsv")

if __name__ == "__main__":
    dir_zip, pliego = Path(sys.argv[1]), sys.argv[2]
    if not LOG.exists():
        LOG.write_text("destino\turl_origen\tfecha_descarga\tsha256_zip\tfilas\tsha256\n", encoding="utf-8")
    for anio in map(int, sys.argv[3:]):
        nombre = NOMBRES.get(anio, f"{anio}-Gasto")
        url, zip_ = URL.format(nombre=nombre), dir_zip / f"{nombre}.zip"
        salida = DESTINO / f"{anio}_mef_gasto_pliego-{pliego}_detalle-mensual.csv"
        if salida.exists():
            continue
        bajar(url, zip_)
        clave = f'"{pliego}"'  # los campos vienen entre comillas
        filas = 0
        z = zipfile.ZipFile(zip_)
        with z.open(z.infolist()[0]) as f, open(salida, "w", newline="", encoding="utf-8") as out:
            texto = io.TextIOWrapper(f, encoding="utf-8", newline="")
            cab = next(csv.reader([texto.readline()]))
            i_pl, i_niv = cab.index("PLIEGO"), cab.index("NIVEL_GOBIERNO")
            w = csv.writer(out)
            w.writerow(cab)
            for linea in texto:
                if clave not in linea:
                    continue
                r = next(csv.reader([linea]))
                if r[i_pl] == pliego and r[i_niv] == "E":
                    w.writerow(r)
                    filas += 1
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(f"{salida.as_posix()}\t{url}\t{date.today().isoformat()}\t{sha(zip_)}\t{filas}\t{sha(salida)}\n")
        print(f"OK {salida.name}: {filas} filas", flush=True)
