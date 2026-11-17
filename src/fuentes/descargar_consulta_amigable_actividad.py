"""Descarga de la Consulta Amigable del MEF el detalle de una actividad del pliego 039 (UE 001) por fuente de
financiamiento y por genérica de gasto, 2019–2025 (REV_IND.3).

Actividad: 5001072 «Promoción de la mujer» (categoría 9002, producto 3999999), vinculada por su denominación con la
conducción de la PNIG. La fuente permite verificar con el POI el financiamiento por donaciones (proyecto de la AECID en la
DGIGND); la genérica, la composición del gasto. Reutiliza la navegación de descargar_consulta_amigable.py; no sobrescribe
archivos existentes y registra cada archivo con su hash.

Uso (desde la raíz del repositorio):
    uv run --no-project --with requests --with lxml python src/fuentes/descargar_consulta_amigable_actividad.py
"""
import csv
import hashlib
from datetime import date
from pathlib import Path

from descargar_consulta_amigable import ANIOS, Consulta, MONTOS, URL, actualizacion, filas
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

RUTA = ("9002", "3999999", "5001072")   # categoría, producto, actividad
DIMENSIONES = {"Fuente": "fuente", "Genérica": "generica"}
DESTINO = Path(f"{EVAL}/03_datos/raw/consulta_amigable")
LOG = Path(f"docs/fuentes/descargas_consulta-amigable_{date.today().isoformat()}.tsv")

if __name__ == "__main__":
    for anio in ANIOS:
        salida = DESTINO / f"{anio}_mef_consulta-amigable_pliego-039_actividad-{RUTA[2]}.csv"
        if salida.exists():
            continue
        q = Consulta(anio)
        d = q.bajar(q.inicio, "Nivel de Gobierno")
        d = q.bajar(d, "Sector", "E")
        d = q.bajar(d, "Pliego", "39")
        d = q.bajar(d, "Ejecutora", "039")
        d = q.bajar(d, "Categoría", "001")
        d = q.bajar(d, "Producto", RUTA[0])
        acts = q.bajar(d, "Act.", RUTA[1])
        registros = []
        for boton, dim in DIMENSIONES.items():
            x = q.bajar(acts, boton, RUTA[2])
            registros += [{"anio": anio, "actividad": RUTA[2], "dimension": dim, "codigo": c, "denominacion": n}
                          | dict(zip(MONTOS, m)) for _, c, n, m in filas(x)]
        with open(salida, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=list(registros[0]))
            w.writeheader()
            w.writerows(registros)
        h = hashlib.sha256(salida.read_bytes()).hexdigest()
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(f"{salida.as_posix()}\t{URL.format(y=anio)}\t{date.today().isoformat()}\t{actualizacion(acts)}\t"
                    f"{len(registros)}\t{h}\n")
        print(f"OK {anio}: {len(registros)} registros")
