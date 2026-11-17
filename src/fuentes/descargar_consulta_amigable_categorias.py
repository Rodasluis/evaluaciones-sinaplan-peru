"""Descarga de la Consulta Amigable del MEF la ejecución de las categorías presupuestales vinculadas a los servicios de la
PNIG, por entidad, producto y actividad (REV_IND.3, presupuesto de los servicios).

Categorías: 1002 «Productos específicos para reducción de la violencia contra la mujer» y 0080 «Lucha contra la violencia
familiar», que financian servicios del objetivo prioritario 1. Recorrido por año: categoría → nivel de gobierno; en el
Gobierno Nacional, sector → pliego → producto → actividad; en los gobiernos regionales y locales, producto → actividad
(sin abrir cada pliego). Montos en soles: PIA, PIM, certificación, compromiso anual, atención de compromiso mensual,
devengado y girado. Reutiliza la navegación de descargar_consulta_amigable.py; no sobrescribe archivos existentes y
registra cada archivo con su hash.

Uso (desde la raíz del repositorio):
    uv run --no-project --with requests --with lxml python src/fuentes/descargar_consulta_amigable_categorias.py
"""
import csv
import hashlib
from datetime import date
from pathlib import Path

from descargar_consulta_amigable import Consulta, MONTOS, URL, actualizacion, filas
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

CATEGORIAS = {"1002": range(2019, 2026), "0080": range(2019, 2026)}
DESTINO = Path(f"{EVAL}/03_datos/raw/consulta_amigable/categorias")
LOG = Path(f"docs/fuentes/descargas_consulta-amigable_{date.today().isoformat()}.tsv")
NIVELES = {"E": "Gobierno Nacional", "R": "Gobiernos Regionales", "M": "Gobiernos Locales"}


def recorrer(q, doc, base):
    """Desde una página con filas seleccionables a nivel de pliego o nivel, baja a producto y actividad."""
    out = []
    prods = q.bajar(doc, "Producto", base["_codigo"])
    for _, pc, pn, _ in filas(prods):
        acts = q.bajar(prods, "Act.", pc)
        for _, ac, an, m in filas(acts):
            out.append({k: v for k, v in base.items() if not k.startswith("_")} |
                       {"producto": f"{pc}: {pn}", "actividad": f"{ac}: {an}"} | dict(zip(MONTOS, m)))
    return out


if __name__ == "__main__":
    DESTINO.mkdir(parents=True, exist_ok=True)
    for cat, anios in CATEGORIAS.items():
        for anio in anios:
            salida = DESTINO / f"{anio}_mef_consulta-amigable_categoria-{cat}.csv"
            if salida.exists():
                continue
            q = Consulta(anio)
            cats = q.bajar(q.inicio, "Categoría")
            if not any(c == cat for _, c, _, _ in filas(cats)):
                print(f"{anio}: la categoría {cat} no tiene presupuesto")
                continue
            niveles = q.bajar(cats, "Nivel de Gobierno", cat)
            registros = []
            for _, nc, _, _ in filas(niveles):
                if nc == "E":
                    sectores = q.bajar(niveles, "Sector", "E")
                    for _, sc, sn, _ in filas(sectores):
                        pliegos = q.bajar(sectores, "Pliego", sc)
                        for _, pc, pn, _ in filas(pliegos):
                            registros += recorrer(q, pliegos, {"anio": anio, "categoria": cat, "nivel": NIVELES["E"],
                                                               "sector": f"{sc}: {sn}", "pliego": f"{pc}: {pn}",
                                                               "_codigo": pc})
                else:
                    registros += recorrer(q, niveles, {"anio": anio, "categoria": cat, "nivel": NIVELES.get(nc, nc),
                                                       "sector": "", "pliego": "", "_codigo": nc})
            with open(salida, "w", newline="", encoding="utf-8") as f:
                w = csv.DictWriter(f, fieldnames=list(registros[0]))
                w.writeheader()
                w.writerows(registros)
            h = hashlib.sha256(salida.read_bytes()).hexdigest()
            with open(LOG, "a", encoding="utf-8") as f:
                f.write(f"{salida.as_posix()}\t{URL.format(y=anio)}\t{date.today().isoformat()}\t{actualizacion(niveles)}\t"
                        f"{len(registros)}\t{h}\n")
            print(f"OK {anio} categoría {cat}: {len(registros)} registros")
