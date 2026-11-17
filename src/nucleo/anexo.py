"""Anexo técnico: piezas para armar la carpeta que acompaña un informe de evaluación (núcleo; extraído de
src/analisis_pnig/anexo_tecnico.py, que arma el del piloto).

- libro_publicado(): copia un libro de análisis con las hojas renombradas sin siglas;
- publico(): cambia en un texto de rastro los nombres de libros y hojas por los publicados;
- tabla_xlsx(): guarda un DataFrame en Excel con anchos de columna legibles;
- Registro: lleva la lista de archivos copiados (carpeta, archivo, descripción, origen) y su SHA-256.
"""
import re
from pathlib import Path

import openpyxl
import pandas as pd

from .fuentes import sha256


def libro_publicado(origen, destino, hoja_publica):
    """Copia un libro con las hojas renombradas (hoja_publica: función nombre -> nombre publicado)."""
    wb = openpyxl.load_workbook(origen)
    for ws in wb.worksheets:
        ws.title = hoja_publica(ws.title)
    wb.save(destino)


def publico(texto, libros, hojas):
    """Cambia los nombres de libros (dict repositorio -> publicado) y de hojas entre « » por los publicados."""
    t = str(texto)
    for rep, pub in libros.items():
        t = t.replace(rep, pub)
    return re.sub(r"«([^»]+)»", lambda m: f"«{hojas.get(m.group(1), m.group(1))}»", t)


def tabla_xlsx(df, destino, hoja):
    with pd.ExcelWriter(destino, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name=hoja, index=False)
        ws = xw.sheets[hoja]
        for col in ws.columns:
            ancho = min(80, max(len(str(c.value or "")) for c in col) + 2)
            ws.column_dimensions[col[0].column_letter].width = ancho


class Registro:
    """Lista de archivos del anexo: cada carpeta es una subcarpeta numerada de la salida."""

    def __init__(self, salida, carpetas):
        self.salida = Path(salida)
        self.d = {c: self.salida / c for c in carpetas}
        self.archivos = []          # (carpeta, archivo relativo, descripción, origen)

    def crear(self):
        for ruta in self.d.values():
            ruta.mkdir(parents=True, exist_ok=True)

    def registrar(self, carpeta, ruta, descripcion, origen=""):
        self.archivos.append((carpeta, Path(ruta).relative_to(self.d[carpeta]).as_posix(), descripcion,
                              str(origen).replace("\\", "/")))

    def lista(self):
        return pd.DataFrame([{"Carpeta": c, "Archivo": a, "Descripción": desc, "Origen en el repositorio": o,
                              "SHA-256": sha256(self.salida / c / a)} for c, a, desc, o in self.archivos])
