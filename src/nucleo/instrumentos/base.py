"""Piezas comunes de los instrumentos genéricos: lectura y validación de las entradas estándar y libro de salida.

Las entradas son CSV en UTF-8 con cabecera. Los valores se leen como texto, sin espacios en los extremos; las celdas
vacías quedan como "". Los problemas de validación se reúnen y se informan juntos (ErrorInsumo), con el archivo, la
fila y la corrección esperada, para que el analista los corrija de una vez.
"""
import re
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill

from ..config import CARPETAS

SI_NO = ("Sí", "No")
PLANTILLAS = "plantillas/nueva_evaluacion/evaluacion"
SIGLAS = re.compile(r"\b(?:REV_|REP_ENT|ENTREV)")
CABECERA = PatternFill("solid", fgColor="034754")


class ErrorInsumo(ValueError):
    """Entradas que no cumplen su esquema; el mensaje lista todos los problemas."""


class Problemas:
    """Acumula los problemas de validación de un instrumento y los lanza juntos."""

    def __init__(self, instrumento):
        self.instrumento = instrumento
        self.lista = []

    def __call__(self, mensaje):
        self.lista.append(mensaje)

    def __bool__(self):
        return bool(self.lista)

    def revisar(self):
        if self.lista:
            raise ErrorInsumo(f"{self.instrumento}: {len(self.lista)} problema(s) en las entradas:\n  - "
                              + "\n  - ".join(self.lista))


def ruta(cfg, carpeta, nombre):
    return cfg.abs(carpeta, nombre)


def leer(cfg, nombre, columnas, problemas, carpeta="insumos", opcional=False):
    """CSV de entrada con sus columnas obligatorias. Si falta el archivo y es opcional, devuelve None."""
    archivo = ruta(cfg, carpeta, nombre)
    rel = cfg.ruta(carpeta, nombre)
    if not archivo.exists():
        if not opcional:
            problemas(f"{rel}: no existe (cabecera de ejemplo en {PLANTILLAS}/{CARPETAS.get(carpeta, carpeta)}/)")
        return None
    df = pd.read_csv(archivo, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    df.columns = [c.strip() for c in df.columns]
    df = df.map(lambda v: v.strip() if isinstance(v, str) else v)
    faltan = [c for c in columnas if c not in df.columns]
    if faltan:
        problemas(f"{rel}: faltan las columnas {', '.join(faltan)}")
        return None
    if df.empty:
        if not opcional:
            problemas(f"{rel}: no tiene filas")
        return None
    return df


def en_escala(df, columna, escala, nombre, problemas, vacio=False):
    """Valores de una columna dentro de su escala (vacío permitido si vacio=True)."""
    for i, v in enumerate(df[columna], start=2):
        if v == "" and vacio:
            continue
        if v not in escala:
            problemas(f"{nombre}, fila {i}: {columna} = «{v}»; valores admitidos: {', '.join(escala)}")


def unicos(df, columna, nombre, problemas):
    rep = df[columna][df[columna].duplicated()].unique()
    if len(rep):
        problemas(f"{nombre}: {columna} repetido: {', '.join(rep)}")


def existentes(df, columna, universo, nombre, problemas, separador=None):
    universo = set(universo)
    for i, v in enumerate(df[columna], start=2):
        valores = lista(v, separador) if separador else [v]
        fuera = [x for x in valores if x and x not in universo]
        if fuera:
            problemas(f"{nombre}, fila {i}: {columna} = «{', '.join(fuera)}» no existe")


def obligatorio(df, columna, nombre, problemas):
    for i, v in enumerate(df[columna], start=2):
        if v == "":
            problemas(f"{nombre}, fila {i}: falta {columna}")


def lista(texto, separador=";"):
    """«1.1.1; 1.1.3» -> ["1.1.1", "1.1.3"]."""
    return [x.strip() for x in str(texto).split(separador) if x.strip()]


def numero(texto):
    """Número de una celda («8,4», «8.4», «1 250») o None si está vacía."""
    t = str(texto).strip().replace(" ", "").replace(" ", "")
    if t == "":
        return None
    if "," in t and "." not in t:
        t = t.replace(",", ".")
    return float(t)


def vacio(x):
    """None en lugar de NaN (pandas convierte los None de una columna numérica)."""
    return None if x is None or pd.isna(x) else x


def numeros(df, columnas, nombre, problemas):
    """Convierte columnas numéricas; informa las celdas que no son números."""
    salida = df.copy()
    for c in columnas:
        if c not in df.columns:
            continue
        valores = []
        for i, v in enumerate(df[c], start=2):
            try:
                valores.append(numero(v))
            except ValueError:
                problemas(f"{nombre}, fila {i}: {c} = «{v}» no es un número")
                valores.append(None)
        salida[c] = valores
    return salida


def criterios(filas):
    """Hoja «Criterios»: (elemento, criterio, plano). Plano: CEPLAN (con página), Literatura o Equipo evaluador."""
    return pd.DataFrame(filas, columns=["Elemento", "Criterio", "Plano"])


def escribir_libro(destino, hojas, criterios_filas):
    """Libro de análisis con sus hojas y la de criterios al final. Nombres de hoja legibles: sin siglas de instrumento
    y de 31 caracteres o menos (ADR-010)."""
    destino = Path(destino)
    destino.parent.mkdir(parents=True, exist_ok=True)
    hojas = {**hojas, "Criterios": criterios(criterios_filas)}
    for nombre in hojas:
        assert len(nombre) <= 31 and not SIGLAS.search(nombre), f"nombre de hoja no publicable: {nombre}"
    with pd.ExcelWriter(destino, engine="openpyxl") as w:
        for nombre, df in hojas.items():
            df.to_excel(w, sheet_name=nombre, index=False)
            ws = w.sheets[nombre]
            for celda in ws[1]:
                celda.font = Font(bold=True, color="FFFFFF")
                celda.fill = CABECERA
                celda.alignment = Alignment(wrap_text=True, vertical="top")
            for j, col in enumerate(df.columns, start=1):
                largo = max([len(str(col))] + [len(str(v)) for v in df[col].head(200)])
                ws.column_dimensions[ws.cell(1, j).column_letter].width = min(max(10, largo * 0.9), 60)
                for celda in ws.iter_cols(min_col=j, max_col=j, min_row=2):
                    for c in celda:
                        c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.freeze_panes = "A2"
    return destino


def anios(cfg):
    p = cfg["periodo"]
    return list(range(int(p["inicio"]), int(p["fin"]) + 1))
