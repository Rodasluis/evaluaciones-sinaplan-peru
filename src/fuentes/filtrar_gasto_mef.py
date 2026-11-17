"""Filtra el conjunto «Presupuesto y Ejecución de Gasto» del Portal de Datos Abiertos del MEF (REV_IND.3, verificación de las
cifras presupuestales reportadas en los informes LIO y en el formulario de seguimiento).

El conjunto es la Consulta Amigable de Gasto en formato abierto, con el detalle que la Consulta Amigable no muestra en
pantalla: secuencia funcional (meta), finalidad y nombre de la meta, por fuente de financiamiento y genérica de gasto. La
meta permite ubicar el gasto de un órgano o de una intervención cuando la entidad lo identifica en su nombre.

Cada archivo anual (un CSV de 8 a 10 GB dentro de un ZIP de 0,5 a 0,7 GB) se lee por bloques desde el ZIP, sin
descomprimirlo, y se guardan tres extractos:
  1. gobierno nacional completo, sumado por meta, fuente y genérica (todos los meses);
  2. gobiernos regionales y locales, solo en los programas presupuestales de PROGRAMAS, sumados por pliego, producto,
     actividad y fuente;
  3. totales por programa presupuestal y nivel de gobierno, para todos los programas.
Montos en soles: PIA, PIM, certificación, compromiso anual, devengado y girado. Los ZIP quedan fuera del repositorio;
el hash de cada ZIP y de cada extracto se registra en el log.

Fuente: https://datosabiertos.mef.gob.pe/dataset/presupuesto-y-ejecucion-de-gasto
Uso (desde la raíz del repositorio; <dir_zip> fuera del repositorio: los ZIP 2019–2025 están en
<carpeta administrativa>/datos_abiertos_mef_gasto; ruta en config/entorno.local.yaml, clave carpeta_administrativa):
    uv run --no-project --with pandas --with requests python src/fuentes/filtrar_gasto_mef.py <dir_zip> 2019 2020 ...
"""
import hashlib
import sys
import zipfile
from datetime import date
from pathlib import Path

import pandas as pd
import requests
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

URL = "https://fs.datosabiertos.mef.gob.pe/datastorefiles/{nombre}.zip"
NOMBRES = {2025: "2025-Gasto-Mensual"}  # desde 2025 el archivo anual se publica con el sufijo «Mensual»
DESTINO = Path(f"{EVAL}/03_datos/raw/datos_abiertos_mef")
LOG = Path(f"docs/fuentes/descargas_datos-abiertos-mef_{date.today().isoformat()}.tsv")
PROGRAMAS = ["0002", "0080", "0103", "0116", "0131", "1002"]
MONTOS = ["MONTO_PIA", "MONTO_PIM", "MONTO_CERTIFICADO", "MONTO_COMPROMETIDO_ANUAL", "MONTO_DEVENGADO", "MONTO_GIRADO"]
NACIONAL = ["ANO_EJE", "SECTOR", "SECTOR_NOMBRE", "PLIEGO", "PLIEGO_NOMBRE", "SEC_EJEC", "EJECUTORA", "EJECUTORA_NOMBRE",
            "SEC_FUNC", "PROGRAMA_PPTO", "PROGRAMA_PPTO_NOMBRE", "TIPO_ACT_PROY", "PRODUCTO_PROYECTO",
            "PRODUCTO_PROYECTO_NOMBRE", "ACTIVIDAD_ACCION_OBRA", "ACTIVIDAD_ACCION_OBRA_NOMBRE", "FUNCION",
            "DIVISION_FUNCIONAL", "GRUPO_FUNCIONAL", "META", "FINALIDAD", "META_NOMBRE", "DEPARTAMENTO_META_NOMBRE",
            "FUENTE_FINANCIAMIENTO", "FUENTE_FINANCIAMIENTO_NOMBRE", "GENERICA", "GENERICA_NOMBRE"]
SUBNACIONAL = ["ANO_EJE", "NIVEL_GOBIERNO", "PLIEGO", "PLIEGO_NOMBRE", "PROGRAMA_PPTO", "PROGRAMA_PPTO_NOMBRE",
               "PRODUCTO_PROYECTO", "PRODUCTO_PROYECTO_NOMBRE", "ACTIVIDAD_ACCION_OBRA", "ACTIVIDAD_ACCION_OBRA_NOMBRE",
               "FUENTE_FINANCIAMIENTO", "FUENTE_FINANCIAMIENTO_NOMBRE"]
TOTALES = ["ANO_EJE", "NIVEL_GOBIERNO", "PROGRAMA_PPTO", "PROGRAMA_PPTO_NOMBRE"]


def sha(p):
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for b in iter(lambda: f.read(1 << 24), b""):
            h.update(b)
    return h.hexdigest()


def bajar(url, zip_):
    if zip_.exists():
        return
    with requests.get(url, stream=True, timeout=120, headers={"User-Agent": "Mozilla/5.0"}) as r:
        r.raise_for_status()
        with open(zip_.with_suffix(".part"), "wb") as f:
            for b in r.iter_content(1 << 22):
                f.write(b)
    zip_.with_suffix(".part").rename(zip_)


def sumar(partes, claves):
    return pd.concat(partes).groupby(claves, as_index=False, dropna=False)[MONTOS].sum()


if __name__ == "__main__":
    dir_zip = Path(sys.argv[1])
    DESTINO.mkdir(parents=True, exist_ok=True)
    if not LOG.exists():
        LOG.write_text("destino\turl_origen\tfecha_descarga\tsha256_zip\tfilas\tsha256\n", encoding="utf-8")
    for anio in map(int, sys.argv[2:]):
        nombre = NOMBRES.get(anio, f"{anio}-Gasto")
        url, zip_ = URL.format(nombre=nombre), dir_zip / f"{nombre}.zip"
        salidas = {k: DESTINO / f"{anio}_mef_gasto_{k}.csv"
                   for k in ("gobierno-nacional_meta", "programas-pnig_subnacional", "programas_totales")}
        if all(s.exists() for s in salidas.values()):
            continue
        bajar(url, zip_)
        z = zipfile.ZipFile(zip_)
        nac, sub, tot = [], [], []
        with z.open(z.infolist()[0]) as f:
            for b in pd.read_csv(f, dtype=str, chunksize=2_000_000, encoding="utf-8", keep_default_na=False):
                b[MONTOS] = b[MONTOS].apply(pd.to_numeric, errors="coerce").fillna(0)
                nac.append(b[b.NIVEL_GOBIERNO == "E"].groupby(NACIONAL, as_index=False)[MONTOS].sum())
                s = b[(b.NIVEL_GOBIERNO != "E") & b.PROGRAMA_PPTO.isin(PROGRAMAS)]
                sub.append(s.groupby(SUBNACIONAL, as_index=False)[MONTOS].sum())
                tot.append(b.groupby(TOTALES, as_index=False)[MONTOS].sum())
                print(f"  {anio}: {sum(len(x) for x in nac)} filas parciales del gobierno nacional", flush=True)
        h_zip = sha(zip_)
        for (k, s), (partes, claves) in zip(salidas.items(), [(nac, NACIONAL), (sub, SUBNACIONAL), (tot, TOTALES)]):
            df = sumar(partes, claves)
            df.to_csv(s, index=False, encoding="utf-8")
            with open(LOG, "a", encoding="utf-8") as f:
                f.write(f"{s.as_posix()}\t{url}\t{date.today().isoformat()}\t{h_zip}\t{len(df)}\t{sha(s)}\n")
            print(f"OK {s.name}: {len(df)} filas")
