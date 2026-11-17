"""Descarga de microdatos del INEI para el análisis de brechas de la evaluación de la PNIG (D1 y D2).

Baja del portal oficial (proyectos.inei.gob.pe) los módulos que no están en disco, con la verificación TLS activa, una
pausa entre archivos y sin suplantar un navegador. Cada zip se guarda en su propia carpeta y se descomprime en una
subcarpeta nueva, sin rutas absolutas ni «..» (los archivos descargados son datos no confiables). Cada descarga se
registra con su hash en docs/fuentes/descargas_microdatos-inei_<fecha>.tsv (regla 7 de CLAUDE.md).

Los códigos de archivo salen del catálogo del portal (src/fuentes/catalogo_microdatos_inei.py). Lo que ya existe en
disco no se vuelve a bajar: ENAHO 2018–2023 y ENDES 2021–2023 están en la carpeta local de microdatos del usuario, y
ENDES 2021–2025 (módulos 1631, 1635 y 1637) en su repositorio «Indicadores ENDES».

Uso (desde la raíz del repositorio):
    uv run --no-project --with requests python src/fuentes/descargar_microdatos_inei.py
"""
import csv
import datetime as dt
import hashlib
import io
import os
import time
import zipfile

import requests
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

URL = "https://proyectos.inei.gob.pe/iinei/srienaho/descarga/{formato}/{archivo}.zip"
RAIZ = f"{EVAL}/03_datos/raw/microdatos_inei"
REGISTRO = f"docs/fuentes/descargas_microdatos-inei_{dt.date.today().isoformat()}.tsv"
PAUSA = 3

# (encuesta, año, archivo, formato, motivo)
ARCHIVOS = [
    *[("ENAHO", "2016", f"546-Modulo{m}", "STATA", "Brechas del diagnóstico de la PNIG (ENAHO 2016)") for m in ("03", "05")],
    *[("ENAHO", "2017", f"603-Modulo{m}", "STATA", "Año del diagnóstico de la PNIG") for m in ("02", "03", "05", "34")],
    ("ENAHO", "2018", "634-Modulo05", "STATA", "Validación del IOP.03.02 con la serie de SERVIR (la carpeta local no "
                                               "trae la base del módulo 500 de 2018)"),
    *[("ENAHO", "2024", f"966-Modulo{m}", "STATA", "Serie 2019–2025") for m in ("02", "03", "05", "34")],
    ("ENAHO", "2024", "966-Modulo04", "STATA", "Seguro de salud (p419a) para el empleo informal"),
    *[("ENAHO", "2025", f"1031-Modulo{m}", "STATA", "Serie 2019–2025") for m in ("02", "03", "05", "34")],
    ("ENAHO", "2025", "1031-Modulo04", "STATA", "Seguro de salud (p419a) para el empleo informal"),
    *[("ENDES", "2019", f"691-Modulo{m}", "SPSS", "Serie 2019–2025, mujeres") for m in ("66", "67", "71", "73")],
    *[("ENDES", "2020", f"739-Modulo{m}", "SPSS", "Serie 2019–2025, mujeres") for m in ("1631", "1632", "1635", "1637")],
    *[("ENDES", a, f"{c}-Modulo1632", "STATA", "Anticoncepción y demanda insatisfecha") for a, c in
      (("2021", "760"), ("2022", "786"), ("2023", "910"), ("2024", "968"), ("2025", "1038"))],
    *[("ENARES", "2015", f"503-Modulo{m}", "SPSS", "Tolerancia social, línea de base") for m in range(761, 766)],
    *[("ENARES", "2019", f"694-Modulo{m}", "STATA", "Tolerancia social") for m in range(1515, 1519)],
    *[("ENARES", "2024", f"976-Modulo{m}", "STATA", "Tolerancia social") for m in range(1951, 1956)],
    *[("ENUT", "2010", f"286-Modulo{m}", "SPSS", "Uso del tiempo, diagnóstico de la PNIG") for m in range(125, 131)],
    *[("ENUT", "2024", f"963-Modulo{m}", "STATA", "Uso del tiempo") for m in range(1849, 1856)],
    ("EPEN", "2022", "790-Modulo76", "SPSS", "Informalidad y empleo, base anual departamental"),
    *[("EPEN", a, f"{c}-Modulo76", "STATA", "Informalidad y empleo, base anual departamental") for a, c in
      (("2023", "874"), ("2024", "935"), ("2025", "1001"))],
    *[("ENA", a, f"{c}-Modulo{m}", "STATA", "Productoras y productores agropecuarios: tierra y crédito")
      for a, c, mods in (("2019", "701", (1547, 1548, 1545, 1529)), ("2021", "868", (1807, 1808, 1793)),
                         ("2022", "844", (1754, 1755, 1752, 1737)), ("2023", "922", (1846, 1828, 1844)),
                         ("2024", "973", (1911, 1893, 1909)), ("2025", "1036", (1911, 1893, 1909))) for m in mods],
]


def extraer_seguro(zip_path, destino):
    """Descomprime sin seguir rutas absolutas ni salir de la carpeta de destino."""
    os.makedirs(destino, exist_ok=True)
    base = os.path.realpath(destino)
    nombres = []
    with zipfile.ZipFile(zip_path) as z:
        for info in z.infolist():
            ruta = os.path.realpath(os.path.join(destino, info.filename))
            if not ruta.startswith(base + os.sep) and ruta != base:
                raise ValueError(f"ruta fuera de la carpeta en {zip_path}: {info.filename}")
            if info.is_dir():
                continue
            z.extract(info, destino)
            nombres.append(info.filename)
    return nombres


def main():
    nuevo = not os.path.exists(REGISTRO)
    with open(REGISTRO, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter="\t")
        if nuevo:
            w.writerow(["url_origen", "destino", "sha256", "fecha_descarga", "bytes", "verificacion_identidad"])
        for enc, anio, archivo, formato, motivo in ARCHIVOS:
            carpeta = f"{RAIZ}/{enc.lower()}/{anio}"
            zip_path = f"{carpeta}/{archivo}.zip"
            if os.path.exists(zip_path) and os.path.getsize(zip_path) > 0:
                print("ya existe", zip_path)
                continue
            os.makedirs(carpeta, exist_ok=True)
            url = URL.format(formato=formato, archivo=archivo)
            time.sleep(PAUSA)
            r = requests.get(url, timeout=600)
            if r.status_code != 200 or r.content[:2] != b"PK":
                print("FALLA", url, r.status_code, r.headers.get("content-type"))
                continue
            sha = hashlib.sha256(r.content).hexdigest()
            try:
                with zipfile.ZipFile(io.BytesIO(r.content)) as prueba:
                    danado = prueba.testzip()
            except zipfile.BadZipFile:
                danado = "archivo zip"
            if danado:
                w.writerow([url, "(no se guarda)", sha, dt.date.today().isoformat(), len(r.content),
                            f"{enc} {anio}: el archivo del portal está dañado (CRC en {danado}); no se usa"])
                print("DAÑADO EN EL ORIGEN", archivo, danado)
                continue
            with open(zip_path, "wb") as z:
                z.write(r.content)
            nombres = extraer_seguro(zip_path, f"{carpeta}/{archivo}")
            datos = [n.split("/")[-1] for n in nombres if n.lower().endswith((".dta", ".sav"))]
            w.writerow([url, zip_path, sha, dt.date.today().isoformat(), len(r.content),
                        f"{enc} {anio}, {motivo}. Contiene: {'; '.join(datos)}"])
            f.flush()
            print(f"{archivo}: {len(r.content) / 1e6:.1f} MB, {len(datos)} bases")


if __name__ == "__main__":
    main()
