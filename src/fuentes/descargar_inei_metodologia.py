"""Descarga de los documentos metodológicos del INEI que sustentan la desagregación departamental (D1 y D2).

Cada documento acredita el nivel de inferencia de una encuesta: ENAHO, ENDES, EPEN, ENA, ENARES 2024 y ENUT 2024;
la Cuenta Satélite de la Economía Informal da la regla del empleo informal con la ENAHO.
Baja cada archivo con la verificación TLS activa y sin suplantar un navegador, comprueba que sea un PDF y lo registra
en docs/fuentes/descargas_inei-metodologia_<fecha>.tsv (regla 7 de CLAUDE.md).

Uso (desde la raíz del repositorio):
    uv run --no-project --with requests python -I -X utf8 src/fuentes/descargar_inei_metodologia.py
"""
import csv
import datetime as dt
import hashlib
import os
import time

import requests

REGISTRO = f"docs/fuentes/descargas_inei-metodologia_{dt.date.today().isoformat()}.tsv"
DESTINO = "docs/fuentes/inei_metodologia"
SRIENAHO = "https://proyectos.inei.gob.pe/iinei/srienaho/Descarga/DocumentosMetodologicos"
LIBROS = "https://www.inei.gob.pe/media/MenuRecursivo/publicaciones_digitales/Est"
ARCHIVOS = [  # (url, nombre de destino, descripción)
    (f"{SRIENAHO}/2025-63/01_Ficha_Tecnica.pdf", "2025_inei_enaho_ficha-tecnica.pdf",
     "INEI, Ficha técnica de la Encuesta Nacional de Hogares 2025"),
    ("https://proyectos.inei.gob.pe/endes/2025/DOCUMENTOS/FICHA_TECNICA_ENDES_2025.pdf",
     "2025_inei_endes_ficha-tecnica.pdf", "INEI, Ficha técnica de la ENDES 2025"),
    (f"{LIBROS}/Lib2029/libro.pdf", "2025_inei_epen_mercado-laboral-departamental-2022-2024.pdf",
     "INEI, «Perú: Indicadores del Mercado Laboral a nivel departamental, 2022-2024» (EPEN, julio de 2025)"),
    (f"{SRIENAHO}/2024-62/03_FICHA_TECNICA_ENA_2024.pdf", "2024_inei_ena_ficha-tecnica.pdf",
     "INEI, Ficha técnica de la Encuesta Nacional Agropecuaria 2024"),
    (f"{LIBROS}/Lib2031/libro.pdf", "2025_inei_enares-2024_principales-resultados.pdf",
     "INEI, «Informe de los Principales Resultados, ENARES 2024» (julio de 2025)"),
    (f"{LIBROS}/Lib2010/libro.pdf", "2025_inei_enut-2024_principales-resultados.pdf",
     "INEI, «Principales Resultados, ENUT - 2024»"),
    ("https://cdn.www.gob.pe/uploads/document/file/9215525/7564428-libro.pdf",
     "2025_inei_produccion-empleo-informal-cuenta-satelite-2022-2024.pdf",
     "INEI, «Producción y Empleo Informal en el Perú. Cuenta Satélite de la Economía Informal 2022-2024» (31-12-2025); "
     "anexo A2.2, pp. 87–90: regla del empleo informal con la ENAHO"),
]


def main():
    os.makedirs(DESTINO, exist_ok=True)
    nuevo = not os.path.exists(REGISTRO)
    with open(REGISTRO, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter="\t")
        if nuevo:
            w.writerow(["url_origen", "destino", "sha256", "fecha_descarga", "bytes", "verificacion_identidad"])
        for url, nombre, desc in ARCHIVOS:
            destino = f"{DESTINO}/{nombre}"
            if os.path.exists(destino):
                print("ya existe", destino)
                continue
            time.sleep(2)
            r = requests.get(url, timeout=600)
            if r.status_code != 200 or not r.content.startswith(b"%PDF"):
                print("FALLA", url, r.status_code, r.headers.get("content-type"))
                continue
            with open(destino, "wb") as g:
                g.write(r.content)
            sha = hashlib.sha256(r.content).hexdigest()
            w.writerow([url, destino, sha, dt.date.today().isoformat(), len(r.content), desc])
            print(f"{nombre}: {len(r.content) / 1e6:.1f} MB, sha256 {sha[:16]}…")


if __name__ == "__main__":
    main()
