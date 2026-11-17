"""Descarga de publicaciones oficiales para el análisis de brechas (D1 y D2) y registro con su hash.

Baja cada archivo con la verificación TLS activa (con agente de usuario de navegador solo para gob.pe y el portal de
leyes del Congreso, que rechazan otros clientes), comprueba que sea un PDF (o el tipo
esperado) y lo registra en docs/fuentes/descargas_fuentes-brechas_<fecha>.tsv (regla 7 de CLAUDE.md).

Uso (desde la raíz del repositorio):
    uv run --no-project --with requests python -I -X utf8 src/fuentes/descargar_fuentes_brechas.py
"""
import csv
import datetime as dt
import hashlib
import os
import time

import requests

REGISTRO = f"docs/fuentes/descargas_fuentes-brechas_{dt.date.today().isoformat()}.tsv"
DESTINO = "docs/fuentes/pnig/04_seguimiento/fuentes_indicadores"
ARCHIVOS = [  # (url, nombre de destino, firma esperada, descripción)
    ("https://cdn.www.gob.pe/uploads/document/file/9089792/"
     "7463589-peru-brechas-de-genero-2025-avances-hacia-la-igualdad-de-mujeres-y-hombres.pdf",
     "2025_inei_brechas-de-genero-2025.pdf", b"%PDF",
     "INEI, «Perú: Brechas de Género, 2025: Avances hacia la igualdad de mujeres y hombres» (2-12-2025)"),
    ("https://portal.jne.gob.pe/portal_documentos/files/7035a123-eb9a-4be0-9c57-72003e46cc3c.pdf",
     "2025_jne_resolucion-0053-2025-jne_distribucion-escanos-eg2026.pdf", b"%PDF",
     "JNE, Resolución N.° 0053-2025-JNE, distribución de escaños del Senado y la Cámara de Diputados, EG 2026"),
    ("https://portal.jne.gob.pe/portal_documentos/files/b9f4ffa0-c103-48eb-be0f-7b76309e630d.pdf",
     "2026_jne_resolucion-1411-2026-jne_proclamacion-diputados.pdf", b"%PDF",
     "JNE, Resolución N.° 1411-2026-JNE, proclamación de los diputados electos para el periodo 2026-2031 (19-6-2026)"),
    ("https://busquedas.elperuano.pe/api/archivo/file/C-0OsDvs4d8AIlIRsdEz-0/*/2528732-1.PDF",
     "2026_jne_resolucion-1412-2026-jne_proclamacion-senadores.pdf", b"%PDF",
     "JNE, Resolución N.° 1412-2026-JNE, proclamación de los senadores electos para el periodo 2026-2031 (El Peruano)"),
    ("https://leyes.congreso.gob.pe/Documentos/Leyes/29470.pdf",
     "2009_congreso_ley-29470_modifica-ley-27683-elecciones-regionales.pdf", b"%PDF",
     "Congreso, Ley N.° 29470, modifica el artículo 5 de la Ley N.° 27683: umbral de 30 % y segunda elección "
     "regional (El Peruano, 14-12-2009, p. 407816; escaneo sin capa de texto)"),
    ("https://www.gob.pe/institucion/jne/noticias/1455074-el-6-de-diciembre-sera-la-segunda-eleccion-regional",
     "2026_jne_nota-prensa_segunda-eleccion-regional-6-diciembre.html", b"<!DOCTYPE",
     "JNE, nota de prensa del 10-10-2026: el Acuerdo del Pleno fija el 6-12-2026 para la Segunda Elección Regional "
     "donde ninguna fórmula obtenga no menos del 30 % de los votos válidos"),
]
# gob.pe y leyes.congreso.gob.pe rechazan los clientes sin agente de usuario de navegador (CLAUDE.md, «Directorio de
# funcionarios»)
AGENTE = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                        "Chrome/130.0 Safari/537.36"}


def main():
    nuevo = not os.path.exists(REGISTRO)
    with open(REGISTRO, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter="\t")
        if nuevo:
            w.writerow(["url_origen", "destino", "sha256", "fecha_descarga", "bytes", "verificacion_identidad"])
        for url, nombre, firma, desc in ARCHIVOS:
            destino = f"{DESTINO}/{nombre}"
            if os.path.exists(destino):
                print("ya existe", destino)
                continue
            time.sleep(2)
            r = requests.get(url, timeout=600, headers=AGENTE if ("www.gob.pe" in url or "congreso.gob.pe" in url) else None)
            if r.status_code != 200 or not r.content.lstrip()[:len(firma)].upper() == firma.upper():
                print("FALLA", url, r.status_code, r.headers.get("content-type"))
                continue
            with open(destino, "wb") as g:
                g.write(r.content)
            sha = hashlib.sha256(r.content).hexdigest()
            w.writerow([url, destino, sha, dt.date.today().isoformat(), len(r.content), desc])
            print(f"{nombre}: {len(r.content) / 1e6:.1f} MB, sha256 {sha[:16]}…")


if __name__ == "__main__":
    main()
