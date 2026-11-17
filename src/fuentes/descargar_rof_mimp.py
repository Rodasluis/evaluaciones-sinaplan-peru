"""Descarga del portal gob.pe el Reglamento de Organización y Funciones (ROF) del MIMP vigente en cada tramo de 2019 a
2025, con la norma que lo aprueba y su organigrama (REV_IND.3 e I4).

Sirve para ubicar, año por año, los órganos que condujeron la PNIG y sus centros de costo en el POI: la estructura cambió
con el ROF de 2021 (D.S. N.° 012-2021-MIMP), el de 2023 (D.S. N.° 010-2023-MIMP) y las modificaciones de 2024 y 2025.
Compendio: https://www.gob.pe/institucion/mimp/colecciones/64-rof-reglamento-de-organizacion-y-funciones-mimp
Usa las funciones de descargar_planeamiento.py, no sobrescribe archivos existentes y registra cada descarga con su hash.

Uso (desde la raíz del repositorio): PYTHONPATH=src/fuentes python src/fuentes/descargar_rof_mimp.py
"""
from pathlib import Path

from descargar_planeamiento import bajar, guardar, verificar

LOG = Path("docs/fuentes/descargas_rof-mimp_2026-10-29.tsv")
CDN = "https://cdn.www.gob.pe/uploads/document/file/"
ROF = Path("docs/fuentes/planeamiento/rof_mimp")
DIRECTOS = [  # (url, texto que debe aparecer en el PDF, destino)
    # ROF 2012 con las modificaciones de 2015, vigente hasta julio de 2021
    (CDN + "315536/ds_003_2012_mimp.pdf", "Reglamento", ROF / "2012_mimp_ds-003-2012-mimp_aprueba-rof.pdf"),
    (CDN + "314884/ds-002-2015-mimp.pdf", "Reglamento", ROF / "2015_mimp_ds-002-2015-mimp_modifica-rof.pdf"),
    (CDN + "314885/ds-004-2015-mimp.pdf", "Reglamento", ROF / "2015_mimp_ds-004-2015-mimp_modifica-rof.pdf"),
    (CDN + "314882/rof-mimp.pdf", "Igualdad", ROF / "2016_mimp_rof-texto-con-modificaciones-2015.pdf"),
    # ROF 2021
    (CDN + "2016789/ds_012_2021_mimp.pdf", "Reglamento", ROF / "2021_mimp_ds-012-2021-mimp_aprueba-rof-seccion-primera.pdf"),
    (CDN + "2041904/rm_201_2021_mimp.pdf", "Segunda", ROF / "2021_mimp_rm-201-2021-mimp_aprueba-rof-seccion-segunda.pdf"),
    (CDN + "2069134/rm_208_2021_mimp.pdf", "Integrado", ROF / "2021_mimp_rm-208-2021-mimp_aprueba-texto-integrado-rof.pdf"),
    (CDN + "2069135/Texto%20Integrado%20del%20Reglamento%20de%20Organizaci%C3%B3n%20y%20Funciones.pdf", "Igualdad",
     ROF / "2021_mimp_rof-texto-integrado.pdf"),
    # ROF 2023
    (CDN + "5247711/Decreto-Supremo-010-2023-mimp.pdf", "Reglamento",
     ROF / "2023_mimp_ds-010-2023-mimp_aprueba-rof-seccion-primera.pdf"),
    (CDN + "5302224/4758990-rm_342_2023_mimp.pdf", "Segunda", ROF / "2023_mimp_rm-342-2023-mimp_aprueba-rof-seccion-segunda.pdf"),
    (CDN + "5305995/4758990-anexo2-seccion-ii-rof-organigrama.pdf", "", ROF / "2023_mimp_rof-organigrama.pdf"),
    (CDN + "5334399/4781405-resolucion-ministerial-n-362-2023-mimp.pdf", "Integrado",
     ROF / "2023_mimp_rm-362-2023-mimp_aprueba-texto-integrado-rof.pdf"),
    (CDN + "5334400/4781405-documento-del-texto-integrado-del-reglamento-de-organizacion-y-funciones-del-ministerio-de-"
     "la-mujer-y-poblaciones-vulnerables.pdf", "Igualdad", ROF / "2023_mimp_rof-texto-integrado.pdf"),
    # Modificaciones de 2024 y 2025
    (CDN + "6949837/5995340-rm_294_2024_mimp.pdf", "Segunda", ROF / "2024_mimp_rm-294-2024-mimp_modifica-rof-seccion-segunda.pdf"),
    (CDN + "6949839/5995340-anexo_2-organigrama.pdf", "", ROF / "2024_mimp_rof-organigrama.pdf"),
    (CDN + "6993367/6028069-rm_306_2024_mimp.pdf", "Integrado", ROF / "2024_mimp_rm-306-2024-mimp_aprueba-texto-integrado-rof.pdf"),
    (CDN + "6993368/6028069-documento-del-texto-integrado-del-reglamento-de-organizacion-y-funciones-2024.pdf", "Igualdad",
     ROF / "2024_mimp_rof-texto-integrado.pdf"),
    (CDN + "8814721/7279785-resolucion-ministerial-000367-2025-mimp.pdf", "Segunda",
     ROF / "2025_mimp_rm-367-2025-mimp_modifica-rof-arts-82-83-104.pdf"),
    (CDN + "8889857/7328510-rm-n-380-2025-mimp.pdf", "Integrado", ROF / "2025_mimp_rm-380-2025-mimp_aprueba-texto-integrado-rof.pdf"),
    (CDN + "8889858/7328510-anexo-rm-n-380-2025-mimp.pdf", "Igualdad", ROF / "2025_mimp_rof-texto-integrado.pdf"),
]

if __name__ == "__main__":
    for url, clave, destino in DIRECTOS:
        if destino.exists():
            continue
        try:
            contenido = bajar(url)
        except Exception as e:  # una fuente caída no detiene el resto
            print(f"FALLA {destino.name}: {e}")
            continue
        verificacion = f"gob.pe (MIMP); «{clave}» {verificar(contenido, clave)} en el texto" if clave else "gob.pe (MIMP); imagen"
        guardar(contenido, destino, url, verificacion, LOG)
