"""Descarga al corpus las evaluaciones anteriores de la PNIG publicadas en gob.pe que no estaban en el corpus.

La evaluación de implementación 2020 (MIMP, junio de 2021) se identificó el 2026-11-06 en la colección «Seguimiento
de la Política Nacional Multisectorial de Igualdad de Género» del MIMP (gob.pe, colección 3111); el usuario decidió
incorporarla a la sección 3.7 del informe (seguimiento de las recomendaciones de evaluaciones anteriores).
Usa las funciones de descargar_planeamiento.py, no sobrescribe archivos existentes y registra cada descarga con su hash.

Uso (desde la raíz del repositorio): python src/fuentes/descargar_evaluaciones_previas.py
"""
from pathlib import Path

from descargar_planeamiento import bajar, guardar, verificar

LOG = Path("docs/fuentes/descargas_evaluaciones-previas_2026-11-06.tsv")
D = Path("docs/fuentes/pnig/05_evaluaciones_previas")
DESCARGAS = [  # (url, texto que debe aparecer en el PDF, destino, fuente)
    ("https://cdn.www.gob.pe/uploads/document/file/1958363/Evaluacion-implementacion-PNIG-2020.pdf",
     "INFORME DE EVALUACIÓN DE IMPLEMENTACIÓN", D / "2020_mimp_informe-evaluacion-implementacion-pnig.pdf",
     "gob.pe, MIMP, informes-publicaciones/1973531 («Evaluación de implementación 2020 de la Política Nacional "
     "Multisectorial de Igualdad de Género (PNIG)»)"),
]

if __name__ == "__main__":
    for url, clave, destino, fuente in DESCARGAS:
        if destino.exists():
            continue
        contenido = bajar(url)
        guardar(contenido, destino, url, f"{fuente}; «{clave}» {verificar(contenido, clave)} en el texto", LOG)
