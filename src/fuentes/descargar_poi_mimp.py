"""Descarga del portal gob.pe un documento del Plan Operativo Institucional (POI) del MIMP por año, 2019–2025 (REV_IND.3).

Para cada año se usa la última versión del POI anual publicada en la colección del MIMP, que trae la programación física
y financiera por centro de costo y actividad operativa; para 2023–2025 se añade el reporte de seguimiento anual del POI,
que informa la ejecución física. Se añaden el POI consistente con el PIA de 2021 a 2025 (versión de apertura, publicada en
diciembre del año anterior) y el informe de evaluación de resultados del PEI y del POI 2022. Sirven para contrastar con la Consulta Amigable
las actividades vinculadas a la conducción de la PNIG (DGIGND), que el clasificador presupuestal no separa por órgano.
Colección: https://www.gob.pe/institucion/mimp/colecciones/16877-plan-operativo-institucional-poi
Usa las funciones de descargar_planeamiento.py, no sobrescribe archivos existentes y registra cada descarga con su hash.

Uso (desde la raíz del repositorio): python src/fuentes/descargar_poi_mimp.py
"""
from pathlib import Path

from descargar_planeamiento import bajar, guardar, verificar

LOG = Path("docs/fuentes/descargas_poi-mimp_2026-10-29.tsv")
CDN = "https://cdn.www.gob.pe/uploads/document/file/"
POI = Path("docs/fuentes/planeamiento/poi_mimp")
DIRECTOS = [  # (url, texto que debe aparecer en el PDF, destino)
    (CDN + "4214277/Plan%20Operativo%20Institucional%202019.pdf", "Operativo", POI / "2019_mimp_poi-anual-2019.pdf"),
    (CDN + "1504527/POI-2020-Modificado-version2.pdf.pdf", "Operativo", POI / "2020_mimp_poi-anual-2020_modificado-version-2.pdf"),
    (CDN + "1762035/Plan%20Operativo%20Institucional%20%28POI%29%202021%20Anual%20Modificado%20Versi%C3%B3n%201%20del%20MIMP.pdf",
     "Operativo", POI / "2021_mimp_poi-anual-2021_modificado-version-1.pdf"),
    (CDN + "3457428/anexo-POI-2022-Modificado-Version-1.pdf", "Operativo", POI / "2022_mimp_poi-anual-2022_modificado-version-1.pdf"),
    (CDN + "5634743/3967745-plan-operativo-institucional-poi-anual-2023-modificado-version-2.pdf", "Operativo",
     POI / "2023_mimp_poi-anual-2023_modificado-version-2.pdf"),
    (CDN + "7413525/6318250-poi_anual_2024_actualizado_version3.pdf", "Operativo",
     POI / "2024_mimp_poi-anual-2024_actualizado-version-3.pdf"),
    (CDN + "9231777/7575955-plan-operativo-institucional-poi-anual-2025-actualizado-version-3.pdf", "Operativo",
     POI / "2025_mimp_poi-anual-2025_actualizado-version-3.pdf"),
    (CDN + "6133315/3967745-reporte-de-seguimiento-anual-2023-del-plan-operativo-institucional-poi-publicado-01-04-2024.pdf",
     "seguimiento", POI / "2024_mimp_reporte-seguimiento-anual-poi-2023.pdf"),
    (CDN + "7832649/4990971-reporte_seguimiento_poi_pliego_2024.pdf", "seguimiento",
     POI / "2025_mimp_reporte-seguimiento-anual-poi-2024.pdf"),
    (CDN + "9701317/7937188-reporte-de-seguimiento-anual-2025-del-plan-operativo-institucional-poi.pdf", "seguimiento",
     POI / "2026_mimp_reporte-seguimiento-anual-poi-2025.pdf"),
    # POI consistente con el PIA (versión de apertura, comparable con el PIA de la Consulta Amigable). El de 2020 no está
    # publicado en la colección: solo la R.M. N.° 336-2019-MIMP que lo aprueba, sin anexo.
    (CDN + "4214915/POI-2021-consistenciado-con-el-PIA-2021.pdf", "Apertura",
     POI / "2020_mimp_poi-2021-consistente-con-el-pia.pdf"),
    (CDN + "4214959/Plan%20Operativo%20Institucional%202022.pdf", "APERTURA",
     POI / "2021_mimp_poi-2022-consistente-con-el-pia.pdf"),
    (CDN + "4215084/Plan%20Operativo%20Institucional%202023.pdf", "APERTURA",
     POI / "2022_mimp_poi-2023-consistente-con-el-pia.pdf"),
    (CDN + "5634707/4990971-plan-operativo-institucional-poi-2024-consistente-con-el-presupuesto-institucional-de-"
     "apertura-pia-2024.pdf", "Operativo", POI / "2023_mimp_poi-2024-consistente-con-el-pia.pdf"),
    (CDN + "7413544/6318263-poi_consistente_con_pia_2025%282%29.pdf", "Operativo",
     POI / "2024_mimp_poi-2025-consistente-con-el-pia.pdf"),
    # Evaluación de resultados del PEI y del POI 2022: ejecución física y financiera por AEI y por actividad operativa
    (CDN + "4500082/1-Informe-de-Evaluacion-de-Resultados-PEI-POI-correspondiente-2022-MIMP.pdf", "Evaluación",
     POI / "2023_mimp_informe-evaluacion-resultados-pei-poi-2022.pdf"),
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
        guardar(contenido, destino, url, f"gob.pe (MIMP); «{clave}» {verificar(contenido, clave)} en el texto", LOG)
