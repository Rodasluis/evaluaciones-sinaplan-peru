"""Descarga las fuentes de datos de indicadores de la PNIG y la guía de políticas nacionales anterior a la de 2024.

Comprende los informes anuales de SERVIR «La mujer en el servicio civil peruano» (ediciones 2017 a 2026), fuente del
IOP.03.02 (participación de servidoras civiles funcionarias y directivas), el anexo de la RPCD 101-2024 y la Guía de
Políticas Nacionales de 2018 en
su versión modificada, para contrastar con la de 2024 cómo se regulaba la modificación de indicadores (REV_IND y D5).
Usa las funciones de descargar_planeamiento.py, no sobrescribe archivos existentes y registra cada descarga con su hash.

Uso (desde la raíz del repositorio): python src/fuentes/descargar_fuentes_indicadores.py
"""
from pathlib import Path

from descargar_planeamiento import bajar, guardar, verificar

LOG = Path("docs/fuentes/descargas_fuentes-indicadores_2026-10-29.tsv")
CDN = "https://cdn.www.gob.pe/uploads/document/file/"
SERVIR = Path("docs/fuentes/pnig/04_seguimiento/fuentes_indicadores")
DIRECTOS = [  # (url, texto que debe aparecer en el PDF, destino, fuente)
    (CDN + "1062917/Informe_La_Mujer_en_el_Servicio_Civil_Peruano_2017.pdf", "servicio civil",
     SERVIR / "2017_servir_informe-mujer-servicio-civil-peruano-2017.pdf", "gob.pe (SERVIR)"),
    (CDN + "1062843/Informe_La_Mujer_en_el_Servicio_Civil_Peruano_2018.pdf", "servicio civil",
     SERVIR / "2018_servir_informe-mujer-servicio-civil-peruano-2018.pdf", "gob.pe (SERVIR)"),
    (CDN + "1062810/Informe_La_Mujer_en_el_Servicio_Civil_Peruano_2019.pdf", "servicio civil",
     SERVIR / "2019_servir_informe-mujer-servicio-civil-peruano-2019.pdf", "gob.pe (SERVIR)"),
    (CDN + "1062754/Informe_La_Mujer_en_el_Servicio_Civil_Peruano_2020.pdf", "servicio civil",
     SERVIR / "2020_servir_informe-mujer-servicio-civil-peruano-2020.pdf", "gob.pe (SERVIR)"),
    (CDN + "1709915/Informe_La_Mujer_en_el_Servicio_Civil_Peruano_2021.pdf.pdf", "servicio civil",
     SERVIR / "2021_servir_informe-mujer-servicio-civil-peruano-2021.pdf", "gob.pe (SERVIR)"),
    (CDN + "2888781/La%20Mujer%20en%20el%20Servicio%20Civil%20Peruano%202022.pdf", "servicio civil",
     SERVIR / "2022_servir_informe-mujer-servicio-civil-peruano-2022.pdf", "gob.pe (SERVIR)"),
    (CDN + "4229152/La%20Mujer%20en%20el%20servicio%20civil%20peruano%202023_SERVIR.pdf", "servicio civil",
     SERVIR / "2023_servir_informe-mujer-servicio-civil-peruano-2023.pdf", "gob.pe (SERVIR)"),
    (CDN + "6014370/5328453-informe-la-mujer-en-el-servicio-civil-peruano%282%29.pdf", "servicio civil",
     SERVIR / "2024_servir_informe-mujer-servicio-civil-peruano-2024.pdf", "gob.pe (SERVIR)"),
    (CDN + "7738049/6543313-estudio-de-la-mujer-en-el-servicio-civil-2025.pdf", "servicio civil",
     SERVIR / "2025_servir_estudio-mujer-servicio-civil-2025.pdf", "gob.pe (SERVIR)"),
    (CDN + "9560669/7823960-la-mujer-en-el-servicio-civil-peruano-2026%282%29.pdf", "servicio civil",
     SERVIR / "2026_servir_informe-mujer-servicio-civil-peruano-2026.pdf", "gob.pe (SERVIR)"),
    (CDN + "2426527/Gu%C3%ADa%20de%20Pol%C3%ADticas%20Nacionales.pdf", "Políticas Nacionales",
     Path("docs/fuentes/ceplan/2018_ceplan_guia-politicas-nacionales_modificada-rpcd-057-2018.pdf"), "gob.pe (CEPLAN)"),
    # Anexo de la RPCD 101-2024, que modifica la aprobación del informe de evaluación (ruta del informe de cierre)
    (CDN + "6989703/6024831-anexos.pdf", "evaluación",
     Path("docs/fuentes/ceplan/2024_ceplan_rpcd-101-2024_anexo-aprobacion-informe-evaluacion.pdf"), "gob.pe (CEPLAN)"),
]

if __name__ == "__main__":
    for url, clave, destino, fuente in DIRECTOS:
        if destino.exists():
            continue
        try:
            contenido = bajar(url)
        except Exception as e:  # una fuente caída no detiene el resto
            print(f"FALLA {destino.name}: {e}")
            continue
        guardar(contenido, destino, url, f"{fuente}; «{clave}» {verificar(contenido, clave)} en el texto", LOG)
