"""Verifica el enlace público de los documentos del corpus que entregó el usuario y que están publicados en portales
oficiales (gob.pe, El Peruano, SINIA y repositorio de la CEPAL).

Para cada documento declara la página de publicación y el archivo publicado, descarga el archivo en memoria (no lo
guarda) y compara su SHA-256 con el de la copia del repositorio. Escribe docs/fuentes/enlaces_publicos_2026-11-06.tsv,
que lee inventario_fuentes.py para las columnas acceso, enlace_publico, pagina_publicacion y verificacion_enlace.
Los documentos sin publicación identificada (internos o no publicados al 2026-11-06) no se declaran.

Uso (desde la raíz del repositorio): python src/fuentes/verificar_enlaces_publicos.py
"""
import csv
import hashlib
import re
import subprocess
import tempfile
from datetime import date
from pathlib import Path

from descargar_planeamiento import bajar, de_elperuano

SALIDA = Path("docs/fuentes/enlaces_publicos_2026-11-06.tsv")
F = Path("docs/fuentes")
G = "https://www.gob.pe/institucion/"
C = "https://cdn.www.gob.pe/uploads/document/file/"
SYE = G + "mimp/informes-publicaciones/{}-informes-de-seguimiento-y-evaluacion-de-planes-y-politicas-{}"
LIO = G + "mimp/informes-publicaciones/{}-{}-informe-de-{}avances-en-el-cumplimiento-de-la-ley-{}"
LIO_A = "los-"
LIO_L = "n-28983-de-igualdad-de-oportunidades-entre-mujeres-y-hombres-periodo-"
LIO_N = "de-igualdad-de-oportunidades-entre-mujeres-y-hombres-periodo-"

# (documento, página de publicación, archivo publicado)
ENLACES = [
    ("ceplan/5614517-ceplan-guia-para-la-elaboracion-de-indicadores-2024(2).pdf", G + "ceplan/informes-publicaciones/5614517",
     C + "6581312/5614517-ceplan-guia-para-la-elaboracion-de-indicadores-2024%282%29.pdf"),
    ("ceplan/5661161-guia-para-el-planeamiento-institucional-ceplan.pdf", G + "ceplan/informes-publicaciones/5661161",
     C + "6483943/5661161-guia-para-el-planeamiento-institucional-ceplan.pdf"),
    ("ceplan/5678657-guia-para-el-seguimiento-y-evaluacion-de-politicas-nacionales-y-planes-del-sinaplan-actualizada-2024.pdf",
     G + "ceplan/informes-publicaciones/5678657",
     C + "6562639/5678657-guia-para-el-seguimiento-y-evaluacion-de-politicas-nacionales-y-planes-del-sinaplan-actualizada-2024.pdf"),
    ("ceplan/5771193-guia-para-el-planeamiento-estrategico-sectorial-2024.pdf", G + "ceplan/informes-publicaciones/5771193",
     C + "6637908/5771193-guia-para-el-planeamiento-estrategico-sectorial-2024.pdf"),
    ("ceplan/5972882-ceplan-guia-para-el-pdlc-provincial-y-distrital.pdf", G + "ceplan/informes-publicaciones/5972882",
     C + "6918912/5972882-ceplan-guia-para-el-pdlc-provincial-y-distrital.pdf"),
    ("ceplan/5973676-guia-para-el-plan-de-desarrollo-regional-concertado-actualizada-2024.pdf",
     G + "ceplan/informes-publicaciones/5973676",
     C + "6920147/5973676-guia-para-el-plan-de-desarrollo-regional-concertado-actualizada-2024.pdf"),
    ("ceplan/6078850-ceplan-guia-de-prospectiva-para-las-politicas-y-planes-del-sinaplan-actualizada-2024.pdf",
     G + "ceplan/informes-publicaciones/6078850",
     C + "7067037/6078850-ceplan-guia-de-prospectiva-para-las-politicas-y-planes-del-sinaplan-actualizada-2024.pdf"),
    ("ceplan/6088033-ceplan-guia-de-politicas-nacionales-actualizada-2024.pdf", G + "ceplan/informes-publicaciones/6088033",
     C + "7077561/6088033-ceplan-guia-de-politicas-nacionales-actualizada-2024.pdf"),
    ("internacional/2017_cepal_lc-crm-13-5_estrategia-montevideo.pdf", "https://hdl.handle.net/11362/41011",
     "https://repositorio.cepal.org/server/api/core/bitstreams/a13d08d0-4481-434c-8fe7-2eb4f482c306/content"),
    ("pnig/01_normativa/2025_congreso_ley-32535_igualdad-oportunidades-mujeres-hombres.pdf",
     "https://busquedas.elperuano.pe/dispositivo/NL/2471571-3", "elperuano:2471571-3"),
    ("pnig/02_pemig/2020_mimp_ds-002-2020-mimp_pemig.pdf",
     "https://sinia.minam.gob.pe/normas/decreto-supremo-que-aprueba-plan-estrategico-multisectorial-igualdad",
     "https://sinia.minam.gob.pe//sites/default/files/sinia/archivos/public/docs/ds_002-2020-mimp.pdf"),
    # reportes de seguimiento (RS) y de cumplimiento (RC): la página es la de la que proviene la copia del corpus
    ("pnig/04_seguimiento/2020_mimp_reporte-seguimiento-pnig.pdf", G + "mimp/informes-publicaciones/5982262",
     C + "6937006/5982262-reporte-de-seguimiento-de-la-pnig-2020.pdf"),
    ("pnig/04_seguimiento/2020_mimp_reporte-cumplimiento-pnig.pdf", G + "mimp/informes-publicaciones/1982352-reporte-de-"
     "cumplimiento-2020-de-la-politica-nacional-multisectorial-de-igualdad-de-genero",
     C + "1967947/Reporte-de-Cumplimiento-PNIG-2020.pdf"),
    ("pnig/04_seguimiento/2021_mimp_reporte-seguimiento-pnig.pdf", G + "mimp/informes-publicaciones/5986171",
     C + "6936858/5986171-reporte-de-seguimiento-pnig-2021.pdf"),
    ("pnig/04_seguimiento/2021_mimp_reporte-cumplimiento-pnig.pdf", SYE.format(4327522, 2021),
     C + "4702064/1-Reporte-de-cumplimiento-PNIG-2021.pdf"),
    ("pnig/04_seguimiento/2022_mimp_reporte-seguimiento-pnig.pdf", G + "mimp/informes-publicaciones/5986450",
     C + "6937275/5986450-1-reporte-de-seguimiento-pnig-2022.pdf"),
    ("pnig/04_seguimiento/2022_mimp_reporte-cumplimiento-pnig.pdf", SYE.format(4327760, 2022),
     C + "4702408/4-Reporte-de-cumplimiento-PNIG-2022.pdf"),
    ("pnig/04_seguimiento/2023_mimp_reporte-seguimiento-pnig.pdf", SYE.format(5622014, 2023),
     C + "6490986/5622014-pnig-rs-2023-f-f.pdf"),
    ("pnig/04_seguimiento/2023_mimp_reporte-cumplimiento-pnig.pdf", SYE.format(5622014, 2023),
     C + "6490969/5622014-pnig-rc-2023-f-f.pdf"),
    ("pnig/04_seguimiento/2024_mimp_reporte-seguimiento-pnig.pdf", SYE.format(6869059, 2024),
     C + "8221972/6869059-reporte_seguimiento_pnig_2024.pdf"),
    ("pnig/04_seguimiento/2024_mimp_reporte-cumplimiento-pnig.pdf", SYE.format(6869059, 2024),
     C + "8221968/6869059-reporte_cumplimiento_pnig_2024.pdf"),
    ("pnig/04_seguimiento/2025_mimp_reporte-seguimiento-pnig.pdf", SYE.format(8241376, 2025),
     C + "10161080/8241376-rs_pnig-_2025-formato-mimp.pdf"),
    ("pnig/04_seguimiento/2025_mimp_reporte-seguimiento-pnig_aplicativo-ceplan.pdf", SYE.format(8241376, 2025),
     C + "10161079/8241376-rs_pnig-_2025-aplicativo-ceplan.pdf"),
    ("pnig/04_seguimiento/2025_mimp_reporte-cumplimiento-pnig.pdf", SYE.format(8241376, 2025),
     C + "10120227/8241376-rc_pnig_2025.pdf"),
    ("pnig/04_seguimiento/fuentes_indicadores/2026_onpe_analitica-electoral-03_resultados-eg-2026.pdf",
     G + "onpe/informes-publicaciones/8436407",
     C + "10392349/8436407-ae-03-resultados-electorales-eg-2026-votos-escanos-y-perfil-de-autoridades%282%29.pdf"),
    ("pnig/05_evaluaciones_previas/2021_mimp_informe-evaluacion-resultados-pnig.pdf", SYE.format(4327522, 2021),
     C + "4702211/2_ER_PNIG-2021.pdf"),
    ("pnig/05_evaluaciones_previas/2023_mimp_informe-evaluacion-pnig.pdf", SYE.format(5622014, 2023),
     C + "7465240/5622014-informe_de_evaluacion_2023_de_la_politica_nacional_de_igualdad_de_genero.pdf"),
    # informes anuales de cumplimiento de la LIO (colección 16 del MIMP en gob.pe)
    ("pnig/06_informes_lio/2011_mimp_informe-cumplimiento-lio.pdf",
     LIO.format(196565, "v", LIO_A, "n-28983-de-igualdad-de-oportunidades-entre-mujeres-y-hombres-balance-2007-2011"),
     C + "199110/v_informe_avances_igualdad.pdf"),
    ("pnig/06_informes_lio/2012_mimp_informe-cumplimiento-lio.pdf", LIO.format(196566, "vi", LIO_A, LIO_L + "2012"),
     C + "199118/vi_informe_avances_igualdad.pdf"),
    ("pnig/06_informes_lio/2013_mimp_informe-cumplimiento-lio.pdf", LIO.format(196567, "vii", LIO_A, LIO_L + "2013"),
     C + "300196/d199119_opt.pdf"),
    ("pnig/06_informes_lio/2014_mimp_informe-cumplimiento-lio.pdf", LIO.format(196571, "viii", LIO_A, LIO_L + "2014"),
     C + "300198/d199123_opt.pdf"),
    ("pnig/06_informes_lio/2015_mimp_informe-cumplimiento-lio.pdf", LIO.format(196572, "ix", LIO_A, LIO_L + "2015"),
     C + "199124/2015_ix_Informe_LIO.pdf"),
    ("pnig/06_informes_lio/2016_mimp_informe-cumplimiento-lio.pdf", LIO.format(200810, "x", LIO_A, LIO_L + "2016"),
     C + "204459/X_INFORME_LIO__.pdf"),
    ("pnig/06_informes_lio/2017_mimp_informe-cumplimiento-lio.pdf", LIO.format(441037, "xi", LIO_A, LIO_L + "2017"),
     C + "521624/xi-Informe-LIO-2017.pdf"),
    ("pnig/06_informes_lio/2018_mimp_informe-cumplimiento-lio.pdf", LIO.format(441038, "xii", LIO_A, LIO_L + "2018"),
     C + "521625/xii-Informe-LIO-2018.pdf"),
    ("pnig/06_informes_lio/2019_mimp_informe-cumplimiento-lio.pdf", LIO.format(2169054, "xiii", LIO_A, LIO_L + "2019"),
     C + "2201730/XIII-Informe-LIO-2019.pdf"),
    ("pnig/06_informes_lio/2020_mimp_informe-cumplimiento-lio.pdf", LIO.format(2169058, "xiv", LIO_A, LIO_L + "2020"),
     C + "2201738/XIV-Informe-LIO-2020.pdf"),
    ("pnig/06_informes_lio/2021_mimp_informe-cumplimiento-lio.pdf", LIO.format(3316719, "xv", LIO_A, LIO_L + "2021"),
     C + "3474197/XV-Informe-LIO-2021.pdf"),
    ("pnig/06_informes_lio/2022_mimp_informe-cumplimiento-lio.pdf", LIO.format(4953974, "xvi", "", LIO_N + "2022"),
     C + "5571431/4953974-xvi-informe-lio-2022.pdf"),
    ("pnig/06_informes_lio/2023_mimp_informe-cumplimiento-lio.pdf", LIO.format(6237149, "xvii", "", LIO_N + "2023"),
     C + "7299235/6237149-xvii_informe_lio_2023.pdf"),
    ("pnig/06_informes_lio/2024_mimp_informe-cumplimiento-lio.pdf", LIO.format(7005343, "xviii", "", LIO_N + "2024"),
     C + "8436309/7005343-xviii_informe_lio_2024.pdf"),
]
# versión publicada distinta de la del corpus: (páginas publicadas, páginas del corpus, descripción)
OTRA_VERSION = {
    "pnig/05_evaluaciones_previas/2021_mimp_informe-evaluacion-resultados-pnig.pdf": (
        49, 54, "La versión publicada (49 p.) difiere de la del corpus (54 p.), que empieza con el Memorando N.° "
                "0180-2022-MTPE/4 y tiene otra paginación; las citas del informe de cierre usan la copia del corpus"),
}


def paginas(contenido):
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as t:
        t.write(contenido)
    try:
        out = subprocess.run(["pdfinfo", t.name], capture_output=True, text=True, encoding="latin-1").stdout
    finally:
        Path(t.name).unlink()
    m = re.search(r"Pages:\s+(\d+)", out)
    return int(m.group(1)) if m else None


filas = []
for doc, pagina, archivo in ENLACES:
    local = (F / doc).read_bytes()
    if archivo.startswith("elperuano:"):
        archivo, _ = de_elperuano(archivo.split(":")[1])
    publicado = bajar(archivo.replace("*", "%2A"))
    h = hashlib.sha256(publicado).hexdigest()
    if doc in OTRA_VERSION:
        p_pub, p_loc, texto = OTRA_VERSION[doc]
        assert h != hashlib.sha256(local).hexdigest() and paginas(publicado) == p_pub and paginas(local) == p_loc, doc
        acceso, verificacion = "público (otra versión)", texto
    else:
        assert h == hashlib.sha256(local).hexdigest(), f"{doc}: el archivo publicado no es idéntico a la copia"
        acceso, verificacion = "público", "SHA-256 idéntico al del archivo publicado"
    filas.append([f"docs/fuentes/{doc}", pagina, archivo, acceso, verificacion, h, date.today().isoformat()])
    print(acceso, doc, flush=True)

with open(SALIDA, "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f, delimiter="\t")
    w.writerow(["documento", "pagina_publicacion", "enlace_archivo", "acceso", "verificacion", "sha256_publicado",
                "fecha_verificacion"])
    w.writerows(filas)
print("OK", len(filas), "enlaces")
