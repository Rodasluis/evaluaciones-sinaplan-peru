"""Descarga las políticas nacionales afines a la PNIG para el alineamiento horizontal (REV_ALI.3) y la base de datos de
políticas nacionales del CEPLAN.

Universo (nota metodológica v14, numeral 5.3):
  - sucesoras vigentes de los planes y políticas con los que la PNIG declaró su alineamiento horizontal en 2019
    (PNAIA 2012–2021, Plan Nacional de Seguridad Ciudadana 2013–2018, política frente a la trata, Plan Nacional de
    Derechos Humanos 2018–2021 y Plan de Acción en Género y Cambio Climático);
  - políticas aprobadas del registro del CEPLAN que comparten población objetivo o servicios con la PNIG.
Los vínculos salen de la base de datos del CEPLAN (corte al 15-10-2026). En tres filas la base cruza los vínculos: en
Salud y en Pueblos Indígenas el del documento y el del decreto están intercambiados, y en Derechos Humanos el vínculo
apunta a la Estrategia ante el Cambio Climático; para esta última se usa el documento publicado en gob.pe.
Usa las funciones de descargar_planeamiento.py, no sobrescribe archivos existentes y registra cada descarga con su hash.

Uso (desde la raíz del repositorio): python src/fuentes/descargar_politicas_afines.py
"""
import hashlib
from datetime import date
from pathlib import Path

from descargar_planeamiento import bajar, guardar, verificar

LOG = Path("docs/fuentes/descargas_politicas-afines_2026-10-30.tsv")
D = Path("docs/fuentes/politicas_afines")
CDN = "https://cdn.www.gob.pe/uploads/document/file/"
BASE_CEPLAN = (CDN + "9351001/7665898-base-de-datos-politicas-nacionales-de-los-sectores-al-15-10-2026.xlsx?v=1792092865",
               Path("docs/fuentes/ceplan/2026_ceplan_base-datos-politicas-nacionales_al-15-10-2026.xlsx"))
DIRECTOS = [  # (url, texto que debe aparecer en el PDF, destino)
    (CDN + "1991214/Politica-Nacional-Multisectorial-para-las-ninas-ninos-y-adolescentes-2030.pdf?v=1625257749",
     "Niñas, Niños y Adolescentes", D / "2021_mimp_ds-008-2021-mimp_pnm-ninas-ninos-adolescentes-al-2030.pdf"),
    (CDN + "3390077/ANEXO%20006-2022-IN.pdf?v=1657555077", "Seguridad Ciudadana",
     D / "2022_mininter_ds-006-2022-in_pnm-seguridad-ciudadana-al-2030.pdf"),
    (CDN + "2050294/Decreto%20Supremo%20N%C2%B0%20009-2021-IN.pdf?v=1628273259", "Trata de Personas",
     D / "2021_mininter_ds-009-2021-in_pn-trata-de-personas-al-2030.pdf"),
    (CDN + "7949468/6684278-politica-nacional-multisectorial-de-derechos-humanos-al-2040.pdf", "Derechos Humanos",
     D / "2024_minjusdh_ds-018-2024-jus_pnm-derechos-humanos-al-2040.pdf"),
    (CDN + "7767750/6562767-politica-nacional-estrategia-nacional-ante-cambio-climatico-al-2050.pdf?v=1741880813",
     "Cambio Climático", D / "2024_minam_ds-012-2024-minam_estrategia-cambio-climatico-al-2050.pdf"),
    (CDN + "1257169/Pol%C3%ADtica%20Nacional%20Multisectorial%20de%20Salud%20al%202030%20.pdf?v=1598319240",
     "Salud", D / "2020_minsa_ds-026-2020-sa_pnm-salud-al-2030.pdf"),
    (CDN + "1948868/Pol%C3%ADtica%20Nacional%20de%20Empleo%20Decente%20-%20PED.pdf?v=1623777587", "Empleo Decente",
     D / "2021_mtpe_ds-013-2021-tr_pn-empleo-decente.pdf"),
    (CDN + "349461/DS255_2019EF.pdf?v=1565109724", "Inclusión Financiera",
     D / "2019_mef_ds-255-2019-ef_pn-inclusion-financiera.pdf"),
    (CDN + "2071814/DECRETO%20SUPREMO%2017-2021-MIDAGRI.pdf?v=1628609014", "Agraria",
     D / "2021_midagri_ds-017-2021-midagri_pn-agraria-2021-2030.pdf"),
    (CDN + "4034822/Pol%C3%ADtica%20Nacional%20de%20Desarrollo%20e%20Inclusi%C3%B3n%20Social%20PNDIS%20al%202030_parte01.pdf.pdf"
     "?v=1673390374", "Inclusión Social", D / "2022_midis_ds-008-2022-midis_pn-desarrollo-inclusion-social-al-2030_parte-1.pdf"),
    # El registro del CEPLAN enlaza solo la parte 1; las partes 2 a 4 (fichas de servicios y anexos) están en la página
    # del decreto: gob.pe/institucion/midis/normas-legales/3823404-008-2022-midis (consultada el 2026-10-30).
    *[(CDN + f"{4034821 + n}/Pol%C3%ADtica%20Nacional%20de%20Desarrollo%20e%20Inclusi%C3%B3n%20Social%20PNDIS%20al%202030"
       f"_parte0{n}.pdf.pdf?v=1673390374", "Inclusión Social",
       D / f"2022_midis_ds-008-2022-midis_pn-desarrollo-inclusion-social-al-2030_parte-{n}.pdf") for n in (2, 3, 4)],
    (CDN + "1934191/Pol%C3%ADtica-Nacional-Multisectorial-para-las-Personas-Adultas-Mayores-al-2030.pdf?v=1623110006",
     "Adultas Mayores", D / "2021_mimp_ds-006-2021-mimp_pnm-personas-adultas-mayores-al-2030.pdf"),
    (CDN + "1932186/POL%C3%8DTICA%20NACIONAL%20MULTISECTORIAL%20EN%20DISCAPACIDAD%20PARA%20EL%20DESARROLLO%20AL%202030..pdf"
     "?v=1622920983", "Discapacidad", D / "2021_mimp_ds-007-2021-mimp_pnm-discapacidad-para-el-desarrollo-al-2030.pdf"),
    (CDN + "375207/DS_N__013-2019-MINEDU__NL_DL_.pdf?v=1569964630", "Juventud",
     D / "2019_minedu_ds-013-2019-minedu_pn-juventud.pdf"),
    (CDN + "300210/d268038_opt.pdf?v=1553182372", "Ámbitos Rurales",
     D / "2018_minedu_ds-013-2018-minedu_pn-atencion-educativa-ambitos-rurales.pdf"),
    (CDN + "1275742/DS%20N%C2%B0%20012-2020-MINEDU%20Normas%20Legales.pdf?v=1598932942", "Educación Superior",
     D / "2020_minedu_ds-012-2020-minedu_pn-educacion-superior-tecnico-productiva.pdf"),
    (CDN + "3531092/POL%C3%8DTICA%20NACIONAL%20DE%20MODERNIZACI%C3%93N%20DE%20LA%20GESTI%C3%93N%20P%C3%9ABLICA%20AL%202030"
     "%281%29.pdf.pdf", "Modernización", D / "2022_pcm_ds-103-2022-pcm_pn-modernizacion-gestion-publica-al-2030.pdf"),
    (CDN + "3202061/PNPA%20al%202030.pdf.pdf?v=1654272957", "Afroperuano",
     D / "2022_mincul_ds-005-2022-mc_pn-pueblo-afroperuano-al-2030.pdf"),
    ("https://repositorio.cultura.pe/postmaster/2026/PNPI%20AL%202040_VF_%5bF%5d.pdf", "Pueblos Indígenas",
     D / "2026_mincul_ds-002-2026-mc_pn-pueblos-indigenas-u-originarios-al-2040.pdf"),
    (CDN + "2017413/Pol%C3%ADtica%20Nacional%20de%20Vivienda%20y%20Urbanismo.pdf?v=1626384939", "Vivienda",
     D / "2021_mvcs_ds-012-2021-vivienda_pn-vivienda-urbanismo-al-2030.pdf"),
    (CDN + "8462829/7003736-anexo-politica-nacional-de-pesca-al-2040-parte-1.pdf?v=1754684034", "Pesca",
     D / "2025_produce_ds-012-2025-produce_pn-pesca-al-2040_parte-1.pdf"),
]

if __name__ == "__main__":
    url, destino = BASE_CEPLAN
    if not destino.exists():
        contenido = bajar(url)
        assert contenido[:2] == b"PK", "la base de datos no es un .xlsx"
        destino.write_bytes(contenido)
        with open(LOG, "a", encoding="utf-8") as f:
            if f.tell() == 0:
                f.write("url_origen\tdestino\tsha256\tfecha_descarga\tverificacion_identidad\n")
            f.write(f"{url}\t{destino.as_posix()}\t{hashlib.sha256(contenido).hexdigest()}\t{date.today()}\t"
                    "gob.pe (CEPLAN), «Base de datos | Políticas Nacionales de los Sectores», última actualización "
                    "del 19-03-2026 según la página; corte del archivo al 15-10-2026\n")
        print(f"OK {destino.name}")
    for url, clave, destino in DIRECTOS:
        if destino.exists():
            print(f"ya existe {destino.name}")
            continue
        try:
            contenido = bajar(url)
        except Exception as e:  # se informa y se sigue con las demás
            print(f"ERROR {destino.name}: {e}")
            continue
        fuente = "repositorio.cultura.pe" if "cultura.pe" in url else "gob.pe"
        guardar(contenido, destino, url, f"{fuente}; «{clave}» {verificar(contenido, clave)} en el texto", LOG)
