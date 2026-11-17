"""Descarga los documentos que la nota metodológica cita y que faltaban en el corpus (revisión del 2026-10-28).

Comprende normas del Sinaplan, la Ley N.° 28983, el reporte de seguimiento 2025 en formato del aplicativo del CEPLAN,
instrumentos de planeamiento del periodo (Plan Bicentenario, Políticas de Estado, PESEM y PEI del MIMP) e instrumentos
internacionales del alineamiento (REV_ALI). Usa las funciones de descargar_planeamiento.py, no sobrescribe archivos
existentes y registra cada descarga con su hash.

Uso (desde la raíz del repositorio): python src/fuentes/descargar_complementos.py
"""
from pathlib import Path

from descargar_planeamiento import bajar, de_elperuano, guardar, verificar

LOG = Path("docs/fuentes/descargas_complementos_2026-10-28.tsv")
F = Path("docs/fuentes")
CDN = "https://cdn.www.gob.pe/uploads/document/file/"
ELPERUANO = [  # (id del dispositivo, texto que debe aparecer en el PDF, destino)
    ("2516641-1", "076-2026-PCM", F / "ceplan/2026_pcm_ds-076-2026-pcm_modifica-reglamento-politicas-nacionales.pdf"),
    ("2325725-1", "099-2024", F / "ceplan/2024_ceplan_rpcd-099-2024_modifica-guia-sye.pdf"),
    ("2327969-1", "101-2024", F / "ceplan/2024_ceplan_rpcd-101-2024_modifica-directiva-y-guia-sye.pdf"),
]
DIRECTOS = [  # (url, texto que debe aparecer en el PDF, destino, fuente)
    # Sinaplan
    (CDN + "1904126/Reglamento-Politicas-Nacionales%20DS%20029-2018-PCM.pdf.pdf", "029-2018-PCM",
     F / "ceplan/2018_pcm_ds-029-2018-pcm_reglamento-politicas-nacionales.pdf", "gob.pe (PCM)"),
    (CDN + "6380197/5373394-directiva-001-2024-ceplan-pcd-modificada-directiva-general-de-planeamiento-estrategico-del-"
     "sinaplan-pdf.pdf", "Directiva General de Planeamiento Estratégico",
     F / "ceplan/2024_ceplan_directiva-001-2024_planeamiento-estrategico-modificada.pdf", "gob.pe (CEPLAN)"),
    (CDN + "10185438/8285878-rpcd_lista-de-pn-y-lineamiento-para-modificacion-de-pn-oajff.pdf", "042-2026",
     F / "ceplan/2026_ceplan_rpcd-042-2026_lista-pn-y-lineamientos-modificacion.pdf", "gob.pe (CEPLAN)"),
    (CDN + "10185439/8285878-anexo-1-lista-de-politicas-nacionalesf.pdf", "Igualdad de Género",
     F / "ceplan/2026_ceplan_rpcd-042-2026_anexo-1-lista-politicas-nacionales.pdf", "gob.pe (CEPLAN)"),
    (CDN + "10185440/8285878-anexo-2-lineamientosf.pdf", "lineamientos",
     F / "ceplan/2026_ceplan_rpcd-042-2026_anexo-2-lineamientos.pdf", "gob.pe (CEPLAN)"),
    # PNIG
    (CDN + "105124/_28983_-_07-11-2012_06_33_15_-LEY-28983.pdf", "28983",
     F / "pnig/01_normativa/2007_congreso_ley-28983_igualdad-oportunidades.pdf", "gob.pe (Congreso)"),
    ("https://www2.congreso.gob.pe/sicr/cendocbib/con2_uibd.nsf/09861DEC915B9C1F0525784D00608583/$FILE/Ley_28983.pdf",
     "LEY Nº 28983", F / "pnig/01_normativa/2007_congreso_ley-28983_texto-con-concordancias.pdf",
     "congreso.gob.pe (texto con concordancias; la edición oficial es la de El Peruano)"),
    (CDN + "10161079/8241376-rs_pnig-_2025-aplicativo-ceplan.pdf", "Igualdad de Género",
     F / "pnig/04_seguimiento/2025_mimp_reporte-seguimiento-pnig_aplicativo-ceplan.pdf", "gob.pe (MIMP)"),
    # Planeamiento del periodo
    ("https://sinia.minam.gob.pe/sites/default/files/sinia/archivos/public/docs/2526.pdf", "Plan Bicentenario",
     F / "planeamiento/2011_ceplan_plan-bicentenario-peru-hacia-el-2021.pdf", "SINIA (MINAM)"),
    ("https://acuerdonacional.pe/wp-content/uploads/2024/01/Politicas_de_Estado_2023.pdf", "Políticas de Estado",
     F / "planeamiento/2023_acuerdo-nacional_politicas-de-estado-vision-2050.pdf", "acuerdonacional.pe"),
    (CDN + "375422/pesem_2018_2020-dcto-prospectivo-al-2030.pdf", "Plan Estratégico Sectorial Multianual",
     F / "planeamiento/2018_mimp_pesem-2018-2022_documento-prospectivo-2030.pdf", "gob.pe (MIMP)"),
    (CDN + "4201736/RM_128-2021-MIMP-PESEM.pdf", "128-2021-MIMP",
     F / "planeamiento/2021_mimp_rm-128-2021-mimp_amplia-pesem-2018-2024.pdf", "gob.pe (MIMP)"),
    (CDN + "4329708/rm_106_2022_mimp.pdf", "106-2022-MIMP",
     F / "planeamiento/2022_mimp_rm-106-2022-mimp_aprueba-pesem-2018-2026.pdf", "gob.pe (MIMP)"),
    (CDN + "4329709/anexo_PESEM_2018_2026.pdf", "2018",
     F / "planeamiento/2022_mimp_pesem-2018-2026_anexo.pdf", "gob.pe (MIMP)"),
    (CDN + "4208420/pei_2019-2022.pdf", "Plan Estratégico Institucional",
     F / "planeamiento/2019_mimp_pei-2019-2022.pdf", "gob.pe (MIMP)"),
    (CDN + "4226799/rm_126_2022_mimp.pdf", "126-2022-MIMP",
     F / "planeamiento/2022_mimp_rm-126-2022-mimp_aprueba-pei-2019-2025.pdf", "gob.pe (MIMP)"),
    (CDN + "4226800/Plan%20Estrat%C3%A9gico%20Institucional%20-%20PEI%202019-2025%20del%20Ministerio%20de%20la%20Mujer%20y%20"
     "Poblaciones%20Vulnerables.pdf", "Plan Estratégico Institucional",
     F / "planeamiento/2022_mimp_pei-2019-2025.pdf", "gob.pe (MIMP)"),
    (CDN + "4221572/rm_097_2023_mimp.pdf", "097-2023-MIMP",
     F / "planeamiento/2023_mimp_rm-097-2023-mimp_aprueba-pei-2019-2025-modificado.pdf", "gob.pe (MIMP)"),
    (CDN + "4221573/Plan%20Estrat%C3%A9gico%20Institucional%202019%20-%202025%20Modificado%20del%20Pliego%20039%3A%20"
     "Ministerio%20de%20la%20Mujer%20y%20Poblaciones%20Vulnerables.pdf", "Plan Estratégico Institucional",
     F / "planeamiento/2023_mimp_pei-2019-2025-modificado.pdf", "gob.pe (MIMP)"),
    (CDN + "6351504/5576545-rm_180_2024_mimp.pdf", "180-2024-MIMP",
     F / "planeamiento/2024_mimp_rm-180-2024-mimp_amplia-pei-2019-2027.pdf", "gob.pe (MIMP)"),
    (CDN + "6351505/5576545-plan-estrategico-institucional-2019-2027-ampliado-del-pliego-039-ministerio-de-la-mujer-y-"
     "poblaciones-vulnerables.pdf", "Plan Estratégico Institucional",
     F / "planeamiento/2024_mimp_pei-2019-2027-ampliado.pdf", "gob.pe (MIMP)"),
    (CDN + "7569709/6427622-plan-estrategico-institucional-2025-2030-del-mimp.pdf", "2025",
     F / "planeamiento/2025_mimp_rm-024-2025-mimp_aprueba-pei-2025-2030.pdf", "gob.pe (MIMP)"),
    (CDN + "7569710/6427622-anexo-rm-n-024-2025-mimp.pdf", "Plan Estratégico Institucional",
     F / "planeamiento/2025_mimp_pei-2025-2030.pdf", "gob.pe (MIMP)"),
    # Instrumentos internacionales
    ("https://www.ohchr.org/sites/default/files/Documents/ProfessionalInterest/cedaw_SP.pdf", "discriminación contra la mujer",
     F / "internacional/1979_onu_cedaw_convencion.pdf", "ACNUDH (ohchr.org)"),
    ("https://documents.un.org/doc/undoc/gen/n22/273/98/pdf/n2227398.pdf", "CEDAW/C/PER/CO/9",
     F / "internacional/2022_onu-cedaw_cedaw-c-per-co-9_observaciones-finales-peru.pdf", "ONU (documents.un.org)"),
    ("https://www.oas.org/es/mesecvi/docs/BelemDoPara-ESPANOL.pdf", "Belém do Pará",
     F / "internacional/1994_oea_convencion-belem-do-para.pdf", "OEA (MESECVI)"),
    ("https://www.unwomen.org/sites/default/files/Headquarters/Attachments/Sections/CSW/BPA_S_Final_WEB.pdf", "Beijing",
     F / "internacional/1995_onu_declaracion-y-plataforma-de-accion-de-beijing.pdf", "ONU Mujeres (unwomen.org)"),
    ("https://documents.un.org/doc/undoc/gen/n15/291/93/pdf/n1529193.pdf", "A/RES/70/1",
     F / "internacional/2015_onu_a-res-70-1_agenda-2030.pdf", "ONU (documents.un.org)"),
]

if __name__ == "__main__":
    for nl, numero, destino in ELPERUANO:
        if destino.exists():
            continue
        url, titulo = de_elperuano(nl)
        contenido = bajar(url.replace("*", "%2A"))
        guardar(contenido, destino, url,
                f"El Peruano NL {nl} («{titulo[:90]}»); «{numero}» {verificar(contenido, numero)} en el texto", LOG)
    for url, clave, destino, fuente in DIRECTOS:
        if destino.exists():
            continue
        try:
            contenido = bajar(url)
        except Exception as e:  # una fuente caída no detiene el resto
            print(f"FALLA {destino.name}: {e}")
            continue
        guardar(contenido, destino, url, f"{fuente}; «{clave}» {verificar(contenido, clave)} en el texto", LOG)
