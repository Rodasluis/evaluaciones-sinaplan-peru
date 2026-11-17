"""Descarga las normas de planeamiento nacional y sectorial para el análisis de alineamiento vertical (REV_ALI).

Para El Peruano se obtiene la ruta del PDF desde la página del dispositivo y se verifica que el título contenga el número
de la norma; para gob.pe y otras fuentes se descarga el PDF directo. Cada descarga se registra con su hash. Los
archivos que ya existen en el destino no se vuelven a descargar. Las funciones bajar, guardar, verificar y
de_elperuano las reutiliza descargar_complementos.py.

Uso (desde la raíz del repositorio): python src/fuentes/descargar_planeamiento.py
"""
import hashlib
import re
import subprocess
import tempfile
import urllib.request
from datetime import date
from pathlib import Path

DEST = Path("docs/fuentes/planeamiento")
LOG = Path("docs/fuentes/descargas_planeamiento_2026-10-28.tsv")
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
ELPERUANO = [  # (id del dispositivo, texto que debe aparecer en el PDF, archivo de destino)
    ("2307974-1", "075-2024-PCM", "2024_pcm_ds-075-2024-pcm_modifica-pgg-2023-2026.pdf"),
    ("2500055-1", "043-2026-PCM", "2026_pcm_ds-043-2026-pcm_pgg-seguridad-continuidad-transicion.pdf"),
    ("2551404-3", "127-2026-PCM", "2026_pcm_ds-127-2026-pcm_pgg-2026-2031.pdf"),
    ("2215042-3", "103-2023-PCM", "2023_pcm_ds-103-2023-pcm_actualiza-pedn-2050.pdf"),
]
DIRECTOS = [  # (url, texto que debe aparecer en el PDF, archivo de destino, fuente)
    ("https://cdn.www.gob.pe/uploads/document/file/155522/D-S-N-056-2018-PCM.pdf", "056-2018-PCM",
     "2018_pcm_ds-056-2018-pcm_pgg-al-2021.pdf", "gob.pe (PCM)"),
    ("https://cdn.www.gob.pe/uploads/document/file/2339712/Decreto%20Supremo%20N%C2%B0%20164-2021-PCM.pdf", "164-2021-PCM",
     "2021_pcm_ds-164-2021-pcm_pgg-2021-2026.pdf", "gob.pe (PCM)"),
    ("https://cdn.www.gob.pe/uploads/document/file/4325305/Decreto%20Supremo%20N%C2%B0042%20-2023-PCM%2C.pdf", "042-2023-PCM",
     "2023_pcm_ds-042-2023-pcm_pgg-2023-2026.pdf", "gob.pe (PCM)"),
    ("https://cdn.www.gob.pe/uploads/document/file/4325307/4033369-politica-gobierno-2023-w.pdf", "Política General",
     "2023_pcm_pgg-2023-2026_documento.pdf", "gob.pe (PCM)"),
    ("https://img.lpderecho.pe/wp-content/uploads/2025/12/Decreto-Supremo-141-2025-PCM-LPDerecho.pdf", "141-2025-PCM",
     "2025_pcm_ds-141-2025-pcm_pgg-2025-2026.pdf", "reproducción de El Peruano en LP Derecho (secundaria)"),
    ("https://cdn.www.gob.pe/uploads/document/file/3459220/DS%20N%C2%B0%20095-2022-PCM.pdf.pdf", "095-2022-PCM",
     "2022_pcm_ds-095-2022-pcm_pedn-2050.pdf", "gob.pe (CEPLAN)"),
    ("https://cdn.www.gob.pe/uploads/document/file/3900981/ANEXO%20DS%20N%C2%B0%20095-2022-PCM.pdf.pdf",
     "Plan Estratégico de Desarrollo Nacional", "2022_pcm_ds-095-2022-pcm_anexo-pedn-2050.pdf",
     "gob.pe (PCM), anexo del D.S. 095-2022-PCM"),
    ("https://cdn.www.gob.pe/uploads/document/file/8574273/4637571-tomo-i-plan-estrategico-de-desarrollo-nacional-al-2050.pdf",
     "Plan Estratégico de Desarrollo Nacional", "2023_ceplan_pedn-2050_tomo-i.pdf", "gob.pe (CEPLAN)"),
    ("https://cdn.www.gob.pe/uploads/document/file/8574274/4637571-pedn-tomo-ii-indicadores-del-plan-estrategico-de-desarrollo-"
     "nacional-al-2050.pdf", "Plan Estratégico de Desarrollo Nacional", "2023_ceplan_pedn-2050_tomo-ii-indicadores.pdf",
     "gob.pe (CEPLAN)"),
    ("https://cdn.www.gob.pe/uploads/document/file/6124107/5412627-rm-109-2024-mimp.pdf", "109-2024-MIMP",
     "2024_mimp_rm-109-2024-mimp_aprueba-pesem-2024-2030.pdf", "gob.pe (MIMP)"),
    ("https://cdn.www.gob.pe/uploads/document/file/6124108/5412627-plan-estrategico-sectorial-multianual-pesem-2024-2030-"
     "del-sector-mujer-y-poblaciones-vulnerables.pdf", "Plan Estratégico Sectorial Multianual",
     "2024_mimp_pesem-2024-2030_sector-mujer-y-poblaciones-vulnerables.pdf", "gob.pe (MIMP)"),
]


def bajar(url):
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=180) as r:
        return r.read()


def guardar(contenido, destino, url, verificacion, log):
    assert contenido[:4] == b"%PDF", f"{destino.name}: no es PDF"
    destino.parent.mkdir(parents=True, exist_ok=True)
    destino.write_bytes(contenido)
    paginas = re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", str(destino)], capture_output=True, text=True,
                                                            encoding="latin-1").stdout)
    h = hashlib.sha256(contenido).hexdigest()
    if not log.exists():
        log.write_text("url_origen\tdestino\tsha256\tfecha_descarga\tverificacion_identidad\n", encoding="utf-8")
    with open(log, "a", encoding="utf-8") as f:
        f.write(f"{url}\t{destino.as_posix()}\t{h}\t{date.today()}\t{verificacion}\n")
    print(f"OK {destino.name} ({paginas.group(1) if paginas else '?'} p.)")


def verificar(contenido, clave):
    with tempfile.TemporaryDirectory() as tmp:
        f = Path(tmp) / "x.pdf"
        f.write_bytes(contenido)
        texto = subprocess.run(["pdftotext", "-enc", "UTF-8", str(f), "-"], capture_output=True).stdout.decode("utf-8", "ignore")
    texto = re.sub(r"\s+", " ", texto)
    return "verificado" if clave.lower() in texto.lower() else "NO hallado"


def de_elperuano(nl):
    """Devuelve (url del PDF, título) de un dispositivo de El Peruano."""
    pagina = bajar(f"https://busquedas.elperuano.pe/dispositivo/NL/{nl}").decode("utf-8", "ignore")
    titulo = re.search(r"<title>([^<]+)", pagina).group(1).split(" - ")[0]
    ruta = re.search(r'(/api/archivo/file/[^"\\]+\.PDF)', pagina)
    return "https://busquedas.elperuano.pe" + ruta.group(1), titulo


if __name__ == "__main__":
    for nl, numero, nombre in ELPERUANO:
        if (DEST / nombre).exists():  # no se sobrescriben archivos ya registrados (algunos llevan OCR)
            continue
        url, titulo = de_elperuano(nl)
        contenido = bajar(url.replace("*", "%2A"))
        guardar(contenido, DEST / nombre, url,
                f"El Peruano NL {nl} («{titulo[:90]}»); «{numero}» {verificar(contenido, numero)} en el texto", LOG)
    for url, clave, nombre, fuente in DIRECTOS:
        if (DEST / nombre).exists():
            continue
        contenido = bajar(url)
        guardar(contenido, DEST / nombre, url, f"{fuente}; «{clave}» {verificar(contenido, clave)} en el texto", LOG)
