"""OCR de PDF escaneados: agrega una capa de texto invisible sobre el PDF original, sin tocar sus imágenes.

Cada página sin texto (menos de 20 palabras) se rasteriza a 300 ppp en escala de grises con pdftoppm y se procesa con
Tesseract (español e inglés) en modo solo texto; esa capa se superpone a la página original con pypdf. Las páginas que ya
tienen texto quedan como están, y el tamaño del archivo apenas cambia. El original queda en el historial de git o en su
URL de descarga, y ambos hashes se registran en el log.

Requiere Tesseract (C:\\Program Files\\Tesseract-OCR), el modelo spa.traineddata en ~/.local/share/tessdata (con
configs/pdf y pdf.ttf copiados desde la carpeta tessdata de Tesseract) y pypdf.

Uso: python src/fuentes/ocr_pdf.py <log.tsv> <pdf> [<pdf> ...]
"""
import hashlib
import subprocess
import sys
import tempfile
from datetime import date
from pathlib import Path

from pypdf import PdfReader, PdfWriter, Transformation

TESSERACT = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
TESSDATA = str(Path.home() / ".local" / "share" / "tessdata")
LOG = Path(sys.argv[1])
MIN_PALABRAS = 20


def sha(p):
    return hashlib.sha256(Path(p).read_bytes()).hexdigest()


def palabras(pdf, pagina=None):
    rango = ["-f", str(pagina), "-l", str(pagina)] if pagina else []
    out = subprocess.run(["pdftotext", "-enc", "UTF-8", *rango, str(pdf), "-"], capture_output=True).stdout
    return len(out.decode("utf-8", "ignore").split())


if not LOG.exists():
    LOG.write_text("archivo\tsha256_original\tsha256_ocr\tpaginas_con_ocr\tpalabras_agregadas\tfecha\n", encoding="utf-8")
for arg in sys.argv[2:]:
    pdf = Path(arg)
    n = len(PdfReader(pdf).pages)
    sin_texto = [i for i in range(1, n + 1) if palabras(pdf, i) < MIN_PALABRAS]
    if not sin_texto:
        print(f"SIN CAMBIOS {pdf.name}: todas las páginas tienen texto")
        continue
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        imgs = []
        for i in sin_texto:
            subprocess.run(["pdftoppm", "-r", "300", "-gray", "-png", "-f", str(i), "-l", str(i), "-singlefile",
                            str(pdf), str(tmp / f"p{i:04d}")], check=True)
            imgs.append(tmp / f"p{i:04d}.png")
        (tmp / "lista.txt").write_text("\n".join(str(x) for x in imgs), encoding="utf-8")
        subprocess.run([TESSERACT, str(tmp / "lista.txt"), str(tmp / "capa"), "--tessdata-dir", TESSDATA,
                        "-l", "spa+eng", "--psm", "1", "-c", "textonly_pdf=1", "pdf"], check=True, capture_output=True)
        capa = PdfReader(tmp / "capa.pdf")
        escritor = PdfWriter(clone_from=pdf)
        for k, i in enumerate(sin_texto):
            pagina, texto = escritor.pages[i - 1], capa.pages[k]
            if pagina.rotation:
                pagina.transfer_rotation_to_content()
            caja = pagina.cropbox
            sx = float(caja.width) / float(texto.mediabox.width)
            sy = float(caja.height) / float(texto.mediabox.height)
            pagina.merge_transformed_page(texto, Transformation().scale(sx, sy).translate(float(caja.left),
                                                                                          float(caja.bottom)))
        salida = tmp / "salida.pdf"
        with open(salida, "wb") as f:
            escritor.write(f)
        agregadas = palabras(salida) - palabras(pdf)
        if agregadas < MIN_PALABRAS * len(sin_texto):
            print(f"REVISAR {pdf.name}: solo {agregadas} palabras en {len(sin_texto)} páginas; no se reemplaza")
            continue
        original = sha(pdf)
        pdf.write_bytes(salida.read_bytes())
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(f"{pdf.as_posix()}\t{original}\t{sha(pdf)}\t{len(sin_texto)}\t{agregadas}\t{date.today()}\n")
    print(f"OK {pdf.name}: {len(sin_texto)} de {n} páginas con OCR, {agregadas} palabras", flush=True)
