"""Rechaza en el XML los cambios de texto propuestos de un .docx revisado y acepta los de formato, sin abrir Word.

Sirve de paso previo a `comparar_word.ps1` cuando la versión revisada trae muchos cambios de formato registrados
(w:rPrChange): rechazarlos con Word por COM (Revisions.RejectAll) puede tardar horas en un informe largo. La
comparación de `comparar_word.ps1` no compara formato, de modo que aceptar esos cambios no altera el resultado.

Qué hace en el cuerpo, los encabezados, los pies y las notas:
  - inserciones (w:ins, w:moveTo): se quitan con su contenido;
  - eliminaciones (w:del, w:moveFrom): se restituye el texto (w:delText pasa a w:t);
  - cambios de formato (w:rPrChange, w:pPrChange y demás): se quitan, con lo que queda el formato nuevo.
Los comentarios se conservan. Si encuentra una inserción de marca de párrafo o de fila, se detiene: rechazarla exige
unir párrafos o borrar filas, y eso se hace con Word.

Uso (desde la raíz del repositorio):
    uv run python src/nucleo/rechazar_cambios.py <docx revisado> <docx de salida>
"""
import re
import sys
import zipfile
from collections import Counter

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
PARTES = re.compile(r"word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$")
FORMATO = ("rPrChange", "pPrChange", "sectPrChange", "tblPrChange", "tblPrExChange", "trPrChange", "tcPrChange",
           "tblGridChange", "numberingChange")
MARCAS = ("rPr", "trPr")   # w:ins o w:del como marca de párrafo o de fila


def _q(tag):
    return "{%s}%s" % (W, tag)


def _nombre(e):
    return etree.QName(e).localname


def procesar(xml, conteo):
    raiz = etree.fromstring(xml)
    for e in list(raiz.iter(_q("ins"), _q("moveTo"))):
        padre = e.getparent()
        if _nombre(padre) in MARCAS:
            raise NotImplementedError(f"inserción de marca en <w:{_nombre(padre)}>: rechazarla con Word")
        padre.remove(e)
        conteo["inserciones rechazadas"] += 1
    for e in list(raiz.iter(_q("del"), _q("moveFrom"))):
        padre = e.getparent()
        if _nombre(padre) in MARCAS:   # marca de párrafo o de fila eliminada: rechazar es conservarla
            padre.remove(e)
        else:
            i = padre.index(e)
            for hijo in list(e):
                padre.insert(i, hijo)
                i += 1
            padre.remove(e)
        conteo["eliminaciones rechazadas"] += 1
    for t in raiz.iter(_q("delText")):
        t.tag = _q("t")
    for t in raiz.iter(_q("delInstrText")):
        t.tag = _q("instrText")
    for tag in FORMATO:
        for e in list(raiz.iter(_q(tag))):
            e.getparent().remove(e)
            conteo["cambios de formato aceptados"] += 1
    return etree.tostring(raiz, xml_declaration=True, encoding="UTF-8", standalone=True)


def rechazar(entrada, salida):
    conteo = Counter()
    with zipfile.ZipFile(entrada) as zin, zipfile.ZipFile(salida, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            datos = zin.read(item.filename)
            if PARTES.match(item.filename):
                datos = procesar(datos, conteo)
            zout.writestr(item, datos)
    return conteo


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    c = rechazar(sys.argv[1], sys.argv[2])
    print("; ".join(f"{k}: {v}" for k, v in c.items()) or "sin cambios registrados")
