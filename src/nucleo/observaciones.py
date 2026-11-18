"""Observaciones de la validación del informe (Fase 4, ADR-017): matriz de los comentarios de un Word revisado y JSON de
respuestas para src/nucleo/comparar_word.ps1.

La matriz lista cada comentario raíz en el orden del documento, con su número («Observación NN», si lo trae), la sección
donde está, el texto comentado y el comentario, y deja las columnas respuesta y cambio para el equipo evaluador. No
lleva el nombre de quien comentó. El documento con observaciones y su matriz van a la carpeta administrativa, fuera del
repositorio (ruta_externa("carpeta_administrativa")).

Uso (desde la raíz del repositorio):
    uv run python src/nucleo/observaciones.py matriz <docx revisado> <matriz.xlsx>
    uv run python src/nucleo/observaciones.py json <matriz.xlsx> <respuestas.json>
El JSON lleva una clave por comentario («orden_001», …) y, si el comentario tiene número, también «01», «02», …; se pasa
a comparar_word.ps1 con -Respuestas.
"""
import json
import re
import sys
import zipfile
from pathlib import Path

import pandas as pd
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
NS = {"w": W, "w15": W15}
NUMERO = re.compile(r"Observaci.n (\d+)")
COLUMNAS = ["orden", "numero", "seccion", "texto_comentado", "comentario", "respuesta", "cambio", "estado"]


def _texto(el):
    return "".join(t.text or "" for t in el.iter(f"{{{W}}}t")).strip()


def comentarios(docx):
    """Comentarios raíz en el orden del documento: orden, número, sección, texto comentado y comentario."""
    with zipfile.ZipFile(docx) as z:
        nombres = set(z.namelist())
        if "word/comments.xml" not in nombres:
            return pd.DataFrame(columns=COLUMNAS)
        com = etree.fromstring(z.read("word/comments.xml"))
        doc = etree.fromstring(z.read("word/document.xml"))
        ext = etree.fromstring(z.read("word/commentsExtended.xml")) if "word/commentsExtended.xml" in nombres else None
    respuestas = set()
    if ext is not None:
        respuestas = {e.get(f"{{{W15}}}paraId") for e in ext.iter(f"{{{W15}}}commentEx")
                      if e.get(f"{{{W15}}}paraIdParent")}
    texto, raiz = {}, {}
    for c in com.iter(f"{{{W}}}comment"):
        cid = c.get(f"{{{W}}}id")
        texto[cid] = _texto(c)
        pids = {p.get("{http://schemas.microsoft.com/office/word/2010/wordml}paraId") for p in c.iter(f"{{{W}}}p")}
        raiz[cid] = not (pids & respuestas)
    seccion, abiertos, anclado, orden = "", {}, {}, []
    for p in doc.iter(f"{{{W}}}p"):
        estilo = p.find("w:pPr/w:pStyle", NS)
        if estilo is not None and re.match(r"(?i)(heading|t[ií]tulo|ttulo)", estilo.get(f"{{{W}}}val", "")):
            seccion = _texto(p) or seccion
        for el in p.iter():
            if el.tag == f"{{{W}}}commentRangeStart":
                abiertos[el.get(f"{{{W}}}id")] = []
            elif el.tag == f"{{{W}}}commentRangeEnd":
                cid = el.get(f"{{{W}}}id")
                anclado[cid] = " ".join(abiertos.pop(cid, [])).strip()
            elif el.tag == f"{{{W}}}t":
                for lista in abiertos.values():
                    lista.append(el.text or "")
            elif el.tag == f"{{{W}}}commentReference":
                cid = el.get(f"{{{W}}}id")
                if raiz.get(cid) and cid not in [x[0] for x in orden]:
                    orden.append((cid, seccion))
    filas = []
    for i, (cid, sec) in enumerate(orden, 1):
        m = NUMERO.search(texto.get(cid, ""))
        filas.append({"orden": i, "numero": m.group(1).zfill(2) if m else "", "seccion": sec,
                      "texto_comentado": re.sub(r"\s+", " ", anclado.get(cid, ""))[:400],
                      "comentario": texto.get(cid, ""), "respuesta": "", "cambio": "", "estado": "Pendiente"})
    return pd.DataFrame(filas, columns=COLUMNAS)


def matriz(docx, destino):
    df = comentarios(docx)
    Path(destino).parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(destino, index=False, sheet_name="Observaciones")
    return df


def respuestas_json(xlsx, destino):
    """JSON de respuestas para comparar_word.ps1, solo con las observaciones que tienen respuesta."""
    df = pd.read_excel(xlsx, dtype=str, engine="openpyxl").fillna("")
    salida = {}
    for r in df.itertuples():
        if r.respuesta.strip():
            salida[f"orden_{int(r.orden):03d}"] = r.respuesta.strip()
            if r.numero:
                salida[r.numero.zfill(2)] = r.respuesta.strip()
    Path(destino).write_text(json.dumps(salida, ensure_ascii=False, indent=1), encoding="utf-8")
    return salida


def main(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    if len(argv) != 3 or argv[0] not in ("matriz", "json"):
        print(__doc__)
        return 2
    if argv[0] == "matriz":
        df = matriz(argv[1], argv[2])
        print(f"{len(df)} observaciones en {argv[2]}; {int((df.numero != '').sum())} con número.")
    else:
        r = respuestas_json(argv[1], argv[2])
        print(f"{sum(k.startswith('orden_') for k in r)} respuestas en {argv[2]}.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
