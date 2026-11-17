"""D6: lista de verificación de los siete enfoques transversales (REV_DOC, Sección C de la nota metodológica).

Paso 1 (extracción, reproducible): busca cada enfoque en cada elemento del diseño y guarda las menciones con su
ubicación exacta en 04_analisis/d6_enfoques/evidencia_enfoques.csv:
  - Diagnóstico: documento de la PNIG, numerales 1.4.1 a 1.4.6 (pp. 10 a 30 del PDF del D.S. N.° 008-2019-MIMP);
  - Objetivos prioritarios y lineamientos: numeral 2 (p. 38) y matriz de consistencia (Anexo 5);
  - Servicios: numeral 3 (pp. 39 a 41), fichas de servicios (Anexo 8) y protocolos aprobados;
  - Indicadores: fichas de los OP (Anexo 4) y fichas de indicadores de servicio del PEMIG (Anexo 5);
  - Marco y glosario: presentación, base legal y glosario (pp. 5 a 9 y 43 a 46), donde una mención no orienta el
    diseño.
Paso 2 (valoración del equipo evaluador): 04_analisis/d6_enfoques/valoracion_enfoques.csv asigna a cada enfoque y
elemento un nivel con su sustento y una cita que debe constar en la evidencia extraída (el script lo verifica).
Escala (nota, Sección C de REV_DOC): ausente, cuando el enfoque no aparece en el diseño; mencionado, cuando se enuncia
sin desarrollo; desarrollado, cuando orienta el diagnóstico o los objetivos; operacionalizado, cuando se traduce en
lineamientos, servicios o indicadores. El nivel del enfoque es el más alto alcanzado en algún elemento.

Uso (desde la raíz del repositorio):
    uv run --no-project --with pandas --with xlrd --with python-calamine --with openpyxl python -I -X utf8 src/analisis_pnig/d6_enfoques.py
Salida: evaluaciones/pnig_2026/04_analisis/D6_enfoques_transversales.xlsx
"""
import glob
import os
import re
import subprocess

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

PNIG = "docs/fuentes/pnig/01_normativa/2019_mimp_ds-008-2019-mimp_pnig.pdf"
ANEXOS = "docs/fuentes/pnig/01_normativa/pnig_anexos"
PEMIG = "docs/fuentes/pnig/02_pemig"
PROTOCOLOS = "docs/fuentes/pnig/01_normativa/protocolos"
D = f"{EVAL}/04_analisis/d6_enfoques"
SALIDA = f"{EVAL}/04_analisis/D6_enfoques_transversales.xlsx"
ENFOQUES = {  # enfoque: patrón de mención explícita del enfoque o de su traducción operativa
    "Género": r"enfoque de g[eé]nero|perspectiva de g[eé]nero",
    "Derechos humanos": r"enfoque (de|basado en) (los )?derechos|enfoque de derechos humanos",
    "Interseccionalidad": r"interseccional",
    "Territorialidad": r"enfoque territorial|territorialidad|pertinencia territorial|enfoque de territorio",
    "Interculturalidad": r"intercultural|pertinencia cultural|adecuaci[oó]n cultural|culturalmente pertinente",
    "Ciclo de vida": r"ciclo de vida|curso de vida|intergeneracional",
    "Gestión por resultados": r"gesti[oó]n (p[uú]blica )?(orientada|basada|por) (a |en )?(los )?resultados|"
                              r"presupuesto por resultados|enfoque de resultados",
}
ORDEN = ["Diagnóstico", "Objetivos prioritarios", "Lineamientos", "Servicios", "Indicadores",
         "Marco, alineamiento y glosario", "Protocolos (implementación)"]
NIVELES = ("Ausente", "Mencionado", "Desarrollado", "Operacionalizado")


def limpiar(t):
    return re.sub(r"\s+", " ", str(t).replace("ﬁ", "fi").replace("ﬂ", "fl")).strip()


def pdf_paginas(ruta, desde=1, hasta=None):
    n = int(re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", ruta], capture_output=True).stdout.decode(
        "utf-8", "replace")).group(1))
    hasta = min(hasta or n, n)
    for p in range(desde, hasta + 1):
        t = subprocess.run(["pdftotext", "-enc", "UTF-8", "-f", str(p), "-l", str(p), ruta, "-"],
                           capture_output=True).stdout.decode("utf-8", "replace")
        yield p, limpiar(t)


def celdas(ruta):
    for hoja, df in pd.read_excel(ruta, sheet_name=None, header=None, dtype=str).items():
        texto = " | ".join(limpiar(v) for v in df.fillna("").to_numpy().ravel() if str(v).strip())
        yield hoja, texto


def anexo5():
    """Matriz de consistencia por columnas, con objetivo, lineamiento y servicio rellenados hacia abajo."""
    a = pd.read_excel(f"{ANEXOS}/2019_mimp_pnig_anexo-5_matriz-consistencia.xlsx", header=None, dtype=str).iloc[4:]
    a = a.rename(columns={1: "causa", 3: "op", 4: "ind_op", 5: "lineamiento", 6: "servicio", 7: "persona",
                          10: "estandar", 11: "definicion", 12: "ind_estandar"})
    for c in ("causa", "op", "lineamiento", "servicio"):
        a[c] = a[c].ffill()
    return a.fillna("")


def codigo(texto):
    m = re.match(r"\s*(OP\d|\d\.\d(?:\.\d)?)", str(texto))
    return m.group(1) if m else limpiar(texto)[:40]


def fragmentos():
    """(elemento, documento, ubicación, texto) de cada unidad de búsqueda."""
    doc = "Documento de la PNIG (D.S. N.° 008-2019-MIMP)"
    for p, t in pdf_paginas(PNIG):
        if 10 <= p <= 30:
            yield "Diagnóstico", doc, f"p. {p}", t
        elif p == 38:
            i = t.find("OBJETIVOS PRIORITARIOS Y LINEAMIENTOS")
            yield "Marco, alineamiento y glosario", doc, f"p. {p}", t[:i]
            yield "Objetivos prioritarios", doc, f"p. {p}", t[i:]
        elif 39 <= p <= 41:
            yield "Servicios", doc, f"p. {p}", t
        elif 5 <= p <= 9 or 31 <= p <= 37 or 42 <= p <= 46:
            yield "Marco, alineamiento y glosario", doc, f"p. {p}", t
    a5 = anexo5()
    m5 = "Anexo 5 de la PNIG (matriz de consistencia)"
    for elemento, cols, clave in (("Diagnóstico", ["causa"], "causa"), ("Objetivos prioritarios", ["op"], "op"),
                                  ("Lineamientos", ["lineamiento"], "lineamiento"),
                                  ("Servicios", ["servicio", "persona", "estandar", "definicion"], "servicio"),
                                  ("Indicadores", ["ind_op", "ind_estandar"], "servicio")):
        for unidad, g in a5.groupby(clave, sort=False):
            texto = limpiar(" | ".join(" | ".join(g[c]) for c in cols))
            yield elemento, m5, f"{codigo(unidad)}", texto
    for f in sorted(glob.glob(f"{ANEXOS}/2019_mimp_pnig_anexo-8_fichas-servicios-op*.xls")):
        for hoja, t in celdas(f):
            yield "Servicios", f"Anexo 8 de la PNIG ({os.path.basename(f)[-7:-4].upper()})", hoja, t
    a4 = pd.read_excel(f"{ANEXOS}/2019_mimp_pnig_anexos-2-3-4-6.xlsx", sheet_name="Anexo N° 4", header=None,
                       dtype=str).fillna("")
    for hoja, t in (("Anexo N° 4", " | ".join(limpiar(v) for v in a4.to_numpy().ravel() if str(v).strip())),):
        yield "Indicadores", "Anexo 4 de la PNIG (fichas de los OP)", hoja, t
    pe = pd.read_excel(f"{PEMIG}/2020_mimp_pemig_anexo-5_fichas-indicadores-servicios.xlsx",
                       sheet_name="e__Seguimiento_de_la_implem", header=None, dtype=str).iloc[6:]
    for _, r in pe.iterrows():
        if isinstance(r[2], str) and isinstance(r[0], str) and re.match(r"I?SS?\.|\d", r[0]):
            yield "Indicadores", "PEMIG, Anexo 5 (nombre y método del indicador de servicio)", r[0], limpiar(
                f"{r[2]} | {r[3]}")
    for f in sorted(glob.glob(f"{PROTOCOLOS}/*.pdf")):
        if "resolucion" in f:
            continue
        texto = " ".join(t for _, t in pdf_paginas(f))
        yield "Protocolos (implementación)", "Protocolo aprobado", os.path.basename(f), texto


def extraer():
    filas = []
    for elemento, doc, ubic, texto in fragmentos():
        for enfoque, patron in ENFOQUES.items():
            hallados = list(re.finditer(patron, texto, re.I))
            if not hallados:
                continue
            m = hallados[0]
            cita = texto[max(0, m.start() - 140): m.end() + 140]
            filas.append({"enfoque": enfoque, "elemento": elemento, "documento": doc, "ubicacion": ubic,
                          "menciones": len(hallados), "fragmento": f"…{cita}…"})
    return pd.DataFrame(filas)


def resumen(ev):
    t = ev.groupby(["enfoque", "elemento"]).agg(unidades=("ubicacion", "nunique"), menciones=("menciones", "sum"))
    t = t.reset_index().pivot(index="enfoque", columns="elemento", values="unidades").reindex(
        index=list(ENFOQUES), columns=ORDEN).fillna(0).astype(int)
    n_prot = len([f for f in glob.glob(f"{PROTOCOLOS}/*.pdf") if "resolucion" not in f])
    prot = ev[ev.elemento == "Protocolos (implementación)"].groupby("enfoque").ubicacion.nunique().reindex(
        list(ENFOQUES)).fillna(0).astype(int)
    t["Protocolos que lo mencionan"] = prot.map(lambda x: f"{x} de {n_prot}")
    return t.reset_index()


def valoracion(ev):
    ruta = f"{D}/valoracion_enfoques.csv"
    if not os.path.exists(ruta):
        return None
    v = pd.read_csv(ruta)
    assert set(v.nivel) <= set(NIVELES), "nivel fuera de la escala"
    assert set(v.enfoque) == set(ENFOQUES), "falta valorar algún enfoque"
    for _, r in v.iterrows():
        if isinstance(r.cita, str) and r.cita.strip():
            mismo = ev[(ev.enfoque == r.enfoque) & (ev.ubicacion.astype(str) == str(r.ubicacion))]
            ok = any(limpiar(r.cita) in limpiar(f) for f in mismo.fragmento)
            assert ok, f"la cita de {r.enfoque} ({r.ubicacion}) no consta en la evidencia extraída"
    return v


def escribir(hojas):
    with pd.ExcelWriter(SALIDA, engine="openpyxl") as w:
        for nombre, df in hojas.items():
            df.to_excel(w, sheet_name=nombre, index=False)
    wb = load_workbook(SALIDA)
    azul = PatternFill("solid", fgColor="1F3864")
    colores = {"Operacionalizado": "C6EFCE", "Desarrollado": "DDEBF7", "Mencionado": "FFEB9C", "Ausente": "F8CBAD"}
    for ws in wb.worksheets:
        for c in ws[1]:
            c.fill, c.font = azul, Font(bold=True, color="FFFFFF")
            c.alignment = Alignment(wrap_text=True, vertical="center")
        for col in ws.columns:
            largo = max(len(str(c.value or "")) for c in col[:40])
            ws.column_dimensions[col[0].column_letter].width = min(max(10, largo * 0.9), 70)
        for fila in ws.iter_rows(min_row=2):
            for c in fila:
                c.alignment = Alignment(wrap_text=True, vertical="top")
                if isinstance(c.value, str) and c.value in colores:
                    c.fill = PatternFill("solid", fgColor=colores[c.value])
        ws.freeze_panes = "B2"
    wb.save(SALIDA)


if __name__ == "__main__":
    os.makedirs(D, exist_ok=True)
    ev = extraer()
    ev.to_csv(f"{D}/evidencia_enfoques.csv", index=False, encoding="utf-8")
    res = resumen(ev)
    hojas = {"Resumen de la evidencia": res}
    v = valoracion(ev)
    if v is not None:
        hojas = {"Lista de verificación": v, **hojas}
    hojas["Evidencia"] = ev
    hojas["Criterios"] = pd.DataFrame([
        ["Nota metodológica", "Enfoques", "Género, derechos humanos, interseccionalidad, territorialidad, "
         "interculturalidad, ciclo de vida y gestión por resultados (numeral 4; REV_DOC, Sección C)."],
        ["Nota metodológica", "Escala", "Ausente; mencionado (se enuncia sin desarrollo); desarrollado (orienta el "
         "diagnóstico o los objetivos); operacionalizado (se traduce en lineamientos, servicios o indicadores)."],
        ["Equipo evaluador", "Búsqueda", "Patrones de mención explícita del enfoque o de su traducción operativa "
         "(p. ej., «pertinencia cultural» para interculturalidad). Una unidad es una página, una hoja o un protocolo."],
        ["Equipo evaluador", "Valoración", "El nivel lo asigna el equipo evaluador leyendo la evidencia: una mención en "
         "la base legal o el glosario no cuenta como desarrollo, y un servicio cuyo nombre o estándar incorpora el "
         "enfoque cuenta como operacionalización."],
    ], columns=["Plano", "Elemento", "Regla"])
    escribir(hojas)
    pd.set_option("display.width", 220)
    print(res.to_string(index=False))
