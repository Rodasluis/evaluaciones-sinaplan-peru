"""Cobertura por servicio en los informes LIO 2019–2025: ubica el encabezado de cada servicio y mide su sección.

La extensión de cada sección (en líneas) y las menciones temáticas son aproximaciones de la riqueza de la información, no
medidas de calidad. Requiere el texto de cada informe extraído con:
    pdftotext -enc UTF-8 -layout docs/fuentes/pnig/06_informes_lio/<año>_mimp_informe-cumplimiento-lio.pdf <dir>/<año>.txt

Uso:
    python cobertura_lio.py <dir_txt> src/formulario_fie03/datos/servicios.csv <csv_salida> [<csv_secciones>]

Con <csv_secciones> (REP_ENT) se exporta además el texto de cada sección con sus páginas de inicio y fin en el PDF; la
página se cuenta por los saltos de página (\\f) que pdftotext inserta al final de cada una, y cada cambio de página se
marca en el texto con «[[p. N]]».
"""
import csv, difflib, re, sys, unicodedata

DIR, SERV, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
OUT_SEC = sys.argv[4] if len(sys.argv) > 4 else None
secciones = []


def norm(t):
    t = unicodedata.normalize("NFKD", t).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[ \t]+", " ", t).strip()


serv = list(csv.DictReader(open(SERV, encoding="utf-8")))
nombres = {s["codigo_corto"]: norm(s["servicio"])[:60] for s in serv}
TEMAS = {
    "presupuesto": r"presupuest|s/ ?\d|\bpim\b|devengad",
    "protocolo": r"protocolo",
    "cobertura": r"atencion|atendid|beneficiari|usuari|cobertura|participant",
    "territorio": r"departament|region|provincia|distrito",
    "indicador": r"\bmeta\b|indicador",
    "articulacion": r"articula|coordina|convenio",
    "dificultad": r"dificult|limitac|brecha de personal|restriccion",
    "monto_soles": r"s/\.?\s?\d",
    "sin_info": r"no (remitio|reporto|brindo|envio|presento|alcanzo)[^.]{0,40}informacion|no se (recibio|cuenta con|reporto) "
                r"informacion|sin informacion|no ha remitido|no registra informacion|no reporto avances",
}
H1 = re.compile(r"^(\d+(\.\d+)+\.?\s+)?()servicio:?\s*(n\s?[o°º]?\.?\s*)?(\d\.\d\.\d)\b")
H2 = re.compile(r"\(servicio (\d\.\d\.\d)\.?\)\s*$")
H3 = re.compile(r"^\d+(\.\d+)+\.?\s+servicio \d+\.?\s+(.*)$")          # 2019: «6.1.2.1. Servicio 1. Atención…»
H4 = re.compile(r"^(\d\.\d\.\d)\.?\s+([a-z].*)$")                          # «1.2.1. Estrategia comunitaria…»


def por_nombre(txt):
    best, score = None, 0
    for c, n in nombres.items():
        r = difflib.SequenceMatcher(None, txt[:45], n[:45]).ratio()
        if r > score:
            best, score = c, r
    return best if score > 0.75 else None


filas, encabezados = [], []
for y in range(2019, 2026):
    raw = open(f"{DIR}/{y}.txt", encoding="utf-8", errors="ignore").read()
    crudas = raw.split("\n")
    lineas = [norm(l) for l in crudas]
    pagina, p = [], 1
    for l in crudas:
        p += l.count("\f")      # el «\f» abre la línea siguiente al salto: esa línea ya es de la página nueva
        pagina.append(p)
    n = len(lineas)
    fin_indice = max((i for i, l in enumerate(crudas[: n * 15 // 100]) if re.search(r"([^\w\s]|_){4,}\s*\d*\s*$", l)),
                     default=0)
    pos = {}
    for i in range(fin_indice + 1, n):
        l = lineas[i]
        c = None
        m = H1.match(l) or None
        if m:
            c = m.group(5)
        elif H2.search(l):
            c = H2.search(l).group(1)
        elif y == 2019 and H3.match(l):
            c = por_nombre(H3.match(l).group(2) + " " + (lineas[i + 1] if i + 1 < n else ""))
        elif H4.match(l):
            c0, resto = H4.match(l).groups()
            c = c0 if c0 in nombres and difflib.SequenceMatcher(None, resto[:30], nombres[c0][:30]).ratio() > 0.6 else None
        if c in nombres and c not in pos:
            pos[c] = i
            encabezados.append((y, c, i, l[:90]))
    orden = sorted((i, c) for c, i in pos.items())
    for k, (i, c) in enumerate(orden):
        j = orden[k + 1][0] if k + 1 < len(orden) else min(i + 600, n)
        j = min(j, i + 1500)
        sec = " ".join(lineas[i:j])
        filas.append({"anio": y, "servicio": c, "seccion": "Sí", "lineas": j - i, "caracteres": len(sec),
                      **{t: len(re.findall(v, sec)) for t, v in TEMAS.items()}})
        if OUT_SEC:
            # cada cambio de página se marca con «[[p. N]]» para ubicar después la página de cada párrafo
            texto = []
            for k2 in range(i, j):
                if k2 > i and pagina[k2] != pagina[k2 - 1]:
                    texto.append(f"[[p. {pagina[k2]}]]")
                texto.append(crudas[k2].replace("\f", ""))
            secciones.append({"anio": y, "servicio": c, "pagina_inicio": pagina[i], "pagina_fin": pagina[j - 1],
                              "texto": "\n".join(texto)})
    for c in nombres:
        if c not in pos:
            filas.append({"anio": y, "servicio": c, "seccion": "No", "lineas": 0, "caracteres": 0, **{t: 0 for t in TEMAS}})
with open(OUT, "w", newline="", encoding="utf-8-sig") as f:
    w = csv.DictWriter(f, fieldnames=list(filas[0])); w.writeheader(); w.writerows(filas)
with open(OUT.replace(".csv", "_encabezados.tsv"), "w", encoding="utf-8") as f:
    for e in encabezados:
        f.write("\t".join(map(str, e)) + "\n")
if OUT_SEC:
    with open(OUT_SEC, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(secciones[0])); w.writeheader(); w.writerows(secciones)
print("ok", len(filas), len(secciones))
