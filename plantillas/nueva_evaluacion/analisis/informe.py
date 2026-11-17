"""Informe de evaluación (capa de evaluación mínima; plantilla de src/nucleo/nueva_evaluacion.py).

Arma el informe en Word con el núcleo (nucleo.documento y nucleo.portada) y la estructura de la tipología guardada en
config.yaml (sección informe):
    portada (portada.yaml de la evaluación), índices, una sección por entrada de la estructura;
    metodología: tabla de preguntas, criterios e instrumentos, y reglas de confianza;
    secciones de análisis: una subsección por pregunta, con sus hallazgos y una nota al pie de fuente por hallazgo;
    conclusiones: valoración y respuesta de cada pregunta;
    recomendaciones: matriz de la Tabla 4 de la guía (p. 56);
    anexo: hallazgos, confianza y fuentes independientes.
Lo que el equipo no ha escrito queda como [POR COMPLETAR] resaltado. Al final corre la auditoría de estilo.

Uso (desde la raíz del repositorio; primero sintesis.py):
    uv run python src/analisis_<id>/informe.py [--word]
    --word: termina el documento con Word (src/nucleo/actualizar_campos_word.ps1): índices, paginación y PDF.
"""
import re
import subprocess
import sys

from comun import ASPECTOS, CFG, INTERNOS, TIPOLOGIA, contar, leer_hallazgos, ruta, verificar
from nucleo import auditoria
from nucleo.config import RAIZ_REPO
from nucleo.documento import Documento
from nucleo.formato import lista_es
from nucleo.portada import agregar_portada

PENDIENTE = "[POR COMPLETAR]"
ANCHO = 9026   # ancho del texto en twips (15,9 cm)
CONFIANZA = {"Sólido": "dos o más fuentes independientes coinciden",
             "Con respaldo parcial": "dos o más fuentes coinciden, con discrepancias que se explican",
             "Indicio": "una sola fuente independiente lo respalda",
             "No determinado": "el análisis está pendiente"}


def _num(seccion):
    m = re.match(r"^(\d+)\.", seccion)
    return m.group(1) if m else None


def metodologia(d, s):
    nombre = CFG["instrumentos"]
    preguntas = CFG["preguntas"]
    d.p(f"La {TIPOLOGIA.lower()} comprende el periodo {CFG['periodo']['inicio']}–"
        f"{CFG['periodo']['fin']}. Responde {len(preguntas)} preguntas, cada una vinculada con un criterio de "
        f"evaluación y con los instrumentos que aportan su evidencia, como resume la Tabla {d.n_tabla + 1}.")
    filas = [[f"{k}. {' '.join(str(q['texto']).split())}", q["criterio"], lista_es([nombre[i] for i in q["instrumentos"]])]
             for k, q in preguntas.items()]
    d.leyenda("Preguntas de evaluación, criterios e instrumentos")
    d.tabla(["Pregunta", "Criterio", "Instrumentos"], filas, [5400, 1600, 2026], size=16)
    d.fuente("Fuente: configuración de la evaluación. Elaboración propia.")
    d.p(f"Cada pregunta se responde con al menos un instrumento documental y, cuando es posible, con una segunda fuente "
        f"independiente. La confianza de cada hallazgo depende del número de fuentes independientes que lo respaldan: "
        + "; ".join(f"{k.lower()}, cuando {v}" for k, v in CONFIANZA.items()) + ".")
    d.p(f"{PENDIENTE} {s.get('contenido', '')}")


def analisis(d, s, P, H):
    pref = str(s["preguntas"])
    num = _num(s["seccion"])
    preguntas = [k for k in CFG["preguntas"] if k.startswith(pref)]
    for j, k in enumerate(preguntas, 1):
        q = CFG["preguntas"][k]
        d.h2(f"{num}.{j} {q['criterio']} (pregunta {k})")
        d.p((f"Pregunta {k}. ", {"b": True}), " ".join(str(q["texto"]).split()), recuadro="pregunta")
        for h in (x for x in H if x["pregunta"] == k):
            cuerpo = h["enunciado"].rstrip() + (f" {h['detalle']}" if h["detalle"] else "")
            if not cuerpo.endswith((".", "?", "!", ")")):
                cuerpo += "."
            piezas = [(f"{h['aspecto']}. ", {"b": True})] if h["aspecto"] else []
            confianza = [] if h["confianza"] == "No determinado" else [(f" Confianza: {h['confianza'].lower()}.",
                                                                          {"i": True})]
            par = d.p(*piezas, cuerpo, *confianza, recuadro="hallazgo")
            fuentes = list(dict.fromkeys(e["fuente"] for e in h["evidencias"]))
            if fuentes:
                d.nota_pie(par, f"Fuente: {'; '.join(fuentes)}. Elaboración propia.")
            if h["salvedad"]:
                d.p(("Salvedad. ", {"i": True}), h["salvedad"])
        v = P.get(k, {})
        if v.get("respuesta"):
            d.p(("Respuesta. ", {"b": True}), v["respuesta"],
                (f" Valoración: {v['valoracion'].lower()}." if v.get("valoracion") else ""), recuadro="respuesta")
        else:
            d.p(f"{PENDIENTE} Respuesta y valoración de la pregunta {k}.", recuadro="respuesta")


def conclusiones(d, P):
    for k in CFG["preguntas"]:
        v = P.get(k, {})
        texto = (f"{v['valoracion']}. {v['respuesta']}" if v.get("respuesta") else f"{PENDIENTE} Conclusión de la pregunta.")
        d.p((f"Pregunta {k}. ", {"b": True}), texto)


def recomendaciones(d, R):
    if not R:
        d.p(f"{PENDIENTE} Recomendaciones según la matriz de la Tabla 4 de la guía: aspecto de mejora, recomendación, "
            f"unidad de organización responsable y plazo.")
        return
    filas = [[str(n), r["aspecto"], r["recomendacion"], r["responsable"], r["plazo"]] for n, r in enumerate(R, 1)]
    d.leyenda("Matriz de recomendaciones")
    d.tabla(["N.°", "Aspecto de mejora", "Recomendación", "Unidad de organización responsable", "Plazo"], filas,
            [500, 1400, 4126, 1700, 1300], size=16)
    d.fuente("Fuente: hallazgos y conclusiones de la evaluación. Elaboración propia.")
    por = {a: sum(r["aspecto"] == a for r in R) for a in ASPECTOS}
    d.p(f"La matriz reúne {len(R)} {'recomendación' if len(R) == 1 else 'recomendaciones'}: {por['Implementación']} de implementación, {por['Diseño']} de "
        f"diseño y {por['Seguimiento y evaluación']} de seguimiento y evaluación. Cada una se deriva de los hallazgos y conclusiones de la "
        f"evaluación, tiene una unidad de organización responsable y un plazo, como pide la guía para las políticas "
        f"nacionales.")


def anexo(d, H):
    d.h1("Anexo 1. Hallazgos, confianza y fuentes")
    filas = [[h["pregunta"], h["aspecto"] or h["enunciado"][:80], h["confianza"], str(contar(h))] for h in H]
    d.leyenda("Hallazgos por pregunta, con su confianza y sus fuentes independientes")
    d.tabla(["Pregunta", "Hallazgo", "Confianza", "Fuentes independientes"], filas, [1100, 5126, 1600, 1200], size=16)
    d.fuente("Fuente: matriz de triangulación de la evaluación. Elaboración propia.")


def construir(P, H, R, entradas=None):
    t = CFG["periodo"]
    elabora = CFG["entidad"].get("elabora") or ""
    d = Documento(pie=f"Informe de {TIPOLOGIA.lower()} · {CFG['politica']['sigla']} {t['inicio']}–{t['fin']}",
                  titulo_doc=CFG["titulo"], autor=elabora if elabora and not elabora.startswith("[") else
                  "OMEP–OGMEPGD, MIMP", informe=True)
    agregar_portada(d.d, str(RAIZ_REPO / CFG["productos"]["portada"]))
    titulos, leyendas = entradas or ([], [])
    d.titulo_indice("Índice")
    d.indice(r'TOC \o "1-2" \h \z \u', [(x, "", n) for x, n in titulos])
    d.salto()
    d.titulo_indice("Índice de tablas")
    d.indice(r'TOC \h \z \c "Tabla"', [(f"{r} {k}. {x}", "", 1) for r, k, x in leyendas if r == "Tabla"],
             nivel_de="Tabla")
    d.cerrar_seccion(numeracion="lowerRoman")
    for s in CFG["informe"]:
        nombre = s["seccion"]
        if nombre.startswith("Anexo"):
            continue
        d.h1(nombre)
        if "preguntas" in s:
            analisis(d, s, P, H)
        elif "metodolog" in nombre.lower():
            metodologia(d, s)
        elif "conclusion" in nombre.lower():
            conclusiones(d, P)
        elif "recomendacion" in nombre.lower():
            recomendaciones(d, R)
        else:
            d.p(f"{PENDIENTE} {s.get('contenido', '')}")
    d.h1("Referencias")
    d.p(f"{PENDIENTE} Referencias en formato APA, con las fuentes citadas en las notas al pie.")
    anexo(d, H)
    d.numerar_cuerpo_desde_1()
    return d


def main(argv):
    P, H, R = leer_hallazgos()
    verificar(P, H, R)
    borrador = construir(P, H, R)
    d = construir(P, H, R, (borrador.titulos, borrador.leyendas))
    salida = ruta("informe", CFG["productos"]["informe"].split("/", 1)[-1])
    salida.parent.mkdir(parents=True, exist_ok=True)
    d.guardar(salida, comillas_inglesas=True)
    rel = salida.relative_to(RAIZ_REPO).as_posix()
    n, falta = auditoria.leyendas_sin_analisis(salida)
    est = auditoria.estilo(salida)
    internos = auditoria.terminos(salida, INTERNOS)
    pend = sum(PENDIENTE in p.text for p in d.d.paragraphs)
    print(f"Informe generado: {rel}; {pend} párrafos {PENDIENTE}; {n} leyendas, {len(falta)} sin análisis; "
          f"{len(est)} avisos de estilo; {len(internos)} términos internos.")
    for x in falta:
        print("  leyenda sin análisis:", x)
    for nombre, frag in est:
        print(f"  {nombre}: …{frag}…")
    for frag in internos:
        print(f"  término interno: …{frag}…")
    if "--word" in argv:
        pdf = salida.with_suffix(".pdf")
        subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
                        str(RAIZ_REPO / "src/nucleo/actualizar_campos_word.ps1"), "-Docx", str(salida), "-Pdf", str(pdf)],
                       check=True)
        print(f"Terminado con Word; PDF de revisión: {pdf.relative_to(RAIZ_REPO).as_posix()}")
    return 1 if falta or internos else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
