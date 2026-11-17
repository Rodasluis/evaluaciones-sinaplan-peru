"""Deck de muestra de la skill presentacion-evaluacion, con datos sintéticos, en los dos estilos: sirve para ver los
componentes de evaluación antes de armar una presentación real y como prueba de que la skill funciona.

Uso (desde la raíz del repositorio):
    uv run python .claude/skills/presentacion-evaluacion/scripts/ejemplo_evaluacion.py <carpeta_de_salida>
Salida: ejemplo_lab.pptx y ejemplo_institucional.pptx en la carpeta indicada (no en el repositorio).
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import componentes as K  # noqa: E402
from mimp_deck import MimpDeck, NAVY, ORANGE, RED as RED_INST  # noqa: E402
from mimp_lab_deck import GOLD, GREEN, RED, TEAL, VIOLET, MimpLabDeck  # noqa: E402

VALORACION = [["D1", "Vigencia del diagnóstico", "Pertinencia", "Parcial", K.resumen_confianza(["Sólido", "Indicio"])],
              ["D2", "Situación futura deseada", "Pertinencia", "Insuficiente", K.resumen_confianza(["Sólido"] * 3)],
              ["I1", "Implementación de los servicios", "Eficacia", "Parcial",
               K.resumen_confianza(["Con respaldo parcial", "Indicio"])]]
ESTADO = [["OP1", "10", "6", "3", "1", "7"], ["OP2", "8", "2", "4", "2", "3"], ["Total", "18", "8", "7", "3", "10"]]


def lab(salida):
    d = MimpLabDeck()
    d.cover("Evaluación de prueba", "Resultados para validación", unidad="OMEP · OGMEPGD", fecha="Noviembre de 2026")
    s = d.content("Mensajes principales", icon="objetivo", accent=GOLD)
    w3 = (d.CW - 0.4) / 3
    for i, (num, lab_, col) in enumerate([("8 de 18", "servicios implementados al cierre", GREEN),
                                          ("3", "preguntas valoradas", VIOLET), ("S/ 3,31 millones", "ejecutados en la "
                                                                                  "conducción", GOLD)]):
        K.tarjeta_cifra(d, s, d.M + i * (w3 + 0.2), 2.3, w3, 1.75, num, lab_, col)
    K.caja(d, s, d.M, 4.3, d.CW, 2.4, "Lectura", ["Los servicios con protocolo aprobado son los que más avanzan.",
                                                  "El objetivo 2 concentra los servicios sin ejecución."], TEAL)
    s = d.content("Valoración de las preguntas", icon="check", accent=VIOLET)
    K.tabla_valoracion(d, s, VALORACION, h=1.8)
    s = d.content("Estado de los servicios por objetivo prioritario", icon="engranaje", accent=GREEN)
    K.tabla_estado(d, s, ESTADO, h=1.8)
    d.icon_bullets(s, d.M + 6.95, 2.35, d.CW - 6.95, 3, [("alerta", RED, "El objetivo 2 no tiene servicios "
                                                                        "implementados.")], size=13)
    d.closing("Gracias", pie="OGMEPGD · OMEP · Noviembre de 2026")
    d.save(salida)


def institucional(salida):
    d = MimpDeck()
    d.cover("Evaluación de prueba", "Resultados para la Alta Dirección", unidad="OMEP · OGMEPGD",
            fecha="Noviembre de 2026")
    s = d.content("Mensajes principales")
    w3 = (d.CW - 0.4) / 3
    for i, (num, lab_, col) in enumerate([("8 de 18", "servicios implementados al cierre", NAVY),
                                          ("3", "preguntas valoradas", ORANGE), ("1", "objetivo sin servicios "
                                                                                  "implementados", RED_INST)]):
        K.tarjeta_cifra(d, s, d.M + i * (w3 + 0.2), 2.3, w3, 1.75, num, lab_, col)
    K.caja(d, s, d.M, 4.3, d.CW, 2.4, "Lectura", ["Los servicios con protocolo aprobado son los que más avanzan."], NAVY)
    s = d.content("Valoración de las preguntas")
    K.tabla_valoracion(d, s, VALORACION, h=1.8)
    s = d.content("Estado de los servicios por objetivo prioritario")
    K.tabla_estado(d, s, ESTADO, h=1.8)
    d.closing("¡Gracias!", pie="OGMEPGD · OMEP · Noviembre de 2026")
    d.save(salida)


if __name__ == "__main__":
    carpeta = sys.argv[1] if len(sys.argv) > 1 else "."
    os.makedirs(carpeta, exist_ok=True)
    lab(os.path.join(carpeta, "ejemplo_lab.pptx"))
    institucional(os.path.join(carpeta, "ejemplo_institucional.pptx"))
    print("Decks de muestra en", carpeta)
