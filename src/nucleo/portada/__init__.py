"""Portada de los informes técnicos de evaluación en Word (python-docx), con el diseño centrado (G), basado en la
portada del informe de evaluación 2023 de la PNIG del MIMP. Es el único diseño vigente desde la ronda 35 (pedido del
usuario); los modelos A a F y H a K se retiraron. Instrucciones en docs/portada.md.

Todo lo que se ve en la portada sale de config/portada.yaml: textos, logos, colores y tipografía.

Uso:
    import sys; sys.path.insert(0, "src")
    from nucleo.portada import agregar_portada
    info = agregar_portada(documento, "config/portada.yaml")   # o evaluaciones/<id>/portada.yaml       # documento: docx.Document vacío
    # la portada queda como primera sección (A4, sin encabezado, pie ni número); el cuerpo sigue en la página 2
"""
import logging

from . import opcion_g
from .comun import DISENO, DatosFaltantes, cargar_config, cerrar_seccion_portada, validar
from .paleta import Paleta, contraste

registro = logging.getLogger("portada")

__all__ = ["agregar_portada", "cargar_config", "DatosFaltantes", "DISENO"]


def agregar_portada(documento, ruta_config="config/portada.yaml", titulo=None):
    """Agrega la portada al final del cuerpo del documento (en la práctica, al inicio de un documento vacío) y la cierra
    como sección propia. titulo sustituye al de la configuración.
    Devuelve un resumen con el diseño, la tipografía usada, el contraste del blanco sobre el rojo y los avisos."""
    cfg = validar(cargar_config(ruta_config))
    paleta = Paleta.desde_config(cfg)
    avisos = []
    cierre = opcion_g.construir(documento, cfg, paleta, avisos, (titulo or cfg["titulo"]).strip()) or {}
    cerrar_seccion_portada(documento, paleta, **cierre)
    return {"opcion": DISENO[0], "diseno": DISENO[1], "fuente": paleta.fuente, "colores": paleta.todos(),
            "contraste_blanco_primario": round(contraste(paleta.blanco, paleta.primario), 2), "avisos": avisos}
