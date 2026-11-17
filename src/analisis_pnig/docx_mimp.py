"""Compatibilidad: el generador de documentos Word y el formato pasaron al núcleo (src/nucleo/documento.py y
src/nucleo/formato.py, Fase 2). Este módulo los reexporta para los scripts del piloto."""
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.documento import *  # noqa: E402,F401,F403
from nucleo.documento import Documento, PLANTILLA, ANCHO_TEXTO_CM  # noqa: E402,F401
from nucleo.formato import (comillas, comillas_parrafo, fmt_num, lista_es, millones, monto, pct,  # noqa: E402,F401
                            soles)
