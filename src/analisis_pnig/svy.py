"""Compatibilidad: la estimación con diseño muestral pasó al núcleo (src/nucleo/encuestas.py, Fase 2)."""
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.encuestas import *  # noqa: E402,F401,F403
from nucleo.encuestas import Z95, _razon, _var_total  # noqa: E402,F401
