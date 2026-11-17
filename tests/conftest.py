"""Configuración de las pruebas: los scripts se ejecutan desde la raíz del repositorio (sus rutas son relativas)."""
import os
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
os.chdir(RAIZ)
sys.path.insert(0, str(Path(__file__).resolve().parent))
os.environ.setdefault("PYTHONUTF8", "1")
