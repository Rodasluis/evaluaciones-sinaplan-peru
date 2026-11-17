"""Paleta y tipografía de la portada, leídas siempre de config/portada.yaml (nunca escritas en las opciones).

- Colores: primario (rojo institucional), gris oscuro, gris claro y texto casi negro, más el blanco. Se rechaza
  cualquier azul: la portada no lleva azul ni combina rojo con azul.
- Tipografía: la de la configuración si está instalada en Windows; si no, Calibri, con un aviso en el registro.
"""
import colorsys
import logging
import os
from dataclasses import dataclass
from pathlib import Path

registro = logging.getLogger("portada")
BLANCO = "FFFFFF"
_AVISADAS = set()            # tipografías ya avisadas en esta ejecución
CARPETAS_FUENTES = [Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts",
                    Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "Windows" / "Fonts"]


def _hex(valor, nombre):
    v = str(valor or "").strip().lstrip("#").upper()
    if len(v) != 6 or any(c not in "0123456789ABCDEF" for c in v):
        raise ValueError(f"color «{nombre}» no válido en la configuración: {valor!r}")
    return v


def es_azul(hexa):
    """Tono entre cian y violeta con saturación apreciable."""
    r, g, b = (int(hexa[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, s, v = colorsys.rgb_to_hsv(r, g, b)
    return 0.47 <= h <= 0.75 and s > 0.20 and v > 0.15


def contraste(c1, c2):
    """Razón de contraste WCAG entre dos colores hexadecimales."""
    def lum(c):
        canales = []
        for i in (0, 2, 4):
            x = int(c[i:i + 2], 16) / 255
            canales.append(x / 12.92 if x <= 0.03928 else ((x + 0.055) / 1.055) ** 2.4)
        r, g, b = canales
        return 0.2126 * r + 0.7152 * g + 0.0722 * b
    a, b = sorted((lum(c1), lum(c2)), reverse=True)
    return (a + 0.05) / (b + 0.05)


def _buscar_fuente(nombre, negrita):
    """Archivo .ttf de la fuente instalada (para medir el título), o None."""
    clave = nombre.lower().replace(" ", "")
    candidatos = []
    for carpeta in CARPETAS_FUENTES:
        if carpeta.exists():
            candidatos += [p for p in carpeta.iterdir() if p.suffix.lower() in (".ttf", ".otf")
                           and p.stem.lower().replace(" ", "").replace("-", "").startswith(clave)]
    if not candidatos:
        return None
    if negrita:
        bold = [p for p in candidatos if p.stem.lower().endswith(("bold", "-bold", "b", "bd")) and "italic" not in
                p.stem.lower() and not p.stem.lower().endswith(("semibold", "extrabold"))]
        return str(sorted(bold or candidatos, key=lambda p: len(p.stem))[0])
    regular = [p for p in candidatos if p.stem.lower() in (clave, f"{clave}-regular", f"{clave}regular")]
    return str((regular or sorted(candidatos, key=lambda p: len(p.stem)))[0])


@dataclass
class Paleta:
    primario: str
    gris_oscuro: str
    gris_claro: str
    texto: str
    fuente: str
    ttf_negrita: str
    ttf_regular: str
    blanco: str = BLANCO

    @classmethod
    def desde_config(cls, cfg):
        col = cfg.get("colores") or {}
        valores = {k: _hex(col.get(k), k) for k in ("primario", "gris_oscuro", "gris_claro", "texto")}
        azules = [k for k, v in valores.items() if es_azul(v)]
        if azules:
            raise ValueError(f"la portada no admite azul: revise los colores {azules} de config/portada.yaml")
        pedida = str(cfg.get("tipografia") or "Calibri").strip()
        ttf_b, ttf_r = _buscar_fuente(pedida, True), _buscar_fuente(pedida, False)
        fuente = pedida
        if not ttf_b:
            if pedida not in _AVISADAS:
                registro.warning("La tipografía «%s» no está instalada; la portada usa Calibri.", pedida)
                _AVISADAS.add(pedida)
            fuente = "Calibri"
            ttf_b, ttf_r = _buscar_fuente("calibri", True), _buscar_fuente("calibri", False)
        return cls(fuente=fuente, ttf_negrita=ttf_b, ttf_regular=ttf_r or ttf_b, **valores)

    def todos(self):
        return {self.primario, self.gris_oscuro, self.gris_claro, self.texto, self.blanco}
