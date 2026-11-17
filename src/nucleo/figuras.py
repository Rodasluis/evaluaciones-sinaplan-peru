"""Motor de figuras de los informes: diagramas en HTML y CSS que Chrome sin ventana convierte en PNG (núcleo;
antes en src/analisis_pnig/figuras_informe.py).

Cada figura es una función que devuelve (html, ancho) con pagina(cuerpo, ancho, css_extra); renderizar() la
convierte en PNG con la escala de impresión. Las comillas « » del cuerpo pasan a “ ” (ronda 35).
"""
import html
import re
import subprocess
from pathlib import Path

from .entorno import herramienta
from .formato import comillas

CHROME = Path(herramienta("chrome"))  # config/entorno.yaml o variable CHROME
ESCALA = 3

# paleta base (validada con validate_palette.js sobre #ffffff)
AZUL, TINTA, TINTA2, TENUE, GRILLA = "#1F3864", "#1d1d1b", "#52514e", "#898781", "#e1e0d9"
ESTADO = {"Implementado": "#0ca30c", "Implementado parcialmente": "#fab219", "No implementado": "#d03b3b"}

CSS = """
*{box-sizing:border-box;margin:0;padding:0}
body{background:#fff;font-family:Arial,Helvetica,sans-serif;color:%(tinta)s;width:%(ancho)spx}
.fig{padding:18px 22px;position:relative}
.ley{display:flex;gap:18px;flex-wrap:wrap;font-size:12px;color:%(tinta2)s;margin-top:14px;align-items:center}
.ley .i{display:inline-flex;align-items:center;gap:6px}
.sw{width:13px;height:13px;border-radius:3px;display:inline-block}
.nota{font-size:11px;color:%(tenue)s;margin-top:8px;line-height:1.35}
.col-t{font-size:11px;font-weight:bold;letter-spacing:.04em;text-transform:uppercase;color:%(tinta2)s}
"""


def mezcla(hexa, t):
    """Mezcla el color con blanco: t = 0 da blanco, t = 1 el color."""
    r, g, b = (int(hexa[i:i + 2], 16) for i in (1, 3, 5))
    return "#" + "".join(f"{round(255 + (c - 255) * t):02x}" for c in (r, g, b))


def e(texto):
    return html.escape(str(texto))


def pagina(cuerpo, ancho, css_extra=""):
    estilo = CSS % {"tinta": TINTA, "tinta2": TINTA2, "tenue": TENUE, "ancho": ancho} + css_extra
    # comillas “ ” en las figuras del informe (ronda 35, pedido del usuario), con ‘ ’ dentro de una cita
    cuerpo = comillas(cuerpo)[0]
    return (f'<!doctype html><html lang="es"><head><meta charset="utf-8"><style>{estilo}</style></head>'
            f'<body><div class="fig">{cuerpo}</div><script>document.body.setAttribute("data-alto",'
            f'Math.ceil(document.documentElement.getBoundingClientRect().height));</script></body></html>')


def _chrome(perfil, *args):
    return subprocess.run([str(CHROME), "--headless=new", "--disable-gpu", "--hide-scrollbars", "--no-first-run",
                           "--no-default-browser-check", "--mute-audio", f"--user-data-dir={perfil}", *args],
                          capture_output=True, timeout=180)


def renderizar(ruta_html, ruta_png, ancho, perfil):
    uri = Path(ruta_html).resolve().as_uri()
    dom = _chrome(perfil, "--dump-dom", f"--window-size={ancho},3000", uri).stdout.decode("utf-8", "replace")
    alto = int(re.search(r'data-alto="(\d+)"', dom).group(1))
    if Path(ruta_png).exists():
        Path(ruta_png).unlink()
    _chrome(perfil, f"--screenshot={Path(ruta_png).resolve()}", f"--window-size={ancho},{alto}",
            f"--force-device-scale-factor={ESCALA}", "--default-background-color=ffffffff", uri)
    assert Path(ruta_png).exists(), f"Chrome no generó {ruta_png}"
    return alto


def leyenda(items):
    return '<div class="ley">' + "".join(
        f'<span class="i"><span class="sw" style="background:{c}"></span>{e(t)}</span>' for t, c in items) + "</div>"
