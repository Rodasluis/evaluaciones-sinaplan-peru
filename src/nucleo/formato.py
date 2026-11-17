"""Formato de cifras y textos para los informes: números con espacio de miles y coma decimal, montos en millones de
soles, porcentajes, enumeraciones en español y comillas tipográficas (“ ” y, dentro de una cita, ‘ ’)."""
from docx.oxml.ns import qn


ABREN, CIERRAN = "«“", "»”"


def comillas(texto, nivel=0):
    """Cambia « » por “ ” y, dentro de una cita, por ‘ ’. Devuelve el texto y el nivel de anidamiento al final, para
    seguir la cuenta en el run siguiente del mismo párrafo."""
    out = []
    for ch in texto:
        if ch in ABREN:
            out.append("“" if nivel == 0 else "‘")
            nivel += 1
        elif ch in CIERRAN:
            nivel = max(nivel - 1, 0)
            out.append("”" if nivel == 0 else "’")
        else:
            out.append(ch)
    return "".join(out), nivel


def comillas_parrafo(par):
    """Aplica comillas() a los textos de un párrafo de Word en orden, sin tocar los códigos de campo."""
    nivel = 0
    for t in par.iter(qn("w:t")):
        if t.text and any(c in t.text for c in ABREN + CIERRAN):
            t.text, nivel = comillas(t.text, nivel)


def fmt_num(x, dec=0):
    """Número con separador de miles en espacio y coma decimal (estilo de la nota)."""
    s = f"{x:,.{dec}f}"
    return s.replace(",", " ").replace(".", ",")


def soles(x):
    return f"S/ {fmt_num(round(x))}"


def monto(x):
    """Monto en soles para el texto del informe (ronda 35, pedido del usuario: sin «miles»): desde un millón, en
    millones con dos decimales (3,31 millones de soles); por debajo, en soles enteros (845 210 soles)."""
    if abs(x) >= 1e6:
        return f"{fmt_num(x / 1e6, 2)} millones de soles"
    return f"{fmt_num(round(x))} soles"


def millones(x, vacio="s. m."):
    """Monto en millones de soles con dos decimales, para tablas cuya cabecera ya dice «millones de soles»."""
    return fmt_num(x / 1e6, 2) if x else vacio


def pct(x, dec=1):
    return f"{fmt_num(x, dec)} %"


def lista_es(items, conj="y"):
    items = list(items)
    if len(items) <= 1:
        return "".join(items)
    return ", ".join(items[:-1]) + f" {conj} " + items[-1]
