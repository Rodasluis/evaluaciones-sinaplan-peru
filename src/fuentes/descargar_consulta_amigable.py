"""Descarga de la Consulta Amigable del MEF la ejecución presupuestal del pliego 039 (MIMP), 2019–2025 (REV_IND.3).

Recorre, para cada año: nivel de gobierno E → sector 39 → pliego 039 → unidad ejecutora. Guarda los totales de todas las
unidades ejecutoras y, para la 001 (Administración Nivel Central, donde están la DGIGND y la OGMEPGD), el detalle por
categoría presupuestal → producto o proyecto → actividad, acción de inversión u obra. Montos en soles: PIA, PIM,
certificación, compromiso anual, atención de compromiso mensual, devengado y girado; avance = devengado / PIM.

La Consulta Amigable es una aplicación ASP.NET: cada nivel se obtiene con un envío del formulario de la página anterior
(el estado viaja en __VIEWSTATE), por lo que cada rama se consulta desde la página guardada de su nivel superior.
No sobrescribe archivos existentes y registra cada archivo con su hash.

Uso (desde la raíz del repositorio):
    uv run --no-project --with requests --with lxml python src/fuentes/descargar_consulta_amigable.py
"""
import csv
import hashlib
import re
import time
from datetime import date
from pathlib import Path

import requests
from lxml import html
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

URL = "https://apps5.mineco.gob.pe/transparencia/Navegador/Navegar_7.aspx?y={y}&ap=ActProy"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
ANIOS = range(2019, 2026)
DESTINO = Path(f"{EVAL}/03_datos/raw/consulta_amigable")
LOG = Path(f"docs/fuentes/descargas_consulta-amigable_{date.today().isoformat()}.tsv")
MONTOS = ["pia", "pim", "certificacion", "compromiso_anual", "atencion_compromiso_mensual", "devengado", "girado"]


def campos(doc):
    d = {i.get("name"): i.get("value", "") for i in doc.xpath("//input[@type='hidden']") if i.get("name")}
    for s in doc.xpath("//select"):
        op = s.xpath("./option[@selected]") or s.xpath("./option")
        d[s.get("name")] = op[0].get("value") if op else ""
    return d


def filas(doc):
    """[(valor del radio, código, denominación, montos)] de la tabla de la página."""
    out = []
    for tr in doc.xpath("//tr[.//input[@type='radio']]"):
        c = [re.sub(r"\s+", " ", td.text_content()).strip() for td in tr.xpath("./td")]
        codigo, _, nombre = c[1].partition(":")
        out.append((tr.xpath(".//input[@type='radio']")[0].get("value"), codigo.strip(), nombre.strip(),
                    [float(x.replace(",", "")) if x else 0.0 for x in c[2:9]]))
    return out


def boton(doc, prefijo):
    return next((b.get("name"), b.get("value")) for b in doc.xpath("//input[@type='submit']")
                if b.get("name") and b.get("value", "").startswith(prefijo))


class Consulta:
    def __init__(self, anio):
        self.s, self.anio = requests.Session(), anio
        self.inicio = html.fromstring(self._pedir())

    def _pedir(self, datos=None):
        for intento in range(4):
            try:
                r = (self.s.post if datos else self.s.get)(URL.format(y=self.anio), data=datos, headers=UA, timeout=120)
                r.raise_for_status()
                time.sleep(0.6)
                return r.content
            except requests.RequestException:
                time.sleep(5 * (intento + 1))
        raise RuntimeError(f"La Consulta Amigable no respondió ({self.anio})")

    def bajar(self, doc, siguiente, codigo=None):
        """Desde la página `doc`, selecciona la fila `codigo` (si se da) y pide el nivel `siguiente`."""
        datos = campos(doc)
        if codigo is not None:
            datos["grp1"] = next(v for v, c, _, _ in filas(doc) if c == codigo or c.startswith(codigo + "-"))
        nombre, valor = boton(doc, siguiente)
        datos[nombre] = valor
        return html.fromstring(self._pedir(datos))


def actualizacion(doc):
    m = re.search(r"\d{1,2} de \w+ de \d{4}", doc.text_content())
    return m.group(0) if m else ""


if __name__ == "__main__":
    DESTINO.mkdir(parents=True, exist_ok=True)
    for anio in ANIOS:
        salida = DESTINO / f"{anio}_mef_consulta-amigable_pliego-039.csv"
        if salida.exists():
            continue
        q = Consulta(anio)
        d = q.bajar(q.inicio, "Nivel de Gobierno")
        d = q.bajar(d, "Sector", "E")
        d = q.bajar(d, "Pliego", "39")
        ues = q.bajar(d, "Ejecutora", "039")
        registros = [{"anio": anio, "nivel": "unidad_ejecutora", "unidad_ejecutora": f"{c}: {n}", "categoria": "",
                      "producto": "", "actividad": ""} | dict(zip(MONTOS, m)) for _, c, n, m in filas(ues)]
        ue = next(f"{c}: {n}" for _, c, n, _ in filas(ues) if c.startswith("001"))
        cats = q.bajar(ues, "Categoría", "001")
        for _, cc, cn, _ in filas(cats):
            prods = q.bajar(cats, "Producto", cc)
            for _, pc, pn, _ in filas(prods):
                acts = q.bajar(prods, "Act.", pc)
                for _, ac, an, m in filas(acts):
                    registros.append({"anio": anio, "nivel": "actividad", "unidad_ejecutora": ue, "categoria": f"{cc}: {cn}",
                                      "producto": f"{pc}: {pn}", "actividad": f"{ac}: {an}"} | dict(zip(MONTOS, m)))
        with open(salida, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=list(registros[0]))
            w.writeheader()
            w.writerows(registros)
        h = hashlib.sha256(salida.read_bytes()).hexdigest()
        nuevo = not LOG.exists()
        with open(LOG, "a", encoding="utf-8") as f:
            if nuevo:
                f.write("archivo\turl\tfecha_consulta\tactualizacion_mef\tregistros\tsha256\n")
            f.write(f"{salida.as_posix()}\t{URL.format(y=anio)}\t{date.today().isoformat()}\t{actualizacion(ues)}\t"
                    f"{len(registros)}\t{h}\n")
        print(f"OK {anio}: {len(registros) - len(filas(ues))} actividades de la UE 001; actualización MEF: {actualizacion(ues)}")
