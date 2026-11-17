"""Catálogo de módulos del portal de microdatos del INEI (proyectos.inei.gob.pe/microdatos).

Consulta los tres formularios del portal (CambiaEnc.asp, CambiaAnio.asp y cambiaPeriodo.asp), igual que lo hace la
página «Consulta por encuestas», y guarda la lista de módulos con su código de descarga. No descarga microdatos: solo
lee la tabla de módulos. Las solicitudes van con la verificación TLS activa y sin suplantar un navegador.

Método tomado del paquete público inei-microdatos (github.com/fiorellarmartins/inei-microdatos, client.py): los
valores se codifican con la convención escape() de JavaScript y el cuerpo en latin-1.

Uso (desde la raíz del repositorio):
    uv run --no-project --with requests python src/fuentes/catalogo_microdatos_inei.py ENAHO ENDES ENARES ENUT
Salida: evaluaciones/pnig_2026/03_datos/raw/microdatos_inei/catalogo_<encuesta>.csv (no se versiona).
"""
import csv
import os
import re
import sys
import time

import requests
import os as _os, sys as _sys  # noqa: E401
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))  # src/, núcleo
from nucleo.config import cargar as _cargar_config  # noqa: E402
EVAL = _cargar_config().raiz  # carpeta de la evaluación (variable EVALUACION; por defecto pnig_2026)

BASE = "https://proyectos.inei.gob.pe/microdatos/"
ENCUESTAS = {
    "ENAHO": "Condiciones de Vida y Pobreza - ENAHO",
    "ENDES": "Encuesta Demográfica y de Salud Familiar - ENDES",
    "ENARES": "Encuesta Nacional Sobre Relaciones Sociales - ENARES",
    "ENUT": "Encuesta Nacional de Uso del Tiempo - ENUT",
    "EPEN": "EPEN  DEPARTAMENTOS",
    "ENA": "ENCUESTA NACIONAL AGROPECUARIA",
    "ENAPRES": "Encuesta Nacional de Programas Presupuestales - ENAPRES",
    "RENAMU": "Registro Nacional de Municipalidades - RENAMU",
    "VICTIMIZACION": "Encuesta Nacional Especializada sobre Victimización",
    "EMPRESAS": "Encuesta Nacional de Empresas",
}
SALIDA = f"{EVAL}/03_datos/raw/microdatos_inei"
DESDE = {"ENAHO": 2016, "ENDES": 2016, "EPEN": 2019, "ENA": 2016, "ENAPRES": 2016, "RENAMU": 2016,
         "VICTIMIZACION": 2016, "EMPRESAS": 2016}  # años previos no se usan en la evaluación
PAUSA = 1.5  # segundos entre solicitudes, para no cargar el portal


def js_escape(s):
    seguro = set("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789@*_+-./")
    return "".join(c if c in seguro else f"%{ord(c):02X}" for c in s)


def post(sesion, recurso, cuerpo):
    time.sleep(PAUSA)
    r = sesion.post(BASE + recurso, data=cuerpo.encode("latin-1"), timeout=120,
                    headers={"Content-Type": "application/x-www-form-urlencoded"})
    r.raise_for_status()
    return r.content.decode("latin-1")


def opciones(html):
    return [(v, re.sub(r"<[^>]+>", "", t).strip())
            for v, t in re.findall(r"<option[^>]*value=['\"]?([^'\" >]*)['\"]?[^>]*>(.*?)</option>", html, re.S)
            if v]


def modulos(html):
    filas = []
    for fila in re.findall(r"<tr[^>]*>(.*?)</tr>", html, re.S):
        tds = re.findall(r"<td[^>]*>(.*?)</td>", fila, re.S)
        if len(tds) < 7:
            continue
        limpia = lambda s: re.sub(r"<[^>]+>", "", s).strip()
        cod_enc, cod_mod, nombre = limpia(tds[3]), limpia(tds[5]), limpia(tds[6])
        if not cod_enc.isdigit() or not cod_mod:
            continue
        stata = re.findall(r"descarga/STATA/([^\"'>]+)\.zip", fila)
        spss = re.findall(r"descarga/SPSS/([^\"'>]+)\.zip", fila)
        filas.append({"codigo_encuesta": cod_enc, "codigo_modulo": cod_mod, "modulo": nombre,
                      "archivo_stata": stata[0] if stata else "", "archivo_spss": spss[0] if spss else ""})
    return filas


def catalogo(clave):
    nombre = ENCUESTAS[clave]
    enc = js_escape(nombre)
    s = requests.Session()
    s.get(BASE + "Consulta_por_Encuesta.asp?CU=19558", timeout=120)
    anios = [v for v, _ in opciones(post(s, "CambiaEnc.asp", f"bandera=1&_cmbEncuesta={enc}"))]
    filas = []
    for anio in anios:
        if not anio.isdigit() or int(anio) < DESDE.get(clave, 2010):
            continue
        periodos = opciones(post(s, "CambiaAnio.asp", f"bandera=1&_cmbEncuesta={enc}&_cmbAnno={anio}"
                                                      f"&_cmbEncuesta0={enc}"))
        for valor, etiqueta in periodos:
            html = post(s, "cambiaPeriodo.asp", f"bandera=1&_cmbEncuesta={enc}&_cmbAnno={anio}"
                                                f"&_cmbTrimestre={js_escape(valor)}")
            for m in modulos(html):
                filas.append({"encuesta": clave, "anio": anio, "periodo": etiqueta, **m})
    os.makedirs(SALIDA, exist_ok=True)
    ruta = f"{SALIDA}/catalogo_{clave.lower()}.csv"
    with open(ruta, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["encuesta", "anio", "periodo", "codigo_encuesta", "codigo_modulo",
                                          "modulo", "archivo_stata", "archivo_spss"])
        w.writeheader()
        w.writerows(filas)
    return ruta, len(filas), anios


if __name__ == "__main__":
    for clave in sys.argv[1:] or ENCUESTAS:
        ruta, n, anios = catalogo(clave)
        print(clave, n, "módulos;", "años:", ", ".join(anios[:30]), "->", ruta)
