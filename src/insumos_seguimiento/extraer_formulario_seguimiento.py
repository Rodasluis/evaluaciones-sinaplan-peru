"""Extrae las respuestas de la PNIG del formulario web «Sistema de Seguimiento de Políticas» de la OMEP (2023–2025).

Los libros originales contienen las cuatro políticas del MIMP, hojas de usuarios y enlaces a Google Drive, por lo que se
mantienen fuera del repositorio. Este script deja una sola hoja por año con las filas de la PNIG y las columnas
relevantes para la evaluación, y construye un panel armonizado para el análisis.

Uso:
    uv run --no-project --with openpyxl --with pandas python extraer_formulario_seguimiento.py <dir_origen> <dir_raw> <csv_panel>
"""
import re
import sys
from pathlib import Path

import pandas as pd

ORIGEN, RAW, PANEL = Path(sys.argv[1]), Path(sys.argv[2]), Path(sys.argv[3])
LIBRO = "Formulario de Seguimiento de Politicas MIMP {}.xlsx"


def adjunto(v):
    """Sustituye un enlace por una marca de existencia: los enlaces de Drive no se publican."""
    return "Sí" if isinstance(v, str) and v.strip().startswith("http") else "No"


def actividades_poi(r):
    """Compacta los cinco bloques de actividades operativas del POI 2023 en un solo texto."""
    partes = []
    for k in range(1, 6):
        ao = str(r.get(f"Actividad operativa {k} (AO{k})", "-")).strip()
        if ao in ("-", "nan", ""):
            continue
        meta = r.get(f"Meta física \nAO{k}" if k == 1 else f"Meta física AO{k}")
        partes.append(f"AO{k}: {ao.lstrip('-').strip()} | AEI: {str(r.get(f'AEI ({k}) asociada al servicio')).lstrip('-').strip()}"
                      f" | unidad: {str(r.get(f'Unidad de medida AO{k}')).lstrip('-').strip()}"
                      f" | meta física: {str(meta).lstrip('-').strip()}"
                      f" | ejecución física: {str(r.get(f'Ejecución física AO{k}')).lstrip('-').strip()}"
                      f" | presupuesto programado: {str(r.get(f'Presupuesto programado AO{k}')).lstrip('-').strip()}"
                      f" | presupuesto ejecutado: {str(r.get(f'Presupuesto ejecutado AO{k}')).lstrip('-').strip()}")
    return "\n".join(partes) if partes else None


# ───────────────────────────── 2023 ─────────────────────────────
d = pd.read_excel(ORIGEN / LIBRO.format(2023), sheet_name="Data_2")
d = d[d["Política"] == "PNIG"].copy()
d["Actividades operativas del POI"] = d.apply(actividades_poi, axis=1)
d["Adjuntó base de datos"] = d["BD de Informacion reportada"].map(adjunto)
y23 = d.rename(columns={
    "Responsable": "Entidad y órgano responsable", "Servicio": "Objetivo prioritario o servicio",
    "Sentido del indicador": "Sentido", "Valor del indicador en el periodo": "Valor del indicador",
    "Factores favorables para alcanzar y/o superar la meta": "Factores favorables",
    "Factores que dificultaron el cumplimiento de la meta": "Factores que dificultaron",
    "Medidas que deben ser adoptadas": "Medidas propuestas", "Servicio se encuentra en el POI?": "Servicio en el POI",
    "¿Se requiere añadir actividades operativas?": "Requiere añadir actividades operativas",
    "¿Que actividades operativas se requerirían añadir?": "Actividades operativas por añadir",
    "¿Se requiere la participación de otras entidades?": "Requiere otras entidades",
    "¿Existe dificultad en la coordinación?": "Dificultad de coordinación",
    "¿Porque se difulta la coordinación?": "Causa de la dificultad de coordinación",
    "¿Se esta generando efectos no previstos?": "Efectos no previstos",
    "Efecto no previsto por la provisión del servicio": "Descripción del efecto no previsto",
    "Motivo por el que no reporta": "Motivo de no reporte", "Estado": "Estado del registro",
    "Fuente 2": "Tipo de fuente", "entidad": "Origen del responsable"})[[
    "ID", "Nivel", "Objetivo prioritario o servicio", "Indicador", "Entidad y órgano responsable", "Origen del responsable",
    "Sentido", "Fuente", "Tipo de fuente", "Motivo de no reporte", "Numerador", "Denominador", "Valor del indicador", "Meta",
    "Nivel de avance", "Nivel de cumplimiento", "Adjuntó base de datos", "Factores favorables", "Buenas prácticas implementadas",
    "Factores que dificultaron", "Medidas propuestas", "Servicio en el POI", "Actividades operativas del POI",
    "Requiere añadir actividades operativas", "Actividades operativas por añadir", "Requiere otras entidades",
    "Dificultad de coordinación", "Causa de la dificultad de coordinación", "Efectos no previstos",
    "Descripción del efecto no previsto", "Contexto económico", "Contexto político", "Contexto social", "Estado del registro"]]

# ───────────────────────────── 2024 ─────────────────────────────
d = pd.read_excel(ORIGEN / LIBRO.format(2024), sheet_name="Data")
d = d[d["Política"] == "PNIG"].copy()
d["Adjuntó base de datos"] = d["BD de Informacion reportada"].map(adjunto)
y24 = d.rename(columns={
    "Servicio": "Objetivo prioritario o servicio", "Entidad responsable": "Entidad", "Dirreción responsable": "Órgano responsable",
    "Sentido del indicador": "Sentido", "¿Dispone de la información para reportar el avance del indicador?": "Dispone de información",
    "Motivo por el que no reporta": "Motivo de no reporte", "Valor del indicador en el periodo": "Valor del indicador",
    "Factores favorables para alcanzar y/o superar la meta": "Factores favorables",
    "Factores que dificultaron el cumplimiento de la meta": "Factores que dificultaron",
    "Causas de la falla de planeación hacia ariba": "Causas de posible falla de planeación hacia arriba",
    "Causas de la falla de planeación hacia abajo": "Causas de posible falla de planeación hacia abajo",
    "Medidas que deben ser adoptadas": "Medidas propuestas",
    "¿Considera que existen oportunidades de mejora?": "Considera oportunidades de mejora",
    "Oportunidades de mejora identificadas": "Oportunidades de mejora", "Estado": "Estado del registro"})[[
    "ID", "Nivel", "Objetivo prioritario o servicio", "Indicador", "Entidad", "Órgano responsable", "Sentido",
    "Dispone de información", "Motivo de no reporte", "Numerador", "Denominador", "Valor del indicador", "Meta",
    "Nivel de avance", "Nivel de cumplimiento", "Adjuntó base de datos", "Factores favorables", "Buenas prácticas implementadas",
    "Factores que dificultaron", "Causas de posible falla de planeación hacia arriba",
    "Causas de posible falla de planeación hacia abajo", "Medidas propuestas", "Considera oportunidades de mejora",
    "Oportunidades de mejora", "Estado del registro"]]

# ───────────────────────────── 2025 ─────────────────────────────
d = pd.read_excel(ORIGEN / LIBRO.format(2025), sheet_name="BD PNIG")
st = pd.read_excel(ORIGEN / LIBRO.format(2025), sheet_name="BD Indicadores").set_index("ID")["Record Status"]
d["Estado en el aplicativo"] = d["ID"].map(st).map({"approved": "Aprobado", "submitted": "Enviado", "draft": "Borrador"})
d["Adjuntó información departamental"] = d["Info departamental"].map(adjunto)
y25 = d.rename(columns={
    "Servicio": "Objetivo prioritario o servicio", "Direccion": "Órgano responsable", "Nivel indicador": "Nivel",
    "Tipo indicador": "Tipo de indicador", "Tiene info": "Dispone de información", "Motivo no reporte": "Motivo de no reporte",
    "Valor indicador": "Valor del indicador", "Avance": "Nivel de avance", "Cumplimiento": "Nivel de cumplimiento",
    "Factores logro": "Factores favorables", "Buenas practicas": "Buenas prácticas implementadas",
    "Factores dificultad": "Factores que dificultaron", "Medidas mejora": "Medidas propuestas",
    "Brecha alta": "Causas de posible falla de planeación hacia arriba",
    "Brecha baja": "Causas de posible falla de planeación hacia abajo", "Requiere otras entidades": "Requiere otras entidades",
    "Dificultad coordinacion": "Dificultad de coordinación", "Causa dificultad": "Causa de la dificultad de coordinación",
    "Contexto": "Contexto", "Efectos no previstos": "Efectos no previstos", "Estado": "Estado del registro"})[[
    "ID", "Nivel", "Objetivo prioritario o servicio", "Indicador", "Entidad", "Órgano responsable", "Sentido",
    "Tipo de indicador", "Dispone de información", "Motivo de no reporte", "Numerador", "Denominador", "Valor del indicador",
    "Meta", "Nivel de avance", "Nivel de cumplimiento", "Adjuntó información departamental",
    "Factores favorables", "Buenas prácticas implementadas", "Factores que dificultaron",
    "Causas de posible falla de planeación hacia arriba", "Causas de posible falla de planeación hacia abajo",
    "Medidas propuestas", "Requiere otras entidades", "Dificultad de coordinación", "Causa de la dificultad de coordinación",
    "Contexto", "Efectos no previstos", "Estado del registro", "Estado en el aplicativo"]]

# ─────────────────────── una hoja por año ───────────────────────
RAW.mkdir(parents=True, exist_ok=True)
for anio, df in [(2023, y23), (2024, y24), (2025, y25)]:
    df = df.sort_values("ID").reset_index(drop=True)
    with pd.ExcelWriter(RAW / f"{anio}_omep_formulario-seguimiento-pnig.xlsx", engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name=f"PNIG_{anio}", index=False)
        ws = xw.sheets[f"PNIG_{anio}"]
        ws.freeze_panes = "B2"
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 14 if col[0].value in ("ID", "Nivel", "Meta") else 28
    print(anio, df.shape)


# ─────────────────────── panel armonizado ───────────────────────
def num(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = str(v).strip().replace("%", "").replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def cumplimiento(v):
    s = str(v).lower()
    if "arriba" in s or "falla alta" in s:
        return "Posible falla hacia arriba"
    if "abajo" in s or "falla baja" in s:
        return "Posible falla hacia abajo"
    if "posible falla" in s:
        return "Posible falla (sin sentido)"
    for k in ("Alto", "Medio", "Bajo"):
        if s == k.lower():
            return k
    return "Sin cálculo"


filas = []
for anio, df in [(2023, y23), (2024, y24), (2025, y25)]:
    for _, r in df.iterrows():
        m = re.match(r"PNIG_I(OP|SS)\.(\d+)\.(\d+)(?:\.(\d+)\.(\d+))?", r["ID"])
        es_op = m.group(1) == "OP"
        if anio == 2023:
            ent, _, org = str(r["Entidad y órgano responsable"]).partition("/")
        else:
            ent, org = r["Entidad"], r["Órgano responsable"]
        ent = str(ent).strip()
        if ent.startswith("MIMP/") or ent.endswith("-MIMP"):      # 2024 registra órganos del MIMP como entidad
            org, ent = ent.replace("MIMP/", "").replace("-MIMP", ""), "MIMP"
        av = num(r["Nivel de avance"])
        if anio == 2023 and av is not None:
            av *= 100                      # en 2023 el avance se registró como razón
        valor = num(r["Valor del indicador"])
        dispone = r.get("Dispone de información")
        if anio == 2023:
            dispone = "No" if isinstance(r["Motivo de no reporte"], str) else "Sí"
        contexto = r.get("Contexto")
        if anio == 2023:
            contexto = " | ".join(f"{k}: {r[k]}" for k in ("Contexto económico", "Contexto político", "Contexto social")
                                  if isinstance(r[k], str) and r[k].strip() not in ("-", ""))
        filas.append({
            "anio": anio, "id": r["ID"], "nivel": "OP" if es_op else "Servicio",
            "op": f"OP.{int(m.group(2)):02d}",
            "servicio": None if es_op else f"{int(m.group(2))}.{int(m.group(3))}.{int(m.group(4))}",
            "indicador": r["Indicador"], "entidad": str(ent).strip(), "organo": str(org).strip() or None,
            "dispone_informacion": dispone, "motivo_no_reporte": r["Motivo de no reporte"],
            "valor": valor, "meta": num(r["Meta"]), "avance_pct": av, "cumplimiento": cumplimiento(r["Nivel de cumplimiento"]),
            "reporta_valor": valor is not None, "adjunto_bd": r.get("Adjuntó base de datos"),
            "factores_favorables": r["Factores favorables"], "buenas_practicas": r["Buenas prácticas implementadas"],
            "factores_dificultad": r["Factores que dificultaron"], "medidas": r["Medidas propuestas"],
            "falla_arriba": r.get("Causas de posible falla de planeación hacia arriba"),
            "falla_abajo": r.get("Causas de posible falla de planeación hacia abajo"),
            "requiere_otras_entidades": r.get("Requiere otras entidades"), "dificultad_coordinacion": r.get("Dificultad de coordinación"),
            "causa_dificultad": r.get("Causa de la dificultad de coordinación"), "contexto": contexto or None,
            "efectos_no_previstos": r.get("Descripción del efecto no previsto") if anio == 2023 else r.get("Efectos no previstos"),
            "oportunidades_mejora": r.get("Oportunidades de mejora"), "servicio_en_poi": r.get("Servicio en el POI"),
            "actividades_poi": r.get("Actividades operativas del POI"), "estado_registro": r["Estado del registro"]})
panel = pd.DataFrame(filas)
for c in ("requiere_otras_entidades", "dificultad_coordinacion", "servicio_en_poi"):
    panel[c] = panel[c].replace({"yes": "Sí", "no": "No", "Si": "Sí"})
panel["dispone_informacion"] = panel["dispone_informacion"].replace({"Si": "Sí"})
PANEL.parent.mkdir(parents=True, exist_ok=True)
panel.to_csv(PANEL, index=False, encoding="utf-8-sig")
print("panel", panel.shape)
