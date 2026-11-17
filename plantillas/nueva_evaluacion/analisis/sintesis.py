"""Síntesis de los hallazgos (capa de evaluación mínima; plantilla de src/nucleo/nueva_evaluacion.py).

Lee evaluaciones/<id>/04_analisis/hallazgos.yaml, verifica los hallazgos con las reglas del núcleo y escribe:
    04_analisis/matriz_triangulacion_preguntas.xlsx   preguntas, hallazgos y evidencias, con la confianza
    06_informe/rastro_hallazgos.csv                    rastro técnico de cada hallazgo, para el anexo técnico

Uso (desde la raíz del repositorio):
    uv run python src/analisis_<id>/sintesis.py
"""
import pandas as pd

from comun import CFG, contar, leer_hallazgos, ruta, verificar


def tablas(P, H):
    preguntas = CFG["preguntas"]
    nombre = CFG["instrumentos"]
    filas_p = []
    for k, q in preguntas.items():
        hs = [h for h in H if h["pregunta"] == k]
        filas_p.append({"pregunta": k, "criterio": q["criterio"], "texto": " ".join(str(q["texto"]).split()),
                        "instrumentos": "; ".join(nombre[i] for i in q["instrumentos"]),
                        "hallazgos": len(hs), "pendientes": sum(h["confianza"] == "No determinado" for h in hs),
                        "valoracion": P.get(k, {}).get("valoracion", ""), "respuesta": P.get(k, {}).get("respuesta", "")})
    filas_h = [{"hallazgo": h["id"], "pregunta": h["pregunta"], "aspecto": h["aspecto"], "enunciado": h["enunciado"],
                "confianza": h["confianza"], "fuentes_independientes": contar(h), "estado": h["estado"],
                "detalle": h["detalle"], "salvedad": h["salvedad"]} for h in H]
    filas_e = [{"hallazgo": h["id"], **e} for h in H for e in h["evidencias"]]
    columnas_e = ["hallazgo", "plano", "instrumento", "texto", "fuente", "base", "n_fuentes"]
    return pd.DataFrame(filas_p), pd.DataFrame(filas_h), pd.DataFrame(filas_e, columns=columnas_e)


def main():
    P, H, R = leer_hallazgos()
    verificar(P, H, R)
    tp, th, te = tablas(P, H)
    salida = ruta("analisis", "matriz_triangulacion_preguntas.xlsx")
    with pd.ExcelWriter(salida, engine="openpyxl") as w:
        tp.to_excel(w, sheet_name="Preguntas", index=False)
        th.to_excel(w, sheet_name="Hallazgos", index=False)
        te.to_excel(w, sheet_name="Evidencias", index=False)
    rastro = th[["hallazgo", "pregunta", "confianza", "fuentes_independientes"]].copy()
    rastro["rastro"] = [" | ".join(f"{e['instrumento']}: {e['base']}" for e in h["evidencias"]) for h in H]
    ruta("informe").mkdir(parents=True, exist_ok=True)
    rastro.to_csv(ruta("informe", "rastro_hallazgos.csv"), index=False, encoding="utf-8-sig", lineterminator="\n")
    pend = int(tp["pendientes"].sum())
    print(f"Síntesis de {CFG.id}: {len(tp)} preguntas, {len(th)} hallazgos ({pend} pendientes), {len(te)} evidencias, "
          f"{len(R)} recomendaciones; verificación sin observaciones.")


if __name__ == "__main__":
    main()
