"""Equivalencia de los instrumentos genéricos (src/nucleo/instrumentos/) con el piloto PNIG (ADR-017).

Convierte los libros e insumos versionados del piloto a las entradas estándar, corre los instrumentos genéricos en una
carpeta temporal y comprueba que reproducen las clasificaciones del piloto:
    revisión documental   clasificación por criterios y estado asignado de los 52 servicios;
    indicadores           avance y nivel de cada indicador y año con valor;
    evidencia             solidez y dirección de cada tipo de intervención;
    coherencia            valoración de cada servicio y alineamiento vertical por objetivo y momento.
El piloto no se toca: si aparece una diferencia, se corrige el genérico.
"""
import pandas as pd
import pytest
import yaml

from nucleo.config import RAIZ_REPO, cargar_desde
from nucleo.instrumentos import base, rev_ali, rev_doc, rev_evi, rev_ind

PILOTO = RAIZ_REPO / "evaluaciones" / "pnig_2026"
A = PILOTO / "04_analisis"
ANIOS = range(2019, 2026)


def libro(nombre, hoja):
    return pd.read_excel(A / nombre, sheet_name=hoja, dtype=str).fillna("")


@pytest.fixture
def ev(tmp_path):
    datos = yaml.safe_load((PILOTO / "config.yaml").read_text(encoding="utf-8"))
    datos["id"] = "pnig_equivalencia"
    carpeta = tmp_path / "evaluaciones" / "pnig_equivalencia"
    (carpeta / "02_insumos").mkdir(parents=True)
    (carpeta / "config.yaml").write_text(yaml.safe_dump(datos, allow_unicode=True), encoding="utf-8")
    return cargar_desde(tmp_path, "pnig_equivalencia")


def escribir(cfg, nombre, df):
    df.to_csv(cfg.abs("insumos", nombre), index=False, encoding="utf-8")


def servicios(cfg, d):
    escribir(cfg, "servicios.csv", pd.DataFrame({"codigo": d["Código"], "op": d["OP"], "lineamiento": d["Lineamiento"],
                                                 "servicio": d["Servicio"]}))


def test_revision_documental(ev):
    d = libro("REV_DOC_estado_implementacion_servicios.xlsx", "Estado por servicio")
    servicios(ev, d)
    reporta = d["Años con valor reportado"].map(lambda s: "Sí" if any(2019 <= int(a) <= 2025
                                                                       for a in base.lista(s, ",")) else "No")
    comp = pd.DataFrame({
        "codigo": d["Código"], "protocolo": d["Protocolo aprobado"],
        "asignacion": d["Asignación presupuestal identificable"], "reporte_indicador": reporta,
        **{f"ejecucion_{a}": d[f"LIO {a}"].map({"Sí": "Sí", "No": "No", "Mínima": "Sin información"}).fillna("")
           for a in ANIOS},
        "ajuste": [e if aj else "" for e, aj in zip(d["Estado asignado (preliminar)"], d["Ajuste por lectura cualitativa"])],
        "sustento_ajuste": d["Ajuste por lectura cualitativa"]})
    escribir(ev, "rev_doc_componentes.csv", comp)
    r = rev_doc.correr(ev)
    g = pd.read_excel(r["libro"], sheet_name="Estado por servicio", dtype=str).fillna("")
    assert list(g["Código"]) == list(d["Código"])
    assert list(g["Clasificación por criterios"]) == list(d["Clasificación por criterios"])
    assert list(g["Estado asignado"]) == list(d["Estado asignado (preliminar)"])
    assert (r["Implementado"], r["Implementado parcialmente"], r["No implementado"]) == (24, 26, 2)


def test_indicadores(ev):
    b = pd.read_excel(A / "REV_IND_base_indicadores.xlsx", sheet_name="Base REV_IND.1")
    # «Ascendente (acumulado anual)» es ascendente; «Promedio» y «N.D.» quedan sin sentido definido, como en el piloto
    sentido = b["Sentido aplicado"].str.extract(r"^(Ascendente|Descendente)", expand=False).fillna("")
    escribir(ev, "indicadores.csv", pd.DataFrame({"codigo": b["Código"], "op": b["OP"], "nombre": b["Indicador"],
                                                  "nivel": b["Nivel"].map(lambda n: "Objetivo prioritario" if n == "OP"
                                                                          else "Servicio"),
                                                  "sentido": sentido}))
    filas = [{"codigo": r["Código"], "anio": a, "meta": "" if pd.isna(r[f"Meta {a}"]) else repr(float(r[f"Meta {a}"])),
              "valor": "" if pd.isna(r[f"Valor {a}"]) else repr(float(r[f"Valor {a}"])), "fuente": "Piloto PNIG"}
             for _, r in b.iterrows() for a in ANIOS if pd.notna(r[f"Meta {a}"]) or pd.notna(r[f"Valor {a}"])]
    escribir(ev, "rev_ind_valores.csv", pd.DataFrame(filas))
    r = rev_ind.correr(ev)
    g = pd.read_excel(r["libro"], sheet_name="Avance por año")
    piloto = b.melt(id_vars=["Código"], value_vars=[f"Avance {a} (%)" for a in ANIOS], value_name="avance")
    piloto["Año"] = piloto.variable.str.extract(r"(\d{4})").astype(int)
    nivel = b.melt(id_vars=["Código"], value_vars=[f"Semáforo {a}" for a in ANIOS], value_name="nivel")
    piloto["nivel"] = nivel.nivel.values
    m = g.merge(piloto, on=["Código", "Año"])
    assert len(m) == len(b) * len(ANIOS)
    assert (m["Avance (%)"].notna() == m.avance.notna()).all()
    con = m[m.avance.notna()]
    assert ((con["Avance (%)"] - con.avance.round(1)).abs() < 0.051).all()
    assert (con["Nivel de avance"] == con.nivel.str.split(" \\(").str[0]).all()


def test_evidencia(ev):
    ins = A / "rev_evi_insumos"
    tip = pd.read_csv(ins / "tipologia.csv", dtype=str).fillna("")
    est = pd.read_csv(ins / "evidencia.csv", dtype=str).fillna("")
    serv = pd.read_csv(PILOTO / "02_insumos" / "servicios.csv", dtype=str).fillna("")
    escribir(ev, "servicios.csv", pd.DataFrame({"codigo": serv.codigo_corto, "op": serv.op,
                                                "lineamiento": serv.lineamiento, "servicio": serv.servicio}))
    escribir(ev, "rev_evi_tipos.csv", tip.rename(columns={"nodo_modelo": "causa_efecto"}))

    def incluida(r):   # criterio del piloto (rev_evi_evidencia.py, Anexo 6)
        if r.diseno in rev_evi.PRIMARIOS:
            return r.ambito.startswith(("Perú", "Ingreso medio", "Ingreso bajo y medio", "Ingreso bajo ("))
        return True

    inc = est.apply(incluida, axis=1)
    escribir(ev, "rev_evi_estudios.csv", est.assign(
        incluido=inc.map({True: "Sí", False: "No"}),
        comparable=(~est.ambito.str.startswith(("Ingreso alto", "No informado"))).map({True: "Sí", False: "No"}),
        motivo_exclusion=inc.map({True: "", False: "Evaluación primaria fuera del ámbito del criterio de inclusión"})))
    r = rev_evi.correr(ev)
    g = pd.read_excel(r["libro"], sheet_name="Matriz por tipo", dtype=str).fillna("")
    p = libro("REV_EVI_matriz_evidencia.xlsx", "Matriz por tipo")
    assert list(g.Tipo) == list(p.Tipo)
    assert list(g.Solidez) == list(p.Solidez)
    assert list(g["Dirección"]) == list(p["Dirección"])
    assert list(g.Plausibilidad) == list(p.Plausibilidad)


def test_coherencia_y_alineamiento(ev):
    c = libro("REV_ALI_coherencia_alineamiento.xlsx", "REV_ALI.1 Por servicio")
    servicios(ev, c)
    vig = [(cod, i) for cod, s in zip(c["Código"], c["Indicadores vigentes del servicio"]) for i in base.lista(s, ",")]
    escribir(ev, "indicadores.csv", pd.DataFrame({"codigo": [i for _, i in vig], "nivel": "Servicio", "op": "",
                                                  "servicio": [cod for cod, _ in vig], "nombre": "", "sentido": ""}))
    escribir(ev, "cadena.csv", pd.DataFrame({"codigo": c["Código"],
                                             "causa_efecto": c["Causa o efecto del modelo (Anexo 5)"],
                                             "tipo_vinculo": c["Tipo de vínculo"].where(
                                                 c["Tipo de vínculo"].isin(["Causa", "Efecto"]), ""),
                                             "estandar": c["Estándares (Anexo 5)"]}))
    ver = pd.read_csv(A / "rev_ali_insumos" / "vertical.csv", dtype=str).fillna("")
    momentos = {"Periodo y transición": ["Periodo", "Transición"]}
    ver = ver.assign(momento=ver.momento.map(lambda m: momentos.get(m, ["Periodo" if m.startswith("Periodo")
                                                                         else m]))).explode("momento")
    escribir(ev, "alineamiento.csv", ver.assign(tipo="Vertical", documento="")[
        ["tipo", "momento", "op", "instrumento", "referencia", "documento", "cita", "correspondencia", "sustento"]])
    r = rev_ali.correr(ev)
    g = pd.read_excel(r["libro"], sheet_name="Coherencia por servicio", dtype=str).fillna("")
    assert list(g["Valoración"]) == list(c["Valoración"])
    assert (r["servicios_completos"], r["servicios"]) == (24, 52)
    a = pd.read_excel(r["libro"], sheet_name="Alineamiento por objetivo", dtype=str).fillna("")
    p = libro("REV_ALI_coherencia_alineamiento.xlsx", "REV_ALI.2 Por objetivo")
    for momento, col in (("Periodo", "Valoración (periodo)"), ("Transición", "Valoración (transición)")):
        gen = a[a.Momento == momento].set_index("Objetivo prioritario")["Valoración"]
        assert gen.to_dict() == p.set_index("OP")[col].to_dict(), momento
