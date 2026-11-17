"""Estimación con diseño muestral complejo para las encuestas del INEI (ENAHO, ENDES, ENARES, ENUT). Núcleo
(antes src/analisis_pnig/svy.py, que lo reexporta).

Estimador de razón ponderado con varianza por linealización de Taylor (conglomerados dentro de estratos, con
corrección n_h/(n_h − 1)), el mismo que valida el repositorio «Indicadores ENDES» del usuario contra los errores
estándar publicados por el INEI. Los estratos con un solo conglomerado no aportan varianza.

Para las brechas se linealiza el contraste entre dos dominios del mismo diseño (mujeres y hombres): la diferencia
M − H y la razón M/H (por el logaritmo), de modo que el error estándar incluye la covarianza entre ambos dominios. Los
cambios entre años se tratan como muestras independientes.

Uso:
    from nucleo.encuestas import Diseno, estimar, brecha
    d = Diseno(df, peso="fac500a", estrato="estrato", psu="conglome")
    estimar(d, "pea", universo=df.pet == 1)               # proporción o promedio
    brecha(d, "pea", grupo=df.sexo, a="Mujer", b="Hombre")  # M, H, M − H y M/H con errores estándar
"""
import numpy as np
import pandas as pd

Z95 = 1.959963984540054


class Diseno:
    def __init__(self, df, peso, estrato, psu):
        self.w = pd.to_numeric(df[peso], errors="coerce").to_numpy(dtype=float)
        self.h = df[estrato].astype(str).to_numpy()
        self.c = (df[estrato].astype(str) + "|" + df[psu].astype(str)).to_numpy()
        self.n = len(df)


def _var_total(z, d, mascara):
    """Varianza del total de z (ya linealizada) entre conglomerados dentro de estratos."""
    t = pd.DataFrame({"h": d.h[mascara], "c": d.c[mascara], "z": z[mascara]}).groupby(["h", "c"], sort=False).z.sum()
    t = t.reset_index()
    g = t.groupby("h").z
    n_h = g.transform("size")
    dev = t.z - g.transform("mean")
    t = t.assign(n_h=n_h, d2=dev ** 2)
    t = t[t.n_h > 1]
    if t.empty:
        return 0.0
    return float((t.n_h / (t.n_h - 1) * t.d2).sum())


def _razon(y, d, dom):
    m = dom & ~np.isnan(y) & ~np.isnan(d.w) & (d.w > 0)
    W = d.w[m].sum()
    if W <= 0:
        return np.nan, m, W
    return float((d.w[m] * y[m]).sum() / W), m, W


def estimar(d, y, universo=None, escala=1.0):
    """Proporción (y en 0/1) o promedio de y en el universo. Devuelve estimación, ee, IC 95 %, cv y n."""
    y = np.asarray(y, dtype=float)
    dom = np.ones(d.n, bool) if universo is None else np.asarray(universo, bool)
    r, m, W = _razon(y, d, dom)
    if np.isnan(r):
        return {"estimacion": np.nan, "ee": np.nan, "ic_inf": np.nan, "ic_sup": np.nan, "cv": np.nan, "n": 0}
    z = np.zeros(d.n)
    z[m] = d.w[m] * (y[m] - r) / W
    ee = np.sqrt(_var_total(z, d, m | (d.w > 0)))
    return {"estimacion": escala * r, "ee": escala * ee, "ic_inf": escala * (r - Z95 * ee),
            "ic_sup": escala * (r + Z95 * ee), "cv": 100 * ee / r if r else np.nan, "n": int(m.sum())}


def brecha(d, y, grupo, a, b, universo=None, escala=1.0):
    """Indicador en los dominios a y b (p. ej., mujeres y hombres), su diferencia a − b y su razón a/b."""
    y = np.asarray(y, dtype=float)
    grupo = np.asarray(grupo)
    base = np.ones(d.n, bool) if universo is None else np.asarray(universo, bool)
    ra, ma, Wa = _razon(y, d, base & (grupo == a))
    rb, mb, Wb = _razon(y, d, base & (grupo == b))
    za = np.zeros(d.n)
    zb = np.zeros(d.n)
    za[ma] = d.w[ma] * (y[ma] - ra) / Wa
    zb[mb] = d.w[mb] * (y[mb] - rb) / Wb
    todo = ma | mb
    ee_a = np.sqrt(_var_total(za, d, todo))
    ee_b = np.sqrt(_var_total(zb, d, todo))
    ee_dif = np.sqrt(_var_total(za - zb, d, todo))
    razon = ra / rb if rb else np.nan
    ee_log = np.sqrt(_var_total(za / ra - zb / rb, d, todo)) if ra and rb else np.nan
    return {
        "valor_a": escala * ra, "ee_a": escala * ee_a, "n_a": int(ma.sum()),
        "valor_b": escala * rb, "ee_b": escala * ee_b, "n_b": int(mb.sum()),
        "diferencia": escala * (ra - rb), "ee_diferencia": escala * ee_dif,
        "razon": razon, "ee_razon": razon * ee_log if ee_log == ee_log else np.nan,
    }


DEPARTAMENTOS = {"01": "Amazonas", "02": "Áncash", "03": "Apurímac", "04": "Arequipa", "05": "Ayacucho",
                 "06": "Cajamarca", "07": "Callao", "08": "Cusco", "09": "Huancavelica", "10": "Huánuco", "11": "Ica",
                 "12": "Junín", "13": "La Libertad", "14": "Lambayeque", "15": "Lima", "16": "Loreto",
                 "17": "Madre de Dios", "18": "Moquegua", "19": "Pasco", "20": "Piura", "21": "Puno",
                 "22": "San Martín", "23": "Tacna", "24": "Tumbes", "25": "Ucayali"}


# Ámbitos de publicación del INEI: Lima se separa en Lima Metropolitana (provincia de Lima, 43 distritos) y Región
# Lima (las otras nueve provincias); el Callao va aparte. La EPEN solo distingue «Provincia de Lima, incluye la
# Provincia Constitucional del Callao» y «Región de Lima» (variable DIVISION_LIMA); la ENA publica Lima entera.
AMBITOS = {**{k: v for k, v in DEPARTAMENTOS.items() if k != "15"}, "LM": "Lima Metropolitana", "RL": "Región Lima",
           "LMC": "Lima Metropolitana y Callao", "15": "Lima"}


def ambito_ubigeo(ubigeo):
    """Ámbito a partir del ubigeo del INEI (al menos cuatro dígitos): «LM», «RL» o el código del departamento."""
    u = str(ubigeo).strip().split(".")[0].zfill(6)
    if u[:2] == "15":
        return "LM" if u[:4] == "1501" else "RL"
    return u[:2]


def codigo_departamento(nombre):
    """Código de dos dígitos a partir del nombre, sin importar tildes, mayúsculas ni espacios («  Huanuco» -> «10»)."""
    import unicodedata

    def norm(x):
        return unicodedata.normalize("NFKD", str(x)).encode("ascii", "ignore").decode().strip().lower()

    inverso = {norm(v): k for k, v in DEPARTAMENTOS.items()}
    return inverso[norm(nombre)]


def por_dominio(df, dominio, peso, estrato, psu, y, grupo=None, a="Mujer", b="Hombre", escala=1.0):
    """Estimación en cada valor de `dominio` (el código de ámbito: departamento, «LM», «RL» o «LMC»).

    Arma el diseño con el subconjunto del dominio: los estratos quedan anidados en él (departamento × estrato), que es
    como el INEI estratifica la ENAHO, la ENDES, la EPEN y la ENA. Con `grupo` devuelve la brecha a − b; sin él, la
    proporción o el promedio. Agrega el coeficiente de variación de cada valor (cv_a, cv_b o cv), en %.
    """
    y = pd.Series(np.asarray(y, dtype=float), index=df.index)
    filas = []
    for clave, idx in df.groupby(dominio, sort=True).groups.items():
        sub = df.loc[idx].reset_index(drop=True)
        d = Diseno(sub, peso, estrato, psu)
        ys = y.loc[idx].to_numpy()
        if grupo is None:
            r = estimar(d, ys, escala=escala)
        else:
            g = np.asarray(df.loc[idx, grupo] if isinstance(grupo, str) else pd.Series(grupo, index=df.index).loc[idx])
            r = brecha(d, ys, g, a, b, escala=escala)
            for k in ("a", "b"):
                v, e = r[f"valor_{k}"], r[f"ee_{k}"]
                r[f"cv_{k}"] = 100 * e / abs(v) if v else np.nan
        filas.append({"ambito": str(clave).zfill(2), **r})
    return filas


def cambio(v1, ee1, v2, ee2):
    """Cambio entre dos años (muestras independientes): diferencia, ee y si es significativo al 95 %."""
    dif = v2 - v1
    ee = float(np.sqrt(ee1 ** 2 + ee2 ** 2))
    return dif, ee, bool(abs(dif) > Z95 * ee)
