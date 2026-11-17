"""Presupuesto: conciliación de montos entre dos fuentes (por ejemplo, la información remitida por un órgano frente a
los datos abiertos del gasto del MEF), por órgano y año."""
import pandas as pd


def conciliar(filas, etiqueta_a, etiqueta_b, tolerancia=1.0, omitir_ceros=True):
    """filas: iterable de (claves: dict, monto_a, monto_b). Devuelve un DataFrame con las claves, ambos montos
    redondeados a céntimos, la diferencia y «Coincide» (Sí si la diferencia absoluta es menor que la tolerancia, en
    soles). Con omitir_ceros, no se listan las combinaciones sin monto en ninguna de las dos fuentes."""
    salida = []
    for claves, a, b in filas:
        a = 0.0 if a is None or pd.isna(a) else float(a)
        b = 0.0 if b is None or pd.isna(b) else float(b)
        if omitir_ceros and a == 0 and b == 0:
            continue
        salida.append({**claves, etiqueta_a: round(a, 2), etiqueta_b: round(b, 2), "Diferencia (S/)": round(a - b, 2),
                       "Coincide": "Sí" if abs(a - b) < tolerancia else "No"})
    return pd.DataFrame(salida)
