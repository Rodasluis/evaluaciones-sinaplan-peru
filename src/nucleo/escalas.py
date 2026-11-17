"""Escalas de clasificación de las evaluaciones: estado de implementación de los servicios y nivel de avance de los
indicadores frente a su meta. Los umbrales y parámetros vienen de la configuración de cada evaluación (config.yaml,
sección escalas) y de las decisiones registradas (ADR-011 y ADR-012 en el piloto)."""

IMPLEMENTADO, PARCIAL, NO_IMPLEMENTADO, SIN_INFO = ("Implementado", "Implementado parcialmente", "No implementado",
                                                    "Sin información suficiente")
# semáforo de avance de los reportes de cumplimiento de la PNIG (RC 2025, Anexo 01, p. 16): umbral inferior y nivel
AVANCE_PNIG = [(150, "Alto con posible falla de planeación"), (95, "Alto"), (75, "Medio"), (15, "Bajo"),
               (float("-inf"), "Bajo con posible falla de planeación")]


def clasificar_implementacion(faltan, ejecucion_reciente, ejecucion_periodo, sin_ejecucion, minimo_componentes=2,
                              n_componentes=3):
    """Estado de un servicio.

    faltan: componentes que le faltan (protocolo, asignación identificable, reporte de indicador);
    ejecucion_reciente: ejecución reportada en los años de cierre (2024 o 2025 en el piloto);
    ejecucion_periodo: ejecución reportada en algún año del periodo;
    sin_ejecucion: todos los años del periodo sin ejecución (no solo sin información).
    Implementado: al menos minimo_componentes de los n_componentes y ejecución reciente (ADR-012)."""
    if n_componentes - len(faltan) >= minimo_componentes and ejecucion_reciente:
        return IMPLEMENTADO
    if ejecucion_periodo:
        return PARCIAL
    if sin_ejecucion:
        return NO_IMPLEMENTADO
    return SIN_INFO


def nivel(v, escala=AVANCE_PNIG):
    """Nivel de una razón de avance (en %) según la escala [(umbral inferior, nivel), …] de mayor a menor."""
    return next(etq for umbral, etq in escala if v >= umbral)


def avance(sentido, meta, valor):
    """Razón de avance sin acotar (valor/meta o meta/valor según el sentido), en %, o None con el motivo."""
    if sentido not in ("Ascendente", "Descendente"):
        return None, "Sin sentido definido"
    if meta is None:
        return None, "Sin meta"
    if valor is None:
        return None, "Sin valor"
    num, den = (valor, meta) if sentido == "Ascendente" else (meta, valor)
    if den == 0:
        return None, "No interpretable (meta o valor igual a 0)"
    if meta <= 0 or valor < 0 or (sentido == "Descendente" and valor <= 0):
        return None, "No interpretable (meta o valor ≤ 0)"
    return num / den * 100, None
