"""Corre los instrumentos genéricos de una evaluación (src/nucleo/instrumentos/).

Uso (desde la raíz del repositorio):
    uv run python src/nucleo/instrumentos/correr.py <id> [REV_DOC REV_IND …] [--validar]
    uv run python src/nucleo/instrumentos/correr.py <id> --guias        guías de entrevista por rol (ENTREV)
    uv run python src/nucleo/instrumentos/correr.py <id> --formulario   formularios a las entidades (REP_ENT)
    uv run python src/nucleo/instrumentos/correr.py <id> --consolidar   respuestas de los formularios a 02_insumos/

Sin instrumentos, corre los de config.yaml (sección instrumentos) en el orden del núcleo; las entrevistas, solo si
modalidad.entrevistas es verdadero. Con --validar, solo revisa las entradas. Un instrumento con entradas incompletas
se informa y no detiene a los demás; el código de salida es 1 si alguno falló.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))   # src/
from nucleo.config import RAIZ_REPO, cargar_desde  # noqa: E402
from nucleo.instrumentos import INSTRUMENTOS, activos, entrev, rep_ent  # noqa: E402
from nucleo.instrumentos.base import ErrorInsumo  # noqa: E402


def correr(cfg, instrumentos=None, solo_validar=False, salida=print):
    """Corre (o valida) los instrumentos; devuelve {instrumento: resumen o ErrorInsumo}."""
    resultados = {}
    for k in instrumentos or activos(cfg):
        modulo = INSTRUMENTOS[k]
        try:
            if solo_validar:
                modulo.validar(cfg)
                resultados[k] = {"validado": True}
                salida(f"{k} ({modulo.NOMBRE}): entradas sin observaciones")
            else:
                r = modulo.correr(cfg)
                resultados[k] = r
                detalle = ", ".join(f"{c} {v}" for c, v in r.items() if c != "libro")
                salida(f"{k} ({modulo.NOMBRE}): {Path(r['libro']).relative_to(cfg.repo).as_posix()}; {detalle}")
        except ErrorInsumo as e:
            resultados[k] = e
            salida(str(e))
    return resultados


def main(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    if not argv or argv[0].startswith("-"):
        print(__doc__)
        return 2
    cfg = cargar_desde(RAIZ_REPO, argv[0])
    banderas = {a for a in argv[1:] if a.startswith("--")}
    pedidos = [a for a in argv[1:] if not a.startswith("--")]
    desconocidos = [a for a in pedidos if a not in INSTRUMENTOS]
    if desconocidos:
        print(f"Instrumentos sin versión genérica: {', '.join(desconocidos)} (hay: {', '.join(INSTRUMENTOS)})")
        return 2
    try:
        if "--guias" in banderas:
            for r in entrev.guias(cfg):
                print("Guía:", r.relative_to(cfg.repo).as_posix())
            return 0
        if "--formulario" in banderas:
            for r in rep_ent.formulario(cfg):
                print("Formulario:", r.relative_to(cfg.repo).as_posix())
            return 0
        if "--consolidar" in banderas:
            n_seg, n_val = rep_ent.consolidar(cfg)
            print(f"Respuestas consolidadas: {n_seg} segmentos y {n_val} valores de indicadores en 02_insumos/")
            return 0
    except ErrorInsumo as e:
        print(e)
        return 1
    resultados = correr(cfg, pedidos or None, "--validar" in banderas)
    return 1 if any(isinstance(r, Exception) for r in resultados.values()) else 0


if __name__ == "__main__":
    sys.exit(main())
