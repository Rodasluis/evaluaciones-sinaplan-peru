---
description: Arma el anexo técnico de la evaluación y lo audita antes de entregarlo
argument-hint: [id de la evaluación, por defecto pnig_2026]
---

Arma el anexo técnico de la evaluación `$ARGUMENTS` (por defecto `pnig_2026`), después del informe y de la
presentación de resultados.

1. Confirma que el informe y la presentación vigentes están en `evaluaciones/<id>/06_informe/` y que el informe ya se
   terminó con Word (índices y paginación al día).
2. Piloto PNIG: `UV_OFFLINE=1 PYTHONUTF8=1 uv run python src/analisis_pnig/anexo_tecnico.py`. Arma
   `evaluaciones/pnig_2026/07_anexo_tecnico/`, que no se versiona, con su índice y la lista de archivos con SHA-256.
   Otra evaluación: usa `nucleo.anexo` siguiendo `anexo_tecnico.py` como modelo.
3. Audita la carpeta: `uv run python src/nucleo/auditoria.py entrega evaluaciones/<id>/07_anexo_tecnico` y, antes de
   enviarla fuera del equipo, lanza el agente `auditor-de-publicacion` sobre la misma carpeta.
4. Informa el número de archivos, el tamaño total y el resultado de la auditoría. El envío del anexo es una acción
   externa: se confirma con el usuario antes de hacerlo.
