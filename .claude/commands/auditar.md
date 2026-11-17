---
description: Corre las auditorías del repositorio (informe, archivos preparados para el commit y, si se indica, una carpeta de entrega)
argument-hint: [informe | commit | entrega <carpeta>]  (sin argumento: informe y commit)
---

Corre las auditorías de `src/nucleo/auditoria.py` desde la raíz del repositorio con `UV_OFFLINE=1 PYTHONUTF8=1`.
Argumentos: $ARGUMENTS

- **informe** (o sin argumento): el informe de la evaluación activa (`productos.informe` de
  `evaluaciones/<id>/config.yaml`, por defecto `pnig_2026`):
  `uv run python src/nucleo/auditoria.py informe <ruta>`.
- **commit** (o sin argumento): los archivos preparados para el commit:
  `uv run python src/nucleo/auditoria.py commit`. Es la misma revisión del gancho `.githooks/pre-commit`.
- **entrega `<carpeta>`**: `uv run python src/nucleo/auditoria.py entrega <carpeta>`, y después el agente
  `auditor-de-publicacion` sobre la misma carpeta.

Resume lo encontrado por severidad. No borres, muevas ni edites archivos para resolverlo: propón la acción y espera la
aprobación del usuario.
