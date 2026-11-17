# ADR-013. Núcleo reutilizable y capa de evaluación

**Fecha:** 2026-11-12 · **Estado:** aprobada por el usuario (plan de la Fase 2); integración a `main` pendiente de su
aprobación

## Contexto
- **El piloto de la PNIG** dejó probadas las piezas de una evaluación del Sinaplan, pero escritas para la PNIG:
  - 51 de los 71 scripts de `src/analisis_pnig/` fijaban `evaluaciones/pnig_2026`;
  - los datos del caso iban dentro del código (ajustes de estado, metas verificadas, sesiones de la CONAIG, hitos);
  - la plantilla de Word era la propia nota de la PNIG;
  - `tipologias/`, `plantillas/`, `tests/` y `.claude/agents|commands` estaban vacíos.
- **Pedido del usuario (2026-11-12):** integrar el piloto en `main`, pasar a la Fase 2 (ADR-004) y después preparar la
  versión pública.
- **Decisiones del usuario:** núcleo genérico más capa de evaluación; trabajo en la rama `fase-2-nucleo`; las cuatro
  tipologías con su estructura; una skill propia de presentaciones basada en `ppt-mimp` y `ppt-mimp-lab`.

## Decisión
- **Tres capas:**
  - núcleo `src/nucleo/`, sin datos de ninguna evaluación (`docs/nucleo.md`);
  - capa de cada evaluación, `src/analisis_<id>/`, con sus análisis y su relato;
  - datos en `evaluaciones/<id>/`: `config.yaml` (preguntas, escalas, productos) y `02_insumos/` (parámetros del caso).
- **Tipologías** en `tipologias/<tipo>/estructura.yaml`, según la Guía de seguimiento y evaluación del CEPLAN (2024),
  con su página: cierre, probada en el piloto; diseño, implementación y resultados, estructura sin probar.
- **Plantilla de evaluación nueva** (`plantillas/nueva_evaluacion/`) y generador `src/nucleo/nueva_evaluacion.py`, con
  una capa mínima que arma el informe desde la estructura de la tipología: preguntas, hallazgos con su pie de fuente,
  conclusiones, matriz de recomendaciones de la Tabla 4 (p. 56) con los aspectos de la Tabla 5 (p. 57) y anexo.
- **Claude Code:** agentes `codificador-ciego`, `revisor-de-informe` y `auditor-de-publicacion`; comandos
  `/nueva-evaluacion`, `/regenerar-informe`, `/anexo-tecnico` y `/auditar`; skill `presentacion-evaluacion`.
- **Control previo al commit** (`.githooks/pre-commit` con `nucleo.auditoria`): bloquea las categorías de la regla 6 de
  CLAUDE.md y las rutas con el nombre de usuario; avisa de los archivos de más de 5 MB.
- **Entorno:** proyecto uv con versiones fijas; rutas de herramientas y de datos externos en `config/entorno.yaml` y
  `config/entorno.local.yaml` (no versionado), nunca en el código.
- **Rutas que cambian:** la portada pasa de `src/portada` a `src/nucleo/portada`, y `actualizar_campos_word.ps1` de
  `src/analisis_pnig/` a `src/nucleo/`. Las menciones anteriores en ADR-010 quedan como registro histórico.

## Regla de la fase
Los productos del piloto deben salir iguales antes y después de cada etapa. Una referencia en `tests/referencia_pnig/`
guarda las huellas de los libros de `04_analisis` (por hoja), de los CSV, del HTML de las figuras y el texto completo
del informe; `uv run pytest` la compara en cada etapa.

## Consecuencias
- **Piloto:** libros, CSV, figuras e informe idénticos a la referencia; 34 pruebas en verde al cierre de la etapa 5.
- **Presentación de resultados generada:** cambia solo en dos correcciones de la skill nueva:
  - quita los `<a:effectLst>` vacíos duplicados (31 formas), que hacían que PowerPoint pidiera reparar el archivo;
  - escribe «1 sólido» y «1 indicio» en singular.

  La presentación del repositorio, editada a mano, no se regeneró: conserva esos dos defectos hasta que el usuario
  decida corregirlos sobre el archivo.
- **CSV:** `.gitattributes` los guarda sin conversión de fin de línea, para que la regeneración no deje cambios falsos.
- **Prueba de humo:** con las cuatro tipologías se crearon evaluaciones de demostración, sus informes salieron sin
  observaciones de la auditoría y se borraron.
- **Matriz de requisitos del CEPLAN** con su estado real: 11 probados en el piloto, 9 también en la plantilla, 5
  parciales, 6 pendientes, 3 fuera del alcance del repositorio y 1 solo con estructura.

## Alternativas descartadas
- **Generador de informes genérico único:** habría obligado a reescribir el relato del piloto, que es propio del caso,
  y a perder la comparación exacta con lo ya validado.
- **Copiar el piloto para cada evaluación:** multiplica el código y los errores; las correcciones del núcleo no
  llegarían a las demás evaluaciones.
