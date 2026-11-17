---
description: Crea una evaluación nueva desde la plantilla y una tipología (cierre, diseño, implementación o resultados)
argument-hint: <id> <tipologia> "<política>" <sigla> <inicio> <fin>
---

Crea una evaluación nueva con `src/nucleo/nueva_evaluacion.py`. Argumentos recibidos: $ARGUMENTS

1. Si falta alguno de estos datos, pídelo antes de seguir: identificador (minúsculas y guion bajo, por ejemplo
   `pnx_2027`), tipología (`cierre`, `diseno`, `implementacion` o `resultados`), nombre de la política o plan, sigla,
   año de inicio y de fin del periodo evaluado, y título de la evaluación. Lee `tipologias/README.md` y dile al
   usuario si la tipología elegida está probada o es una estructura sin probar.
2. Ejecuta desde la raíz del repositorio:
   `uv run python src/nucleo/nueva_evaluacion.py <id> --tipologia <tipo> --titulo "<título>" --politica "<política>" --sigla <SIGLA> --inicio <AAAA> --fin <AAAA>`
3. Muestra lo que se creó (`evaluaciones/<id>/` y `src/analisis_<id>/`) y las preguntas tipo de `config.yaml`.
4. Recuerda los pasos siguientes de `plantillas/nueva_evaluacion/LEEME.md`. Las preguntas, el alcance y el diseño
   metodológico los aprueba el usuario antes de avanzar (CLAUDE.md, regla 5).
5. No hagas commit salvo que el usuario lo pida.
