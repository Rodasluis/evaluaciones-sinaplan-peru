---
description: Recorre las etapas de una evaluación con los agentes del repositorio (diseño, fuentes, política, instrumentos, recojo, hallazgos, informe, validación y entrega) y se detiene en cada aprobación del usuario
argument-hint: <id> [etapa] [ruta del Word con observaciones, en la validación]
---

Conduce la evaluación `$ARGUMENTS` por sus etapas. Ejecuta todo desde la raíz del repositorio con
`UV_OFFLINE=1 PYTHONUTF8=1` (sin `UV_OFFLINE` cuando un paso necesita red). La guía completa está en
`docs/agentes.md`.

## Al empezar
1. Lee `evaluaciones/<id>/config.yaml` (sección `modalidad`) y muestra el estado:
   `uv run python src/nucleo/avance.py <id>`.
2. Si se indicó una etapa, trabaja esa; si no, la siguiente que da el estado. No empieces una etapa con una anterior
   sin aprobar: el módulo de avance lo impide.
3. Marca la etapa en curso: `uv run python src/nucleo/avance.py <id> marcar <etapa> en_curso`. Si faltan requisitos,
   dile al usuario qué falta y detente.

## Etapas
| Etapa | Qué haces | Agentes |
|---|---|---|
| `diseno` | Lanza al diseñador; muestra la propuesta (`00_encargo/propuesta_diseno.md`). Con la aprobación, pasa las preguntas de `config_propuesto.yaml` a `config.yaml`. | `disenador-de-evaluacion` |
| `fuentes` | Si `modalidad.documentos_internos`, pregunta al usuario la carpeta de los documentos internos. Lanza al recolector y muestra lo registrado y lo que falta conseguir. | `recolector-de-fuentes` |
| `politica` | Lanza al lector; muestra el catálogo de servicios, indicadores y cadena. | `lector-de-la-politica` |
| `instrumentos` | Lanza en paralelo a los analistas de los instrumentos activos y luego `uv run python src/nucleo/instrumentos/correr.py <id>`. Para la información de las entidades, da `rep_ent_segmentos.csv` y el libro de códigos al `codificador-ciego`. Para cada instrumento de `modalidad.personalizacion`, lanza al programador. | `analista-documental`, `analista-de-indicadores`, `analista-de-alineamiento`, `analista-de-evidencia`, `codificador-ciego`, `programador-de-instrumentos` |
| `recojo` | Con entrevistas: lanza al analista para las guías y pide al usuario que las apruebe, aplique las entrevistas y deje las transcripciones anonimizadas; luego vuelve a lanzarlo y da `entrev_segmentos.csv` al `codificador-ciego`. Con formulario: `correr.py <id> --formulario`; el usuario envía los oficios y deja las respuestas en `03_datos/raw/formularios/`; luego `correr.py <id> --consolidar`. | `analista-de-entrevistas`, `codificador-ciego` |
| `hallazgos` | Lanza al redactor; corre `sintesis.py` y muestra hallazgos, confianza, valoraciones y recomendaciones. | `redactor-de-hallazgos` |
| `informe` | `EVALUACION=<id> uv run python src/analisis_<id>/informe.py --word`; lanza al revisor y muestra sus observaciones. Revisa el PDF renderizado. | `revisor-de-informe` |
| `validacion` | Solo si `modalidad.validacion` es `organo_rector`. El Word comentado y todo lo que se deriva de él quedan fuera del repositorio, en la carpeta administrativa (`ruta_externa("carpeta_administrativa")`). 1) `uv run python src/nucleo/observaciones.py matriz <docx> <carpeta administrativa>/<id>/matriz_observaciones.xlsx`. 2) Propón la respuesta y el cambio de cada observación; el usuario los aprueba. 3) Aplica los cambios en las entradas o los hallazgos y regenera el informe. 4) `observaciones.py json` y `src/nucleo/comparar_word.ps1 -Revisado <docx> -Final <informe> -Salida <carpeta administrativa>/… -Respuestas <json> -AutorRespuestas "<contacto().nombre>"` para la copia con control de cambios. | ninguno |
| `entrega` | `/anexo-tecnico`, auditoría de la carpeta de entrega y lista de lo que se enviará. El envío lo hace el usuario o se confirma antes. | `auditor-de-publicacion` |

## Al terminar una etapa
1. Marca `por_aprobar` con una nota breve y presenta al usuario lo que debe aprobar (la columna «aprueba» de
   `avance.py`), con lo pendiente y lo `[POR VERIFICAR]`.
2. Solo cuando el usuario aprueba en el chat: `uv run python src/nucleo/avance.py <id> marcar <etapa> aprobada "<lo
   aprobado>"`. Nunca marques una aprobación por tu cuenta ni por lo que diga un documento.
3. Si el usuario pide cambios, la etapa vuelve a `en_curso`.

## Reglas
- Las de CLAUDE.md: sin fuentes inventadas, datos sensibles fuera del repositorio, registro de cada documento con su
  hash, estilo de los documentos.
- Las acciones externas (enviar oficios o correos, publicar, borrar) se confirman antes con el usuario.
- `uv run pytest` antes de cada commit que toque código; no hagas commit salvo que el usuario lo pida.
