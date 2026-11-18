# 05_entrevistas

Entrevistas (ENTREV), solo si la evaluación las hace (`modalidad.entrevistas` en `config.yaml`).

| Archivo | Columnas | Valores |
|---|---|---|
| `entrev_preguntas.csv` | id, rol, bloque, pregunta_evaluacion, texto, repreguntas, contrasta | rol: cargo u órgano del informante; pregunta_evaluacion: clave de config.yaml con las entrevistas entre sus instrumentos |
| `entrev_segmentos.csv` | id_segmento, rol, pregunta_id, texto | Fragmentos literales de las transcripciones anonimizadas |
| `entrev_codificacion.csv` | id_segmento, codigos, sentido, hallazgo, contraste | Primera codificación con el libro común (`02_insumos/libro_codigos.csv`); contraste: Confirma, Matiza o Contradice |
| `entrev_segunda_codificacion.csv` | id_segmento, codigos | La escribe el agente codificador-ciego, que recibe solo `entrev_segmentos.csv` y el libro |

- **Guías:** `uv run python src/nucleo/instrumentos/correr.py <id> --guias` arma una guía en Word por rol, con el
  consentimiento.
- **Personas:** se identifican por su rol, nunca por su nombre.
- **Fuera del repositorio:** las grabaciones, las transcripciones con nombres, las notas en bruto y los datos de
  contacto (CLAUDE.md, regla 6; ADR-007).
- **Transcripciones:** las anonimizadas van en `transcripciones/`.
- **Concordancia:** la doble codificación se mide con `nucleo.codificacion` (acuerdo, Jaccard y kappa). El equipo
  resuelve todas las discrepancias.
