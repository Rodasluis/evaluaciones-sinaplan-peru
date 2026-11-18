# Agentes y etapas de una evaluación (Fase 4, ADR-017)

El repositorio elabora una evaluación con agentes de Claude Code, etapa por etapa. Los agentes leen las fuentes,
llenan las entradas estándar y corren los instrumentos genéricos del núcleo. El usuario aprueba cada etapa antes de
pasar a la siguiente (CLAUDE.md, regla 5). Lo que solo puede hacer una persona se pregunta al crear la evaluación y
queda en la sección `modalidad` de `config.yaml`.

## Cómo se usa
1. `/nueva-evaluacion <id> <tipología> "<política>" <sigla> <inicio> <fin>`: crea la evaluación y hace las preguntas
   iniciales:
   - recojo primario: entrevistas, formulario a las entidades o ninguno;
   - si hay documentos internos adicionales;
   - validación con el órgano rector o solo informe final;
   - si se personalizan los instrumentos.
2. `/evaluar <id>`: corre la etapa pendiente, lanza sus agentes y se detiene a pedir la aprobación.
3. `uv run python src/nucleo/avance.py <id>`: estado de las etapas, la siguiente y lo que le falta.

## Etapas
| N.° | Etapa | Agentes | Qué produce | Qué aprueba el usuario | Se omite si |
|---|---|---|---|---|---|
| 1 | Diseño | `disenador-de-evaluacion` | Preguntas adaptadas, matriz pregunta–instrumento–fuente, alcance y limitaciones | Preguntas, alcance y matriz | Nunca |
| 2 | Fuentes | `recolector-de-fuentes` | Documentos de portales oficiales e internos, con su hash y su registro | Lista de fuentes y lo que falta | Nunca |
| 3 | Lectura de la política | `lector-de-la-politica` | Catálogo de servicios, lineamientos, indicadores y cadena | Catálogo | Nunca |
| 4 | Instrumentos | `analista-documental`, `analista-de-indicadores`, `analista-de-alineamiento`, `analista-de-evidencia`, `codificador-ciego`; `programador-de-instrumentos` si se personaliza | Entradas llenas y libros de análisis | Reglas, umbrales y libros | Nunca |
| 5 | Recojo primario | `analista-de-entrevistas`, `codificador-ciego` | Guías por rol o formularios; después, el registro codificado | Guías y formulario antes de aplicarlos | No hay entrevistas ni formulario |
| 6 | Hallazgos | `redactor-de-hallazgos` | Hallazgos, valoraciones y recomendaciones verificados por `sintesis.py` | Hallazgos y valoraciones | Nunca |
| 7 | Informe | `revisor-de-informe` | Informe en Word y PDF de revisión | Versión del informe | Nunca |
| 8 | Validación | Ninguno | Matriz de observaciones, respuestas y copia con control de cambios, fuera del repositorio | Respuestas y versión final | Solo se entrega el informe final |
| 9 | Entrega | `auditor-de-publicacion` | Anexo técnico auditado | Envío (acción externa) | Nunca |

## Lo que queda a cargo de la persona
- **Entrevistas:** las aplica el usuario y deja las transcripciones anonimizadas. Las grabaciones y la relación entre
  el rol y la persona quedan fuera del repositorio.
- **Formulario a las entidades:** el usuario envía los oficios y deja las respuestas en `03_datos/raw/formularios/`.
- **Documentos internos:** el usuario indica la carpeta; el recolector los copia con su hash.
- **Validación:** el órgano rector comenta el Word; el usuario aprueba las respuestas.
- **Portales que bloquean robots y microdatos con registro:** el recolector los lista y el usuario los descarga.
- **Toda acción externa** (enviar, publicar, borrar) se confirma antes.

## Instrumentos genéricos y personalización
- **Genéricos:** los seis instrumentos del núcleo (`src/nucleo/instrumentos/`) leen entradas estándar
  (`plantillas/nueva_evaluacion/evaluacion/02_insumos/LEEME.md`) y escriben sus libros con una hoja de criterios.
- **Equivalencia con el piloto:** reproducen las clasificaciones de la PNIG (`tests/test_equivalencia_pnig.py`):
  - estado de los 52 servicios;
  - avance de cada indicador y año;
  - solidez de los 24 tipos de intervención;
  - coherencia por servicio y alineamiento vertical por objetivo.
- **Personalización:** si el usuario la pide, el `programador-de-instrumentos` escribe en `src/analisis_<id>/` un
  script que lee el libro genérico y le agrega lo particular de la política, con los scripts del piloto como modelo,
  sin reemplazar el genérico.

## Reglas comunes de los agentes
- **Fuentes:** sin fuentes, cifras ni citas inventadas; las citas, literales y con página; lo no verificado,
  `[POR VERIFICAR]`.
- **Datos sensibles:** fuera del repositorio; las personas, por su cargo.
- **Estilo:** el de los documentos (CLAUDE.md), sin guiones explicativos.
- **Dónde escriben:** solo en `evaluaciones/<id>/` (el recolector, además, en `docs/fuentes/`).
- **Decisiones:** no cambian `config.yaml` ni `avance.yaml`; las aprobaciones las registra el comando cuando el
  usuario las da en el chat.
- **Preguntas al usuario:** los subagentes no preguntan; lo hace `/evaluar`.
