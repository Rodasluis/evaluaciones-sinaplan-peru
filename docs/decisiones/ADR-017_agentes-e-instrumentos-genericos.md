# ADR-017. Agentes por etapa e instrumentos genéricos (Fase 4)

**Fecha:** 2026-11-18 · **Estado:** aprobada por el usuario

## Contexto
- **Observación del usuario (2026-11-18):** al usar el repositorio público, comprobó que no funciona como un conjunto
  de agentes que elaboren la evaluación, sino como un repositorio que se alimenta de evaluaciones ya hechas.
- **Qué faltaba:** el procedimiento del piloto (diseño, fuentes, lectura de la política, instrumentos, hallazgos,
  validación) quedó en la conversación. `src/analisis_pnig/` guarda el resultado, pero está escrito para las fuentes
  de la PNIG. Los tres agentes existentes (`codificador-ciego`, `revisor-de-informe`, `auditor-de-publicacion`)
  actúan sobre trabajo hecho.
- **Lo que no se puede automatizar:** las entrevistas, los oficios a las entidades, los documentos internos sin
  publicación, la validación del órgano rector y algunos portales.

## Decisión
1. **Preguntas iniciales** (decisión del usuario). `/nueva-evaluacion` pregunta lo que queda a cargo de la persona y
   lo guarda en `config.yaml`, sección `modalidad`:
   - recojo primario (entrevistas, formulario a las entidades o ninguno);
   - documentos internos adicionales;
   - validación con el órgano rector o solo informe final;
   - personalización de los instrumentos.

   El generador aplica las respuestas:
   - sin entrevistas, quita ese instrumento de las preguntas;
   - una pregunta sin instrumentos recibe la revisión documental;
   - escribe las cabeceras de las entradas y `avance.yaml`.

   Las mismas respuestas se pasan con banderas para usarlo sin Claude Code.
2. **Instrumentos genéricos** (decisión del usuario). `src/nucleo/instrumentos/` reúne los seis instrumentos sobre
   entradas estándar en CSV:
   - revisión documental;
   - indicadores y recursos;
   - coherencia y alineamiento, con ubicación de las citas en su documento;
   - revisión de evidencia;
   - información de las entidades, con formulario por oficio;
   - entrevistas, con guías por rol y doble codificación.

   Cada uno valida sus entradas, informa todos los problemas juntos y escribe su libro con una hoja de criterios
   por plano. Las reglas son las del núcleo (escalas, evidencia, codificación, presupuesto) y sus parámetros, los de
   `config.yaml`. `src/nucleo/busqueda_academica.py` toma del piloto la búsqueda replicable con estrategias en YAML.
3. **Personalización** (decisión del usuario). Si la evaluación la pide, el `programador-de-instrumentos` escribe en
   `src/analisis_<id>/` código propio que complementa al genérico, sin reemplazarlo, con el piloto como modelo.
4. **Agentes por etapa.** Nueve agentes nuevos y el comando `/evaluar`, que recorre nueve etapas: diseño, fuentes,
   lectura de la política, instrumentos, recojo primario, hallazgos, informe, validación y entrega. Se detiene en
   cada aprobación del usuario.
   - **Avance:** `src/nucleo/avance.py` guarda el estado de cada etapa, impide empezar una con otra anterior sin
     aprobar y omite las que no aplican a la modalidad.
   - **Preguntas al usuario:** solo las hace el comando; los subagentes no preguntan ni aprueban.
5. **Validación.** `src/nucleo/observaciones.py` arma la matriz de observaciones de un Word revisado, sin el nombre de
   quien comenta, y el JSON de respuestas. `comparar_word.ps1` responde ahora también por el orden del comentario. El
   documento con observaciones y su matriz van a la carpeta administrativa.
6. **Informe de la plantilla.** Suma un anexo con una tabla resumen y su análisis por cada libro genérico.
7. **Piloto:** `src/analisis_pnig/` no cambia y su referencia sigue igual.

## Planos
- **CEPLAN:** tipologías, estructura del informe, matriz de recomendaciones y atributos de las fichas de indicadores
  (`tipologias/comun/atributos_ficha.yaml`, con página).
- **Literatura:** solidez de la evidencia, búsquedas replicables, análisis temático y concordancia.
- **Equipo evaluador:** reglas de clasificación del piloto (ADR-011 y ADR-012), parámetros de `config.yaml` y la
  inclusión de los estudios.

## Verificación (2026-11-18)
- **Equivalencia con el piloto** (`tests/test_equivalencia_pnig.py`): con los libros versionados del piloto
  convertidos a las entradas estándar, los genéricos reproducen:
  - la clasificación por criterios y el estado asignado de los 52 servicios (24, 26 y 2);
  - el avance y el nivel de cada indicador y año con valor;
  - la solidez, la dirección y la plausibilidad de los 24 tipos de intervención;
  - la valoración de coherencia de cada servicio (24 completos de 52);
  - el alineamiento vertical por objetivo en el periodo y en la transición.

  La única adaptación fue de etiqueta: «Ascendente (acumulado anual)» se lee como ascendente.
- **Caso ficticio** (`tests/test_instrumentos.py`, `tests/datos/evaluacion_ficticia/`):
  - libros esperados;
  - errores de validación: columna faltante, ajuste sin sustento, cita no ubicada y segmento sin codificar;
  - guías, formulario y consolidación;
  - búsqueda sin red;
  - corrida completa hasta el informe con su anexo de resultados.
- **Generador y avance:** `test_nueva_evaluacion.py` y `test_avance.py`. Matriz de observaciones: `test_nucleo.py`.
  `comparar_word.ps1` se probó con Word en un documento de dos comentarios sin número.
- **Pruebas:** 58 en verde (37 antes de esta fase); la referencia del piloto sigue igual.
- **Prueba manual:** la línea de comandos, de la creación al informe, con una evaluación ficticia que luego se borró.

## Consecuencias
- **Uso:** una evaluación nueva se hace con `/nueva-evaluacion` y `/evaluar` (`docs/agentes.md`).
- **Carga del usuario:** aprobar cada etapa, aplicar las entrevistas, enviar los oficios, entregar los documentos
  internos y conseguir lo que bloquea a los robots.
- **Tipologías sin probar:** diseño, implementación y resultados usan los mismos instrumentos. El estado de
  implementación de los servicios solo corre si la tipología tiene `escalas.implementacion`.
- **Pendiente:** probar el flujo completo con una segunda política real.
