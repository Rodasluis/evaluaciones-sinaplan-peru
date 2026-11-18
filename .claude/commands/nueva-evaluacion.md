---
description: Crea una evaluación nueva desde la plantilla y una tipología (cierre, diseño, implementación o resultados), con las preguntas iniciales sobre lo que queda a cargo del equipo
argument-hint: <id> <tipologia> "<política>" <sigla> <inicio> <fin>
---

Crea una evaluación nueva con `src/nucleo/nueva_evaluacion.py`. Argumentos recibidos: $ARGUMENTS

1. **Datos de la evaluación.** Si falta alguno, pídelo antes de seguir: identificador (minúsculas y guion bajo, por
   ejemplo `pnx_2027`), tipología (`cierre`, `diseno`, `implementacion` o `resultados`), nombre de la política o plan,
   sigla, año de inicio y de fin del periodo evaluado, y título de la evaluación. Lee `tipologias/README.md` y dile al
   usuario si la tipología elegida está probada o es una estructura sin probar.
2. **Preguntas iniciales.** Hazlas con AskUserQuestion, en una sola llamada, con estas opciones:
   - «¿Qué recojo de información primaria hará?» (varias): Entrevistas a los órganos que conducen la política;
     Formulario a las entidades por oficio; Ninguno, solo documentos.
   - «¿Cuenta con documentos internos adicionales, sin publicación oficial?»: Sí (se pedirán en la etapa de
     fuentes); No.
   - «¿Se validará el informe con el órgano rector o solo se entregará el informe final?»: Validación con el órgano
     rector; Solo informe final.
   - «¿Desea personalizar los instrumentos?»: No, solo los genéricos (recomendado); Sí, con código propio.
   Si responde que sí a la personalización, pregunta en una segunda llamada cuáles (varias), entre los instrumentos
   que quedan activos según la primera respuesta.
3. **Creación.** Ejecuta desde la raíz del repositorio:
   `uv run python src/nucleo/nueva_evaluacion.py <id> --tipologia <tipo> --titulo "<título>" --politica "<política>" --sigla <SIGLA> --inicio <AAAA> --fin <AAAA>`
   con las banderas de las respuestas: `--sin-entrevistas` si no hay entrevistas, `--formulario-entidades`,
   `--documentos-internos`, `--validacion informe_final` si solo se entrega el informe final y
   `--personalizar REV_IND,REV_EVI` con los instrumentos elegidos.
4. **Resumen.** Muestra lo que se creó (`evaluaciones/<id>/`, `src/analisis_<id>/`, las cabeceras de las entradas y
   `avance.yaml`), la sección `modalidad` y las preguntas tipo de `config.yaml`, y los avisos del generador.
5. **Siguiente paso.** Indica que `/evaluar <id>` recorre las etapas con los agentes (`docs/agentes.md`) y que lo
   primero es el diseño, que el usuario aprueba (CLAUDE.md, regla 5).
6. No hagas commit salvo que el usuario lo pida.
