# ADR-006. El bloque 4 de GES-02 se adapta a la función de cada órgano

**Fecha:** 2026-11-03 · **Estado:** aprobada por el usuario

## Contexto
El Anexo 4 de la nota define los bloques y las preguntas orientadoras de GES-02. Las orientadoras son comunes a los tres informantes. La diferenciación por órgano estaba solo en las repreguntas de la guía, que el informante no ve.

Así, los cuestionarios que se remiten antes de la entrevista eran casi idénticos. El usuario revisó los tres cuestionarios y pidió preguntas propias de la función de cada órgano:
- **DGIGND:** el PEMIG (D.S. N.° 002-2020-MIMP), instrumento de implementación de la Política;
- **OGPPM:** el planeamiento y la coordinación con el CEPLAN, ante el que es punto focal (Oficios N.° D000054-2026-MIMP-OGPPM y D000289-2026-MIMP-OGPPM);
- **OGMEPGD:** las limitaciones del seguimiento y de la evaluación.

En su versión, el usuario quitó además las dos orientadoras genéricas del bloque 4 (capacidades y riesgos) y la pregunta «¿Qué recomendación priorizaría?».

## Decisión
- **Núcleo común:** se mantiene para los tres. Comprende la apertura, la rectoría y la articulación, la pregunta de la Tabla A7.1 del CEPLAN y las lecciones. Así se pueden contrastar las respuestas de los tres informantes a la misma pregunta.
- **CONAIG:** su orientadora y sus tres repreguntas se formulan solo a la DGIGND (ajuste pedido por el usuario el mismo día).
  - En su lugar, la OGPPM y la OGMEPGD reciben esta orientadora: «¿Qué papel cumplió el PEMIG en la articulación de las entidades responsables de los servicios y qué limitaciones tuvo como instrumento de implementación de la Política?».
  - Vínculos: I4; códigos I4.COORD e I4.RECT; contrasta I4-H3 y D3-H2.
- **Bloque 4:** pasa de «Capacidades y sostenibilidad» a un bloque propio de cada órgano:
  - DGIGND, «Implementación y PEMIG», con 2 orientadoras;
  - OGPPM, «Planeamiento y coordinación con el CEPLAN», con 3;
  - OGMEPGD, «Seguimiento y evaluación», con 3.
- **Sostenibilidad (I5):** la última orientadora de cada bloque propio pregunta por las condiciones que esa función necesita en la nueva política. Así I5 conserva su fuente GES-02.
- **Preguntas genéricas:** las dos del bloque 4 y «¿Qué recomendación priorizaría?» salen del cuestionario. En la guía quedan como repreguntas B.
- **Pregunta de la Tabla A7.1:** nombra al órgano en cada cuestionario.
- **Vínculos:** cada pregunta nueva declara su pregunta de evaluación (I3, I4 o I5), sus códigos del libro común y el hallazgo que contrasta. Las de la OGPPM sobre el CEPLAN y la de la OGMEPGD sobre el seguimiento aportan además contraste a D5.
- **Matriz y tiempos:** no cambia la matriz de preguntas de evaluación. La duración sigue entre 45 y 60 minutos: DGIGND 59, OGPPM 55 y OGMEPGD 51.
- **Planos:** el bloque 4 adaptado y sus repreguntas son decisión del equipo evaluador. Los bloques y las demás orientadoras siguen siendo los del Anexo 4 de la nota.

## Consecuencias
- **Implementación:** `src/analisis_pnig/ges02_guias_entrevista.py` (tipo «Orientadora por función»).
- **Cuestionarios:** DGIGND pasa a 14 preguntas, OGPPM a 12 y OGMEPGD a 10.
- **Repreguntas nuevas con evidencia y página:**
  - el cronograma de protocolos del PEMIG;
  - los indicadores de cobertura del PEMIG que recomendó el CEPLAN;
  - la ruta de actualización de indicadores comunicada al CEPLAN;
  - las conclusiones del Informe N.° 000006-2026 del CEPLAN.
- **Ids:** cambian los del bloque 4 en `GES-02_preguntas.csv`.
  - La pregunta sobre el PEMIG entra como G2.3, así que los ids del bloque 2 desde G2.3 suben en uno.
  - `valoracion_recomendaciones.csv` se actualizó: G2.12 a G2.17 pasan a G2.13 a G2.18.
  - El registro de codificación se rehízo vacío y la matriz de triangulación se regeneró.
- **Nota metodológica:** no se modifica. El ajuste se documenta en esta nota de decisión y en la guía de cada informante.
