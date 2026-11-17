# ADR-008. GES-02 queda con lo que se preguntó y el informe incorpora las entrevistas

**Fecha:** 2026-11-06 · **Estado:** aprobada por el usuario

## Contexto
Después de codificar las entrevistas (ADR-007), el instrumento GES-02 seguía con su diseño: 117 preguntas en las guías de los tres órganos. Muchas de ellas no se formularon. El usuario pidió:
1. eliminar las preguntas que no se formularon y ajustar el cuestionario a lo que realmente se preguntó;
2. no hacer consultas adicionales por correo, sino deducir de las entrevistas lo que no se preguntó completo;
3. actualizar el informe de evaluación con los hallazgos de las entrevistas.

## Decisión
- **Instrumento según lo aplicado:**
  - el catálogo (`GES-02_preguntas.csv`), las guías y los cuestionarios quedan solo con las preguntas formuladas (literales, reformuladas o en parte) y las surgidas;
  - los cuestionarios aplicados muestran las preguntas en el orden de la entrevista, con su redacción real y su minuto;
  - las guías aplicadas muestran la pregunta de la guía, cómo se formuló y su evidencia de contraste.
- **Versión diseñada:** las guías y los cuestionarios enviados se conservan en la carpeta administrativa (`cuestionarios\GES-02_version_disenada_2026-11-03\`). `preguntas(C)` sigue construyendo el diseño en el código, y `aplicada()` deriva de él lo aplicado.
- **Segmentos reasignados:** los 7 segmentos que remitían a preguntas no formuladas pasan a la pregunta que se estaba respondiendo cuando se dijeron.
- **Deducciones** (`GES-02_deducciones.csv`), con los segmentos que las sustentan:
  - DED-01: ante las recomendaciones de la evaluación de 2023 y del CEPLAN, la medida fue coordinar la actualización de la PNIG, que no llegó a concretarse;
  - DED-02: la conducción no contó con los recursos necesarios.
- **Recomendaciones de evaluaciones anteriores:** las 10 que remitían a preguntas no formuladas (G2.15 a G2.18) quedan sin pregunta GES-02, con una nota, y conservan su estado.
- **Informe, borrador v1:**
  - 20 hallazgos llevan evidencia del nivel conductor, una por órgano, con plano «Nivel conductor»;
  - nuevo hallazgo I4-H5, sobre la rectoría y la conducción;
  - ningún hallazgo ni valoración queda «Sujeto a contraste GES-02» y se eliminaron las marcas [PENDIENTE: contraste GES-02];
  - D7 pasa de «Sin valorar» a «Parcial», como propuesta para validar;
  - nueva recomendación sobre priorización y ajustes en tres recomendaciones;
  - nuevo Anexo 7 con la aplicación de GES-02.
- **Planos:** las preguntas eliminadas eran del diseño del equipo evaluador y del Anexo 4. Que no se formularan es un hecho de la aplicación y se declara como limitación. Las deducciones y la redacción de los hallazgos son del equipo evaluador.

## Consecuencias
- **Hallazgos sin contraste:** siete hallazgos cuyo contraste preveía la guía no se contrastaron con el nivel conductor: D1-H2, D4-H3, I2-H1, I3-H2, I4-H2, I5-H3 e I5-H4. Quedan como limitación en el informe y en la hoja «Vacíos» del registro.
- **Implementación:**
  - `src/analisis_pnig/ges02_guias_entrevista.py`: funciones `aplicacion()`, `aplicada()`, `guia()`, `cuestionario()` y `registro()`, con las hojas «Deducciones» y «Vacíos»;
  - `src/analisis_pnig/sintesis_hallazgos.py`: función `ges()` y lectura del registro;
  - `src/analisis_pnig/sintesis_triangulacion.py`: la columna del nivel conductor con la evidencia GES-02;
  - `src/analisis_pnig/informe_cierre.py`.
- **Verificaciones:**
  - cada segmento citado por un hallazgo o una deducción existe;
  - la confianza es compatible con las fuentes independientes, y cada entrevista cuenta como una fuente;
  - las citas breves constan en las transcripciones anonimizadas;
  - no quedan marcas de GES-02 en el informe.
- **Pendiente:** el informe de la OMEP sobre los indicadores del PEMIG, solicitado y fuera del corpus. El 60 % de indicadores del PEMIG solo de cobertura se verificó el 2026-11-07 con el anexo 5 (45 de 85 de cobertura, 57 con los de participación de las mujeres; 15 de calidad): la cifra de la DGIGND queda en ese rango.
- **Registro:** `docs/fuentes/reorganizacion_2026-11-06.tsv`, con las copias de la versión diseñada y del borrador v0 del informe.
- **Actualización (2026-11-06):** el análisis de pertinencia que refirió la OGPPM no entra en el corpus documental (decisión del usuario). El informe ya no lo menciona: el hallazgo I5-H2 deja de citar el segmento NC-OGPPM-024, y la nota de ese segmento lo explica.
- **Actualización (2026-11-07):** el usuario validó todos los hallazgos y valoraciones (2026-11-06). D3-H2 e I4-H4 cambiaron después: D3-H2 corrige «sin indicador en el PEMIG» por «sin indicador vigente» (el PEMIG asignó 85 indicadores a los 52 servicios) e I4-H4 incorpora las 23 recomendaciones de la evaluación de implementación 2020.
