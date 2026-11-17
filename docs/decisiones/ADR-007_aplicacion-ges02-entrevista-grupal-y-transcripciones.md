# ADR-007. Aplicación de GES-02: entrevista grupal de la DGIGND y transcripciones anonimizadas

**Fecha:** 2026-11-06 · **Estado:** aprobada por el usuario

## Contexto
Las tres entrevistas GES-02 se hicieron por Google Meet:

| Órgano | Fecha y hora | Duración | Modalidad | Transcripción |
|---|---|---|---|---|
| OGPPM | 2026-11-04, 16:01 | 34 min | Individual, con dos entrevistadores | Meet no la generó; se transcribió con faster-whisper (large-v3-turbo, local) y se contrastó con la de Pinpoint |
| OGMEPGD | 2026-11-05, 11:56 | 25 min | Individual | Meet |
| DGIGND | 2026-11-05, 15:04 | 85 min | Grupal | Meet |

Lo que la aplicación cambió frente a lo previsto:
- **Entrevista grupal:** el Anexo 4 de la nota prevé entrevistas individuales a un profesional designado por órgano. La DGIGND designó a tres especialistas de la Dirección de Políticas (DPIGND), que respondieron juntas desde una sola cuenta de Meet. La transcripción no distingue quién habla: los cambios de voz quedan dentro de un mismo bloque.
- **Conducción:** las especialistas aclararon (37:46) que la conducción operativa de la Política corresponde a la Dirección General y que la DPIGND tiene a su cargo la formulación, el diseño y la asistencia técnica.
- **Transcripciones:** las de Meet y Whisper confunden siglas y cambian el sentido en algunos pasajes. Los resúmenes de Gemini atribuyen mal lo dicho y tienen errores de contenido.
- **Datos personales:** el LEEME de `05_entrevistas/` solo admitía en el repositorio el registro codificado. Las grabaciones y las transcripciones tienen nombres.

## Decisión
- **Codificación de la DGIGND:** como entrevista del órgano, con un solo código (NC-DGIGND) y sin distinguir voces (decisión del usuario, 2026-11-06). Se descartó separar las tres voces, porque eso exigía reconstruir los turnos con el audio de 85 minutos y las preguntas de evaluación son institucionales.
- **Corrección de las transcripciones:**
  - siglas corregidas con reglas declaradas;
  - lecturas dudosas que el usuario verificó en el audio, marcadas [audio]: 6 en la OGPPM, 4 en la OGMEPGD y 17 en la DGIGND;
  - lecturas adoptadas de Pinpoint, marcadas [Pinpoint];
  - lecturas aún dudosas, marcadas [¿?].
  - El resumen de Gemini no se usa como fuente.
- **Transcripciones anonimizadas en el repositorio** (decisión del usuario, 2026-11-06), en `evaluaciones/pnig_2026/05_entrevistas/transcripciones/`:
  - E es el equipo evaluador y NC-órgano, la persona informante;
  - los nombres pasan a [informante], [evaluador] o [especialista DPIGND], y el cargo personal que se declara en las presentaciones, a [cargo];
  - se conservan los cargos institucionales de un órgano, como «la directora», porque la atribución es por órgano;
  - la cabecera de cada transcripción indica el método, las reglas y el SHA-256 de la versión de base.
- **Fuera del repositorio**, en la carpeta administrativa (`cuestionarios\Google Meet\`): las grabaciones, los docx de Gemini, las versiones corregidas con nombres y los scripts de transcripción, corrección y anonimización. El script de anonimización no puede ir a `src/` porque contiene la lista de nombres que retira. Comprueba que ninguno quede en las salidas y barre los archivos versionados en busca de los nombres de las personas entrevistadas.
- **Registro de la aplicación:** cada pregunta de la guía de cada informante se registra como literal, reformulada, en parte, no formulada o espontánea (el tema salió sin que se preguntara), con su minuto. Las preguntas que no estaban en la guía se registran como surgidas.
- **Codificación:**
  - primera codificación de Claude sobre la transcripción anonimizada, que valida el usuario;
  - un segmento por idea, con códigos del libro común, sentido, hallazgo que contrasta y resultado del contraste;
  - las citas breves se comprueban contra la transcripción anonimizada.
  - Los segmentos que pueden cambiar una valoración se marcan como candidatos a la doble codificación ciega.
- **Planos:** la modalidad individual y los bloques son de la nota. La entrevista grupal es un hecho de la aplicación, y su tratamiento como entrevista del órgano y la anonimización son decisión del equipo evaluador.

## Consecuencias
- **Implementación:** `src/analisis_pnig/ges02_guias_entrevista.py` lee `GES-02_aplicacion.csv` y `GES-02_codificacion.csv` y rehace en cada corrida el libro `GES-02_registro_codificacion.xlsx`, con las hojas:
  - «Registro»;
  - «Consentimiento», con fecha, hora, duración y modalidad reales;
  - «Aplicación»;
  - «Preguntas», con la formulación por informante;
  - «Vacíos».
  Lo que el usuario anota en «validacion» y «observacion_evaluador» se conserva entre corridas.
- **Resultado:**
  - 109 segmentos (OGPPM 30, OGMEPGD 23, DGIGND 56), de los que 13 son candidatos a la doble codificación;
  - 7 preguntas surgidas;
  - 47 filas de vacíos.
- **Sin consulta adicional por correo** (decisión del usuario, 2026-11-06). Dos orientadoras que no se formularon completas se responden con lo que dijeron las entrevistas. La deducción, con los segmentos que la sustentan, va en la columna «deduccion» de `GES-02_aplicacion.csv`:
  - **Tabla A7.1** (sección 3.7): ante las recomendaciones de la evaluación de 2023 y del informe técnico del CEPLAN, la medida fue coordinar la actualización de la PNIG, que no llegó a concretarse. No se mencionan otras medidas (NC-OGMEPGD-016, NC-OGPPM-017, NC-DGIGND-050).
  - **Limitaciones presupuestales:** la conducción no contó con los recursos necesarios (NC-DGIGND-023, -026, -034 y -037). La OGPPM no se refirió a ellas.
- **Validación:** el usuario validó los 13 segmentos candidatos (2026-11-06).
- **Vacíos que quedan como hallazgo:** cuatro hallazgos sujetos a contraste no tienen ningún segmento: D1-H2, I3-H2, I4-H2 e I5-H1.
- **Atribución en el informe:** por nivel de informante y no por persona (nota, numeral 10). En la DGIGND, lo dicho se atribuye al órgano.
- **Nota metodológica:** no se modifica. La desviación de la modalidad se declara en el informe de evaluación de cierre y en esta nota de decisión.
- **Registro de movimientos:** `docs/fuentes/reorganizacion_2026-11-06.tsv`.
- **Actualización:** el instrumento quedó después con las preguntas formuladas, las deducciones pasaron a `GES-02_deducciones.csv` y el informe incorporó las entrevistas (ADR-008).
- **Doble codificación (2026-11-06):** el usuario revisó y validó los 13 segmentos candidatos y los 96 restantes. Para GES-02, esa revisión reemplaza la doble codificación ciega que el numeral 9 de la nota prevé para los segmentos críticos (decisión del usuario). El informe lo declara en la descripción de la aplicación de GES-02.
