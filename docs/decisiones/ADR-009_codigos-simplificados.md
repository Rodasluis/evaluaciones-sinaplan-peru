# ADR-009. Códigos de instrumentos e identificadores que se entienden solos

**Fecha:** 2026-11-07 · **Estado:** aprobada por el usuario

## Contexto
- **Instrumentos:** los códigos GRD-01, GES-02, FIE-03, MAC-04, PSI-05 y REV-06 nombran el tipo de documento (guía, ficha, matriz, protocolo) y no lo que se revisa. El número no sigue el orden de los anexos de la nota: MAC-04 es el Anexo 2 y GES-02 el Anexo 4.
- **Identificadores de traza:** L2025-…, F2025-…, CIV-024, NC-OGPPM-010, R20-01, RC-01, A001, M015, B01, G2.13 y S-OGPPM-01 no se entienden sin una clave.
- **Momento:** la nota metodológica todavía se puede editar y los memorándums enviados no llevan los códigos.
- **Decisión del usuario (2026-11-07):** siglas nuevas para los seis instrumentos, ENTREV en lugar de ITW, e identificadores autoexplicativos para todo lo demás, con las propuestas presentadas.

## Decisión
**Regla de separadores:** guion bajo dentro de la sigla del instrumento (REV_DOC); guion entre las partes de un identificador (LIO-2025-5.3.1-04).

| Tipo | Anterior | Nuevo |
|---|---|---|
| Revisión documental | GRD-01 | REV_DOC |
| Entrevistas al nivel conductor | GES-02 | ENTREV |
| Información reportada por las entidades | FIE-03 | REP_ENT |
| Coherencia y alineamiento | MAC-04 (MAC-04.1 a .5) | REV_ALI (REV_ALI.1 a .5) |
| Indicadores y recursos | PSI-05 (PSI-05.1 a .4) | REV_IND (REV_IND.1 a .4) |
| Revisión de evidencia | REV-06 | REV_EVI |
| Archivos, scripts y carpetas | GRD-01_…xlsx, grd01_…py, mac04_insumos/ | REV_DOC_…xlsx, rev_doc_…py, rev_ali_insumos/ (igual con las seis siglas) |
| Segmento de entrevista | NC-OGPPM-010 | ENTREV-OGPPM-010 |
| Informante y hablante de la transcripción | NC-OGPPM; E | OGPPM; EVAL |
| Transcripción | GES-02_transcripcion_NC-OGPPM.txt | ENTREV_transcripcion_OGPPM.txt |
| Segmento de informe LIO | L2025-5.3.1-04 | LIO-2025-5.3.1-04 |
| Segmento del formulario de seguimiento | F2025-ISS.04.03.05.01-medidas | FORM-2025-ISS.04.03.05.01-medidas |
| Segmento de informe de la CONAIG | CI, CII, CIII y CIV-024 | CONAIG-2022, CONAIG-2023, CONAIG-2024 y CONAIG-2025-024 |
| Recomendación de evaluación anterior | R20-01, R21-01, R23-01, RC-01 | EV2020-01, EV2021-01, EV2023-01, CEPLAN-01 |
| Estrategia de búsqueda | A001; M001 a M046; B01 | ACA-001; MAN-01 a MAN-46, con DEP-15 a DEP-22 para M015 a M022 (3ie DEP); WEB-01 |
| Pregunta de la entrevista | G2.13; S-OGPPM-01 | P2.13; PS-OGPPM-01 |
| Medio de verificación (seguimiento de recomendaciones) | FIE\|, GRD\|, PSI\|, GES\| | REP_ENT\|, REV_DOC\|, REV_IND\|, ENTREV\| |

- **Números:** las estrategias conservan su número (M023 → MAN-23; M015 → DEP-15), para leerlas directo contra los registros de descarga, que no cambian. La CONAIG usa el año del informe en lugar del número romano, como los LIO y el formulario.
- **Sin repetición:** «GES-02, NC-OGPPM-026» pasa a «ENTREV-OGPPM-026».
- **Nombre de REV_IND:** «Protocolo de seguimiento de indicadores y recursos» pasa a «Revisión de indicadores y recursos», para que el nombre corresponda a la sigla. Los demás nombres no cambian.
- **Lo que no cambia:** preguntas de evaluación (D1 a I5), hallazgos (D3-H2), tipos de intervención (T01 a T24), fuentes de evidencia (E01 a E173), códigos del libro de codificación (I1.PRES), deducciones (DED-01), situación futura deseada (SFD1), campos del formulario y códigos oficiales (OP, servicios, IOP, ISS).
- **Planos:** el CEPLAN no fija códigos para los instrumentos; los códigos son decisión del equipo evaluador.

## Lo que conserva los códigos anteriores
Se conserva como registro histórico:
- **ADR-001 a ADR-008**, con sus nombres de archivo, y el memo de suficiencia del 2026-10-27 de `04_analisis/`.
- **Registros de movimientos:** los `docs/fuentes/descargas_*.tsv` y `reorganizacion_*.tsv` anteriores, con sus nombres y su contenido. Las rutas viejas se encadenan con las nuevas en `reorganizacion_2026-11-07.tsv`.
- **Carpeta administrativa:**
  - cuestionarios enviados y `cuestionarios\GES-02_version_disenada_2026-11-03\`;
  - notas y borradores anteriores, incluida la nota v15 (`notas_metodologicas_anteriores\22_2026-11-07_previa-simplificacion-codigos\`);
  - la copia de la nota adjunta a los memorándums;
  - `archivo_FIE-03_no_enviado\`.
- **`src/formulario_fie03/`:** el aplicativo de recojo que no se envió (ADR-003), desplegado en Apps Script; es otro objeto que REP_ENT.

## Consecuencias
- **Implementación:**
  - **Script de migración:** `src/mantenimiento/adr009_renombrar_codigos.py`, con la función de traducción, los nombres protegidos y los modos `vista`, `renombrar` y `traducir`. Es idempotente.
  - **Renombres:** 479, registrados con su hash en `docs/fuentes/reorganizacion_2026-11-07.tsv`. De ellos, 75 son archivos versionados (con `git mv`) y 404 están fuera de git: las respuestas crudas de la búsqueda de evidencia en `03_datos/raw/rev_evi_busqueda*` y `procesado/rep_ent_segmentos.csv`.
  - **Ajustes manuales en la lógica que dependía del prefijo:**
    - generadores de los identificadores: `rep_ent_segmentos.py`, `rev_evi_busqueda_academica.py`, `rev_evi_estrategias_manuales.py`, `rev_evi_insumos.py` y `seguimiento_recomendaciones.py`;
    - informante sin «NC-», preguntas «P» y «PS-», segmentos «ENTREV-»: `entrev_guias_entrevista.py`, `sintesis_hallazgos.py`, `sintesis_triangulacion.py` e `informe_cierre.py`;
    - abreviatura de las rutas de los libros: `sintesis_hallazgos.corto()`;
    - filtros por estrategia DEP y MAN: `rev_evi_evidencia.py`, `informe_cierre.py` y `figuras_informe.py`.
  - **Anonimizador** (carpeta administrativa, `cuestionarios\Google Meet\scripts\`): escribe `ENTREV_transcripcion_<órgano>.txt`, con hablantes EVAL y órgano, y traduce las siglas de las notas de cabecera que copia de las bases con nombres, que no cambian.
  - **Nota v16:**
    - siglas nuevas;
    - nuevo nombre de REV_IND en la Tabla 6 y en el Anexo 3, e «instrumento REV_IND» donde decía «protocolo PSI-05»;
    - una oración después de la Tabla 6 con la correspondencia con los códigos de la versión remitida con los memorándums;
    - columna «Código» de la Tabla 6 más ancha, para que REV_DOC quepa en una línea.

    REV_IND.4 (recálculo del OP3) sigue sin figurar en la nota, porque agregarlo sería un cambio de diseño.
  - **Presentaciones de la nota:** solo cambian las siglas, sobre los archivos del repositorio.
- **Verificaciones:**
  - **Comparación con la línea base:** se volcaron a texto las 176 salidas antes del cambio y se tradujeron con el script. Después de regenerar todo, 91 coinciden. Las 5 diferencias están explicadas:
    - la nota, por los cambios de texto previstos;
    - la matriz de triangulación y el informe, solo por el orden alfabético de las listas de instrumentos (ENTREV, REP_ENT…);
    - el índice de figuras, porque la nota al pie de la figura de fichas pasa a dos líneas;
    - en un resumen de la búsqueda manual, los códigos JEL B51 y B54 no se traducen: es lo correcto, y la diferencia la genera el comparador.
  - **Transcripciones:** el anonimizador actualizado reproduce byte a byte la traducción de las tres transcripciones.
  - **Barrido:** no quedan códigos anteriores en nombres ni contenidos (texto, xlsx, docx y pptx), salvo en lo histórico listado arriba, en este ADR, en el script de migración y en la oración de correspondencia de la nota.
  - **Scripts:** pasan todas sus comprobaciones (citas de ENTREV en las transcripciones, medios de verificación, frases de los segmentos, conteos de la síntesis).
  - **Inventario:** sigue con 285 documentos y los mismos hashes.
  - **Office:** `validate.py` aprueba el informe, la nota, las dos presentaciones, las guías y los cuestionarios, y se revisó su renderizado.
- **No se volvieron a correr** los scripts que consultan servicios externos (búsquedas académicas, 3ie DEP, descargas de `src/fuentes/`) ni el importador manual. Sus salidas versionadas se tradujeron, y el código ya genera los identificadores nuevos.

## Actualización (2026-11-07): nombres en el informe
- **Decisión del usuario:** el informe y la presentación de resultados usan los nombres de los instrumentos y de sus análisis, no sus siglas ni los identificadores de traza.
- **Informe:**
  - **Glosario:** el numeral 2.3 presenta las siglas en dos tablas, una de instrumentos y análisis y otra de identificadores del rastro. Así no se necesita la sigla entre paréntesis en la primera mención.
  - **Línea de sustento:** cada hallazgo nombra sus fuentes primarias (por ejemplo, «consolidado de indicadores de la OMEP, PEMIG (anexo 5) y entrevistas a la OGPPM, la OGMEPGD y la DGIGND»), con la función `sintesis_hallazgos.fuentes_legibles()`.
  - **Anexo 1:** la columna «Rastro» conserva libro y hoja, página o segmento de cada evidencia (`sintesis_hallazgos.rastro()`).
  - **Notas de fuente de tablas y figuras:** siguen la forma «Fuente: <fuente primaria>. Elaboración: equipo evaluador (<análisis>).». El libro y la hoja de cada una quedan en `06_informe/rastro_tablas_figuras.csv` (parámetro `rastro=` de `docx_mimp`) y en la columna «rastro» del índice de figuras.
  - **Anexo 6:** el medio de verificación es legible; `seguimiento_recomendaciones_previas.xlsx` lleva además la columna «Rastro del medio de verificación».
  - **Dónde quedan siglas e identificadores:** solo en el glosario, en la columna de rastro del Anexo 1, en la columna N.° del Anexo 6 y en las deducciones del Anexo 7.
- **No cambia:** la nota metodológica y sus presentaciones, donde las siglas definen los instrumentos, ni los libros de análisis y la matriz de triangulación, que son técnicos.
- **Verificación:**
  - el barrido del informe da cero siglas fuera de esos lugares;
  - toda referencia de las líneas de sustento anteriores está en el Anexo 1;
  - el cuerpo solo cambia en las menciones reemplazadas, el glosario y el título del Anexo 7;
  - `validate.py` aprueba el informe y se revisó su renderizado (88 páginas).

## Segunda actualización (2026-11-08): informe sin siglas (ADR-010)
- **Qué cambia:** el informe publicable ya no lleva siglas de instrumentos ni identificadores de traza en ninguna parte: salieron el glosario de siglas del numeral 2.3 y la columna de rastro del anexo de hallazgos.
- **Respaldo:** el anexo de hallazgos (ahora Anexo 2) cita documentos con su página, libros con nombre legible y hoja, y entrevistas con su fecha.
- **Rastro técnico:** queda en `06_informe/rastro_hallazgos.csv` y `rastro_tablas_figuras.csv`, para el anexo técnico.
- **Lo que no cambia:** las siglas siguen en los libros de análisis, los scripts, la nota metodológica y sus presentaciones.
