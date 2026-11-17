# ADR-015. Versión final del informe: criterio de la OGAJ y revisión institucional

**Fecha:** 2026-11-13 · **Estado:** aprobada por el usuario

## Contexto
- **Informe N.° D000926-2026-MIMP-OGAJ (11-11-2026):** la Oficina General de Asesoría Jurídica concluye que la PNIG, el PEMIG, la CONAIG y otras normas dictadas en el marco de la Ley N.° 28983 perdieron vigencia el 25 de diciembre de 2025, porque su objeto es incompatible con la Ley N.° 32535. Los contenidos compatibles siguen ejecutándose con sustento directo en esa ley hasta que se aprueben los nuevos instrumentos, y el MIMP debe informar el criterio al CEPLAN. Se incorporó al corpus en `d10830f` (`docs/fuentes/pnig/07_transicion_ley_32535/`).
- **Revisión institucional del informe preliminar, versión 3 (13-11-2026):** un documento en control de cambios con 65 observaciones y 20 cambios propuestos, generados con IA y sin revisión humana confirmada. No va al repositorio: queda en la carpeta administrativa.
- **Decisiones del usuario (2026-11-13), sobre el cuadro de observaciones que le presenté:**
  - aplicar en general lo propuesto;
  - 01: no cambiar el título; la respuesta indica que el informe de la OGAJ se incluyó en el análisis;
  - 42: el servicio 6.4.1 pasa a «No implementado»;
  - 20 y 60: revisar el origen de las fuentes de los hallazgos sólidos;
  - 65: la comparación de la ENA de 2019 a 2025 se marca como indicativa;
  - 17: la respuesta justifica que el plazo de la evaluación no permite revisar los textos completos de la literatura;
  - 59: la respuesta remite al anexo técnico;
  - la versión final, limpia y publicable, y una copia con control de cambios y respuestas en la carpeta administrativa.

## Decisión
1. **Criterio de la OGAJ:** se presenta como el criterio jurídico del órgano competente del MIMP, que la evaluación adopta como referencia normativa, no como un juicio propio. El periodo evaluado (2019–2025) coincide con la vigencia de la PNIG; el corte de la evidencia es noviembre de 2026, y lo ocurrido en 2026 se lee como contexto. Afecta a D1 (objeto incompatible con la ley), D4 (la correspondencia con la ley es la del contenido de los objetivos y servicios), D7, I4 (la CONAIG agotó su objeto) e I5 (no hay vacío normativo, los protocolos aprobados conservan su vigencia y los mecanismos para la igualdad quedan sin mandato).
2. **Servicio 6.4.1:** queda como «No implementado», porque el servicio siguió en diseño todo el periodo y las actividades del INABIF y del SIEF son conexas. Su asignación presupuestal se descarta: la meta que coincidía con el monto del LIO 2022 financia servicios de violencia familiar. Las dos decisiones están en `02_insumos/estado_servicios.yaml` (`ajustes` y la clave nueva `asignacion_descartada`), que aplica `rev_doc_estado_servicios.py`.
3. **Confianza por origen independiente:** un hallazgo es sólido cuando coinciden al menos dos fuentes con productores primarios distintos; dos documentos que reproducen el dato de un mismo productor cuentan como un solo origen. La regla vale para toda la escala, no solo para los sólidos.
   - **Núcleo:** `nucleo.hallazgos.fuentes_independientes(..., origen=)` cuenta orígenes, y `evidencia(..., origen=)` admite un origen explícito. Sin el parámetro, cada fuente es su propio origen, como antes, de modo que otras evaluaciones no cambian.
   - **Capa del piloto:** el mapa `ORIGEN` de `sintesis_hallazgos.py` agrupa los orígenes:
     - información reportada por las entidades: informes LIO, formulario, informes de gobernanza de la CONAIG, y los reportes oficiales de indicadores (RS, RC y consolidado);
     - SIAF: datos abiertos del MEF e información presupuestal de la DGIGND;
     - organismos electorales: JNE y ONPE;
     - instrumentos de planeamiento, con las normas de aprobación de la PGG;
     - cada encuesta, cada entrevista (sin la OGMEPGD) y el informe de la OGAJ, por separado.
4. **Comparaciones indicativas:** las de propiedad y conducción de parcelas de la ENA de 2019 a 2025 (Tabla 10, Anexo 6 y D1-H1) y la de la ENUT (2010–2024). El conteo de brechas se presenta como 14 comparaciones con el mismo instrumento más la ENUT, aparte.
5. **Sensibilidad de la escala:** el libro del estado de los servicios agrega los totales con la regla anterior de tres de tres componentes (ADR-011), y el informe los cita. La regla vigente sigue siendo la de ADR-012.
6. **Otras observaciones aceptadas:**
   - las tres categorías de indicador (formulado, aprobado y con valor): «indicador aprobado» en lugar de «indicador vigente»;
   - las precisiones de método: validación de las entrevistas, kappa, varianza y covarianza, búsqueda de literatura y sus límites;
   - la Tabla 21 con el último valor, el escenario lineal al 2030, el dato 2026 y el resultado;
   - las recomendaciones con plazos concretos, y una nueva, «Implementación: transición normativa» (matriz de incompatibilidades e informe al CEPLAN): pasan de 15 a 16;
   - la versión del informe pasa a «Versión final».
7. **Observaciones no aceptadas, con su respuesta en la copia con control de cambios:**
   - 01: no se cambia el título;
   - 11: no se renombra el criterio;
   - 27: se mantiene «departamentos».
8. **Cambios propuestos no aceptados** (en el texto, sin comentario):
   - la participación que la revisión corrigió a 48 % se mantiene en 47 %: con la estimación sin redondear (77,551) es 46,8 %, y el 48 % sale de los valores redondeados;
   - «sobre del gasto público» se mantiene porque es la denominación oficial del servicio en el anexo del D.S. N.° 008-2019-MIMP.
9. **Copia con control de cambios:**
   - `src/nucleo/rechazar_cambios.py` rechaza en el XML los 20 cambios de texto propuestos y acepta los 764 cambios de formato registrados. Con Word por COM, `Revisions.RejectAll` seguía sin terminar a los 35 minutos. La base resultante coincide con la versión 3 salvo en espacios y en números de leyenda, que la comparación no considera.
   - `src/nucleo/comparar_word.ps1` (Word por COM) compara esa base, con sus 65 comentarios, con la versión final y responde cada comentario con el texto de un JSON.
   - El JSON y la copia quedan en la carpeta administrativa.

## Consecuencias
- **Estados de los servicios:** 24 implementados, 26 parcialmente y 2 no implementados (antes 24, 27 y 1). Asignación identificable: 33 de 52 (antes 34). Con tres de tres componentes: 17, 33 y 2.
- **Confianza de los 43 hallazgos:** 24 sólidos, 10 con respaldo parcial y 9 indicios (antes 30, 11 y 2).

| Hallazgo | Antes | Ahora | Motivo |
|---|---|---|---|
| D1-H2 | Sólido | Indicio | El recálculo con la ENDES y el reporte oficial vienen del mismo dato |
| D1-H4 | Indicio | Sólido | Se suma el informe de la OGAJ |
| D2-H2 | Sólido | Indicio | Reporte oficial y consolidado, un solo origen |
| D7-H1 | Indicio | Con respaldo parcial | Se suma el informe de la OGAJ, que contrasta con la entrevista |
| I2-H1 | Sólido | Indicio | Reporte oficial y consolidado, un solo origen |
| I2-H2 | Con respaldo parcial | Indicio | RS y RC, un solo origen |
| I2-H4 | Sólido | Indicio | Reporte oficial y formulario, un solo origen |
| I3-H2 | Sólido | Indicio | MEF e información de la DGIGND, ambos del SIAF |
| I4-H3 | Con respaldo parcial | Indicio | Informes de gobernanza y formulario, un solo origen |
| I5-H3 | Sólido | Indicio | Solo información reportada por las entidades |
| I5-H4 | Sólido | Indicio | Solo información reportada por las entidades |

- **D4-H2** registra como evidencia el documento de la PNIG, que el análisis ya usaba, y se mantiene sólido.
- **Valoraciones:** no cambian (11 parciales y 1 insuficiente); se actualizan los textos de D1, D3, D4, D5, D7, I1, I4 e I5.
- **Archivos que cambian:**
  - `src/nucleo/hallazgos.py` y los scripts nuevos `src/nucleo/rechazar_cambios.py` y `src/nucleo/comparar_word.ps1`; pruebas nuevas en `tests/test_nucleo.py` (origen independiente y rechazo de cambios);
  - `src/analisis_pnig/`: `rev_doc_estado_servicios.py`, `sintesis_hallazgos.py`, `informe_cierre.py`, `informe_referencias.py`, `informe_evidencia.py`, `rev_evi_insumos.py` y `d1_d2_departamental.py`;
  - `evaluaciones/pnig_2026/02_insumos/`: `estado_servicios.yaml` e `hitos.yaml`;
  - los libros de `04_analisis`, cinco figuras, el rastro y el informe;
  - la presentación de resultados, editada sobre el archivo (láminas 2, 3, 4, 7, 11, 13, 15 y 16).
- **Referencia del piloto:** se recapturó, y `test_cifras_clave` fija las cifras nuevas.
- **Productos derivados:** el anexo técnico y la versión pública se rearman; la publicación sigue esperando la validación del informe.
