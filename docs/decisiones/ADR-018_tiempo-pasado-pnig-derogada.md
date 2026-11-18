# ADR-018. Informe y presentaciones en tiempo pasado (PNIG derogada)

**Fecha:** 2026-11-18 · **Estado:** aprobada por el usuario

## Contexto
- **Observación de la DGIGND:** la PNIG perdió vigencia el 25 de diciembre de 2025 (criterio de la OGAJ, ADR-015 y
  ADR-016), así que los productos deben referirse a ella en pasado. Ejemplo del usuario: «De los siete enfoques
  transversales, solo el de género se operacionaliza con amplitud» pasa a «se operacionalizó».
- **Archivos para enviar:** en la carpeta administrativa (`Para enviar`), el usuario dejó el Word del informe y una
  presentación de 11 láminas. El Word tenía dos ediciones suyas:
  - en el resumen, el formulario de seguimiento de la OMEP, sin el conteo de segmentos codificados;
  - un pie de página más breve.
- **Pedidos del usuario:**
  - pasar a pasado lo que corresponde a la PNIG;
  - mejorar la presentación, sobre todo sus dos últimas láminas;
  - agregar al Word la barra de desplazamiento vertical.

## Decisión
1. **Criterio verbal** (equipo evaluador):
   - **Pasado:** lo que la PNIG fue, tuvo, previó y estableció, y cómo se implementó, midió, condujo y articuló en
     2019–2025. Incluye los servicios, los protocolos, las fichas, los indicadores y sus reportes, el PEMIG, la CONAIG y
     la rectoría.
   - **Presente:**
     - los hechos actuales: las brechas persisten, el diagnóstico conserva vigencia, las normas vigentes;
     - lo que hace el propio informe y su metodología;
     - las definiciones y escalas, y la evidencia de la literatura;
     - el análisis de la transición;
     - las recomendaciones;
     - las citas textuales.
2. **Informe** (decisión del usuario: en el código y regenerado):
   - unos 300 cambios en `informe_cierre.py`, `sintesis_hallazgos.py`, `informe_evidencia.py`,
     `informe_referencias.py` y `figuras_informe.py`;
   - en `04_analisis/seguimiento_recomendaciones/valoracion_recomendaciones.csv`, el estado de las recomendaciones
     previas al cierre;
   - las dos ediciones del usuario se llevan al código;
   - se regeneran el seguimiento de recomendaciones, la triangulación, las figuras con texto y el informe, que Word
     termina.
3. **Preguntas** (decisión del usuario):
   - pasan a pasado las que describen la PNIG: D2, D3, D4, D5, D6 e I4 (`config.yaml`);
   - quedan en presente D1 (vigencia actual), D7 (nuevo marco) e I5 (capacidades actuales);
   - I1 a I3 ya estaban en pasado.
4. **Presentaciones e infografía** (decisión del usuario: también las del repositorio):
   - la completa se edita sobre el archivo (19 cambios);
   - la resumida y la infografía se regeneran con sus scripts.
5. **Presentación de 11 láminas de `Para enviar`**, editada sobre el archivo:
   - 75 cambios a pasado en las láminas 2 a 9 y en sus notas;
   - lámina 10, rehecha en el estilo institucional de las láminas 2 a 9: tres columnas por aspecto, cada
     recomendación con su unidad responsable y su plazo (Tabla 25) y un recuadro de quién responde y cuándo;
   - lámina 11, rehecha como cierre: tres mensajes finales en pasado sobre el fondo de la portada;
   - la versión anterior y la del Word están en `informe_borradores_anteriores\2026-11-18_antes-tiempo-pasado\`.
6. **Barra de desplazamiento:** la barra vertical (y el panel de navegación) es una opción de Word en cada equipo, no
   del documento; se probó con Word que el archivo no la guarda. En una instalación normal aparece sola; si no, se
   activa en Archivo > Opciones > Avanzadas > Mostrar.

## Verificación
- **Diff del texto del informe:** solo cambian verbos y las dos ediciones del usuario. Las únicas cifras distintas son
  los números de página de los índices (el informe pasa de 213 a 214 páginas).
- **Auditorías:**
  - la de estilo, sin observaciones nuevas;
  - la del anexo técnico rearmado, sin problemas (208 archivos, 0 posibles datos personales).
- **Pruebas:** 58 en verde. La referencia del piloto se recapturó por el texto del informe, tres figuras, dos libros
  y un CSV.
- **Presentaciones:** las tres pasan la validación y se revisaron renderizadas; la de 11 láminas, sin `effectLst`
  duplicados.
- **SHA-256 de lo que quedó en `Para enviar`:**
  - Word: `7bf1c6abba85850b539bad927a69bb321760ff89f160646d83b703f5866d26b5`;
  - presentación: `272fe09f924d92d0424185103210187d82f4076eba38c95f228cbca9f262b16a`.
