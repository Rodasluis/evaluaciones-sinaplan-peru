# ADR-016. Informes técnicos de la transición, contraste con la matriz de incompatibilidades, presentación resumida e infografía

**Fecha:** 2026-11-17 · **Estado:** aprobada por el usuario

## Contexto
- **Documentos entregados por el usuario (2026-11-17)** en la carpeta administrativa «informe de derrogacion PNIG»: los informes que sustentan la pérdida de vigencia de la PNIG.
  - El Informe N.° D000926-2026-MIMP-OGAJ ya estaba en el corpus (ADR-015), con el mismo SHA-256.
  - **Informe N.° D000810-2026-MIMP-OGAJ (16-09-2026):** las nueve normas dictadas en el marco de la Ley N.° 28983 siguen vigentes, con sus disposiciones incompatibles inaplicables y el marco conceptual de la PNIG desplazado; encarga a los órganos técnicos la matriz de incompatibilidades.
  - **Informe N.° D000041-2026-MIMP-DGIGND (09-11-2026)**, que consolida el Informe N.° D000067-2026-MIMP-DPIGND y el Informe Técnico N.° D000040-2026-MIMP-DTEG: la incompatibilidad alcanza al objeto de las normas, por lo que hubo derogación tácita de la PNIG, el PEMIG, la CONAIG y otros cuatro decretos supremos, sin vacío normativo. Además:
    - cita el reporte de la PCM que registra a la CONAIG como extinguida;
    - recomienda comunicar la derogación tácita conforme a la Ley N.° 31772 a más tardar el 31-12-2026.
  - **Matriz técnica de incompatibilidades de la DPIGND (09-11-2026):** 108 disposiciones (PNIG 86, PEMIG 17 y CONAIG 5), todas marcadas como incompatibles; para cada grupo de servicios, la base de la Ley N.° 32535 o de la Ley N.° 30364 con la que su contenido puede seguir ejecutándose, o la razón de su incompatibilidad.
- Hasta ahora, el informe citaba estos documentos solo a través de la OGAJ.
- **Pedido del usuario:** incorporarlos y analizarlos, agregar una versión resumida de la presentación de resultados y una infografía de una o dos caras, y revisar si el repositorio está listo para publicarse.

## Decisión
1. **Corpus:** los cinco documentos nuevos se copian a `docs/fuentes/pnig/07_transicion_ley_32535/`, con su hash en `movimientos_derogacion_2026-11-17.tsv` y sus filas en `registro_fuentes.md`. Son documentos internos sin publicación oficial: no salen en la versión pública, salvo en el inventario.
2. **Contraste de D7 con la matriz** (decisión del usuario):
   - `src/analisis_pnig/rev_doc_matriz_incompatibilidades.py` expande los grupos de servicios de la matriz a sus 52 códigos y los clasifica con una regla fija, guardada en `02_insumos/matriz_incompatibilidades.yaml`:
     - la referencia a la educación sexual integral → «se rediseña»;
     - la continuidad sin elementos del enfoque de género o con contenidos por adecuar → «se ajusta»;
     - lo demás → «se mantiene».
   - Esa clasificación se cruza con la de REV_ALI.5, hecha antes de conocer la matriz, y se mide el acuerdo simple y el kappa de Cohen.
   - La matriz no cambia la clasificación de la evaluación.
   - Salida: `04_analisis/REV_DOC_matriz_incompatibilidades.xlsx`, publicado como «Contraste_matriz_incompatibilidades.xlsx».
3. **Informe:**
   - fuentes primarias en lugar de las citas a través de la OGAJ: MIMP (2026f) es la OGAJ D000810, (2026g) la DGIGND D000041 con su matriz, (2026h) la DPIGND D000067 y (2026i) la DTEG D000040;
   - el numeral 1.1 cuenta la secuencia de criterios (setiembre, 9 de noviembre y 11 de noviembre);
   - el numeral 3.2 presenta el contraste en una tabla nueva (Tabla 11) con su análisis;
   - se actualizan D1-H4, D7-H1, D7-H2, I4-H2 e I5-H2, y las valoraciones de D7, I4 e I5;
   - la recomendación «Implementación: transición normativa» agrega la comunicación a las entidades y la de la Ley N.° 31772;
   - el glosario agrega «Derogación tácita».
4. **Origen:** los tres informes técnicos y la matriz son un solo origen («informes técnicos de la DGIGND»), aparte de la OGAJ (regla de ADR-015).
5. **Presentación resumida:** `src/analisis_pnig/presentacion_resumida.py`, ocho láminas con el estilo del MIMP Lab y las cifras leídas de las mismas fuentes que el informe. Después de la primera versión se edita a mano, como la completa.
6. **Infografía:** `src/analisis_pnig/infografia.py`, dos caras A4.
   - Arma un HTML autocontenido y Chrome sin ventana lo imprime en PDF.
   - Los estados usan los colores de las figuras del informe, con el ámbar oscurecido a #e8a317 para pasar el validador de la skill `dataviz` y con etiqueta en cada marca.
   - Sin los logos de `assets/` (versión pública), la banda queda en blanco.
   - Se regenera después del informe.
7. **Componente de la skill:** la nota de confianza de `K.tabla_valoracion` usa la regla del origen independiente.
8. **Copia con control de cambios:** la versión final del 13-11 contra la nueva, en la carpeta administrativa (`Informe_evaluacion_PNIG_version_final_2026-11-17_control_cambios.docx`), con `comparar_word.ps1`. Las respuestas a comentarios se firman con `-AutorRespuestas` (parámetro nuevo).

## Consecuencias
- **Contraste:** 42 de 52 servicios coinciden (acuerdo de 80,8 %, kappa de 0,59).
  - Coinciden en los dos que se rediseñan (2.2.1 y 2.2.2).
  - Las 10 diferencias son entre mantener y ajustar:
    - la evaluación ajusta y la matriz mantiene: 1.2.1, 2.1.1 a 2.1.4, 4.1.1 y 6.2.1;
    - la matriz ajusta y la evaluación mantiene: 3.1.2, 4.3.3 y 6.4.1.
- **Confianza y valoraciones:** no cambian (24 sólidos, 10 con respaldo parcial y 9 indicios; 11 parciales y 1 insuficiente).
- **Corpus:** 293 documentos (antes 288).
- **Informe:** 213 páginas y una tabla más, que corre la numeración de las siguientes. La referencia del piloto se recapturó.
- **Presentación completa:** se editó sobre el archivo (láminas 2, 9, 13 y 15).
- **Anexo técnico:** suma la presentación resumida, la infografía y el libro del contraste.
- **Versión pública:** se rearma sin publicar. El libro del contraste sale con el texto de las observaciones de la matriz, que trata normas y servicios, sin datos personales; la matriz misma no sale.

## Adenda (2026-11-17): evaluación culminada y presentaciones autoexplicativas
- **Validación:** el usuario informó que la DGIGND validó el informe; la evaluación se presenta como culminada. El informe no cambia.
- **Presentaciones** (pedido del usuario): las dos se leen sin expositor.
  - Cada lámina de contenido abre con una bajada, componente nuevo `K.bajada` de la skill, que dice qué muestra y qué concluye.
  - Las recomendaciones llevan una frase que las explica.
  - Las siglas se escriben completas en su primer uso.
  - Se quitó la sección de próximos pasos.
- **Presentación completa:** se editó sobre el archivo. Tiene 17 láminas; la portada dice «Resultados de la evaluación».
- **Presentación resumida:** como no tenía ediciones a mano, se regeneró con su script. Reemplaza los próximos pasos por «Qué deja la evaluación para la nueva política».
- **Corpus:** las presentaciones y la infografía no usan la palabra: dicen «documentos revisados (normas, planes e informes)».
- **Infografía:** reemplaza los próximos pasos por «Productos de la evaluación».
