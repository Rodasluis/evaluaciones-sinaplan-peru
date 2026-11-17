---
description: Regenera los análisis, las figuras y el informe de una evaluación en el orden correcto, con pruebas y auditoría
argument-hint: [id de la evaluación, por defecto pnig_2026]
---

Regenera el informe de la evaluación `$ARGUMENTS` (si no se indica, `pnig_2026`). Ejecuta todo desde la raíz del
repositorio con `UV_OFFLINE=1 PYTHONUTF8=1` y detente en el primer error.

**Piloto PNIG (`pnig_2026`)**, en este orden (el de la verificación de la Fase 2):
1. Solo si cambiaron sus insumos, los scripts de cada instrumento en `src/analisis_pnig/` (revisión documental,
   coherencia y alineamiento, indicadores y recursos, información de las entidades, evidencia). Después de la
   coherencia y alineamiento, `rev_doc_matriz_incompatibilidades.py` (contraste con la matriz de la DPIGND, ADR-016).
2. `uv run python src/analisis_pnig/rev_doc_estado_servicios.py`
3. `uv run python src/analisis_pnig/rev_ind_indicadores.py`
4. `uv run python src/analisis_pnig/rev_ind_3_presupuesto_dgignd.py`
5. `uv run python src/analisis_pnig/rev_evi_evidencia.py`
6. `uv run python src/analisis_pnig/sintesis_triangulacion.py`
7. `uv run python src/analisis_pnig/figuras_informe.py`
8. `uv run python src/analisis_pnig/informe_cierre.py` (lo termina Word con `src/nucleo/actualizar_campos_word.ps1`)
9. `uv run pytest`: la referencia del piloto (`tests/referencia_pnig/`) tiene que seguir igual. Si una diferencia es
   intencional, muéstrala al usuario y vuelve a capturar con `uv run python tests/referencia.py capturar` solo con su
   aprobación.

**Otra evaluación** (`src/analisis_<id>/`):
1. Los scripts de sus instrumentos, si los tiene.
2. `uv run python src/analisis_<id>/sintesis.py`
3. `uv run python src/analisis_<id>/informe.py --word`

**Al final, en los dos casos:**
- `uv run python src/nucleo/auditoria.py informe <ruta del informe>` y un resumen de lo que reporta.
- Los .xlsx y .docx regenerados sin cambios de contenido se restauran con `git restore` para no ensuciar el historial.
- Las presentaciones de resultados (completa y resumida) no se regeneran: se editan a mano sobre el archivo del
  repositorio. La infografía sí: `uv run python src/analisis_pnig/infografia.py`, después del informe.
