# Plantilla de evaluación nueva

`src/nucleo/nueva_evaluacion.py` copia esta plantilla y la completa con la tipología elegida (`tipologias/`):

```
uv run python src/nucleo/nueva_evaluacion.py pnx_2027 --tipologia cierre \
    --titulo "Evaluación de cierre de la Política Nacional de … (2020–2026)" \
    --politica "Política Nacional de …" --sigla PNX --inicio 2020 --fin 2026
```

| Origen en la plantilla | Destino |
|---|---|
| `evaluacion/` | `evaluaciones/<id>/`: carpetas 00 a 06 con su LEEME |
| `config.yaml.plantilla` | `evaluaciones/<id>/config.yaml`: preguntas tipo, instrumentos, escalas y estructura del informe de la tipología |
| `config/portada.yaml` del repositorio | `evaluaciones/<id>/portada.yaml`: etiqueta, título y fecha de la evaluación |
| `hallazgos.yaml.plantilla` | `evaluaciones/<id>/04_analisis/hallazgos.yaml`: un hallazgo pendiente por pregunta |
| `analisis/` | `src/analisis_<id>/`: capa de evaluación mínima sobre el núcleo |

## Flujo
1. **Encargo y diseño.** Completar `config.yaml` (política, periodo, entidad) y ajustar las preguntas tipo a la política: cada pregunta conserva su criterio y sus instrumentos. Redactar la nota metodológica en `00_encargo/`. Los cambios de diseño metodológico los aprueba el usuario (CLAUDE.md, regla 5).
2. **Fuentes.** Descargar o mover los documentos a `docs/fuentes/`, con su hash en `docs/fuentes/*.tsv` y la fuente en `registro_fuentes.md` (regla 7). Lo no verificado se marca `[POR VERIFICAR]`.
3. **Insumos del caso.** Los ajustes, verificaciones y catálogos van en `02_insumos/` (YAML o CSV), no en el código.
4. **Instrumentos.** Un script por instrumento en `src/analisis_<id>/`, sobre los módulos del núcleo. Los del piloto PNIG (`src/analisis_pnig/`) sirven de modelo:

   | Instrumento | Módulos del núcleo | Modelo en el piloto |
   |---|---|---|
   | Revisión documental | `escalas.clasificar_implementacion` | `rev_doc_estado_servicios.py` |
   | Entrevistas | `codificacion.kappa` | `entrev_*.py` |
   | Información reportada por las entidades | `codificacion` | `rep_ent_*.py` |
   | Coherencia y alineamiento | `hallazgos.hoja_publica` | `rev_ali_*.py` |
   | Indicadores y recursos | `escalas.avance`, `escalas.nivel`, `presupuesto.conciliar`, `encuestas` | `rev_ind_*.py` |
   | Revisión de evidencia | `evidencia.solidez`, `evidencia.plausibilidad` | `rev_evi_*.py` |

5. **Hallazgos.** El equipo los escribe en `04_analisis/hallazgos.yaml` con sus evidencias. `sintesis.py` los verifica (fuentes independientes para la confianza, instrumentos de cada pregunta, sin guiones explicativos ni términos internos) y escribe la matriz de triangulación.
6. **Informe.** `informe.py` lo arma con la estructura de la tipología: preguntas en recuadro, un pie de fuente por hallazgo («Fuente: … Elaboración propia.»), conclusiones, la matriz de recomendaciones de la Tabla 4 y el anexo de hallazgos. Lo pendiente queda como `[POR COMPLETAR]`. Con `--word`, Word actualiza los índices y exporta el PDF.
7. **Auditoría y entrega.** `uv run python src/nucleo/auditoria.py informe <docx>` antes de entregar y `uv run python src/nucleo/auditoria.py entrega <carpeta>` antes de enviar el anexo técnico.

```
uv run python src/analisis_<id>/sintesis.py
uv run python src/analisis_<id>/informe.py --word
```

## Qué es de la plantilla y qué es de cada evaluación
- **Núcleo (`src/nucleo/`):** documento, portada, figuras, formato, escalas, hallazgos, evidencia, presupuesto, encuestas, codificación, anexo y auditoría. No lleva datos de ninguna evaluación.
- **Capa de evaluación (`src/analisis_<id>/`):** los análisis y el relato propios. Crece a partir de `sintesis.py` e `informe.py`; el piloto muestra hasta dónde puede llegar (`src/analisis_pnig/informe_cierre.py`).
- **Datos (`evaluaciones/<id>/`):** configuración, insumos, datos, análisis, entrevistas, informe y anexo técnico.
