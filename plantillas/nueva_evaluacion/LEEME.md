# Plantilla de evaluación nueva

`src/nucleo/nueva_evaluacion.py` copia esta plantilla y la completa con la tipología elegida (`tipologias/`) y las
respuestas a las preguntas iniciales. En Claude Code, `/nueva-evaluacion` hace las preguntas y crea la evaluación, y
`/evaluar <id>` recorre las etapas con los agentes (`docs/agentes.md`). Sin Claude Code:

```
uv run python src/nucleo/nueva_evaluacion.py pnx_2027 --tipologia cierre \
    --titulo "Evaluación de cierre de la Política Nacional de … (2020–2026)" \
    --politica "Política Nacional de …" --sigla PNX --inicio 2020 --fin 2026 \
    [--sin-entrevistas] [--formulario-entidades] [--documentos-internos] [--validacion informe_final] \
    [--personalizar REV_IND]
```

| Origen en la plantilla | Destino |
|---|---|
| `evaluacion/` | `evaluaciones/<id>/`: carpetas 00 a 06 con su LEEME |
| `config.yaml.plantilla` | `evaluaciones/<id>/config.yaml`: preguntas tipo, instrumentos, escalas, estructura del informe y `modalidad` |
| `config/portada.yaml` del repositorio | `evaluaciones/<id>/portada.yaml`: etiqueta, título y fecha de la evaluación |
| `hallazgos.yaml.plantilla` | `evaluaciones/<id>/04_analisis/hallazgos.yaml`: un hallazgo pendiente por pregunta |
| `analisis/` | `src/analisis_<id>/`: síntesis e informe sobre el núcleo |
| esquemas de `src/nucleo/instrumentos/` | cabeceras de las entradas en `02_insumos/` y `05_entrevistas/`, y `avance.yaml` |

## Preguntas iniciales (sección `modalidad` de config.yaml)
| Pregunta | Efecto |
|---|---|
| ¿Hará entrevistas? | Sin entrevistas, las preguntas pierden ese instrumento; una pregunta que se queda sin ninguno recibe la revisión documental. La etapa de recojo se omite si tampoco hay formulario |
| ¿Pedirá información a las entidades con un formulario por oficio? | La información de las entidades arma el formulario y consolida las respuestas |
| ¿Cuenta con documentos internos adicionales? | En la etapa de fuentes se pide la carpeta; los documentos se copian con su hash y se marcan sin publicación oficial |
| ¿Validación con el órgano rector o solo informe final? | Activa u omite la etapa de validación |
| ¿Desea personalizar los instrumentos? | Para los elegidos, un agente escribe código propio en `src/analisis_<id>/` que complementa al genérico |

## Flujo
1. **Diseño.** Preguntas adaptadas a la política, matriz pregunta–instrumento–fuente y nota metodológica en
   `00_encargo/`. Los cambios de diseño metodológico los aprueba el usuario (CLAUDE.md, regla 5).
2. **Fuentes.** Documentos en `docs/fuentes/`, con su hash en `docs/fuentes/*.tsv` y su fila en
   `registro_fuentes.md` (regla 7). Lo no verificado se marca `[POR VERIFICAR]`.
3. **Lectura de la política.** Catálogo de servicios, lineamientos, indicadores y cadena de resultados en
   `02_insumos/` (esquemas en `evaluacion/02_insumos/LEEME.md`).
4. **Instrumentos.** Los genéricos del núcleo, sobre esas entradas:
   `uv run python src/nucleo/instrumentos/correr.py <id> [--validar]`.

   | Instrumento | Módulo genérico | Libro | Modelo del piloto para personalizar |
   |---|---|---|---|
   | Revisión documental | `rev_doc` | `REV_DOC_estado_servicios.xlsx` | `rev_doc_estado_servicios.py` |
   | Indicadores y recursos | `rev_ind` | `REV_IND_indicadores_recursos.xlsx` | `rev_ind_*.py` |
   | Coherencia y alineamiento | `rev_ali` | `REV_ALI_coherencia_alineamiento.xlsx` | `rev_ali_*.py` |
   | Revisión de evidencia | `rev_evi` y `busqueda_academica.py` | `REV_EVI_evidencia.xlsx` | `rev_evi_*.py` |
   | Información de las entidades | `rep_ent` | `REP_ENT_informacion_entidades.xlsx` | `rep_ent_*.py` |
   | Entrevistas | `entrev` | `ENTREV_registro_codificado.xlsx` | `entrev_guias_entrevista.py` |

5. **Recojo primario,** si hay entrevistas o formulario: guías por rol (`correr.py <id> --guias`) o formularios por
   entidad (`--formulario`). Después, las transcripciones anonimizadas o las respuestas (`--consolidar`).
6. **Hallazgos.** En `04_analisis/hallazgos.yaml`, con sus evidencias. `sintesis.py` los verifica (fuentes de origen
   independiente para la confianza, instrumentos de cada pregunta, sin guiones explicativos ni términos internos) y
   escribe la matriz de triangulación.
7. **Informe.** `informe.py` lo arma con la estructura de la tipología. Lleva las preguntas en recuadro, un pie de
   fuente por hallazgo («Fuente: … Elaboración propia.»), las conclusiones y la matriz de recomendaciones de la
   Tabla 4. Cierra con dos anexos: los hallazgos y, con una tabla y su análisis por libro genérico, los resultados de
   los instrumentos. Lo pendiente queda como `[POR COMPLETAR]`. Con `--word`, Word actualiza los índices y exporta el
   PDF.
8. **Validación,** si corresponde: `src/nucleo/observaciones.py` arma la matriz de observaciones y el JSON de
   respuestas para `comparar_word.ps1`, fuera del repositorio.
9. **Auditoría y entrega.** `uv run python src/nucleo/auditoria.py informe <docx>` antes de entregar y
   `uv run python src/nucleo/auditoria.py entrega <carpeta>` antes de enviar el anexo técnico.

```
uv run python src/analisis_<id>/sintesis.py
uv run python src/analisis_<id>/informe.py --word
```

## Qué es de la plantilla y qué es de cada evaluación
- **Núcleo (`src/nucleo/`):**
  - documento, portada, figuras, formato, escalas, hallazgos, evidencia, presupuesto, encuestas, codificación,
    anexo y auditoría;
  - los instrumentos genéricos, el avance por etapas y las observaciones de la validación.

  No lleva datos de ninguna evaluación.
- **Capa de evaluación (`src/analisis_<id>/`):**
  - la síntesis y el informe;
  - el código propio de los instrumentos personalizados.

  El piloto muestra hasta dónde puede llegar (`src/analisis_pnig/`).
- **Datos (`evaluaciones/<id>/`):** configuración, entradas, datos, análisis, entrevistas, informe, anexo técnico y
  avance.
