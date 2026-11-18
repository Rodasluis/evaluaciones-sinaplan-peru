# 02_insumos

Entradas de la evaluación que no son código, en CSV (UTF-8, con cabecera) o YAML. Las leen los instrumentos genéricos
(`src/nucleo/instrumentos/`, ADR-017) y, si la evaluación tiene código propio, los scripts de `src/analisis_<id>/`
(`nucleo.config.insumo("<archivo>")`). El generador deja las cabeceras de las entradas de los instrumentos activos; los
agentes de `/evaluar` las llenan y el usuario las aprueba.

**Reglas:** cada dato lleva su fuente (documento y página); un vacío de información se deja vacío y se consigna como
hallazgo, no se estima; las personas se identifican por su cargo. Las celdas Sí o No se escriben con tilde.

## Catálogo de la política (lector de la política)

| Archivo | Columnas | Valores |
|---|---|---|
| `servicios.csv` | codigo, op, objetivo_prioritario, lineamiento, servicio, responsable | Código como lo numera la política; responsables separados por «; » |
| `lineamientos.csv` | lineamiento, op | Opcional: detecta lineamientos sin servicio |
| `indicadores.csv` | codigo, nivel, op, servicio, nombre, sentido, linea_base, anio_base, responsable, vigente | nivel: Objetivo prioritario o Servicio; sentido: Ascendente, Descendente o vacío; vigente: Sí o No |
| `cadena.csv` | codigo, causa_efecto, tipo_vinculo, estandar, fuente | tipo_vinculo: Causa, Efecto o vacío |
| `documentos.csv` | clave, titulo, ruta, enlace | ruta relativa a la raíz del repositorio; la usa la ubicación de citas |

## Revisión documental (REV_DOC)

| Archivo | Columnas | Valores |
|---|---|---|
| `rev_doc_componentes.csv` | codigo; una columna por componente de `escalas.implementacion` y su `_fuente`; ejecucion_AAAA por año del periodo; ejecucion_fuente; ajuste; sustento_ajuste | Componentes: Sí o No. Ejecución: Sí, No o Sin información. Ajuste: un estado de implementación, solo con sustento |

## Indicadores y recursos (REV_IND)

| Archivo | Columnas | Valores |
|---|---|---|
| `rev_ind_valores.csv` | codigo, anio, meta, valor, fuente, pagina, comparable, nota | Una fila por indicador y año; comparable: Sí o No |
| `rev_ind_fichas.csv` | codigo, atributo, cumple, severidad, observacion | Opcional. atributo: clave de `tipologias/comun/atributos_ficha.yaml`; cumple: Sí, Parcial o No; severidad: Crítica, Mayor o Menor |
| `rev_ind_presupuesto.csv` | organo, anio, pia, pim, devengado, fuente, monto_otra_fuente, otra_fuente | Opcional. Montos en soles |

## Coherencia y alineamiento (REV_ALI)

| Archivo | Columnas | Valores |
|---|---|---|
| `alineamiento.csv` | tipo, momento, op, instrumento, referencia, documento, pagina, cita, correspondencia, sustento | tipo: Vertical, Horizontal, Internacional, Normativa o Conceptual. Correspondencia: Directa, Indirecta o Sin correspondencia (vertical, internacional y normativa); Complementariedad, Duplicidad o Vacío de atención (horizontal); Continuidad, Precisión o Divergencia (conceptual). La cita es literal y se ubica en su documento |

## Revisión de evidencia (REV_EVI)

| Archivo | Columnas | Valores |
|---|---|---|
| `rev_evi_tipos.csv` | tipo, nombre, servicios, causa_efecto, resultado_esperado, pertinencia | Cada servicio en un solo tipo; pertinencia: Directa o Parcial |
| `rev_evi_estudios.csv` | id, tipos, referencia, enlace, verificacion, diseno, ambito, anio, calidad, resultado_principal, direccion, incluido, comparable, motivo_exclusion | diseno: Revisión sistemática, Revisión narrativa, Experimental, Cuasi experimental o Descriptivo; direccion: Favorable, Mixto, Nulo, Desfavorable o Sin estimación |
| `rev_evi_estrategias.yaml` | revisados, estrategias (id, tipos, base, cadena) | Lo ejecuta `src/nucleo/busqueda_academica.py` |
| `rev_evi_busquedas.csv` | id, base, tipos, cadena, filtros, fecha, total, revisados, seleccionados, incluidos, url | Lo escribe la búsqueda; las bases sin interfaz pública se agregan a mano |
| `rev_evi_resultados.csv`, `rev_evi_descriptores.csv` | | Registros revisados para el cribado y descriptores usados |

## Información de las entidades (REP_ENT) y libro de códigos

| Archivo | Columnas | Valores |
|---|---|---|
| `libro_codigos.csv` | codigo, nombre, definicion, incluye, excluye | Común a las entrevistas y a las entidades; lo aprueba el usuario |
| `rep_ent_segmentos.csv` | id_segmento, entidad, servicio, fuente, pagina, texto | Fragmentos literales, sin nombres de personas |
| `rep_ent_codificacion.csv` | id_segmento, codigos, sentido, critico | Primera codificación; sentido: Facilitador, Obstáculo o Neutro |
| `rep_ent_segunda_codificacion.csv` | id_segmento, codigos | La escribe el agente codificador-ciego |
| `rep_ent_segmentos_formulario.csv`, `rev_ind_valores_formulario.csv` | | Los escribe `correr.py <id> --consolidar` desde las respuestas del formulario |

Las entradas de las entrevistas están en `05_entrevistas/LEEME.md`.
