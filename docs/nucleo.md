# Núcleo reutilizable (`src/nucleo/`)

El núcleo reúne lo que el piloto de la PNIG dejó probado y que sirve a cualquier evaluación de políticas nacionales y
planes del Sinaplan. No lleva datos de ninguna evaluación: los datos van en `evaluaciones/<id>/` (configuración e
insumos) y el relato de cada evaluación en su capa, `src/analisis_<id>/`. La decisión está en ADR-013.

## Tres capas

| Capa | Ruta | Qué contiene |
|---|---|---|
| Núcleo | `src/nucleo/` | Funciones genéricas, probadas en `tests/` |
| Capa de evaluación | `src/analisis_<id>/` | Análisis, hallazgos e informe de una evaluación; la del piloto es `src/analisis_pnig/` |
| Datos | `evaluaciones/<id>/` | `config.yaml`, `02_insumos/`, datos, análisis, entrevistas, informe y anexo técnico |

Una evaluación nueva se crea con `src/nucleo/nueva_evaluacion.py` desde una tipología (`tipologias/`) y la plantilla
(`plantillas/nueva_evaluacion/`, con su LEEME).

## Módulos

| Módulo | Funciones | Para qué |
|---|---|---|
| `config` | `cargar()`, `Config.ruta()`, `insumo()` | Configuración de la evaluación (variable `EVALUACION`, por defecto `pnig_2026`), rutas de sus carpetas y lectura de `02_insumos/` |
| `entorno` | `herramienta()`, `ruta_externa()` | Rutas de Chrome y LibreOffice, y de los datos fuera del repositorio, desde variables de entorno, `config/entorno.local.yaml` o `config/entorno.yaml` |
| `formato` | `fmt_num`, `soles`, `monto`, `millones`, `pct`, `lista_es`, `comillas`, `comillas_parrafo` | Números con coma decimal, montos en millones de soles, listas en español y comillas “ ” |
| `documento` | `Documento` | Informe en Word sobre `plantillas/word/informe_institucional.docx`: títulos, párrafos con recuadro, tablas y figuras numeradas, notas al pie, índices como campos de Word y guardado con comillas |
| `portada` | `agregar_portada()` | Portada de diseño único (G) desde un `portada.yaml` (`docs/portada.md`) |
| `figuras` | `pagina`, `renderizar`, `leyenda`, `mezcla`, `e` | Figuras en HTML renderizadas a PNG con Chrome sin ventana, con la paleta de estados |
| `encuestas` | `Diseno`, `estimar`, `brecha`, `cambio`, `por_dominio`, `ambito_ubigeo`, `codigo_departamento` | Estimaciones con diseño muestral complejo, brechas por sexo con coeficiente de variación y cambio entre años, por ámbito (departamentos y Lima) |
| `escalas` | `clasificar_implementacion`, `avance`, `nivel` | Estado de implementación de un servicio (parámetros de `config.yaml`, ADR-012 en el piloto) y avance de indicadores frente a su meta |
| `evidencia` | `solidez`, `plausibilidad` | Solidez de la evidencia por servicio y plausibilidad causal |
| `hallazgos` | `evidencia`, `fuentes_independientes`, `verificar`, `hoja_publica`, `libro_publico` | Modelo de hallazgo, reglas de confianza por fuentes de origen independiente (parámetro `origen`, ADR-015), verificación y nombres publicados de libros y hojas |
| `codificacion` | `kappa`, `jaccard`, `acuerdo` | Concordancia de la doble codificación |
| `presupuesto` | `conciliar` | Conciliación de montos entre dos fuentes con tolerancia |
| `fuentes` | `sha256`, `slug`, `a_pdf` | Hash de los documentos, nombres de archivo y conversión a PDF con LibreOffice |
| `anexo` | `Registro`, `libro_publicado`, `publico`, `tabla_xlsx` | Armado del anexo técnico con la lista de archivos y su SHA-256 |
| `auditoria` | `leyendas_sin_analisis`, `estilo`, `terminos`, `archivos_sensibles`, `datos_personales` | Revisión del informe y de los archivos antes de un commit o una entrega |
| `nueva_evaluacion` | `crear()` | Evaluación nueva desde la plantilla y una tipología |

`actualizar_campos_word.ps1` termina un informe con Word por COM: actualiza índices y numeración con su paginación,
incrusta las fuentes y exporta el PDF de revisión.
`comparar_word.ps1` arma, también por COM, la copia con control de cambios de un informe revisado: compara la versión
que se revisó con la versión final y responde cada comentario con el texto de un JSON (ADR-015).
Si la versión revisada trae muchos cambios de formato registrados, `rechazar_cambios.py` rechaza antes, en el XML,
las inserciones y eliminaciones y acepta los cambios de formato; su salida se pasa como `-Revisado`.

## Línea de comandos

```
uv run python src/nucleo/nueva_evaluacion.py <id> --tipologia cierre --titulo "…" --politica "…" --sigla X --inicio 2020 --fin 2025
uv run python src/nucleo/auditoria.py informe <ruta.docx>
uv run python src/nucleo/auditoria.py commit             # lo usa .githooks/pre-commit
uv run python src/nucleo/auditoria.py entrega <carpeta>
```

## Reglas del núcleo

- **Sin datos de una evaluación:** nada de siglas, cifras, servicios ni rutas de un caso. Lo propio de la evaluación va
  en `config.yaml` o en `02_insumos/`.
- **Compatibilidad del piloto:** los módulos que el piloto importaba con su nombre anterior quedan como reexportación
  (`src/analisis_pnig/docx_mimp.py` y `svy.py`), para que sus scripts no cambien.
- **Pruebas:** cada función nueva lleva su prueba en `tests/test_nucleo.py`; cada cambio se verifica con
  `uv run pytest`, que incluye la referencia del piloto (`tests/referencia_pnig/`): libros, CSV, figuras y texto del
  informe tienen que seguir iguales. Si una diferencia es intencional, se revisa y se vuelve a capturar con
  `uv run python tests/referencia.py capturar`.
- **Rutas:** relativas a la raíz del repositorio; las de fuera del repositorio, con `entorno.ruta_externa`.
