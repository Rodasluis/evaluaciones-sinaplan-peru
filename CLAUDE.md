# CLAUDE.md

Repositorio para evaluar políticas nacionales y planes del Perú según la metodología del CEPLAN (Sinaplan), complementada con literatura de evaluación.

- **Piloto:** «Evaluación de cierre de la Política Nacional de Igualdad de Género (2019–2025)» (PNIG), en `evaluaciones/pnig_2026/`.
- **Evaluación nueva:** `/nueva-evaluacion` (con las preguntas iniciales sobre entrevistas, formulario, documentos internos, validación y personalización) y `/evaluar <id>`, que recorre las etapas con los agentes y se detiene en cada aprobación (`docs/agentes.md`, ADR-017).

## Reglas de trabajo
1. **Fuentes:** no inventar fuentes, cifras ni citas. Lo que no esté verificado en un documento propio o en un portal oficial se marca `[POR VERIFICAR]`. Las citas textuales se copian del documento, con página.
2. **Planos separados:** distinguir siempre lo que exige el CEPLAN, lo que aporta la literatura y lo que decide el equipo evaluador.
3. **Vigencia:** antes de citar una norma o guía, verificar que esté vigente y con qué modificatorias.
4. **Resultados:** no fabricar resultados. Los vacíos de información se consignan como hallazgo, no se estiman.
5. **Aprobación:** los cambios de fase, de diseño metodológico o de destinatarios requieren la aprobación del equipo evaluador. Las acciones externas (enviar, publicar, desplegar, borrar) se confirman antes de ejecutarse.
6. **Datos sensibles fuera del repositorio:**
   - CV, órdenes de servicio, recibos, contactos y puntos focales;
   - credenciales y documentos con contraseñas;
   - borradores administrativos y versiones anteriores de la nota.

   Todo eso va en una carpeta fuera del repositorio (`carpeta_administrativa` en `config/entorno.local.yaml`), y se trata conforme a la Ley N.° 29733, Ley de Protección de Datos Personales. Los datos del especialista de contacto de los documentos van en la sección `contacto` del mismo archivo.
7. **Registro de movimientos:** todo movimiento o descarga de documentos se registra con su hash en `docs/fuentes/`.

## Estilo de los documentos
- Español técnico, sin relleno.
- No usar guiones (« - », «—») para explicar términos, introducir incisos o definir. Se usan comas, dos puntos, paréntesis u oraciones completas.
- Los cronogramas van en días: día N y fecha.
- Cada pregunta de evaluación se vincula de forma explícita con su instrumento y sus fuentes.
- El informe no usa preguntas retóricas ni describe las figuras por su disposición; usa comillas “ ” y montos en millones de soles con coma decimal.
- **Memorándums:** referencias a), b), c) citadas en el texto y archivadas junto al memorándum con el nombre exacto que llevan en él.
- **Formato institucional:**
  - presentaciones con la skill del repositorio `presentacion-evaluacion`, en estilo institucional o del MIMP Lab;
  - portadas de informes con el módulo `src/nucleo/portada`, de diseño único (centrado): textos, logos y colores en `evaluaciones/<id>/portada.yaml` (`docs/portada.md`).

## Estructura
| Ruta | Contenido |
|---|---|
| `src/nucleo/` | Núcleo reutilizable sin datos de ninguna evaluación, con los instrumentos genéricos (`instrumentos/`), la búsqueda académica, el avance por etapas y las observaciones de la validación (`docs/nucleo.md`) |
| `tipologias/` | Estructura de cada tipo de evaluación: cierre (probada en el piloto), diseño, implementación y resultados |
| `plantillas/` | Plantilla de Word del informe y plantilla de evaluación nueva (`nueva_evaluacion/`, con su LEEME) |
| `evaluaciones/<id>/` | Configuración (`config.yaml`), insumos del caso (`02_insumos/`), datos (`03_datos/`, no se versiona), análisis, entrevistas, informe y anexo técnico |
| `src/analisis_pnig/` y `src/analisis_<id>/` | Capa de cada evaluación: sus análisis y su relato sobre el núcleo |
| `src/fuentes/` | Descargas y extracciones de fuentes, e inventario (`descargar_inventario.py`) |
| `docs/fuentes/` | Inventario público de las fuentes con enlace oficial y SHA-256 (`LEEME.md`) |
| `docs/marco/` y `docs/decisiones/` | Matriz de requisitos del CEPLAN y notas de decisión (ADR) |
| `tests/` | Pruebas (`uv run pytest`) y referencia del piloto (`tests/referencia_pnig/`) |
| `.claude/` | Agentes por etapa y de apoyo, comandos (`/nueva-evaluacion`, `/evaluar` y otros) y skill de presentaciones (`docs/agentes.md`) |

- **Códigos de los instrumentos (ADR-009):** REV_DOC (revisión documental), ENTREV (entrevistas), REP_ENT (reportes de las entidades), REV_ALI (coherencia y alineamiento), REV_IND (indicadores y recursos) y REV_EVI (revisión de evidencia). El informe y la presentación de resultados usan los nombres, no las siglas; el rastro técnico queda en el anexo técnico (ADR-010).
- **Informe (ADR-010):** autocontenido; cada hallazgo lleva su respaldo en una nota al pie y los pies dicen «Fuente: … Elaboración propia.».
- **Nombres de archivo:** `AAAA_organismo_titulo-corto.ext`.
- **Versiones de la nota y la presentación:** solo la vigente queda en el repositorio; las anteriores, en la carpeta administrativa.
- **Edición de las presentaciones:** después de la primera versión se editan a mano; los cambios se hacen sobre el archivo, no regenerándolas. La infografía se regenera con su script después del informe.

## Entorno técnico (Windows)
- **Python:** proyecto uv (`pyproject.toml` y `uv.lock`, versiones fijas): `uv run python …` desde la raíz. Con pandas 3, `astype(str)` conserva los nulos; usar `map(str)`.
- **Pruebas y control previo al commit:** `uv run pytest` antes de cada commit que toque código; el gancho `.githooks/pre-commit` (se activa con `git config core.hooksPath .githooks`) bloquea las categorías de la regla 6, las rutas con el nombre de usuario y los celulares escritos en el código.
- **Rutas y datos propios del equipo:** `nucleo.entorno.ruta_externa` y `nucleo.entorno.contacto`, con los valores en `config/entorno.local.yaml`; nunca escritos en el código.
- **Consola:** fijar `PYTHONUTF8=1` y `PYTHONIOENCODING=utf-8`.
- **Extracción de PDF:** `pdftotext -enc UTF-8 -layout`. Los escaneos sin capa de texto requieren OCR.
- **Conversión a PDF:** LibreOffice, llamando a `soffice.exe` directamente.
- **Informe:** se termina con Word (`src/nucleo/actualizar_campos_word.ps1`, por COM), que actualiza índices y numeración con su paginación y exporta el PDF de revisión.
- **Borrados:** usar rutas literales en `rm`, nunca variables sin `${VAR:?}`.
- **Tipografía:** Poppins, instalada en el equipo; el script de Word la incrusta en el informe.
