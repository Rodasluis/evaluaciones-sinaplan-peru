# Evaluación de políticas nacionales y planes del Sinaplan

Herramientas para evaluar políticas nacionales y planes del Perú con la metodología del CEPLAN (Guía para el
seguimiento y evaluación de políticas nacionales y planes del Sinaplan, 2024), complementada con literatura de
evaluación. Reúne un núcleo reutilizable, las estructuras de los cuatro tipos de evaluación, una plantilla para
empezar una evaluación nueva y un caso completo: la evaluación de cierre de la Política Nacional de Igualdad de Género
(PNIG, 2019–2025), elaborada por la Oficina de Monitoreo y Evaluación de Políticas del Ministerio de la Mujer y
Poblaciones Vulnerables.

## Arquitectura

| Capa | Ruta | Contenido |
|---|---|---|
| Núcleo | `src/nucleo/` | Informe en Word, portada, figuras, formato, escalas, hallazgos y confianza, evidencia, presupuesto, encuestas con diseño muestral, doble codificación, anexo técnico y auditorías (`docs/nucleo.md`) |
| Tipologías | `tipologias/` | Cierre (probada en el piloto), diseño, implementación y resultados: preguntas tipo, instrumentos, escalas, estructura del informe y requisitos del CEPLAN, con la página de la guía |
| Capa de evaluación | `src/analisis_<id>/` | Análisis y relato de cada evaluación; la del piloto es `src/analisis_pnig/` |
| Datos | `evaluaciones/<id>/` | Configuración, insumos del caso, análisis, instrumentos e informe |

La decisión de diseño está en `docs/decisiones/ADR-013_nucleo-y-capa-de-evaluacion.md`.

## Requisitos
- Windows con Microsoft Word (cierre del informe por COM), Google Chrome (figuras) y LibreOffice (conversiones a PDF).
- Python 3.12 o posterior con [uv](https://docs.astral.sh/uv/); las versiones de las bibliotecas están fijas en
  `uv.lock`.
- Tipografía Poppins para la portada y las presentaciones.

## Puesta en marcha
```
uv sync                                     # entorno con las versiones fijas
git config core.hooksPath .githooks         # control previo al commit
uv run pytest                               # pruebas del núcleo, del generador y de la referencia del piloto
uv run python src/fuentes/descargar_inventario.py --carpeta ceplan   # guías del CEPLAN (docs/fuentes/LEEME.md)
```
Las rutas propias del equipo (Chrome, LibreOffice, microdatos, carpeta administrativa) y los datos del especialista de
contacto van en `config/entorno.local.yaml`, con la estructura de `config/entorno.yaml`; ese archivo no se versiona.

## Identidad institucional
El repositorio no trae logos ni encabezados de ninguna institución. Para usar los suyos:
- copie los logos en las rutas de `logos` de `config/portada.yaml` (o del `portada.yaml` de cada evaluación);
- reemplace el encabezado de `plantillas/word/informe_institucional.docx` por el de su órgano, sin tocar los estilos;
- ponga los fondos y logos de las presentaciones en `.claude/skills/presentacion-evaluacion/assets/` (sin ellos, las
  diapositivas usan colores sólidos).

## Crear una evaluación
```
uv run python src/nucleo/nueva_evaluacion.py pnx_2027 --tipologia cierre \
    --titulo "Evaluación de cierre de la Política Nacional de … (2020–2026)" \
    --politica "Política Nacional de …" --sigla PNX --inicio 2020 --fin 2026
uv run python src/analisis_pnx_2027/sintesis.py
uv run python src/analisis_pnx_2027/informe.py --word
```
El flujo completo (encargo, fuentes, insumos, instrumentos, hallazgos, informe y entrega) está en
`plantillas/nueva_evaluacion/LEEME.md`. En Claude Code, el comando `/nueva-evaluacion` hace lo mismo.

## El caso del piloto
`evaluaciones/pnig_2026/` contiene la configuración, los insumos, los libros de análisis, el banco de preguntas de las
entrevistas, el informe de evaluación de cierre, la presentación de resultados (completa y resumida) y la infografía. El estado del informe (versión y
validación) es el que indica el propio documento.
- **Lo que no se publica:** las transcripciones, las guías y los cuestionarios aplicados y los registros de las
  entrevistas, los documentos sin publicación oficial, los microdatos y la nota metodológica de trabajo. La referencia del piloto
  (`tests/referencia_pnig/excluidos_publico.txt`) lista los archivos que no están.
- **Regeneración:** con los microdatos y las fuentes descargados, `/regenerar-informe` en Claude Code o los pasos de
  `.claude/commands/regenerar-informe.md`.

## Claude Code
- **Agentes** (`.claude/agents/`): `codificador-ciego` (segunda codificación sin ver la primera), `revisor-de-informe`
  (requisitos del CEPLAN, estilo y cifras) y `auditor-de-publicacion` (datos personales y documentos no públicos).
- **Comandos** (`.claude/commands/`): `/nueva-evaluacion`, `/regenerar-informe`, `/anexo-tecnico` y `/auditar`.
- **Skill** `presentacion-evaluacion`: presentaciones en estilo institucional o del MIMP Lab, con componentes de
  evaluación.

Las reglas de trabajo y de estilo están en `CLAUDE.md`.

## Licencias
El código va con licencia MIT (`LICENSE`); la documentación, las tipologías, las plantillas y los contenidos del
piloto, con CC BY 4.0 (`LICENSE-docs.md`). Los documentos de las fuentes no se distribuyen.

## Cómo citar
Rodas, L. (2026). *Evaluación de políticas nacionales y planes del Sinaplan* [repositorio de código].
https://github.com/Rodasluis/evaluaciones-sinaplan-peru
