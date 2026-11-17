---
name: presentacion-evaluacion
description: "Presentaciones PowerPoint de las evaluaciones de políticas nacionales y planes de este repositorio (nota metodológica completa o resumida, resultados, reuniones de validación, Alta Dirección, CEPLAN), en dos estilos del MIMP: el institucional (portada roja, divisores azules) y el del MIMP Lab (multicolor, con íconos). Incluye componentes de evaluación: tarjetas de cifra, tabla de valoración de las preguntas con su confianza y estado de los servicios por objetivo prioritario. Úsala siempre que se pida una presentación, deck, diapositivas o pptx de una evaluación del repositorio, o cambios a una presentación de evaluaciones/<id>/00_encargo o 06_informe."
---

# Presentaciones de las evaluaciones

Skill propia del repositorio, basada en las skills `ppt-mimp` (estilo institucional) y `ppt-mimp-lab` (estilo MIMP Lab).
Sus librerías son copias de las de esas skills, con la corrección de los `effectLst` duplicados (ver *Control de
calidad*), más los componentes de evaluación que probó el piloto PNIG.

## Qué estilo usar

| Estilo | Librería | Cuándo |
|---|---|---|
| MIMP Lab | `scripts/mimp_lab_deck.py` | Nota metodológica y resultados de las evaluaciones (decisiones del usuario del 2026-10-29 y del 2026-11-11 para el piloto PNIG) |
| Institucional | `scripts/mimp_deck.py` | Alta Dirección, Direcciones Generales, oficios, rendición de cuentas |

Si el encargo no dice el estilo, se usa el del Lab para la nota y los resultados, y se pregunta en los demás casos.

## Flujo

1. **Contenido antes que maqueta.** Las cifras se leen de las mismas fuentes que el informe (libros de `04_analisis`,
   hallazgos y valoraciones), nunca se copian a mano, para que coincidan. Lo no verificado se marca `[POR VERIFICAR]`.
2. **Un script por presentación** en `src/analisis_<id>/presentacion_*.py`, que importa la librería y los
   componentes (modelos: `src/analisis_pnig/presentacion_resultados.py`, completa, y
   `src/analisis_pnig/presentacion_resumida.py`, de ocho láminas):
   ```python
   import sys
   from pathlib import Path
   sys.path.insert(0, str(Path(__file__).resolve().parents[2] / ".claude/skills/presentacion-evaluacion/scripts"))
   import componentes as K
   from mimp_lab_deck import MimpLabDeck, GOLD, GREEN, RED, TEAL, VIOLET
   ```
3. **Renderizar y mirar** cada lámina (ver *Control de calidad*).
4. **Después de la primera versión, la presentación se edita a mano** sobre el archivo del repositorio: no se
   regenera (CLAUDE.md). Los cambios posteriores se hacen en el .pptx con python-pptx o en su XML, conservando lo que
   el usuario editó.
5. **Versiones:** solo la vigente queda en el repositorio; las anteriores van a la carpeta administrativa.

## Librerías

Las dos comparten geometría (13,333 × 7,5 in; título en y 1,02; cuerpo de `d.BODY_TOP` 2,15 a `d.BODY_BOT` 6,90;
margen `d.M` 0,92 y ancho útil `d.CW` 11,49) y la misma API: `cover`, `content`, `divider`, `closing`, `text`,
`bullets`, `subhead`, `card`, `header_bar`, `columns`, `steps`, `gantt`, `rect`, `keep_slides` y `save`. El estilo Lab
agrega `icon`, `icon_bullets`, `table`, `tint`, un ícono y un acento por lámina (`content(título, icon=, accent=)`).
El detalle de cada estilo está en `references/formato_institucional.md` y `references/formato_lab.md`.

## Componentes de evaluación (`scripts/componentes.py`)

| Componente | Para qué |
|---|---|
| `K.bajada(d, s, texto)` | Oración de lectura bajo el título (qué muestra la lámina y qué concluye), para que la presentación se entienda sin expositor; devuelve el tope del cuerpo |
| `K.tarjeta_cifra(d, s, x, y, w, h, numero, etiqueta, color)` | Cifra grande con su etiqueta (mensajes principales) |
| `K.caja(d, s, x, y, w, h, titulo, items, color)` | Tarjeta con encabezado de color y viñetas |
| `K.resumen_confianza(confianzas)` | «5 (4 sólidos, 1 indicio)» para la columna de hallazgos |
| `K.tabla_valoracion(d, s, filas)` | Pregunta, tema, criterio, valoración y hallazgos con su confianza, con la nota de confianza al pie |
| `K.tabla_estado(d, s, filas)` | Servicios por objetivo prioritario: implementados, parciales, no implementados y con protocolo |

Funcionan en los dos estilos. `scripts/ejemplo_evaluacion.py <carpeta>` arma un deck de muestra de cada estilo con
datos sintéticos.

## Estructuras tipo

- **Resultados** (piloto PNIG, 18 láminas): portada; objeto, alcance y método; mensajes principales; valoración de las
  preguntas; divisor de diseño; una lámina por pregunta o grupo de preguntas de diseño; divisor de implementación y
  resultados; estado de los servicios por objetivo; recursos; coordinación; indicadores; sostenibilidad;
  recomendaciones; próximos pasos; cierre.
- **Nota metodológica completa:** contexto y decisión; objeto y alcance; marco analítico (criterios y preguntas);
  alineamiento; revisión de evidencia; matriz pregunta, instrumento y fuentes; instrumentos; fuentes; plan de
  ejecución en días; productos; riesgos.
- **Nota metodológica resumida:** objeto y alcance, preguntas, instrumentos, plan en días y productos.

## Reglas del repositorio

- **Nombres, no siglas ni identificadores:** los instrumentos y libros se nombran como en el informe («revisión
  documental», no REV_DOC); el rastro técnico va en el anexo (ADR-009 y ADR-010). La pregunta puede citarse por su
  código (D1, I2) en los títulos.
- **Comillas “ ”** en la presentación de resultados (como el informe de cierre); `N.°`, nunca `Nº`.
- **Montos en millones de soles** con coma decimal: «S/ 3,31 millones».
- **Sin guiones para explicar** términos ni incisos: comas, dos puntos, paréntesis u oraciones.
- **Cronogramas en días:** día N y fecha.
- **Cada pregunta con su instrumento y sus fuentes**, de forma explícita.
- **El rojo solo para riesgos y alertas**; un acento por lámina según su tema, sin repetirlo en más de dos láminas
  seguidas (estilo Lab).
- **Logos siempre a color**, sin recolorear. **Poppins** en todo el deck: recordar al usuario incrustarla al guardar
  en Windows.
- **Tamaños fijos:** cuerpo de 10 a 13 pt, nunca menos de 8,5 pt; títulos de una línea; ~120 palabras por lámina
  como máximo.

## Control de calidad, obligatorio antes de entregar

```
uv run python <skill pptx>/scripts/office/validate.py salida.pptx
"C:\Program Files\LibreOffice\program\soffice.exe" --headless --convert-to pdf --outdir <carpeta> salida.pptx
pdftoppm -png -r 80 salida.pdf lamina
```

Mirar cada imagen: desbordes, títulos en dos líneas, columnas descuadradas, acentos repetidos, numeración y logos.

**effectLst duplicados:** en las skills originales, `rect(..., shadow=True)` deja dos `<a:effectLst>` en el mismo
`spPr` y PowerPoint pide reparar el archivo; `validate.py` no lo detecta. Las copias de esta skill lo corrigen en
`_shadow`. Para un deck hecho con las skills originales, contar los `spPr` con dos `effectLst` y quitar el vacío.
