# Formato MIMP Lab — referencia completa

Contenido:
1. Los cinco tipos de diapositiva
2. Cómo elegir acento e ícono por lámina
3. Composiciones recomendadas por tipo de contenido
4. Diferencias deliberadas con `ppt-mimp`
5. Cómo se obtuvieron los activos gráficos

---

## 1. Los cinco tipos de diapositiva

**Portada** (`d.cover`). Fondo `bg_cover.jpg` (textura clara + franja de tres
colores al pie, a sangre), título en `PETROL` 44 pt centrado, subtítulo en
`PETROL2` 20 pt, filete `TEAL`, unidad responsable en gris y fecha en `TEAL`
negrita. Sin número de página.

**Divisor de sección** (`d.divider`). Fondo sólido de un acento (por defecto
`TEAL`, pero acepta `accent=` para variar entre secciones), número en una
variante pálida del mismo acento (46 pt), título en blanco (30 pt), ambos
alineados a la derecha, filete blanco debajo. A diferencia de `ppt-mimp`
(siempre azul), aquí cada sección puede tomar un acento distinto si eso
ayuda a distinguir bloques temáticos (p. ej. `GREEN` para la sección de
resultados, `VIOLET` para la de evidencia) — pero sigue la misma regla de
`ppt-mimp`: si usas un divisor, usa uno por sección, no uno suelto.

**Contenido** (`d.content`). Fondo `bg_content.jpg` (textura sutil), banda,
título en `PETROL`, línea de acento (parámetro `accent=`, `TEAL` por
defecto) y franja de pie de tres colores. Ícono opcional junto al título.
Es el 80 % de un deck.

**Cierre** (`d.closing`). Mismo fondo que la portada, con "¡Gracias!" en
`PETROL` y una línea de pie con unidad y fecha.

**Cronograma** (`d.content` + `d.gantt`). Caso especial de contenido: columna
de ejes a la izquierda (cada eje toma un acento distinto de la paleta),
rejilla de columnas arriba y barras por actividad.

---

## 2. Cómo elegir acento e ícono por lámina

Esta es la decisión de diseño que más distingue a este skill de `ppt-mimp`:
ahí el acento es fijo (rojo) y no hay que decidir nada; aquí el acento
*comunica*. Antes de llamar a `d.content(...)`, pregúntate de qué trata la
lámina y elige de esta tabla (no es exhaustiva, pero cubre los casos más
comunes de un deck del Lab):

| La lámina trata de… | `accent=` | `icon=` |
|---|---|---|
| Presentación del Lab, contexto, pertinencia | `TEAL` | `idea` |
| Objetivos, alcance | `TEAL` o `GOLD` | `objetivo` / `alcance` |
| Metodología, proceso, mecanismo | `GOLD` | `engranaje` |
| Investigación, evidencia, evaluación | `VIOLET` | `investigacion` |
| Datos, indicadores, resultados numéricos | `VIOLET` | `datos` |
| Alianzas, actores, consejo consultivo | `GOLD` | `colaboracion` |
| Riesgos, alertas, pendientes críticos | `RED` | `alerta` |
| Requisitos, cumplimiento, checklist | `TEAL` | `check` |
| Cronograma, plazos | `PETROL` (el `gantt` ya rota internamente) | `calendario` |
| Premiación, reconocimiento, cierre de hito | `GOLD` | `premio` |
| Presupuesto, financiamiento | `GREEN` | `dinero` |
| Resultados finales, balance, cierre de ciclo | `GREEN` | `premio` o `check` |

Si dos láminas consecutivas caen en la misma fila, está bien repetir el
acento — la regla de "no repetir más de dos seguidas" (ver SKILL.md) es
sobre monotonía visual, no una prohibición estricta cuando el contenido
realmente es continuación directa del mismo tema.

---

## 3. Composiciones recomendadas

| Si el contenido es… | Composición |
|---|---|
| Una idea con tres o cuatro puntos | `subhead` + `icon_bullets` a ancho completo (el marcador de color reemplaza la viñeta plana) |
| Varios elementos comparables (proyectos, líneas de trabajo, ejes) | `columns` con 2–4 tarjetas — deja `palette=None` para que cada una tome un acento distinto |
| Un proceso o ciclo | `steps` con 3–5 etapas — mismo criterio, deja que rote la paleta |
| Datos tabulares con categorías (una columna = una línea/eje/entidad) | `table` con `stripe=[...]` — una columna, un acento |
| Datos tabulares sin categorías claras (fila = un caso, sin agrupar por columna) | `table` sin `stripe` (usa el degradado gris por defecto, más neutro) |
| Un antes/después o dos momentos | dos tarjetas de acentos distintos (p. ej. `RED` "antes" → `GREEN` "después") con una flecha (`d.rect` con `MSO_SHAPE.RIGHT_ARROW`) entre ellas |
| Programación temporal | `gantt` (cada eje rota de acento automáticamente) |

Reglas prácticas de altura: idénticas a `ppt-mimp` (`BODY_TOP` 2.15 a
`BODY_BOT` 6.90 → 4.75 in de cuerpo).

---

## 4. Diferencias deliberadas con `ppt-mimp`

Si ya conoces `ppt-mimp`, esto es lo que cambia aquí y por qué:

- **Fondo con textura, no color plano.** `ppt-mimp` usa `EDEDED` sólido en
  contenido; aquí es `bg_content.jpg` (diagonales muy sutiles + trama de
  puntos) para que el deck se sienta menos "plantilla de oficina" y más
  "laboratorio". La textura es intencionalmente tenue — si en un render se
  ve intrusiva o interfiere con la lectura de una tabla, es un bug, no una
  característica: avisa y considera bajar la opacidad o recortar más la
  imagen antes de generar el asset.
- **Franja de pie de tres colores en cada lámina de contenido.** `ppt-mimp`
  no tiene equivalente; es la firma visual tomada del deck original del
  concurso del Lab (`Concurso v10.12.25`).
- **Cuatro acentos rotando (`PALETTE`), no dos — y el rojo no rota.**
  `ppt-mimp` alterna `NAVY`/`NAVY2` en `steps()` y no varía en `columns()`.
  Aquí `columns()`, `steps()` y `gantt()` rotan sobre `PALETTE = [TEAL, GOLD,
  VIOLET, GREEN]`; `RED` queda fuera de esa rotación a propósito (es un
  color de alerta, no un color decorativo) y solo aparece cuando se pasa a
  mano porque el contenido de ese elemento específico es un riesgo. Además
  `divider()` acepta un acento variable.
- **Insignias de ícono junto al título.** No existen en `ppt-mimp`. Cada
  lámina de contenido debería llevar una (ver tabla de la sección 2).
- **Línea de acento más gruesa (0.045 in vs 0.035 in)** y título 27 pt en
  vez de 28 pt, para dejar espacio visual a la insignia sin que compita con
  el logo del ministerio en la banda superior.
- **Lo que NO cambia:** geometría del lienzo, banda superior con los dos
  logos oficiales, tipografía Poppins, tamaños mínimos de fuente (nunca por
  debajo de 8.5 pt), alineación izquierda, `N.°` y comillas tipográficas.
  Este sigue siendo un documento institucional del MIMP — "colorido" no
  significa informal.

---

## 5. Activos gráficos

`assets/` contiene:

| Archivo | Origen | Uso |
|---|---|---|
| `logo_mimp.png` | Igual que `ppt-mimp` | Banda superior, izquierda |
| `logo_gob.png` | Igual que `ppt-mimp` | Banda superior, derecha |
| `bg_cover.jpg` | Recortado de `Concurso v10.12.25` (fondo con franja de tres colores al pie) | Portada y cierre |
| `bg_content.jpg` | Recortado de `Concurso v10.12.25` (misma textura, sin la franja de color, para no competir con la franja de pie que dibuja `_footer_bar`) | Fondo de diapositivas de contenido |

Si se necesita una textura adicional (por ejemplo para un divisor con
imagen de fondo en vez de color sólido), extraer del mismo pptx de origen
con el mismo criterio: recortar cualquier franja de color ya existente en
la imagen para que no choque con los elementos que dibuja la librería.
