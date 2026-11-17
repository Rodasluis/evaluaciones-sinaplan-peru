# Formato institucional MIMP — referencia completa

Contenido:
1. Los cinco tipos de diapositiva
2. Composiciones recomendadas por tipo de contenido
3. Revisión de un deck existente: qué mirar y en qué orden
4. Errores frecuentes al adaptar decks de otras entidades
5. Cómo se obtuvieron los activos gráficos

---

## 1. Los cinco tipos de diapositiva

**Portada** (`d.cover`). Fondo rojo institucional a sangre, banda blanca con los
logos, título centrado en 54 pt, subtítulo en 22 pt, filete blanco, unidad
responsable y fecha. Sin número de página.

**Divisor de sección** (`d.divider`). Fondo azul, número de sección en azul claro
(46 pt) y título en blanco (30 pt), ambos alineados a la derecha, con filete
debajo. Sirve para marcar los bloques del deck; si se usa uno, se usan para todas
las secciones — un solo divisor suelto descuadra la lectura.

**Contenido** (`d.content`). Fondo `EDEDED`, banda, título azul y línea roja. Es
el 80 % de un deck.

**Cierre** (`d.closing`). Mismo fondo rojo de la portada, con "¡Gracias!" y una
línea de pie con unidad y fecha.

**Cronograma** (`d.content` + `d.gantt`). Caso especial de contenido: columna de
ejes a la izquierda, rejilla de meses arriba y barras por actividad.

---

## 2. Composiciones recomendadas

| Si el contenido es… | Composición |
|---|---|
| Una idea con tres o cuatro puntos | `subhead` + `bullets` a ancho completo |
| Varios elementos comparables (proyectos, servicios, unidades) | `columns` con 2–4 tarjetas de la misma altura |
| Un proceso o ciclo | `steps` con 3–5 etapas; si son más de 5, agrupar |
| Una ficha detallada + su seguimiento | tarjeta grande a la izquierda (≈7.5 in) y tarjeta de "Próximos pasos" a la derecha (≈3.7 in) |
| Un antes/después o dos momentos | dos tarjetas con una flecha (`d.rect` con `MSO_SHAPE.RIGHT_ARROW`) entre ellas |
| Definiciones o marcos normativos | tarjetas apiladas a ancho completo, una por idea, con icono opcional a la izquierda |
| Programación temporal | `gantt` |

Reglas prácticas de altura: con `BODY_TOP` 2.15 y `BODY_BOT` 6.90 hay 4.75 in de
cuerpo. Una línea de texto de 10 pt ocupa ≈0.15 in; una de 12 pt, ≈0.18 in. Una
tarjeta con cabecera de 0.6 in y cinco viñetas de dos líneas necesita ≈2.2 in.

Cuando un bloque tenga etiqueta y contenido en la misma línea, la etiqueta va en
naranja y negrita y el resto en azul normal:

```python
d.text(s, d.M, 2.15, d.CW, 0.5,
       [("Objetivo. ", True, ORANGE),
        ("Crear o fortalecer el sistema de Monitoreo y Aprendizaje…", False)], size=12)
```

---

## 3. Revisión de un deck existente

Cuando el encargo es *revisar* y no *construir*, este es el orden que encuentra
más problemas en menos pasos:

1. **Render y vista**: convertir a PDF, pasar a imágenes y mirar todas las
   láminas. Lo visible pesa más que lo que diga el XML.
2. **Geometría desde el XML**: extraer posición y tamaño de títulos, cuerpos y
   logos por diapositiva. Diferencias de décimas de pulgada entre láminas son el
   defecto más común y no siempre se ven de a una.

   ```python
   import zipfile, re
   z = zipfile.ZipFile("deck.pptx")
   for i in range(1, 30):
       try: x = z.read(f"ppt/slides/slide{i}.xml").decode("utf8")
       except KeyError: break
       for m in re.finditer(r'<p:(sp|pic)>.*?</p:\1>', x, re.S):
           off = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"/><a:ext cx="(\d+)" cy="(\d+)"', m.group(0))
           if off:
               print(i, [round(int(v)/914400, 2) for v in off.groups()])
   ```

3. **Autoajuste**: buscar `<a:normAutofit` en cada slide. Un `fontScale` distinto
   por lámina significa que el cuerpo se ve de un tamaño diferente en cada una;
   `lnSpcReduction` indica interlineado comprimido. Ambos son síntoma de
   sobrecarga de contenido, no de formato.
4. **Fuentes**: contar `typeface="…"` por diapositiva y revisar
   `ppt/theme/theme1.xml`. Si el tema declara Calibri y las láminas usan Poppins,
   todo lo que herede del tema (títulos añadidos a mano) se verá distinto.
5. **Logos**: comparar la saturación media de las imágenes. Una versión en escala
   de grises conviviendo con la versión a color es un error de identidad gráfica.
6. **Ortotipografía**: `Nº` vs `N.°`, comillas rectas vs tipográficas, comillas
   sin cerrar, puntos duplicados.

---

## 4. Errores frecuentes al adaptar decks de otras entidades

- **La paleta de origen se cuela.** Verdes, morados o corporativos ajenos deben
  mapearse a la paleta MIMP: color principal → `NAVY`, acento → `RED` o `ORANGE`,
  rellenos suaves → los tonos pálidos.
- **Los diagramas llegan como imágenes.** Rehacerlos como formas nativas: el
  usuario necesita poder editarlos y una imagen reescalada se ve borrosa en
  proyección.
- **Los logos de terceros.** La plantilla no reserva espacio para logos ajenos.
  Si el convenio exige darles visibilidad, ubicarlos en la portada bajo la línea
  de la unidad o en una franja de créditos al cierre, nunca en la banda superior.
- **Diapositivas que son documentos.** Al migrar suele descubrirse que una lámina
  tenía 300 palabras porque el autor la escribió para leerse. Proponer partirla y
  decir explícitamente qué se recortó.
- **Contenido heredado con errores.** Si al leer el original aparece un dato que
  no calza (una cifra copiada de otra lámina, una fecha vieja), no lo arregles
  inventando: márcalo y avísale al usuario.

---

## 5. Activos gráficos

`assets/` contiene los cuatro archivos de la plantilla, extraídos de un deck
institucional del MIMP y reescalados:

| Archivo | Uso | Colocación |
|---|---|---|
| `bg_red.jpg` | Portada y cierre | 0, 0, 13.333 × 7.5 |
| `bg_blue.jpg` | Divisores de sección | 0, 0, 13.333 × 7.5 |
| `logo_mimp.png` | Banda superior, izquierda | 0.83, 0.19, 2.32 × 0.566 |
| `logo_gob.png` | Banda superior, derecha | 10.423, 0.15, 2.08 × 0.654 |

La banda blanca es un rectángulo redondeado que sobresale por arriba del lienzo
(x −0.3, y −0.5, 13.93 × 1.41), de modo que solo se ven redondeadas sus esquinas
inferiores.
