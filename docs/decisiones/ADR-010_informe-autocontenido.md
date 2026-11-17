# ADR-010. Informe de cierre autocontenido y renumeración de hallazgos

**Fecha:** 2026-11-08 · **Estado:** aprobada por el usuario (plan de la ronda 25, 2026-11-07)

## Contexto
- **Pedido del usuario:** instrucciones de mejora para que el informe de cierre sea autocontenido y autoentendible. Un funcionario que solo tenga el documento publicado debe entender qué se hizo, qué se encontró, con qué respaldo y qué se recomienda, sin la nota metodológica, el repositorio ni el equipo evaluador.
- **Reglas del pedido:**
  - la explicación va en el cuerpo y el detalle técnico, en anexos;
  - no se inventan datos.
- **Qué revierte:** parte de la actualización de ADR-009 de la ronda 24, que dejaba en el informe un glosario de siglas y una columna de rastro con identificadores.

## Decisiones del usuario (2026-11-07)
1. **Libros Excel:** el informe los cita con nombres legibles de libro y de hoja, por ejemplo Brechas_situacion_futura.xlsx, hoja «Situación futura deseada».
   - El repositorio conserva los nombres con sigla (ADR-009).
   - El anexo técnico publicará copias con los nombres legibles; la correspondencia está en `06_informe/correspondencia_libros.csv`.
   - Las reglas viven en `sintesis_hallazgos.LIBROS_PUBLICOS` y `hoja_publica()`.
2. **Hallazgos:** se renumeran por orden de aparición en el informe.
3. **I3-H3:** se mantiene «sólido» en lo que respaldan las cuatro fuentes. «La conducción no contó con los recursos necesarios» se atribuye a la DGIGND como lectura de una sola fuente, en el enunciado, el resumen ejecutivo, la valoración de I3 y el juicio de desempeño.
4. **Validación:** D3-H2 e I4-H5 (antes I4-H4) quedan validados el 7 de noviembre de 2026; la ficha técnica lo consigna.

## Renumeración de hallazgos
| Pregunta | Anterior → nuevo |
|---|---|
| D1 | D1-H3 → D1-H2; D1-H4 → D1-H3; D1-H2 → D1-H4 |
| D2 | D2-H3 → D2-H1; D2-H1 → D2-H2; D2-H2 → D2-H3 |
| I2 | I2-H4 → I2-H3; I2-H3 → I2-H4 |
| I4 | I4-H5 → I4-H1; I4-H1 → I4-H2; I4-H2 → I4-H3; I4-H3 → I4-H4; I4-H4 → I4-H5 |

- **Script:** `src/mantenimiento/adr010_renumerar_hallazgos.py`, con reemplazo simultáneo; se corrió una vez.
- **Alcance:**
  - los scripts que definen o citan hallazgos;
  - la codificación y las deducciones de las entrevistas;
  - la valoración de las recomendaciones.
- **Regeneradas:** las salidas (catálogo de preguntas, registro codificado, guías de entrevista, matriz de triangulación e informe).
- **Registro histórico:** los ADR-006 a 009 y la carpeta administrativa conservan los números anteriores.

## Regla de independencia
- **Regla:** la entrevista a la OGMEPGD no cuenta como fuente independiente, porque la OMEP, que conduce la evaluación, es parte de ese órgano.
- **Conteo:** las fuentes se cuentan por documento (`sintesis_hallazgos.PARTES`); así, por ejemplo, «formulario de seguimiento e informes LIO» son dos documentos y la PNIG con su Anexo 5 es uno.
- **Efecto:** ningún hallazgo cambia de nivel de confianza. El informe la declara en el numeral 2.7.

## Informe
- **Preliminares:**
  - portada;
  - ficha técnica;
  - índice general, de tablas y de figuras, como campos de Word (TOC y SEQ);
  - abreviaturas y siglas.
- **Paginación de los índices:** se escribe en caché con tres vueltas (títulos, índices provisionales y páginas leídas del PDF de LibreOffice), y `updateFields` hace que Word los actualice.
- **Presentación:** con guía de lectura y la definición de «nivel conductor» y «equipo evaluador».
- **Metodología completa (sección 2), sin mencionar la nota metodológica:**
  - tipo y enfoque;
  - criterios y texto de las doce preguntas;
  - procedimiento de cada instrumento;
  - reglas de clasificación en tablas: implementación, nivel de avance con su fórmula, confianza, valoración, escalas de diseño y evidencia;
  - triangulación;
  - fases en días;
  - independencia;
  - limitaciones;
  - disponibilidad de los datos;
  - correspondencia con la Tabla A7.1.
- **Secciones 3 a 5:**
  - cada numeral abre con su pregunta y cierra con «Respuesta a la pregunta»;
  - cada tabla y figura se introduce y se interpreta;
  - los servicios, los indicadores y las siglas llevan su nombre en la primera mención (`Nombrador` en `informe_cierre.py`).
- **Hallazgos:**
  - el enunciado queda en el texto;
  - la nota al pie lleva el identificador, la confianza y el respaldo, en 50 palabras como máximo;
  - las salvedades de fondo van en el cuerpo y las técnicas, en el Anexo 2;
  - los campos nuevos en la síntesis son `detalle`, `salvedad` y `nota_tabla`.
- **Numeral 3.4:** marco por objetivo prioritario y tipo de intervención (`informe_evidencia.py`), con mecanismo, estudios, consistencia, aplicabilidad y conclusión.
- **Pies:** «Fuente: <fuente pública con su cita>. Elaboración propia.», con «Nota:» cuando hace falta. Las fuentes de las figuras están en `figuras_informe.FUENTE_PUBLICA`.
- **Referencias:** APA 7 en orden alfabético, con las 172 fuentes de la revisión de evidencia. Las 23 referencias que decían «et al.» se completaron con Crossref, PubMed, OpenAlex o doi.org:
  - script: `src/fuentes/completar_autores_evidencia.py`;
  - registro: `docs/fuentes/descargas_autores-evidencia_2026-11-08.tsv`;
  - aplicación: en `rev_evi_insumos.py`, sin tocar los literales verificados el 2026-10-31.
- **Anexos:**
  1. matriz de evaluación;
  2. hallazgos con detalle, confianza con color y texto, fuentes independientes y respaldo legible;
  3. estado de cada servicio con nombre, entidad completa, reporte de indicadores y justificación con el monto que coincide con el MEF;
  4. evidencia por tipo con la regla de solidez, el motivo de cada clasificación y todos los estudios;
  5. reporte de seguimiento;
  6. brechas por departamento, con un párrafo por grupo de mapas;
  7. recomendaciones anteriores, agrupadas por evaluación;
  8. entrevistas;
  9. catálogo de servicios e indicadores (nuevo);
  10. documentos de respaldo (nuevo);
  11. glosario (nuevo).
- **Rastro técnico:** `06_informe/rastro_hallazgos.csv` (nuevo) y `rastro_tablas_figuras.csv`, para el anexo técnico.

## Cifras corregidas
| Caso | Corrección |
|---|---|
| Indicadores de servicio 2025 | I2-H4 enumera todas las categorías (19 + 7 + 7 + 4 + 2 + 2 + 8 = 49) y dice que 40 tienen valor, como el Informe N.° D000039-2026-MIMP-DPIGND |
| Enunciados de la situación futura | Son 7; el SFD4 tiene dos metas |
| I4-H2 | La nota dice «informes LIO 2024 y 2025» |
| Articulación de los servicios | Los porcentajes de articulación (38,7 % y 30,6 %) se atribuyen al informe LIO 2025 con sus conteos (19 y 15), y se indica que el informe no precisa el denominador |
| Nivel de avance | El numeral 5.2 explica que no considera la línea de base, con los casos IOP.02.01 e IOP.04.03 |
| IOP.05.02 | No se estima, porque la serie de la ENAHO no reproduce los valores reportados |
| Trabajo no remunerado | La comparación indicativa aparece en el resumen, la síntesis y las conclusiones |

## Verificación
- **Calidad del documento:**
  - `validate.py` aprueba el informe;
  - el barrido final (`informe_cierre.barrido`) no encuentra siglas de instrumentos ni identificadores internos;
  - se revisó el renderizado.
- **Hallazgos:**
  - las 43 notas al pie tienen 36 palabras o menos;
  - los hallazgos aparecen en orden dentro de cada pregunta;
  - frente a la línea base, solo cambian los enunciados previstos (D2-H3, I2-H4, I3-H3 e I4-H4);
  - no se pierde ninguna referencia de rastro y ninguna confianza cambia.
- **Citas:** toda cita autor y año del texto tiene su referencia. Las que el comprobador no empareja son citas dentro de textos literales, rótulos o apellidos compuestos.

## Actualización (2026-11-08, ronda 26): formato y portada
- **Problema reportado por el usuario:** al abrir el informe en Word aparecía el aviso de actualización de campos; si lo aceptaba, el índice quedaba con «1» en todas las páginas y los números de tabla y figura cambiaban de formato.
- **Causas:**
  - el ajuste `updateFields` hacía que Word actualizara los índices antes de paginar;
  - los índices regenerados tomaban estilos que la plantilla no definía;
  - los campos SEQ simples perdían el formato directo;
  - además, la subclase del informe escribía 11 pt y negro como formato directo en todos los párrafos.
- **Decisión:** el informe se termina con Word. `src/analisis_pnig/actualizar_campos_word.ps1` lo abre de forma invisible por COM, repagina, actualiza los campos SEQ, el índice general y los índices de tablas y figuras, guarda el archivo y exporta un PDF para la revisión.
  - Ya no se usa `updateFields`, así que el documento abre sin aviso.
  - Si Word no responde, quedan en caché las páginas leídas del PDF de LibreOffice.
- **Formato** (`docx_mimp.Documento(informe=True)`):
  - **Estilos en lugar de formato directo:**
    - Normal: Arial 11 pt, justificado, 1,15;
    - Título 1: 16 pt azul, empieza página;
    - Títulos 2 y 3;
    - Descripción para las leyendas, con campos SEQ complejos y `\* MERGEFORMAT`;
    - «Fuente MIMP» y «Texto de tabla MIMP»;
    - TDC 1, TDC 2, Tabla de ilustraciones y Título de TDC;
    - nota al pie, encabezado y pie de página.
  - **Secciones:** portada sin encabezado ni pie, preliminares en romanos y cuerpo desde la página 1.
  - **Letra mínima de las tablas:** 7,5 pt; los anchos de columna se ajustaron para que las palabras de las cabeceras no se partan.
  - **Recuadros:**
    - pregunta de evaluación en azul claro;
    - «Respuesta a la pregunta» en gris con borde azul;
    - hallazgos con filete azul a la izquierda.
  - **Figuras:** 15,9 cm como ancho máximo.
- **Portada:**
  - fondo a sangre de `src/analisis_pnig/portada_informe.py`: SVG renderizado con Chrome, con bloque azul institucional, dos círculos iguales que se cruzan y filete rojo;
  - logos del MIMP y del Gobierno del Perú, copiados de los recursos de la skill `ppt-mimp` a `06_informe/figuras/`;
  - textos editables en tablas flotantes ancladas a la página.
- **Verificación:**
  - en el PDF de Word, el índice muestra números de página correctos (romanos en los preliminares, arábigos en el cuerpo);
  - las leyendas conservan el formato;
  - `validate.py` aprueba el archivo guardado por Word;
  - el barrido y las comprobaciones de formato (`revisar_formato`) dan cero;
  - las guías de entrevista no cambian.

## Actualización (2026-11-08, ronda 27): portada y diagramas
- **Pedido del usuario:**
  - quitar de los diagramas las barras laterales más oscuras de las cajas, que delataban el HTML, y cambiar los colores muy oscuros por tonos más suaves;
  - armar la portada con un módulo de tres diseños, según instrucciones propias: paleta roja y gris sin azul, todo leído de una configuración, textos reales de Word y elección del usuario.
- **Alcance de la paleta sin azul:** solo la portada (decisión del usuario). El resto del informe conserva su gama.
- **Portada:**
  - módulo `src/portada` con los diseños A (institucional clásico), B (banda lateral) y C (minimalista);
  - datos en `config/portada.yaml` y logos oficiales en `assets/logos/`; instrucciones en `docs/portada.md`;
  - el usuario eligió la opción A sobre la hoja comparativa con los datos reales (`06_informe/portada/previsualizacion.png`);
  - la portada no lleva versión ni estado; esos datos quedan en la ficha técnica;
  - tipografía Poppins: a pedido del usuario se instaló para su cuenta (18 estilos del repositorio oficial de Google Fonts, licencia OFL, registro en `docs/fuentes/descargas_tipografia-poppins_2026-11-08.tsv`), y el usuario decidió incrustarla completa en el .docx al terminarlo con Word; el rojo `#E1251B` queda por confirmar con el manual de identidad;
  - contraste del blanco sobre el rojo: 4,69:1.
- **Reemplaza** el fondo de la ronda 26 (`portada_informe.py`, SVG con bloque azul y círculos), que se borró junto con las copias de los logos en `06_informe/figuras/` (registro en `docs/fuentes/movimientos_logos-portada_2026-11-08.tsv`).
- **Diagramas** (`figuras_informe.py`):
  - sin barras laterales ni borde superior de color; la categoría se lee por el relleno de la caja;
  - los rellenos oscuros con texto blanco pasan a rellenos claros con texto oscuro del mismo tono, y las barras de datos, a tonos medios;
  - los puntos de confianza de la Figura 1 conservan los colores de las tablas;
  - se corrigió la cuarta línea que asomaba en las tarjetas del mapa de servicios.

## Actualización (2026-11-09, ronda 28): más modelos de portada
- **Pedido del usuario:**
  - más modelos para elegir;
  - la portada de la ronda 26;
  - una como la institucional pero con toda la hoja roja;
  - propuestas basadas en las portadas de las evaluaciones de 2023 del CEPLAN y del MIMP;
  - en A, B y C, letra un 40 % más grande y logos un 20 % más pequeños, en los extremos.
- **Modelos:** siete en `src/portada`: A, B y C ajustados; D (bloque azul de la ronda 26); E (institucional en rojo pleno); F (técnica con recuadro, como el CEPLAN 2023) y G (centrada, como el MIMP 2023). Hoja comparativa y un PDF por modelo en `06_informe/portada/`.
- **Decisiones del usuario:**
  - en E, los logos van sobre placas blancas pequeñas;
  - D reproduce la ronda 26 con los textos de la configuración, sin «Versión 3» ni «Informe preliminar para validación». Es la única excepción a la regla sin azul (y a la de no usar degradados, adornos ni imagen de fondo);
  - los tamaños nuevos superan los rangos de las instrucciones originales.
- **Recuperación:** el código de D se recuperó del historial de la sesión, porque `portada_informe.py` se había borrado en la ronda 27 sin estar en git.
- **Técnica:** E es una tabla anclada de una sola fila con la composición anidada, porque Word partía la tabla de varias filas y pasaba el pie a otra página.

## Actualización (2026-11-09, ronda 29): siglas y portada E
- **Siglas:** CV e IC salen de la lista (pedido del usuario).
  - Las notas de las figuras de brechas y la leyenda de los mapas dicen «coeficiente de variación» (`d1_d2_departamental.py`, regenerado; solo cambiaron esos textos).
  - «IC 95 %» queda como notación estándar en la revisión de evidencia, igual que OR, RR y DME.
- **Portada E:**
  - es igual a A, con el rojo hasta el borde superior;
  - los logos van sin fondo blanco, con el de la institución derivado por `src/portada/logo_sin_fondo.py`; registro en `docs/fuentes/movimientos_logos-portada_2026-11-09.tsv`;
  - lleva una marca de agua con el símbolo de igualdad de género (dibujo propio) sobre un fondo renderizado como imagen. Es una excepción pedida por el usuario.

## Actualización (2026-11-09, ronda 30): colores originales en los diagramas
- **Pedido del usuario:** devolver a los diagramas los colores de la primera versión, sin devolver las barras laterales de color.
- **Diagramas** (`figuras_informe.py`):
  - vuelven los colores de la ronda 22: rellenos plenos de los OP, del azul institucional y de las escalas, con texto blanco sobre los oscuros. Deja sin efecto el suavizado de la ronda 27;
  - las barras siguen fuera. En el problema público y la cadena de valor, causas y efectos se distinguen por el fondo; en el mapa de servicios, el estado va en un punto junto al código; en la línea de tiempo, la categoría se lee por el punto del eje.

## Actualización (2026-11-09, ronda 31): cuatro modelos de portada a partir de un HTML del usuario
- **Pedido del usuario:** revisar los estilos de cuatro portadas de un HTML propio (minimalismo editorial, fotografía documental, glassmorphism y formas planas) e incluirlos entre los modelos.
- **Decisiones del usuario:**
  - paleta: rojo y gris de la configuración, no el morado y dorado del HTML;
  - foto: una institucional suya, no la foto de banco del HTML; mientras falte, solo el velo y un aviso;
  - tipografía: Georgia por Playfair Display y Poppins por Montserrat e Inter; no se instala nada;
  - logos oficiales en lugar del texto del encabezado, sobre placas blancas en los fondos oscuros.
- **Modelos nuevos** en `src/portada`: H (editorial geométrica), I (fotográfica con velo), J (degradado con tarjeta) y K (formas planas). Las medidas siguen al HTML escalado a la A4; el fondo es una imagen renderizada con Chrome y todo el texto es texto de Word (excepciones registradas en `docs/portada.md`).
- **Contraste:** medido palabra por palabra sobre el fondo renderizado. El pie de K pasó a blanco (el gris claro y el rojo claro del HTML quedaban en 3,6:1 y 2,0:1 sobre el resplandor rojo) y el degradado de J no se aclara más allá del primario.
- **Configuración:** claves nuevas `titulo_resaltado`, `foto` y `tipografia_serif`; `logos_sobre_rojo` pasa a `logos_sin_fondo` (la usan E y H). El informe sigue con la opción A hasta que el usuario elija.

## Actualización (2026-11-09, ronda 32): portada D en tonos rojos
- **Pedido del usuario:** pasar la portada D (la de la ronda 26) a tonos rojos, la línea inferior al plomo institucional, la letra un 50 % más grande y el nombre de la institución al rojo institucional en lugar del azul.
- **Cambios** (`src/portada/opcion_d.py`):
  - bloque con degradado del primario a un rojo más oscuro y trama en rojos claros, todo derivado de la configuración;
  - filete bajo el bloque y filete corto junto a la etiqueta en `gris_oscuro`;
  - etiqueta y periodo en rojo muy claro, en lugar del celeste;
  - letra ×1,5; el nombre de la institución va a 14 pt para que el rojo sobre el gris cuente como texto grande;
  - el texto del bloque pasa a 16,4 cm de ancho para que la etiqueta quepa en una línea, y la franja inferior sube si hace falta.
- **Consecuencia:** D ya no lleva azul (`PERMITE_AZUL` vacío) y pasa a llamarse «bloque rojo (ronda 26)».
- **Contraste medido** palabra por palabra: todo el texto cumple el mínimo de su tamaño (3:1 para el grande y 4,5:1 para el normal).

## Actualización (2026-11-10): portada elegida
- El usuario eligió la portada G (centrada) para el informe de cierre (`config/portada.yaml`, `opcion: "G"`). El informe se regeneró con ella: 198 páginas, validación de Office aprobada, solo Poppins incrustada.

## Actualización (2026-11-10, ronda 33): reestructuración del informe y narrativa de figuras y tablas
- **Pedido del usuario:**
  - quitar el subtítulo de la portada, la ficha técnica, la presentación y la guía de lectura;
  - resumen ejecutivo como *abstract* extenso, solo texto, de dos páginas como máximo;
  - quitar los numerales de fases y cronograma, independencia, disponibilidad de datos y correspondencia con la Tabla A7.1; la independencia pasa a las limitaciones;
  - escala de implementación sobre 2019–2025 (ADR-011);
  - cada figura y tabla con descripción y análisis contados como un relato;
  - instrumentos del numeral 1.8 descritos; revisión de literatura de lectura continua;
  - no presentar al «equipo evaluador» como actor, sino explicar los criterios; «selección de documentos» en lugar de «cribado».
- **Figuras:**
  - línea de tiempo solo con hitos normativos y de implementación; la Ley N.° 32535 es hito normativo;
  - título nuevo para la figura de indicadores;
  - azul medio (#4A6698) en las figuras de indicadores y de actores;
  - matriz de preguntas e instrumentos con «con evidencia» y «sin evidencia»;
  - figura de valoración trasladada a la metodología (numeral 2.5);
  - las escalas de implementación y de avance pasan a diagramas.
- **Cuerpo:**
  - la Figura 3 se explica con la matriz de consistencia (Anexo 5 de la PNIG): 27 servicios se vinculan solo con una causa, 14 solo con un efecto, 9 con ambos y 2 sin vínculo consignado;
  - los criterios de correspondencia de la situación futura se explican en el numeral 1.3.

## Actualización (2026-11-10, ronda 34): figuras, nuevas fuentes y términos
- **Pedido del usuario:**
  - «información detallada» en lugar de «sección sustantiva» y «documentación» en lugar de «acervo»;
  - matriz de preguntas e instrumentos en celeste; pregunta 2 de la escala de implementación reformulada;
  - figura de valoración con las pastillas sin color y los puntos de confianza en verde, ámbar y rojo, también en las marcas de las tablas (decisión del usuario);
  - marco conceptual en dos figuras y mapas de servicios en seis, uno por objetivo prioritario, cada uno con su entrada y su análisis;
  - gasto de la conducción en tonos más claros;
  - figura y tabla de indicadores de OP en 2025 con una sola columna, «Nivel 2025»: el oficial o, si falta, el calculado en esta evaluación con asterisco y nota;
  - incorporar la ejecución presupuestal de la actividad 5001072 remitida por la DGIGND y la ayuda memoria de la CONAIG.
- **Fuentes nuevas** (copias de la carpeta administrativa, `movimientos_dgignd-conaig_2026-11-10.tsv`):
  - presupuesto: `rev_ind_3_presupuesto_dgignd.py` concilia la ejecución por finalidad con los datos abiertos del MEF (30 de 30 combinaciones de órgano y año coinciden al sol) y calcula la composición del gasto de la conducción (entre 76 % y 93 % en personal, CAS y locación de 2021 a 2025);
  - CONAIG: figura «Sesiones y acuerdos de la CONAIG, 2021–2026» en el numeral 4.4; el acuerdo de culminar los protocolos se repite cada año de 2021 a 2025.
- **Hallazgos con respaldo nuevo:** I3-H1 y I3-H2 suman la fuente de la DGIGND (I3-H2 pasa de indicio a sólido) e I4-H2 suma la ayuda memoria. Las cifras de los hallazgos no cambian.
- **Referencias:** MIMP, DGIGND (2026a) para la ejecución presupuestal y (2026b) para el marco conceptual, por orden de título (APA 7); MIMP y CONAIG (2026) para la ayuda memoria.
- **Correcciones encontradas al revisar:**
  - la CONAIG se llama «Comisión Nacional de Alto Nivel para la Igualdad de Género», según el D.S. N.° 018-2021-MIMP; el informe usaba otra denominación;
  - `pl()` repetía el número («1 1 no implementado»): ahora devuelve solo la palabra concordada;
  - el glosario remitía al numeral 2.7, que ya no existe, y conservaba la entrada «Equipo evaluador»;
  - la tabla de evidencia por tipo de intervención no tenía párrafo de análisis propio.
- **Nota metodológica y presentaciones:** se editaron directamente (escala 2019–2025, nuevas fuentes y términos); la versión previa está en `notas_metodologicas_anteriores\23_2026-11-10_previa-escala-2019-2025\` (`reorganizacion_2026-11-10.tsv`).

## Actualización (2026-11-11, ronda 35): redacción profesional, comillas, montos y portada única
- **Pedido del usuario:**
  - redacción de informe técnico, sin preguntas retóricas ni descripciones de cómo se lee cada figura;
  - nombres completos de los indicadores;
  - comillas “ ” en lugar de « » en el informe y en la presentación de resultados;
  - montos en millones de soles, con coma decimal, en lugar de miles;
  - sin el informe de la OGAJ, que no llegará;
  - un solo modelo de portada en el repositorio;
  - presentación de resultados con ppt-mimp-lab.
- **Redacción:**
  - se reescribieron unos 40 párrafos del cuerpo, poniendo primero el contenido y no la disposición de la figura; las claves de lectura indispensables quedan en una oración breve;
  - los mapas usan el nombre completo del indicador, tomado de las definiciones de los scripts `d1_brechas_*.py`;
  - el mapa de tolerancia social, que tiene una sola medición departamental (2024), ya no muestra «nan».
- **Comillas:** `docx_mimp.comillas()` convierte « » en “ ” al guardar el informe, en el cuerpo, los encabezados, los pies y las notas al pie; dentro de una cita usa ‘ ’. `figuras_informe.pagina()` hace lo mismo en las figuras. La nota y sus presentaciones conservan « ».
- **Montos:** `docx_mimp.monto()` (por ejemplo, «3,31 millones de soles») y `millones()` para las tablas. La Tabla 17 y la figura del gasto pasan a millones de soles con dos decimales.
- **Escala:** implementado con al menos dos de los tres componentes (ADR-012).
- **Portada:** quedan solo el diseño G y sus auxiliares. Se borraron los modelos A a F y H a K, la hoja comparativa, el logo derivado sin fondo (registro en `reorganizacion_2026-11-11.tsv`) y las previsualizaciones. Se borraron también las figuras que el informe ya no usaba.
- **OGAJ:** se quitó de la nota (párrafo de la posición institucional, fuente de transición, riesgo y referencia), de las dos presentaciones y de los pendientes de la triangulación. Se conservan las exclusiones de alcance que atribuyen el análisis jurídico a la OGAJ.
- **Presentación de resultados:** `src/analisis_pnig/presentacion_resultados.py` genera `06_informe/2026_mimp-ogmepgd_presentacion-resultados-evaluacion-cierre-pnig.pptx` (18 láminas) con las cifras del informe. Desde ahora se edita sobre el archivo.

## Actualización (2026-11-11, ronda 36): títulos, recomendación 9 y anexo técnico
- **Títulos:** Título 1 y título de los índices en 14 pt, Título 2 en 12 pt (pedido del usuario); el Título 3 queda en 11 pt.
- **Recomendación 9:** «Preservar en la nueva política, con sus protocolos, los servicios implementados» lleva una nota al pie con los 24 servicios, en lugar de listarlos en la celda. `docx_mimp.tabla()` acepta piezas con nota al pie (`{"pie": …}`). La matriz pasó a ocupar las páginas 97 a 99.
- **Hojas publicadas:** `sintesis_hallazgos.HOJAS_PUBLICAS` fija nombres sin siglas y de 31 caracteres o menos (límite de Excel), por ejemplo «Coherencia, por servicio». Los usan el informe y el anexo técnico, y cada cita de hoja del Anexo 2 existe en el libro entregado.
- **Anexo técnico** (decisiones del usuario: carpeta `07_anexo_tecnico/` fuera de git; incluye la nota, el informe y la presentación, y el código):
  - `src/analisis_pnig/anexo_tecnico.py` arma `07_anexo_tecnico/Anexo_tecnico_PNIG_2019-2025/`, de nombre corto porque las rutas de Windows no pueden pasar de 260 caracteres;
  - subcarpetas: índice, informe y presentación, nota metodológica, instrumentos, matrices y libros, entrevistas anonimizadas, rastro, fuentes, figuras numeradas y código;
  - el índice va en Word y PDF, y la lista de archivos lleva su SHA-256;
  - el informe registra sus figuras en `06_informe/indice_figuras.csv`, para copiarlas al anexo con su número;
  - en el código, las rutas locales se reemplazan por `%USERPROFILE%`; quedan fuera los memorándums, el mantenimiento y el formulario FIE-03;
  - sin datos personales: se verificaron las columnas de los libros y las transcripciones.
