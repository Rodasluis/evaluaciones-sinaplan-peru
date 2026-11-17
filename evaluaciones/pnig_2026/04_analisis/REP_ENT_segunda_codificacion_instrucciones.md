# Segunda codificación ciega, REP_ENT (evaluación de cierre de la PNIG 2019–2025)

Eres el segundo codificador. Codificas de forma independiente 81 segmentos de texto que reportaron las entidades responsables de los servicios de la Política Nacional de Igualdad de Género (PNIG). Las fuentes son el formulario de seguimiento 2023–2025, los informes anuales de la Ley de Igualdad de Oportunidades (LIO) 2019–2025 y los informes I a IV de la CONAIG.

## Archivos
- `segmentos.csv`: id, fuente, anio, pagina, servicio, indicador, entidad, campo, bloque, texto.
- `libro_codigos.csv`: código, pregunta, bloque, categoría, nombre, definición, incluye, excluye, fuentes.

Lee solo estos dos archivos. No abras ningún otro archivo del equipo ni del repositorio: la codificación debe ser ciega respecto de la primera.

## Qué registrar por segmento
1. **codigos**: todos los códigos del libro que el texto sustenta, separados por «; ». Respeta las columnas «incluye» y «excluye». Usa C.EJEC solo si el segmento describe ejecución sin un factor, y EMER si hay un factor que el libro no prevé (descríbelo en la justificación).
2. **sentido**: Favorable, Dificultad, Mixto (favorable y dificultad), Propuesta (recomendación o ajuste) o Neutro (sin valencia).
3. **critico**: Sí o No. Es hallazgo crítico el segmento que, solo o junto con otros, sustenta una conclusión capaz de cambiar la valoración de una pregunta de evaluación (D5, I1, I4 o I5) o una recomendación sobre la continuidad, el rediseño o el retiro de un servicio o indicador en la política sustituta. Por ejemplo: el servicio no se presta o se interrumpió; falta presupuesto, proveedor o personal y eso bloquea la entrega; el indicador no puede medirse; un cambio normativo afecta al servicio; una falla estructural de coordinación o rectoría; una caída sustantiva de cobertura. No es crítico una dificultad menor, puntual o ya resuelta, ni una descripción de ejecución.
4. **corresponde_a**: si el texto trata de un servicio (código como 4.3.7), de un objetivo prioritario (OP1 a OP6) o de la CONAIG distinto del indicado en las columnas servicio o indicador, anótalo; si no, déjalo vacío.
5. **justificacion**: una oración que diga qué parte del texto sustenta los códigos y la decisión sobre «critico».

## Salida
Escribe `segunda_codificacion.csv` en esta misma carpeta, en UTF-8, con todos los campos entre comillas y las columnas:
`id, codigos, sentido, critico, corresponde_a, justificacion`. Debe tener exactamente 81 filas, una por id de `segmentos.csv`, y solo códigos que existan en el libro.
