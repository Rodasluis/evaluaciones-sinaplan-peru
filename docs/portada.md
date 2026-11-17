# Portada de los informes técnicos

El módulo `src/nucleo/portada` arma la portada de los informes en Word con un solo diseño, el centrado (G). Todo lo que se ve en la portada sale de un archivo de configuración: `config/portada.yaml` en el piloto PNIG y `evaluaciones/<id>/portada.yaml` en cada evaluación nueva, que crea `src/nucleo/nueva_evaluacion.py` con su etiqueta, título y fecha. Para cambiar un texto, un logo o un color no hace falta tocar código; las rutas de los logos son relativas a la raíz del repositorio.

## Diseño

- **Base:** la portada del informe de evaluación 2023 de la PNIG del MIMP.
- **Composición:**
  - logos arriba, en los extremos;
  - todo lo demás centrado: la etiqueta en mayúsculas, el título en rojo y en negrita, el subtítulo si lo hay, la institución y el órgano elaborador en gris, y el lugar y fecha.
- **Historia:**
  - el usuario lo eligió el 2026-11-10 entre once modelos (A a K) para el informe de cierre de la PNIG;
  - en la ronda 35 (2026-11-11), a pedido del usuario, los demás modelos se retiraron del repositorio;
  - su historia está en ADR-010.

## Archivos

| Ruta | Contenido |
|---|---|
| `config/portada.yaml` y `evaluaciones/<id>/portada.yaml` | Textos, rutas de los logos, colores y tipografía |
| `assets/logos/` | Logos oficiales: `logo_institucion.png` (MIMP) y `logo_gobierno.png` (Gobierno del Perú) |
| `src/nucleo/portada/__init__.py` | `agregar_portada(documento, ruta_config)`: valida los datos y arma la portada |
| `src/nucleo/portada/opcion_g.py` | El diseño centrado |
| `src/nucleo/portada/paleta.py` | Colores y tipografía leídos de la configuración |
| `src/nucleo/portada/comun.py` | Validación, tablas ancladas a la página, logos, ajuste del título y cierre de la sección |

## Datos de la configuración

- **Obligatorios:** `titulo`, `institucion`, `organo_elaborador` y `lugar_fecha`. Si falta alguno, el módulo se detiene y pide completarlo; no inventa datos.
- **Opcionales:** `etiqueta` y `subtitulo`. El subtítulo está vacío desde la ronda 33, a pedido del usuario. Si están vacíos, la portada los omite.
- **Órgano elaborador:** si son varias unidades, se separan con « · » y cada una va en su propia línea.
- **Colores:**
  - `primario` es el rojo institucional. Por defecto vale `#E1251B` (Pantone 485 C), valor que hay que confirmar con el manual de identidad de la institución o con la guía de identidad visual de la PCM.
  - Se leen siempre de la configuración, y el módulo rechaza cualquier azul.
- **Tipografía:**
  - Si la indicada no está instalada, se usa Calibri y el registro lo avisa.
  - Poppins está instalada para el usuario desde el 2026-11-08 (`docs/fuentes/descargas_tipografia-poppins_2026-11-08.tsv`).
  - Al terminar el informe, Word la incrusta completa en el .docx (`src/nucleo/actualizar_campos_word.ps1`).

## Reglas de diseño

- **Página:** hoja A4 en su propia sección, sin encabezado, pie ni número de página. El cuerpo empieza en la página siguiente con su numeración.
- **Texto:** todo es texto de Word, en bloques que son tablas ancladas a la página.
- **Lo que no lleva:**
  - azul, fotos, degradados, sombras, adornos ni lemas inventados;
  - versión, créditos, estado ni resultados.
- **Título:** ocupa como máximo 4 líneas. Si no cabe, la letra baja de 1 en 1 punto hasta 20 pt; si aún no cabe, el módulo avisa y hay que acortarlo.
- **Logos:**
  - no se redibujan, no se recolorean y no se deforman;
  - van en una fila, con el mismo alto y con texto alternativo («Logo de {institución}» y «Logo del Gobierno del Perú»);
  - si falta un archivo, va un recuadro punteado y un aviso;
  - el escudo nacional tiene celeste en el campo del blasón: es parte del logo oficial y no se modifica, porque la regla sin azul se aplica a los colores de la portada.

## Verificación

- **El informe informa:**
  - el diseño;
  - la tipografía;
  - el contraste del blanco sobre el rojo (con `#E1251B`, 4,69:1);
  - los avisos (logos que faltan, título que no cabe).
- **En el PDF de Word hay que revisar:**
  - que la portada ocupe una sola página;
  - que los logos estén nítidos y sin deformar;
  - que no haya número de página.

## Detalle técnico

Word mide la posición horizontal de una tabla flotante desde el texto de su primera celda. Por eso los márgenes laterales del texto van como sangrías de párrafo y no como márgenes de celda: un margen izquierdo de celda correría el bloque fuera de su sitio.
