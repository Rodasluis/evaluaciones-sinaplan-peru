# ADR-014. Versión pública del repositorio

**Fecha:** 2026-11-12 · **Estado:** aprobada por el usuario; publicada el 2026-11-17 con su confirmación expresa

## Contexto
- **Pedido del usuario:** después de la Fase 2, preparar la versión pública del repositorio.
- **El repositorio de trabajo no puede publicarse tal cual:**
  - su historia de git pesa 790 MB y contiene datos personales en commits anteriores;
  - `docs/fuentes` pesa 933 MB e incluye 61 documentos sin publicación oficial;
  - guarda las transcripciones y los registros de las entrevistas;
  - tres scripts tenían escritos el nombre, el correo y el celular del especialista de contacto;
  - el FIE-03 lleva los identificadores de su hoja de cálculo y de su proyecto de Apps Script;
  - la plantilla de Word, la portada y la skill llevan la identidad del MIMP y del Gobierno del Perú.

## Decisión
- **Decisiones del usuario (2026-11-12):**
  - piloto completo: informe, presentación de resultados y libros de análisis, sin entrevistas ni documentos internos;
  - fuentes por inventario y descarga, sin los PDF;
  - identidad neutral del kit; los productos del piloto conservan la identidad con que el MIMP los emite;
  - licencias MIT para el código y CC BY 4.0 para la documentación, las tipologías, las plantillas y los contenidos
    del piloto.
- **Repositorio aparte y sin historia**, armado con `src/mantenimiento/version_publica.py` desde una lista de inclusión
  (`publico/manifiesto.yaml`) sobre los archivos versionados. Los archivos propios de la versión pública (README,
  CLAUDE.md, licencias, flujo de pruebas, portada neutral) están en `publico/`. Se puede volver a armar en cada
  actualización.
- **Transformaciones al armar:**
  - encabezado neutral de la plantilla de Word;
  - inventario público de fuentes con el SHA-256 del archivo publicado (el original, cuando el corpus aplicó OCR);
  - portada del piloto en `evaluaciones/pnig_2026/portada.yaml`;
  - autor institucional en los metadatos de los productos del piloto;
  - lista de archivos de la referencia del piloto que no se publican.
- **Fuentes:** `src/fuentes/descargar_inventario.py` las descarga desde su enlace oficial y verifica el SHA-256.
- **Repositorio de trabajo (regla 6):**
  - los datos de contacto pasan a `config/entorno.local.yaml` (`nucleo.entorno.contacto`);
  - la auditoría detecta los celulares escritos en el código;
  - la skill funciona sin los fondos ni los logos.

## Lo que no se publica
- **Del piloto:** la nota metodológica de trabajo y los comentarios del CEPLAN (`00_encargo/`); las transcripciones y
  los registros de las entrevistas; los microdatos (`03_datos/`); el anexo técnico.
- **Guías y cuestionarios de entrevista:** salen de la versión pública en la revisión de segunda línea (2026-11-13).
  Son las versiones aplicadas, con el consentimiento, la fecha y hora, la modalidad y las preguntas como se formularon
  a cada informante, y a los informantes se les informó que la grabación era de uso interno. Del piloto sale solo el
  banco de preguntas (`ENTREV_preguntas.csv`).
- **Fuentes:** los documentos sin publicación oficial y los PDF; los registros de descargas, reorganizaciones y OCR.
- **Herramientas administrativas:** memorándums (`src/documentos_evaluacion/`); el FIE-03, por sus identificadores de
  despliegue; los scripts de mantenimiento.
- **Identidad y configuración:** logos, `config/entorno.local.yaml` y los documentos de trabajo del repositorio
  (CLAUDE.md, README y encargo inicial), reemplazados por los de `publico/`.

## Verificación del armado (2026-11-12)
- **Contenido:** 358 archivos, 46,6 MB.
- **Auditoría del armado:** sin problemas.
- **Pruebas:** 35 en verde desde un entorno nuevo.
- **Demostración:** una evaluación de prueba, creada y borrada, salió con la plantilla neutral.
- **Descarga:** documentos con y sin OCR, con el hash verificado.

## Decisiones del usuario sobre la publicación (2026-11-13)
- **Titular de las licencias:** el autor. Su nombre y la forma de la cita están en `publico/manifiesto.yaml`, de donde
  los toma el armado; en la versión pública solo aparecen en LICENSE, LICENSE-docs.md y README.md.
- **Informe y presentación del piloto:** el MIMP autorizó su publicación y el informe es de autoría íntegra del
  usuario (2026-11-13); salen con CC BY 4.0 y el mismo titular.
- **Momento:** se publica después de la validación del informe (días 25 a 30), con la versión final rearmada.
- **Entrevistas:** la matriz de triangulación y el rastro de los hallazgos salen como en el informe, con lo dicho
  atribuido a cada órgano.
- **Identidad de los commits:** el nombre del autor con su dirección `noreply` de GitHub.

## Publicación (2026-11-17)
- **Condiciones cumplidas:** el informe incorpora el criterio de la OGAJ y la revisión institucional (ADR-015) y los
  informes técnicos de la transición (ADR-016), y la DGIGND lo validó; la evaluación se presenta como culminada.
- **Destino:** repositorio público `evaluaciones-sinaplan-peru` en la cuenta de GitHub del autor, con nombre elegido por
  el usuario. Su URL va en `publico/manifiesto.yaml` (`url_repositorio`), de donde el armado la toma para «Cómo citar».
- **Historia:** la carpeta armada se inicia con git y un solo commit, con la identidad del autor y su dirección
  `noreply` de GitHub. Las actualizaciones se rearman con `version_publica.py` (que conserva `.git`) y se publican con
  un commit nuevo, con la confirmación del usuario.
- **Escaneo de secretos:** la protección de GitHub bloqueó el primer push por 95 identificadores de registros de LA
  Referencia («LAREF:ES_» y 32 caracteres hexadecimales), que confunde con claves de hCaptcha; se revisó que no hay
  claves reales (la de OpenAlex se lee de una variable de entorno). Por decisión del usuario, `.github/secret_scanning.yml`
  (superposición `publico/`) excluye solo los cinco archivos de la búsqueda académica que los contienen.
- **Registro:** publicada el 2026-11-17 en `evaluaciones-sinaplan-peru` (público), con dos commits: `3d3c850`
  (configuración del escaneo de secretos, que GitHub lee de la rama por defecto y por eso va primero) y `71d0ed5`
  (versión pública inicial, 363 archivos, armada desde `d58b4d4`). La prueba automática de GitHub (`pruebas.yml`)
  pasó en el primer push.

## Actualizaciones
- **2026-11-18, Fase 4 (ADR-017):** commit `62a8a2c`, armado desde `ce3775f` (416 archivos, auditoría sin problemas,
  58 pruebas en verde desde un entorno nuevo). Agrega los agentes por etapa, `/evaluar`, las preguntas iniciales de
  `/nueva-evaluacion`, los instrumentos genéricos y `docs/agentes.md`. Push confirmado por el usuario; la prueba
  automática de GitHub pasó.
