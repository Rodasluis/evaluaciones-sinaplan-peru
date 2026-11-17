# Suficiencia de los insumos existentes para la evaluación de implementación de la PNIG

Análisis del 27 de octubre de 2026 sobre dos fuentes:
- el formulario web «Sistema de Seguimiento de Políticas» de la OMEP (2023, 2024 y 2025);
- los informes anuales de la Ley de Igualdad de Oportunidades (LIO), 2019 a 2025.

Se valora si ambas permiten responder las preguntas de implementación sin una nueva solicitud a las entidades.

Datos y scripts:
- `03_datos/raw/formulario_seguimiento_omep/`: una hoja por año, solo PNIG, más los cuestionarios.
- `03_datos/procesado/panel_formulario_seguimiento_pnig_2023_2025.csv`: 159 registros armonizados.
- `03_datos/procesado/cobertura_lio_servicios_2019_2025.csv`: 364 registros, 52 servicios por 7 años.
- `src/insumos_seguimiento/`: scripts que reproducen ambos productos.

## 1. Qué contiene cada fuente

| Fuente | Unidad de registro | Cobertura | Información cuantitativa | Información cualitativa |
|---|---|---|---|---|
| Formulario de seguimiento OMEP 2023–2025 | Indicador de OP o de servicio, por entidad y año | 45, 53 y 61 indicadores; 17 entidades; 24 de los 52 servicios | Numerador, denominador, valor, meta, avance y nivel de cumplimiento; en 2023, actividades operativas del POI con metas y presupuesto | Factores favorables y de dificultad, buenas prácticas, medidas, causas de falla de planeación, coordinación con otras entidades, contexto y efectos no previstos |
| Informes LIO 2019–2025 | Servicio, por objetivo prioritario y año | Los 52 servicios desde 2021; 26 entidades informantes en 2025 | Coberturas, atenciones, metas físicas y, en la mayoría de secciones, referencias presupuestales | Acciones, normas emitidas, articulación, dificultades, análisis por OP, conclusiones, recomendaciones y funcionamiento de la CONAIG |

Ambas fuentes recogieron la información de las mismas entidades responsables de los servicios, mediante solicitud formal del MIMP. Las preguntas del formulario de seguimiento coinciden en buena parte con los bloques B, C y D del FIE-03.

## 2. Formulario de seguimiento 2023–2025

### 2.1 Cobertura y reporte

| | 2023 | 2024 | 2025 |
|---|---|---|---|
| Indicadores registrados (OP / servicio) | 12 / 33 | 13 / 40 | 13 / 48 |
| Servicios con al menos un indicador | 17 | 20 | 24 |
| Servicios con valor reportado | 16 | 19 | 19 |
| Indicadores de OP con valor | 3 | 4 | 7 |
| Indicadores sin información disponible | 10 | 5 | 4 |
| Entidades informantes | 13 | 15 | 17 |

- **Cobertura:** el formulario cubre los 24 servicios que tienen protocolo aprobado con indicadores. Los otros 28 servicios nunca registraron indicador, porque no tienen ficha técnica medible, y para ellos la fuente es el LIO.
- **Continuidad:** 17 servicios tienen información en los tres años, lo que permite trayectoria. Los servicios 1.2.6, 3.1.1, 4.3.6 y 5.5.1 se incorporan recién en 2025, tras aprobar su protocolo.
- **Motivo de no reporte:** casi siempre es que la encuesta nacional no había sido publicada (ENDES, ENAHO o ENARES). Los indicadores del OP2, a cargo del MINSA con la ENDES, no tienen valor en ningún año.

### 2.2 Nivel de cumplimiento de los indicadores de servicio con valor

| Nivel | 2023 | 2024 | 2025 |
|---|---|---|---|
| Alto | 13 | 16 | 19 |
| Medio | 4 | 4 | 7 |
| Bajo | 5 | 2 | 4 |
| Posible falla de planeación | 10 | 16 | 9 |
| Proporción con posible falla | 31 % | 42 % | 23 % |

- **Peso de las fallas:** la posible falla de planeación afecta entre la cuarta parte y dos quintos de los indicadores con valor. Es evidencia directa para la pregunta D5 (calidad de los indicadores).
- **Causas que declaran las entidades:**
  - metas fijadas con información insuficiente;
  - metas fijadas en función de recursos limitados;
  - métodos de cálculo que distorsionan el avance;
  - presupuesto adicional;
  - agenda política;
  - indicadores fuera del ámbito de control de la entidad.
- **Problemas de diseño documentados por las propias entidades:**
  - el sentido del indicador 1.1.2 está invertido respecto de su ficha;
  - SERVIR usa la planilla electrónica y no la ENAHO que fija la ficha;
  - el aplicativo calculó mal la meta del indicador de ONPE;
  - el indicador 1.2.3 se supera en más de 400 %;
  - SERVIR propone incluir las llamadas devueltas en el servicio 1.1.8.
- **Propuesta de retiro:** COFOPRI pide evaluar el retiro del servicio 4.2.2 de la PNIG. Sostiene que titula sin distinción de sexo y no puede comprometer una mayor titulación de mujeres.

### 2.3 Factores que explican el desempeño (pregunta I1)

**Factores de dificultad** (opciones marcadas, recuento de respuestas):

| Factor | 2023 | 2024 | 2025 |
|---|---|---|---|
| Limitado presupuesto | 7 | 5 | 6 |
| Insuficiente personal especializado | 6 | 4 | 5 |
| Débil articulación | 4 | 4 | 4 |
| Cambios de gestión | 4 | 1 | 3 |
| Limitación de competencias funcionales | 5 | 2 | 1 |
| Conflictos sociales y desastres | 10 | 1 | 1 |

**Factores favorables:** predominan la programación y planificación, el trabajo en equipo y el fortalecimiento de capacidades. El presupuesto y la articulación intersectorial aparecen con menor frecuencia.

**Patrón cualitativo consistente en los tres años:** la implementación descansa en la gestión interna de cada entidad y no en la articulación. Las restricciones recurrentes son cuatro:
- **Personal:** brecha de 1 141 plazas en los CEM en 2023, renuncias y contratación tardía.
- **Presupuesto:**
  - recortes;
  - el Decreto de Urgencia N.° 008-2025;
  - una asignación del servicio CEM menor a la prevista en el PPoR 1002.
- **Rotación de autoridades.**
- **Baja prioridad del enfoque de género** en las entidades asistidas.

### 2.4 Articulación (pregunta I4)

- **Necesidad de coordinar:** en 2023, 15 de 37 respuestas declaran requerir a otras entidades para proveer el servicio; en 2025, 12 de 47. Entre quienes lo requieren, cerca de la mitad reporta dificultad de coordinación (6 de 15 en 2023; 7 de 13 en 2025).
- **Causas de la dificultad en 2025:**
  - inexistencia de convenios o protocolos de articulación (la más citada);
  - duplicidad de funciones;
  - demora institucional;
  - discrepancias técnicas;
  - falta de un espacio intersectorial para los lactarios.
- **Pedidos al MIMP:** MINCETUR solicita expresamente una mayor articulación con el MIMP: difusión, préstamo de espacios y codiseño de servicios formativos.
- **Vacío de 2024:** el formulario de ese año no incluyó las preguntas de coordinación y de contexto.

### 2.5 Presupuesto y POI (pregunta I1, componente presupuestal)

- **Solo 2023:** solo el formulario de 2023 preguntó si el servicio está en el POI y por sus actividades operativas.
- **Respuestas:** 16 indicadores responden que sí, 21 que no y 8 no responden. Hay 15 registros con detalle de actividades, metas físicas y presupuesto programado y ejecutado.
- **Problema de consistencia:** los nueve indicadores de AURORA responden «no», pese a ser servicios financiados por el PP 1002.
- **Conclusión:** la asignación presupuestal identificable por servicio no puede establecerse de forma sistemática con esta fuente.

### 2.6 Calidad del dato

- **2025 está cerrado:** 10 registros figuran como «Pendiente» o en borrador solo porque el aplicativo se cerró antes de que se marcaran como enviados. Su contenido es lo reportado y no habrá más información (confirmado por la OMEP el 27-10-2026).
- **Escala del avance:** en 2023 se registró como razón y en 2024 y 2025 como porcentaje; el panel lo armoniza.
- **Indicador IOP.05.02 (brecha salarial en el sector público, a cargo de SERVIR):** se contrastó con la ficha vigente, el consolidado de la OMEP y los reportes de cumplimiento 2023–2025.
  - **Fórmula de la ficha:** P = ((A − B) / B) × 100, donde A es el salario promedio de los servidores y B el de las servidoras. La fuente que fija la ficha es la ENAHO.
  - **Serie con la ENAHO:** 12,3 (2019), 10,0 (2020), 11,75 (2021), sin dato (2022) y 11,7 (2023). El reporte de cumplimiento 2023 valora 11,7 frente a una meta de 14,0 como cumplimiento alto.
  - **2024, error de cálculo:** el valor registrado, 101,98, es el cociente A/B × 100 (3 612 / 3 542), no la brecha. Con la fórmula de la ficha corresponde 1,98. Así lo reconoce el reporte de cumplimiento 2024, que no analiza el indicador porque SERVIR remitió una estimación de 2,0 % con registros administrativos. El error se trasladó al formulario, al consolidado de la OMEP y a la serie del reporte de seguimiento 2025 («102,0»).
  - **2025, cambio de fuente:** el valor −1,03 aplica bien la fórmula ((4 399 − 4 445) / 4 445 × 100). Pero proviene de la planilla electrónica (Informe SERVIR «Mujer en el servicio civil 2025»), no de la ENAHO.
  - **Consecuencia:** los valores de 2024 y 2025 no son comparables con la serie 2019–2023 ni con las metas, que se calcularon sobre la ENAHO. La clasificación «bajo con posible falla de planeación» y el avance de −1 149,9 % son artefactos de ese cambio y no reflejan desempeño.
  - **Inconsistencia de diseño:** la meta 2019 (20,21) supera la línea de base de 2015 (16) en un indicador descendente. Esto obedece a que el sentido original era «ascendente» y se cambió en la propuesta de 2022, cuya acta no fue ratificada.
  - **Tratamiento en la evaluación:** se registra como hallazgo de calidad técnica (D5). Para 2024 se corrige a 1,98 con sustento en la ficha y se declara la ruptura de serie desde 2024.
- **Nombres de entidades:** en 2024 los órganos del MIMP figuran como entidades; el panel los agrupa bajo MIMP.
- **Compromisos:** el campo está en «No» para todos los indicadores de 2024 y 2025. No hubo seguimiento de la Matriz de Compromisos de 2023 por esta vía.

## 3. Informes LIO 2019–2025

- **Cobertura:** desde 2021 cada informe tiene una sección para cada uno de los 52 servicios; en 2019 y 2020 se ubicaron 49 y 47. Es la única fuente con información anual de los servicios sin indicador.
- **Contenido de las secciones** (2022–2025, número de servicios cuya sección lo menciona):

  | Contenido | Servicios |
  |---|---|
  | Presupuesto | 37 a 46 |
  | Protocolo | 32 a 47 |
  | Articulación | 30 a 43 |

  La extensión de cada sección es muy desigual. Las de 6.3.5 y 2.1.1 ocupan pocas líneas; las de 4.4.1, 5.5.1 y 1.2.6, varias páginas.
- **Análisis propio:** desde 2024 incluyen un análisis de la implementación por OP y recomendaciones generales y específicas.
  - Las recomendaciones de 2025 reiteran las restricciones del formulario: incorporar actividades operativas específicas en los instrumentos de planificación para garantizar recursos, y focalizar en ámbitos rurales y de la selva.
  - Las conclusiones de 2025 registran:
    - la cobertura distrital de los CEMF (18 %);
    - el piloto de taxonomía del gasto en tres programas presupuestales del MIMP;
    - los servicios del INABIF en el 6.4.1;
    - las acciones de cuidado en tres distritos para el 4.1.1.
- **CONAIG** (capítulo de articulación interinstitucional): número de sesiones y asistencia; en 2025 fueron tres sesiones, con asistencia de 57 %, 43 % y 91 %. También informa:
  - el avance de las 12 medidas urgentes de Beijing+30;
  - las conclusiones del IV Informe anual de propuestas de mejora de la gobernanza.

  Ese informe se elaboró con información de 21 entidades y 44 servicios. Mide la capacitación del personal (solo 34,7 % de los servicios participó en el curso del MIMP), la articulación intergubernamental (38,7 % coordina con el nivel nacional y 30,6 % con gobiernos regionales y locales) y la articulación con la sociedad civil (20,4 %).
- **Naturaleza de la información:** es declarativa y descriptiva. El LIO no clasifica el estado de implementación del servicio ni separa el presupuesto por servicio de forma homogénea, y el formato cambió entre años.

## 4. Suficiencia por pregunta de implementación

| Pregunta | Qué requiere | Formulario 2023–2025 | LIO 2019–2025 | Otras fuentes ya disponibles | Valoración |
|---|---|---|---|---|---|
| I1 Estado de implementación por servicio y factores | Protocolo, ejecución, reporte, asignación y factores | 24 servicios con indicador; factores en 3 años | 52 servicios en 7 años | Protocolos aprobados; Informe D000039-2026-MIMP-DPIGND con el estado de 58 protocolos; reportes de cumplimiento | **Suficiente** para el estado y los factores; **parcial** para la asignación presupuestal por servicio |
| I2 Trayectoria de indicadores | Valores y metas anuales | 2023–2025 | Valores de contexto | Consolidado OMEP; reportes de seguimiento 2020–2025 | **Suficiente** |
| I3 Recursos del MIMP para la conducción | Ejecución presupuestal | No aplica | No aplica | Consulta Amigable; POI del MIMP | Sin cambio: no dependía de las entidades |
| I4 Rectoría, CONAIG y articulación | Coordinación efectiva y valoración del rol del MIMP | Coordinación y causas (2023 y 2025) | Capítulo CONAIG; medidas urgentes; gobernanza | Actas de la CONAIG; informes anuales de gobernanza I a IV | **Suficiente** para la articulación efectiva; **parcial** para la valoración que hacen las entidades del rol rector del MIMP |
| I5 Capacidades y sostenibilidad ante la transición | Brechas de capacidad y riesgos ante la Ley N.° 32535 | Personal, presupuesto, rotación y contexto | Recomendaciones | Hoja de ruta de la Ley N.° 32535; Informe D000039-2026 | **Suficiente** para las capacidades; **insuficiente** para la perspectiva de las entidades sobre la transición de 2026 |
| Recomendación por componente | Qué mantener, ajustar o reformular | Propuestas de rediseño de indicadores; pedido de retiro del 4.2.2 | Recomendaciones anuales | Evaluaciones de 2021 y 2023 | **Parcial** |

## 5. Conclusión

**Lo que se puede sostener:** los insumos existentes permiten responder las preguntas I1, I2 e I4 y la mayor parte de I5, sin una nueva solicitud a las entidades. El sustento técnico es triple:
- **Mismo contenido:** ambas fuentes ya recogieron, de las mismas entidades y mediante solicitud formal del MIMP, la información que el FIE-03 pedía. En el formulario de seguimiento son los factores, la coordinación y el contexto por servicio; en el LIO, la ejecución anual de los 52 servicios.
- **Periodo completo:** juntas cubren todo el periodo con información anual. El LIO cubre 2019–2025 para todos los servicios y el formulario 2023–2025 para los servicios con indicador.
- **Evita duplicar:** una nueva solicitud repetiría preguntas ya respondidas y añadiría carga a las entidades, sin cambiar el periodo evaluado.

**Vacíos que la evaluación debe cubrir con otros medios:**
1. **Valoración del rol rector del MIMP:** la que hacen las entidades sobre la asistencia técnica y la CONAIG no se preguntó en ninguna fuente. Puede aproximarse con los informes anuales de gobernanza de la CONAIG y las entrevistas al nivel conductor.
2. **Transición a la Ley N.° 32535:** la perspectiva de las entidades sobre la transición y sobre qué servicios mantener o ajustar no está en fuentes anteriores a 2026. Se cubre parcialmente con las propuestas de rediseño registradas en el formulario y debe declararse como limitación.
3. **Asignación presupuestal por servicio:** fuera del MIMP, no es identificable de forma sistemática. Se valora con el POI de 2023, las referencias presupuestales del LIO y la Consulta Amigable donde la estructura lo permita, y se declara como aproximación.

## 6. Insumos por incorporar o verificar

- **Informes anuales de gobernanza de la CONAIG** (I a IV) y sus actas: no están en el corpus. Son documentos internos del MIMP.
- **Consistencia de valores:** extender al resto de indicadores el contraste hecho para el IOP.05.02 (sección 2.6), comparando el formulario con el consolidado de la OMEP y los reportes de cumplimiento.
- **Órganos responsables de los servicios del MIMP sin órgano precisado en la Política:** se identifican a partir de los informes LIO (secciones por servicio, 2022–2025). La OMEP confirmó la identificación el 27-10-2026.

  | Servicio | Órgano responsable (fuente: informes LIO) |
  |---|---|
  | 4.1.1 | Dirección General de Promoción y Desarrollo de la Autonomía Económica de las Mujeres (DGPDAEM) |
  | 6.2.1 | Dirección de Transversalización del Enfoque de Género (DTEG) y DPIGND |
  | 6.3.3 y componente MIMP del 6.3.5 | DPIGND |
  | 6.4.1 | Programa Integral Nacional para el Bienestar Familiar (INABIF) y Dirección General de la Familia y la Comunidad (DGFC) |
