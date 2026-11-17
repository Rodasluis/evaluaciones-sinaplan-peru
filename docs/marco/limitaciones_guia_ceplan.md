# Limitaciones teóricas y metodológicas del marco CEPLAN para evaluar políticas nacionales

**Estado:** BORRADOR Fase 0 — no aprobado. Destino: `docs/marco/limitaciones_guia_ceplan.md`.
**Fecha de corte:** 2026-10-22.

## 0. Alcance y convenciones

Documentos analizados (lectura completa del texto extraído; las páginas citadas son las impresas, que en estos PDF coinciden con la paginación del archivo):

| ID | Documento | Versión / norma | Archivo local |
|---|---|---|---|
| **SyE** | Guía para el seguimiento y evaluación de políticas nacionales y planes del Sinaplan | RPCD N.° 0056-2024-CEPLAN/PCD; 3.ª ed., set. 2024. Modificada por RPCD 099-2024 y 0101-2024 (El Peruano, 19 y 24-09-2024); falta verificar si el PDF ya incorpora la 0101 `[POR VERIFICAR]` | `docs/fuentes/ceplan/5678657-…actualizada-2024.pdf` |
| **PN** | Guía de Políticas Nacionales | Aprobada por RPCD N.° 113-2024/CEPLAN/PCD (El Peruano, 12-10-2024). El PDF dice "Noviembre de 2024" y no consigna la norma en la carátula; falta confirmar que es el anexo de la RPCD 113-2024 `[POR VERIFICAR]` | `docs/fuentes/ceplan/6088033-…actualizada-2024.pdf` |
| **IND** | Guía para la elaboración de indicadores de políticas nacionales y planes estratégicos | RPCD N.° 0046-2024-CEPLAN/PCD; 1.ª ed., jun. 2024. **La definición de línea base fue modificada por el art. 5 de la RPCD 042-2026/CEPLAN/PCD (El Peruano, 20-06-2026); el PDF cargado no recoge ese cambio** | `docs/fuentes/ceplan/5614517-…2024(2).pdf` |

Cada hipótesis se presenta en tres planos que no deben mezclarse:
- **(a) CEPLAN:** lo que dice el texto, con referencia.
- **(b) Literatura:** qué aporta la literatura complementaria. **Ninguna referencia de este plano está verificada aún** (faltan edición y páginas): todas quedan `[POR VERIFICAR]` y no entran a `registro.bib` hasta verificarlas.
- **(c) Decisión del evaluador:** cómo lo resuelve el repositorio.

Veredictos: **Confirmada**, **Parcialmente confirmada**, **Refutada**.

---

## H1. Débil fundamento en teoría del cambio y supuestos causales — *Parcialmente confirmada*

**(a) CEPLAN**
- SyE no usa los términos "teoría del cambio", "teoría del programa" ni "modelo lógico". "Causal" aparece una sola vez: la evaluación de diseño valora la "lógica causal entre las acciones y los objetivos" (SyE §1.3.2, p. 20). Cita a Mideplan (2017a), *Guía de la teoría de la intervención* (p. 141), pero no la aplica en el texto.
- En el informe, la "lógica de intervención" solo se describe en la sección 1 (Tabla 3, p. 55; Tabla A7.1, p. 183). No se pone a prueba.
- **Contrapunto (refuta en parte):** PN sí exige, en la formulación, un **modelo del problema** que muestre relaciones de causalidad sustentadas en evidencia (PN §2.3, pp. 46–47, Tabla 2). También pide alternativas de solución respaldadas por meta-análisis y estudios experimentales o cuasiexperimentales (PN §3.2.1, p. 54) y lineamientos que "conservan la lógica de causalidad" del modelo (PN §4.2.2, pp. 64–65). La cadena de resultados está definida (PN §7, pp. 20–21).
- Brecha: el modelo es del **problema** (árbol de causas), no del **programa**. No se piden supuestos explícitos de cada eslabón, ni mecanismos, ni condiciones de contexto. Los "riesgos" se tratan en la prospectiva (PN p. 50), no como riesgos causales de la cadena. SyE no recupera el modelo de la formulación para contrastarlo ex post.
- Hay una prueba rudimentaria de teoría en el seguimiento: el ejemplo de la Tabla A11 (SyE pp. 160) alerta cuando las acciones cumplen pero el objetivo no, porque las acciones podrían ser insuficientes o no incidir en él.

**(b) Literatura `[POR VERIFICAR]`:** evaluación basada en teoría (Weiss 1997; Chen 2015); teoría del programa (Funnell y Rogers 2011); análisis de contribución (Mayne 2012); evaluación realista y configuraciones CMO (Pawson y Tilley 1997).

**(c) Decisión:** `01_teoria_del_cambio.md` debe tener, para cada eslabón, el supuesto, el mecanismo, la evidencia previa, el riesgo y la fuente. Cada elemento se etiqueta como "texto de la política" o "reconstrucción del evaluador". Una prueba automática exige que cada pregunta de resultados apunte a un eslabón de la TdC.

## H2. Tratamiento limitado de identificación causal y contrafactuales — *Confirmada*

**(a) CEPLAN**
- SyE no menciona "contrafactual" en ninguna página. Sin embargo, propone preguntas de impacto: "¿Cuál ha sido el impacto de la política nacional en la calidad de vida…?" (p. 82) o "¿Qué impacto ha generado… respecto a las relaciones de poder dentro de los hogares?" (Anexo 5, p. 173). No ofrece ningún diseño para responderlas.
- Las técnicas de análisis que lista son "análisis estadísticos, visualización de datos, correlaciones, regresiones, data mining, proyecciones, análisis de contenido" (p. 46). El análisis cuantitativo se presenta como descriptivo o como inferencia muestral (pp. 51–52).
- El ejemplo de metodología de la PNMS (p. 83) dice "analizar de manera integral el impacto y la eficacia" solo con triangulación y estadística: confunde cumplimiento con impacto.
- La evaluación de resultados incluye la "revisión de resultados como contribución" a la situación futura deseada (p. 21), pero no explica cómo hacerla.
- Los ejemplos de redacción atribuyen causalmente resultados a factores de contexto sin ningún diseño que lo sostenga, como el efecto de la pandemia en el consumo de drogas (p. 176).
- IND p. 20 asigna a la evaluación "identificar las causas y efectos", sin método.
- Asimetría: PN valora la evidencia causal **ex ante** (p. 54), pero SyE no prevé producirla **ex post**.

**(b) Literatura `[POR VERIFICAR]`:** Imbens y Rubin (2015); Angrist y Pischke (2009); Gertler et al. (2016); Cunningham (2021). DiD con adopción escalonada: Callaway y Sant'Anna (2021), Sun y Abraham (2021), de Chaisemartin y D'Haultfœuille (2020), Goodman-Bacon (2021). Sensibilidad: Rosenbaum (2002), Oster (2019), Rambachan y Roth (2023).

**(c) Decisión:** un árbol de decisión de diseño según el dato disponible. Prueba automática: ningún hallazgo con verbos causales ("efecto", "impacto", "causó", "gracias a") pasa sin un diseño declarado con sus supuestos y diagnósticos. Si no lo tiene, se reformula como asociación o contribución. El módulo `impacto` solo se activa con identificación defendible.

## H3. Escaso desarrollo de la evaluación de procesos e implementación — *Parcialmente refutada*

**(a) CEPLAN**
- Refuta en parte: la implementación pesa mucho. SyE define cinco elementos: cumplimiento de servicios, procesos de transformación, mecanismos de entrega, asignación de recursos y mecanismos de datos (p. 21). Es la sección más extensa del informe (20 páginas; preguntas p3.1–p3.6 sobre recursos, actividades operativas, servicios que "no están funcionando", coordinación de actores, lineamientos sin servicios y seguimiento de recomendaciones; Tabla A7.1, p. 185).
- Confirma en parte: el "grado de implementación" se mide como "cumplimiento de los logros esperados de los indicadores" (p. 185). No aparecen fidelidad, dosis, alcance, adaptación ni calidad de la entrega como constructos. La capacidad institucional se reduce a recursos. No hay perspectiva de la burocracia de calle.
- PN aporta un punto de apoyo: los **estándares nacionales de cumplimiento** de los servicios (PN §4.2.3, pp. 68–71) se pueden operacionalizar como criterios de fidelidad.

**(b) Literatura `[POR VERIFICAR]`:** Moore et al. (2015, guía MRC de evaluación de procesos); Carroll et al. (2007, fidelidad); Proctor et al. (2011, *implementation outcomes*); CFIR (Damschroder et al. 2009; 2022); RE-AIM (Glasgow et al. 1999); Lipsky (1980); Pressman y Wildavsky (1973).

**(c) Decisión:** el módulo `implementacion` usa como constructos centrales la fidelidad (frente a los estándares de cumplimiento y los protocolos de servicio), el alcance y la adopción. Los "factores" de CEPLAN se reorganizan con dominios tipo CFIR adaptados al Estado peruano (niveles de gobierno, modelos de provisión).

## H4. Abordaje limitado de coherencia, complejidad, efectos no previstos y contexto — *Parcialmente confirmada*

**(a) CEPLAN**
- **Coherencia:** está definida como articulación interna y externa (vertical y horizontal) solo dentro de la evaluación de diseño (SyE p. 20, nota 6). En el informe se reduce a describir el alineamiento con otras políticas, la PGG, el PEDN y los ODS (p1.2, p. 183). En PN es alineamiento formal (fichas del Anexo 8, p. 63) y no duplicación (criterio de exclusividad, p. 33). No se analizan sinergias ni contradicciones sustantivas entre políticas.
- **Complejidad:** ausente en SyE. PN recomienda **no** plantear el problema con complejidad demasiado alta (p. 41) y usa la complejidad solo para fijar el horizonte (p. 28): la evita en lugar de evaluarla.
- **Efectos no previstos:** se mencionan (p. 21; pregunta p2.1.2, p. 184; Anexo 5, p. 173), pero no hay procedimiento para detectarlos.
- **Contexto:** factores económicos, sociales, políticos e institucionales tratados como narrativa (p2.2, p. 184; ejemplo en p. 178).

**(b) Literatura `[POR VERIFICAR]`:** criterio de coherencia de CAD-OCDE (2019/2020; SyE cita la versión en español OCDE 2020, p. 142); coherencia de políticas para el desarrollo sostenible (OCDE); Rogers (2008); Ramalingam (2013).

**(c) Decisión:** un módulo `coherencia` separado del diseño CEPLAN. El diseño cubre la consistencia interna y el alineamiento formal; `coherencia` cubre las interacciones sustantivas con otras políticas y servicios compartidos. PN pp. 74–75 documenta que el servicio CEM es compartido entre la PNIG y la política frente a la trata. Los efectos no previstos se buscan de forma sistemática con una "teoría de daños" en la TdC.

## H5. Ausencia de evaluación económica y de criterios de calidad de la evidencia — *Confirmada (con matiz)*

**(a) CEPLAN**
- **Económica:** en SyE, "costo" solo aparece como "centro de costo" del POI (p. 117). La eficiencia figura como criterio principal (p. 19), pero el informe no tiene sección de eficiencia. La pregunta p3.1.3 (p. 185) trata de ejecución presupuestal, no de eficiencia ni de costo-efectividad.
- **Calidad de la evidencia:** SyE define "evidencias concretas" como "declaraciones de hechos basados en el análisis de información objetiva" (p. 54, nota 25) y usa la triangulación como garantía de credibilidad (p. 47). No gradúa la confianza en los hallazgos.
- **Matiz:** PN adopta criterios de evidencia de Sutcliffe y Court (exactitud, objetividad, credibilidad, generalización, relevancia, disponibilidad; pp. 43–44). Recomienda "no presuponer que algunas formas de evidencia tienen más relevancia que otras" (p. 44), pero en la p. 54 jerarquiza implícitamente: meta-análisis, experimentos, cuasiexperimentos. La tensión no se resuelve.

**(b) Literatura `[POR VERIFICAR]`:** Levin et al. (2018); Boardman et al. (2018); GRADE y escalas de confianza; *evidence gap maps* de 3ie (Snilstveit et al. 2016); revisiones Campbell.

**(c) Decisión:** el módulo `economica` incluye reglas explícitas de cuándo **no** aplicar el análisis (sin efecto identificado, no hay costo-efectividad; sin costeo por servicio, solo análisis de costos). Cada hallazgo lleva un nivel de confianza en escala explícita (alto, medio o bajo) con criterios publicados en `docs/marco/`.

## H6. Equidad, género, interseccionalidad e interculturalidad poco tratados como dimensiones analíticas — *Confirmada en SyE; parcialmente refutada en PN*

**(a) CEPLAN**
- **SyE:** "interculturalidad", "interseccionalidad" y "equidad" no aparecen. "Género" aparece cuatro veces, como desagregación de tablas (p. 51) y en ejemplos (pp. 142, 176, 180). La desagregación por sexo, etnia o edad se **sugiere** en preguntas de ejemplo (pp. 45, 173); no se exige y no hay método de heterogeneidad.
- **PN:** los enfoques transversales son obligatorios en la formulación. Las políticas se redactan con enfoque de género e interculturalidad, según el Anexo 1 del D.S. 029-2018-PCM citado en PN p. 42, nota 33. Hay preguntas orientadoras (p. 43), y los servicios deben incorporar género, interculturalidad, discapacidad, enfoque territorial, derechos humanos e interseccionalidad como atributos (pp. 67, 69–70).
- **IND:** la "capacidad de desagregación" es un aspecto **adicional**, no una característica deseable central (p. 46).
- Novedad normativa: el D.S. 076-2026-PCM (2.ª disposición) y la RPCD 042-2026 disponen la **desagregación territorial** progresiva de los indicadores de objetivos prioritarios. Para la heterogeneidad territorial es una oportunidad; no cubre etnia, edad ni nivel socioeconómico. `[POR VERIFICAR texto oficial]`

**(b) Literatura `[POR VERIFICAR]`:** ONU Mujeres (evaluación con enfoque de género); Bamberger y Segone (2011, evaluación con equidad); interseccionalidad cuantitativa (análisis de estratos MAIHDA); Ellsberg y Heise (2005) sobre medición de la violencia contra las mujeres.

**(c) Decisión:** para la PNIG, cada indicador de objetivo prioritario requiere un análisis de brechas y de heterogeneidad (territorial, étnica, etaria y por nivel socioeconómico) con su incertidumbre, o bien la declaración "evidencia insuficiente: desagregación no disponible". El "enfoque de género" se trata como dimensión medible: brechas, cambios en normas y relaciones de poder. No basta con declararlo.

## H7. Foco en cumplimiento de metas y en indicadores como proxies sin validación de constructo — *Confirmada*

**(a) CEPLAN**
- El **Avance tipo I** es la razón entre el valor obtenido y el logro esperado, acotada a [0, 100] (SyE p. 37; Tabla A1, pp. 144–145). El semáforo usa umbrales de 95 y 75 (p. 42). El **índice de gestión** es el promedio simple del Avance tipo I de los objetivos prioritarios y se presenta como "medida de desempeño… que refleja la eficacia" (pp. 38, 154). La pregunta p2.1 del informe ("¿la política… está cumpliendo con sus propósitos?") se responde "apoyándose en el índice de gestión" (p. 184).
- Problemas que se pueden verificar en el texto:
  1. El cumplimiento depende de la calidad de la meta, que se fija por proyección y consenso (PN pp. 59–61; IND paso 6).
  2. El tope de 100 descarta información sobre sobrecumplimiento, aunque el propio ejemplo PNIG reporta "193,7 %" (p. 180): **la guía es inconsistente con su métrica**.
  3. Se agregan indicadores heterogéneos con pesos ordinales arbitrarios (relevancia R = 1, 2, 3; pp. 150–153).
  4. No se considera la incertidumbre muestral: el semáforo se aplica a estimaciones puntuales de encuesta.
  5. El Avance tipo II, que sí considera la línea de base, es solo "complementario" (pp. 37–38, 146).
- **Validez de constructo:** IND no menciona "constructo". Las características deseables son de tipo SMART (específico, relevante, medible, realizable, temporal; pp. 45–46, nota 13), más confiabilidad, simplicidad y desagregación. El Recuadro 3 (p. 97) admite que los indicadores son "representaciones imperfectas", pero no da un procedimiento de validación.
- Ejemplo PNIG en PN (Tabla 7, pp. 74–75; versión "adaptada" de la PNIG, que debe contrastarse con la ficha técnica oficial `[POR VERIFICAR]`): la calidad del servicio CEM con estándar de "accesibilidad geográfica" se mide como porcentaje de atenciones en la Estrategia Rural. Ese indicador mide la composición de la demanda atendida, no la accesibilidad. Además, los conteos de atenciones mezclan incidencia con propensión a denunciar (subregistro).

**(b) Literatura `[POR VERIFICAR]`:** Cronbach y Meehl (1955) y Messick (1995) sobre validez de constructo; ley de Campbell (1979) sobre la distorsión de indicadores usados para decidir; Kusek y Rist (citados por CEPLAN, p. 143).

**(c) Decisión:** el repo recalcula las métricas CEPLAN (plano a, obligatorias en el reporte) y, por separado, métricas complementarias (plano b): trayectoria frente a contrafactual tendencial, intervalos de confianza para indicadores de encuesta y cumplimiento sin tope. Cada ficha técnica pasa por una auditoría de constructo (qué mide, qué no mide, sesgos, sensibilidad al registro).

## H8. Débil guía sobre uso, evaluabilidad y calidad metodológica del informe — *Parcialmente confirmada*

**(a) CEPLAN**
- **Uso — refutada en parte:** la Etapa 3, "Gestión de la evidencia", está bien desarrollada: aprobación, estrategia de difusión por audiencias, reportes gerenciales, matriz de compromisos con estados y plazos, incorporación a los planes y seguimiento en la evaluación siguiente (pp. 58–68, 88). Es un uso instrumental y procedimental; no prevé uso conceptual ni diseño orientado a usuarios desde el encargo.
- **Evaluabilidad — confirmada:** el término no aparece. Lo más cercano es el punto 5 de la evaluación de diseño, "elementos esenciales para el SyE" (indicadores, metas, líneas de base, responsables; p. 20).
- **Calidad metodológica — confirmada:**
  1. **El Paso 8 pide informar "¿Cómo se diseñó y realizó la evaluación?" (p. 53), pero la estructura obligatoria (Tabla 3, p. 55; Tabla A7.1, pp. 183–187) no tiene sección de metodología ni de limitaciones.**
  2. Es autoevaluación: el informe lo elabora la unidad que conduce la política (p. 84). Las evaluaciones externas o colaborativas solo se prevén para políticas que CEPLAN seleccione en el marco del PEDN (p. 74).
  3. La aprobación revisa formato, claridad y respaldo (p. 60), pero no hay estándares de calidad (UNEG, CAD-OCDE, JCSEE).
  4. La ética se limita a pedir permiso para grabar entrevistas (p. 48).
  5. Las recomendaciones deben surgir de la "consulta o discusión con actores claves" (p. 56). Esto refuerza la apropiación, pero entra en tensión con la independencia.

**(b) Literatura `[POR VERIFICAR]`:** Wholey (evaluabilidad); Patton (2008, *utilization-focused*); Normas y estándares UNEG (2016); estándares de calidad CAD-OCDE (2010); JCSEE (Yarbrough et al. 2011).

**(c) Decisión:** el módulo `evaluabilidad` es el paso 0 obligatorio del flujo. El informe conserva la estructura A7.1 e incorpora en anexo la metodología, la validez y limitaciones, y la reproducibilidad. La revisión de calidad la hace un subagente independiente.

---

## Limitaciones adicionales detectadas (no estaban en las hipótesis)

| # | Limitación | Evidencia CEPLAN | Implicancia |
|---|---|---|---|
| L1 | **Alcance seleccionado por la variable dependiente:** la evaluación prioriza los elementos con alerta (rojo) | SyE pp. 44, 82; A7.1 p1.3, p. 183 | Los éxitos y el sobrecumplimiento no se examinan, y se pierde poder explicativo. El repo exige justificar la selección de casos. |
| L2 | **Ciclo corto y fijo:** evaluación bienal con plazo el 15 de junio y recomendaciones de 24 meses como máximo | SyE pp. 57, 71, 77, 85 | No calza con los rezagos de resultados finales. Hay que separar lo evaluable cada dos años de lo que requiere horizonte largo. |
| L3 | **Calidad del dato:** la validación queda a cargo de las unidades de planeamiento, sin un marco de calidad de datos | SyE p. 35 (remite a Kusek y Rist) | Crítico para indicadores de violencia (subregistro) y registros administrativos. |
| L4 | **"Eficiencia" sin método** pese a ser un criterio principal | SyE pp. 19, 45 | Se cubre con el módulo `economica` y con reglas de no aplicación. |
| L5 | **Límites de páginas por sección** (resumen 1, descripción 2, resultados 10, implementación 20, conclusiones 1, recomendaciones 2) | A7.1 pp. 183–187 | Las secciones complementarias van en anexos para no incumplir. |
| L6 | **Línea de base redefinida en 2026** (periodo previo a la vigencia, o previo a la modificación de la ficha técnica) | RPCD 042-2026, art. 5 `[POR VERIFICAR texto oficial]` | Puede cambiar las líneas de base de la PNIG a mitad de la evaluación; hay que versionar las fichas con fecha de corte. |
