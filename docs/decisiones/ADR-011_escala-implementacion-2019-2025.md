# ADR-011. Escala de implementación de los servicios sobre todo el periodo 2019–2025

**Fecha:** 2026-11-10 · **Estado:** aprobada por el usuario

## Contexto
- **La escala anterior:** venía de la nota metodológica (numeral 9) y medía la ejecución de los servicios solo en 2023–2025.
  - Implementado: ejecución en 2024 o 2025, además de los tres componentes.
  - Parcial: ejecución en 2023–2025.
  - No implementado: sin ejecución en 2023–2025.
  - Los recursos se identificaban con los montos reportados en 2023–2025.
- **Observación del usuario (2026-11-10):** el periodo evaluado es 2019–2025. Los servicios reportan desde 2019 en los informes LIO, en los reportes de seguimiento y de cumplimiento y en el consolidado de indicadores. Leer solo 2023–2025 subestima la ejecución.
- **Datos disponibles:** la base ya tenía la información de 2019 a 2025.
  - Cobertura de los informes LIO de 2019 a 2025 por servicio y año. El informe LIO «2019» corresponde al periodo 2019.
  - Valores de los indicadores de servicio de 2019 a 2025 en el consolidado de la OMEP.
  - Montos verificados en los datos abiertos del MEF para todos los años: 53, 16, 90 y 105 de 2019 a 2022.

## Decisión
- **Componentes:** el protocolo aprobado, los recursos identificados en los datos del MEF y el reporte de indicadores se buscan en todo el periodo 2019–2025.
- **Implementado:** reúne los tres componentes y sigue operando al cierre, con ejecución reportada en 2024 o 2025. La exigencia del cierre es una decisión del usuario.
- **Implementado parcialmente:** ejecución reportada en algún año de 2019–2025 sin alguno de los componentes, o interrumpida antes de 2024.
- **No implementado:** sin ejecución en ningún año de 2019–2025, o solo con acciones preparatorias.
- **Sin información suficiente:** solo secciones mínimas del informe LIO, sin valores de indicador.

## Lectura de los servicios antes clasificados como no implementados
Se revisaron sus secciones de los informes LIO de 2019 a 2022 (páginas del PDF):

| Servicio | Antes | Lectura de 2019–2022 | Ahora |
|---|---|---|---|
| 2.2.1 | No implementado | Orientación en educación sexual integral a familias en 2021 y 2022 (LIO 2021, p. 157; LIO 2022, p. 94). El LIO 2023 dice que el servicio no se implementó como tal | Implementado parcialmente |
| 5.2.1 | No implementado | Acciones preparatorias en 2019 (LIO 2019, p. 59), «sin logro» en 2021 (LIO 2021, p. 517) y sin información en 2022 por ser un servicio nuevo (LIO 2022, p. 151) | No implementado |
| 6.1.1 | No implementado | Programas educativos en 2021 y 2022 (LIO 2021, p. 246; LIO 2022, p. 163) | Implementado parcialmente |

## Consecuencias
- **Estados:** pasan de 16 implementados, 33 parciales y 3 no implementados a 17, 34 y 1.
- **Cambios:**
  - 4.4.1 pasa a implementado, porque tiene montos verificados antes de 2023;
  - 2.2.1 y 6.1.1 pasan a parciales.
- **Recursos identificados:** los servicios con recursos pasan de 30 a 34.
- **Archivos:**
  - `src/analisis_pnig/rev_doc_estado_servicios.py` (criterios y ajustes);
  - el filtro de montos de `informe_cierre.py`;
  - la escala del numeral 2.4, ahora un diagrama de decisión;
  - el Anexo 3 y los hallazgos que dependen del estado.
- **Nota metodológica:** al principio quedó fuera de alcance, porque la edita el usuario. En la ronda 34 el usuario pidió actualizarla: la Tabla 9, el criterio de asignación identificable (numeral 9), la limitación sobre el grado de implementación y la sección B del Anexo 1 aplican ahora la escala de 2019–2025. El informe la declara en su numeral 2.4.

## Terminología (2026-11-10, ronda 34)
- **«Información detallada»** reemplaza a «sección sustantiva» (pedido del usuario). Un año tiene ejecución reportada cuando el informe LIO trae información detallada del servicio, es decir, una sección que describe sus acciones sin declarar falta de información, o cuando alguno de sus indicadores tiene valor. El criterio no cambia; solo el nombre.
- **Ronda 35 (2026-11-11):**
  - a pedido del usuario, los documentos ya no mencionan la extensión mínima de la sección; el umbral queda solo como regla interna de `rev_doc_estado_servicios.py` (`MIN_LINEAS`);
  - la regla de los componentes cambió a dos de tres (ADR-012).
