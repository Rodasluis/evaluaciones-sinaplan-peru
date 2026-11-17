# ADR-012. Implementado con al menos dos de los tres componentes

**Fecha:** 2026-11-11 · **Estado:** aprobada por el usuario

## Contexto
- **La escala de ADR-011** consideraba implementado un servicio solo si reunía los tres componentes:
  - protocolo aprobado;
  - recursos identificados en los datos del MEF;
  - reporte de indicadores con valor.

  Además, el servicio debía seguir operando al cierre, con ejecución reportada en 2024 o 2025. Si le faltaba cualquiera de los componentes, quedaba como implementado parcialmente.
- **Pedido del usuario (2026-11-11):** hacer más laxa la frontera entre implementado e implementado parcialmente. El servicio queda como parcial solo si le faltan al menos dos de los tres componentes.
- **Decisión del usuario sobre el cierre:** se mantiene la exigencia de seguir funcionando en 2024 o 2025.

## Decisión
- **Implementado:** al menos dos de los tres componentes, en algún año de 2019 a 2025, y ejecución reportada en 2024 o 2025.
- **Implementado parcialmente:** ejecución reportada en algún año de 2019 a 2025 con dos o más componentes faltantes, o interrumpida antes de 2024.
- **No implementado y sin información suficiente:** sin cambios respecto de ADR-011.

## Consecuencias
- **Estados:** pasan de 17 implementados, 34 parciales y 1 no implementado a 24, 27 y 1.
- **Servicios que suben a implementados:** a todos les falta un solo componente y todos reportan ejecución en 2024 y 2025.

| Servicio | Componente que le falta |
|---|---|
| 1.1.2 | Recursos identificados en el MEF |
| 1.1.5 | Reporte de indicadores |
| 1.2.6 | Reporte de indicadores |
| 3.1.2 | Recursos identificados en el MEF |
| 4.2.3 | Recursos identificados en el MEF |
| 4.3.7 | Recursos identificados en el MEF |
| 5.5.1 | Reporte de indicadores |

- **Lectura por objetivo prioritario:**
  - los implementados pasan a ser mayoría en el OP1, el OP3, el OP4 y el OP5;
  - la implementación parcial predomina en el OP2 y el OP6;
  - el OP2 sigue sin servicios implementados, porque a todos sus servicios les faltan el protocolo y el reporte de indicadores.
- **Protocolo:** los 24 implementados tienen protocolo aprobado; de los 27 parciales, solo uno lo tiene.
- **Hallazgos con cifras nuevas:** I1-H1 (estado de los servicios) y los textos que citan los conteos. Las valoraciones de las preguntas no cambian.
- **Archivos que cambian:**
  - `src/analisis_pnig/rev_doc_estado_servicios.py` (regla, docstring y hoja «Criterios»);
  - la figura de la escala;
  - el numeral 2.4 y el numeral 4.1 del informe;
  - la Tabla 9 de la nota metodológica (cambio directo; la versión previa está en `notas_metodologicas_anteriores\24_2026-11-11_previa-escala-dos-de-tres-sin-ogaj\`);
  - la presentación de resultados.
