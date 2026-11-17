# Tipologías de evaluación

Cada carpeta describe un tipo de evaluación de políticas nacionales y planes del Sinaplan en `estructura.yaml`:
- definición y elementos que se valoran, según la Guía para el seguimiento y evaluación de políticas nacionales y planes del Sinaplan (CEPLAN, 2024, versión actualizada en `docs/fuentes/ceplan/`), con su página;
- criterios y preguntas tipo;
- instrumentos sugeridos (códigos de `config.yaml`);
- escalas;
- estructura del informe;
- requisitos de `docs/marco/matriz_requisitos_ceplan.csv` que aplican.

| Tipología | Estado | Base |
|---|---|---|
| `cierre/` | Probada en el piloto PNIG 2026 | Guía, p. 73; informe de cierre de la PNIG (2019–2025) |
| `diseno/` | Estructura sin probar | Guía, pp. 20–21 |
| `implementacion/` | Estructura sin probar | Guía, p. 21 |
| `resultados/` | Estructura sin probar | Guía, p. 21 |

**Cómo se usan:**
- `src/nucleo/nueva_evaluacion.py <id> --tipologia <tipo>` copia la plantilla de evaluación y precarga en `config.yaml` las preguntas tipo y la estructura del informe.
- El equipo evaluador ajusta las preguntas a la política, porque las preguntas tipo son un punto de partida.

**Planos:** lo que exige el CEPLAN (con página), lo que aporta la experiencia del piloto y lo que decide el equipo evaluador se distinguen en cada archivo. Lo no verificado en la guía se marca `[POR VERIFICAR]`.

**Integralidad:** según la guía (p. 21), los informes de evaluación de las políticas nacionales y planes son integrales: aunque la evaluación se centre en un tipo, el informe incluye información de los análisis de resultados, de implementación y de diseño.
