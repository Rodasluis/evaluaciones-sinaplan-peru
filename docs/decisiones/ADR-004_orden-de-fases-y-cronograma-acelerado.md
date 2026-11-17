# ADR-004. Orden de fases: Fase 1 corta, piloto PNIG y después las Fases 2 a 4

**Fecha:** 2026-10-27 · **Estado:** aprobada por el usuario

## Contexto
El plan original preveía, antes del piloto de la PNIG (F5), cuatro fases de infraestructura reutilizable:
- F1: entorno de trabajo;
- F2: núcleo técnico;
- F3: marco teórico;
- F4: módulos por tipo de evaluación.

El piloto, sin embargo, ya está en ejecución y el usuario fijó como meta terminarlo en la semana del 2 al 6 de noviembre de 2026.

## Decisión
1. **Fase 0:** se aprueba con dos pendientes: completar el registro de fuentes y poner el repositorio bajo control de versiones.
2. **Fase 1 corta:** `CLAUDE.md` definitivo, notas de decisión y git. Los subagentes y comandos se postergan.
3. **Piloto:** se ejecuta de inmediato con un cronograma interno acelerado. El plan de 30 días de la nota metodológica, hasta el 18-11-2026, se mantiene como compromiso formal.
4. **Fases 2 a 4:** se retoman después del piloto, sobre lo que haya quedado probado.

## Consecuencias
- **Riesgo:** los módulos genéricos se construirán a partir de un caso real, con el riesgo de sesgarlos hacia la PNIG.
- **Mitigación:** registrar en cada script qué parte es específica del piloto.
- **Validaciones externas:** las reuniones con la DGIGND, el Despacho Viceministerial y el CEPLAN dependen de sus agendas y pueden quedar fuera de la semana objetivo.
