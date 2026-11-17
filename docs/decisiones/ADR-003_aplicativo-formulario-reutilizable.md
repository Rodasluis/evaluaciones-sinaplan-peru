# ADR-003. Se conserva el aplicativo de formulario electrónico para otras evaluaciones

**Fecha:** 2026-10-27 · **Estado:** aprobada por el usuario

## Contexto
El aplicativo FIE-03 (Google Apps Script vía clasp) quedó construido, desplegado y probado, pero no se usó en la PNIG (ADR-002). Sus características:
- acceso con usuario y contraseña por entidad;
- formulario prellenado;
- consolidación automática;
- cierre programado.

## Decisión
- El código se mantiene en `src/formulario_fie03/`, con una guía de reutilización (`LEEME.md`).
- `Semilla.js`, que contiene los hashes de contraseñas, sigue fuera del control de versiones.
- Las credenciales legibles se guardan fuera del repositorio.

## Consecuencias
- **Partes específicas de la PNIG:** las identifica el `LEEME.md` y deben adaptarse en cada evaluación: `SHEET_ID`, `scriptId`, preguntas, semilla y textos de los documentos.
- **Despliegue:** el de la PNIG sigue publicado. Retirarlo es decisión del usuario.
