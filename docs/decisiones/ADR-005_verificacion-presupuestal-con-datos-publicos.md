# ADR-005. La información presupuestal se verifica con los datos abiertos del MEF, sin solicitar información adicional

**Fecha:** 2026-10-30 · **Estado:** aprobada por el usuario

## Contexto
La nota v13 trataba la información presupuestal como aproximación:
- **I3 (conducción):** identificaba las actividades del MIMP por su denominación en la Consulta Amigable, que no separa el gasto por órgano.
- **I1 (asignación identificable en GRD-01):** bastaba con que el servicio figurara en el POI 2023 del formulario o con que el LIO trajera montos en soles, sin contrastarlos.

Los datos abiertos del MEF «Presupuesto y Ejecución de Gasto» son la Consulta Amigable con el detalle de meta, y cuadran al sol con ella en el pliego 039 entre 2019 y 2025. Con ellos se pueden separar los órganos y verificar las cifras reportadas (`src/analisis_pnig/psi05_3_verificacion_presupuestal.py`).

## Decisión
- **PSI-05.3:** el gasto de la conducción se separa por las metas de los órganos conductores, según el ROF de cada año, y se contrasta con el POI consistente con el PIA. Las cifras de los informes LIO 2019–2025 y del formulario 2023 se verifican contra los mismos datos.
- **GRD-01:** la asignación presupuestal es identificable solo si al menos un monto reportado para el servicio en 2023–2025 coincide con una meta, una actividad o un producto de la entidad responsable. No cuentan:
  - los totales de programa;
  - los pliegos de otro sector;
  - los montos redondos marcados como posiblemente casuales.
- **Solo datos públicos:** no se piden notas de modificación presupuestaria a la OGPPM ni aclaraciones a AURORA.
  - Los traslados entre órganos se infieren de tres señales: el saldo PIM − PIA, las metas creadas en el año y la certificación anulada al cierre.
  - Lo que esas señales no explican se consigna como diferencia identificada en las metas: por ejemplo, las metas de Lima que los informes LIO 2021 y 2022 excluyen del total de los CEM, o el devengado de los CEM de 2025.
- **Documentos:** nota v14 y presentación completa actualizadas; la resumida no se modifica.

## Consecuencias
- **GRD-01:** pasa a tener 16 servicios implementados, 33 implementados parcialmente y 3 no implementados (antes, 19, 30 y 3).
  - Bajan a «implementado parcialmente» 1.1.2, 3.1.2 y 4.4.1, porque ninguno de sus montos 2023–2025 coincide con una estructura propia en el MEF.
  - Ganan asignación verificada 4.2.1, 4.3.2 y 6.3.2.
- **I3:** deja de ser una aproximación por denominación; su límite es que el PIA y el PIM solo existen como cifra anual por meta.
- **Archivos:** los ZIP anuales del MEF (4,4 GB) están en `datos_abiertos_mef_gasto\` de la carpeta administrativa, con sus hashes en `docs/fuentes/reorganizacion_2026-10-30.tsv`.
