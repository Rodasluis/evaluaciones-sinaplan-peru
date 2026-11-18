# Caso ficticio de prueba

**Datos de prueba, no de una política real.** La «Política Nacional Ficticia» (PNF, 2021–2024), sus servicios,
indicadores, estudios, documentos y entrevistas se inventaron para probar los instrumentos genéricos
(`tests/test_instrumentos.py`). No se citan en ningún informe.

| Carpeta | Contenido |
|---|---|
| `02_insumos/` | Entradas estándar: 3 objetivos, 6 servicios, 5 indicadores, 3 tipos de intervención y 8 estudios |
| `05_entrevistas/` | Preguntas, segmentos y dos codificaciones de dos roles ficticios |
| `documentos/` | Documento de texto donde se ubican las citas del alineamiento |

La prueba crea la evaluación con el generador en una carpeta temporal, copia estas entradas y corre los instrumentos.
