"""Núcleo reutilizable para evaluar políticas nacionales y planes del Sinaplan (Fase 2, ADR-013).

Cada evaluación vive en evaluaciones/<id>/ con su config.yaml; su capa propia (hallazgos, relato del informe) vive en
src/analisis_<id>/ y usa estos módulos. La evaluación activa se elige con la variable de entorno EVALUACION (por
defecto, pnig_2026, el piloto).
"""
