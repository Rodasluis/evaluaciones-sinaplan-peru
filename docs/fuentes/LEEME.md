# Fuentes documentales

Los documentos de las fuentes (guías y directivas del CEPLAN, normas, documentos de la política, informes de
seguimiento y evaluación, planes, instrumentos internacionales y fichas técnicas del INEI) no se distribuyen con el
repositorio: pesan cerca de 900 MB y cada uno tiene su portal oficial.

`inventario_publico.csv` lista cada documento con:
- `ruta`: dónde lo esperan el código y las citas del informe;
- `enlace` y `pagina_publicacion`: el archivo y la página oficiales;
- `sha256`: la huella del archivo publicado, para verificar que es el mismo que se usó;
- `sha256_corpus`: la huella de la copia del equipo evaluador (distinta cuando se le aplicó OCR para extraer el texto);
- `nota`: los casos en que la versión publicada difiere de la usada.

## Descarga
```
uv run python src/fuentes/descargar_inventario.py                  # todo el inventario
uv run python src/fuentes/descargar_inventario.py --carpeta ceplan  # una carpeta
uv run python src/fuentes/descargar_inventario.py --muestra 5      # prueba
```
El script no vuelve a descargar lo que ya está con su huella e informa al final los enlaces caídos y los archivos con
otra huella. Los portales cambian: si un enlace deja de responder, la página de publicación permite ubicar el documento.

## Lo que no está
- **Documentos sin publicación oficial** que la evaluación del piloto usó (actas, documentos de trabajo y
  información remitida por órganos de la entidad). El informe los cita; pueden solicitarse a la entidad conforme a la
  normativa de transparencia y acceso a la información pública.
- **Microdatos** (ENAHO, ENDES, ENARES, ENUT) y **datos abiertos del MEF**: se descargan de los portales del INEI y
  del MEF con los scripts de `src/fuentes/` y van en `evaluaciones/<id>/03_datos/`, que no se versiona.
