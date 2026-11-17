"""Diseño G, centrado (ronda 28; único vigente desde la ronda 35), basado en la portada del informe de evaluación 2023 de la PNIG del MIMP: logos arriba,
en los extremos; todo centrado: la etiqueta en mayúsculas y en negrita, el título en rojo y en negrita (el informe de
2023 lo usaba en azul; aquí no hay azul), el subtítulo, la institución en gris oscuro y en negrita, el órgano
elaborador en gris y el lugar y fecha en negrita. Sin número de página."""
from .comun import A4_CM, ajustar_titulo, fila_logos, lineas_organo, logos, parrafo, separador, tabla

X = 2.5
Y_LOGOS, ALTO_LOGOS = 2.0, 1.36
Y_TITULO, Y_INSTITUCION, Y_FECHA = 6.4, 16.8, 24.0
PT_ETIQUETA, PT_TITULO, PT_SUBTITULO, PT_INSTITUCION, PT_ORGANO, PT_FECHA = 17, 35, 17.5, 17, 14, 17


def construir(documento, cfg, paleta, avisos, titulo):
    ancho_pag, _ = A4_CM
    f = paleta.fuente
    ancho = ancho_pag - 2 * X
    fila_logos(documento, X, Y_LOGOS, logos(documento, cfg, paleta, ALTO_LOGOS, avisos), ALTO_LOGOS, paleta,
               ancho_total_cm=ancho)
    separador(documento, f)
    pt, _ = ajustar_titulo(titulo, paleta.ttf_negrita, PT_TITULO, ancho, avisos)
    textos = []
    if str(cfg.get("etiqueta") or "").strip():
        textos.append(parrafo(f, [(cfg["etiqueta"].strip(), {"pt": PT_ETIQUETA, "b": True, "color": paleta.texto,
                                                             "caps": True})], jc="center", despues=24))
    textos.append(parrafo(f, [(titulo, {"pt": pt, "b": True, "color": paleta.primario})], jc="center", despues=16))
    if str(cfg.get("subtitulo") or "").strip():
        textos.append(parrafo(f, [(cfg["subtitulo"].strip(), {"pt": PT_SUBTITULO, "color": paleta.gris_oscuro})],
                              jc="center"))
    tabla(documento, X, Y_TITULO, [ancho], [{"celdas": [{"parrafos": textos}]}])
    separador(documento, f)
    bloque = [parrafo(f, [(cfg["institucion"].strip(), {"pt": PT_INSTITUCION, "b": True,
                                                        "color": paleta.gris_oscuro})], jc="center", despues=8)]
    bloque += [parrafo(f, [(t, {"pt": PT_ORGANO, "color": paleta.gris_oscuro})], jc="center", despues=3)
               for t in lineas_organo(cfg)]
    tabla(documento, X, Y_INSTITUCION, [ancho], [{"celdas": [{"parrafos": bloque}]}])
    separador(documento, f)
    tabla(documento, X, Y_FECHA, [ancho], [{"celdas": [{"parrafos": [
        parrafo(f, [(cfg["lugar_fecha"].strip(), {"pt": PT_FECHA, "b": True, "color": paleta.texto})],
                jc="center")]}]}])
