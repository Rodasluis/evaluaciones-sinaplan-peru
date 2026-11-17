/*
Resultados de las Elecciones Regionales y Municipales 2026 (ONPE): ganador por provincia y distrito, y su sexo por el
primer nombre. Se ejecuta en la consola del navegador, en una pestaña abierta en
https://resultadoelectoral.onpe.gob.pe/main/resumen: el portal tiene verificación antibots, así que no se consulta con
scripts fuera del navegador. Usa la misma API que la página (presentacion-backend), una consulta a la vez y con pausa.

Todo va dentro de un objeto (window.__onpe) para no pisar variables globales del portal: con nombres cortos la
pestaña se congeló. Pasos (ejecutados el 2026-11-02 con Claude in Chrome en el navegador del usuario):
  1. await __onpe.preparar();                // catálogo de ubigeos del portal (códigos del RENIEC)
  2. await __onpe.lote(100) ... hasta el total // 196 provincias y 1892 distritos; 204 = capital sin elección distrital
  3. __onpe.conteo();                        // mujeres, hombres y no determinados por ámbito
  4. __onpe.dudosos();                       // ganadores sin sexo determinado, para revisarlos a mano (2026-11-03)
  5. await __onpe.regional();                // gobernador regional: dos fórmulas más votadas por circunscripción
  6. await __onpe.actas();                   // porcentaje de actas contabilizadas por elección
__onpe.correr() hace el paso 2 completo en segundo plano dentro de la página; __onpe.avance() informa el progreso.

Regla de sexo (decisión del usuario, 2026-11-02): por el primer nombre; si es ambiguo, por el segundo nombre, siempre
que el nombre completo tenga al menos cuatro palabras (con tres, la segunda es un apellido). JOSE y JUAN son masculinos
aunque sigan MARIA o MERCEDES. Terminación en A = mujer, salvo MASC_A. Las listas se revisaron a mano sobre los 1847
ganadores con nombre: se corrigieron ELIA (mujer), RUNTSMA (hombre), MERCY y NATALY (mujer); ARLEN y GABY con un solo
nombre y ACENCION CRUZ quedan no determinados. La validación con los 130 diputados de la Resolución N.° 1411-2026-JNE
(d2_autoridades_2026.py) obligó a agregar once nombres femeninos y la regla del segundo nombre femenino; con ellas se
reproduce la Tabla 3 de la ONPE en los seis partidos. Si la ONPE no publica el nombre del candidato de la organización
ganadora (44 casos), el sexo queda no determinado. Solo salen conteos: ningún nombre se guarda en el repositorio.

Segunda consulta (2026-11-02, tarde): mismas 2088 ubicaciones, con el mismo conteo por ámbito. La pestaña del portal
se congeló dos veces y la ONPE respondió 403 una vez (verificación antibots vencida): se recargó la página principal
para renovarla, sin eludirla, y se siguió desde /assets/ubig/v1.json, con el avance guardado en localStorage (borrado
al terminar). Chrome frena los temporizadores de las pestañas en segundo plano, así que la pausa real fue mayor que la
indicada. De los 47 casos sin sexo, solo Lima Metropolitana (provincial) se resolvió con la prensa; los otros 46 quedan
sin determinar. regional() y actas() se agregaron en esta consulta.
*/
window.__onpe = (() => {  const norm = s => (s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase().trim();
  const REGLAS = {
    FEM: new Set(('CARMEN ROSARIO MARISOL RUTH BEATRIZ ISABEL INES RAQUEL SOLEDAD MERCEDES DOLORES NIEVES LUZ ESTHER ESTER ' +
      'MIRIAM MIRIAN NOEMI MARLENY MARLENI MARLENE NELLY GLADYS ROCIO CONSUELO PILAR MARIBEL ELIZABETH ELISABETH SHIRLEY ' +
      'KAREN LISBETH LIZBETH LIZ YENNY JENNY YENI ROSMERY JUDITH JUDIT EDITH MILAGROS ANGELES IRIS ABIGAIL LOURDES MARY ' +
      'BETTY NANCY EVELYN EVELING JAKELIN JAQUELINE JACQUELINE KATHERINE LEILY LEINY HEYDE DEICE EILEN ELENITH FLOR MARGOT ' +
      'ANALI LEYDI KELLY SUSAN YANET JANET MAGALY ARACELY ANABEL DORIS LUCY CLARIBEL PILY ROSSMERY MIRTHA GRISELDA MERCY ' +
      'NATALY NATALI CATHERIN CATHERINE HAYDEE JEENNY JENIFER YENIFER MADY MERY NARY NATHALY PIERANGELI YULI KIM').split(' ')),
    MASC_A: new Set('JOSHUA LUCA BAUTISTA GARCIA RUNTSMA'.split(' ')),
    MASC_SEG: new Set('JOSE JUAN WILMER'.split(' ')),   // pueden ir seguidos de un segundo nombre femenino
    AMBIG: new Set('ROSARIO GUADALUPE TRINIDAD CRUZ ASUNCION ACENCION CONCEPCION ADIS ESLI ESLY NIVE EALY ARLEN ARLA SANTA GABY'.split(' ')),
  };

  function sexo(nom) {
    const t = norm(nom).split(/\s+/);
    const una = x => {
      if (!x) return '?';
      if (x === 'MARIA') return 'F';
      if (REGLAS.AMBIG.has(x)) return '?';
      if (REGLAS.FEM.has(x)) return 'F';
      if (x.endsWith('A') && !REGLAS.MASC_A.has(x)) return 'F';
      return 'M';
    };
    if (t[0] === 'JOSE' || t[0] === 'JUAN') return 'M';
    const s = una(t[0]);
    // primer nombre poco frecuente y segundo nombre femenino (Mery Eliana, Yuli Liliana)
    if (s === 'M' && t.length >= 4 && !REGLAS.MASC_SEG.has(t[0]) && una(t[1]) === 'F') return 'F';
    if (s !== '?') return s;
    if (t.length <= 3) return '?';
    return una(t[1]);
  }

  const E = {tareas: [], i: 0, filas: [], errores: []};

  async function preparar() {
    const ub = await fetch('/assets/ubig/v1.json').then(r => r.json());
    E.nombres = Object.fromEntries(ub.map(x => [x.c_ubigeo, x.c_nombre]));
    for (const x of ub) {
      if (!/^\d{6}$/.test(x.c_ubigeo)) continue;
      if (x.c_ubigeo.slice(2, 4) !== '00' && x.c_ubigeo.slice(4) === '00') E.tareas.push({e: 3, u: x.c_ubigeo});
    }
    for (const x of ub) if (/^\d{6}$/.test(x.c_ubigeo) && x.c_ubigeo.slice(4) !== '00') E.tareas.push({e: 4, u: x.c_ubigeo});
    return E.tareas.length;
  }

  async function lote(n) {
    const B = '/presentacion-backend/resumen-general/participantes?';
    const fin = Math.min(E.i + n, E.tareas.length);
    while (E.i < fin) {
      const t = E.tareas[E.i];
      const dep = t.u.slice(0, 2) + '0000', pro = t.u.slice(0, 4) + '00';
      const q = t.e === 3
        ? `idEleccion=3&tipoFiltro=ubigeo_nivel_02&idAmbitoGeografico=1&idUbigeoDepartamento=${dep}&idUbigeoProvincia=${pro}`
        : `idEleccion=4&tipoFiltro=ubigeo_nivel_03&idAmbitoGeografico=1&idUbigeoDepartamento=${dep}&idUbigeoProvincia=${pro}&idUbigeoDistrito=${t.u}`;
      const r = await fetch(B + q);
      if (r.status === 204) E.filas.push({e: t.e, u: t.u, n: 0});
      else if (r.status !== 200) { E.errores.push(`${t.u}:${t.e}:${r.status}`); return 'detenido'; }   // no se insiste
      else {
        const d = ((await r.json()).data || []).sort((a, b) => b.totalVotosValidos - a.totalVotosValidos);
        E.filas.push({e: t.e, u: t.u, n: d.length, nom: d[0] ? d[0].nombreCandidato : '',
                      org: d[0] ? d[0].nombreAgrupacionPolitica : '',
                      v1: d[0] ? d[0].totalVotosValidos : 0, v2: d[1] ? d[1].totalVotosValidos : 0});
      }
      E.i++;
      await new Promise(res => setTimeout(res, 100));
    }
    return `${E.i} de ${E.tareas.length}`;
  }

  function conteo() {
    const c = {};
    for (const f of E.filas) {
      if (f.n === 0) continue;
      const dep = norm(E.nombres[f.u.slice(0, 2) + '0000']), prov = norm(E.nombres[f.u.slice(0, 4) + '00']);
      const amb = dep === 'LIMA' ? (prov === 'LIMA' ? 'LM' : 'RL') : dep;
      const s = f.nom ? sexo(f.nom) : 'S';   // S: la ONPE no publica el nombre del candidato ganador
      const k = `${amb}|${f.e === 3 ? 'P' : 'D'}`;
      c[k] = c[k] || {F: 0, M: 0, '?': 0, S: 0};
      c[k][s]++;
    }
    return c;
  }
  function correr() {
    E.estado = 'en curso';
    lote(E.tareas.length).then(x => { E.estado = x; }).catch(e => { E.estado = 'error ' + e; });
    return 'iniciado';
  }

  const avance = () => `${E.i} de ${E.tareas.length}; ${E.estado}; errores: ${E.errores.join(' ')}`;

  // Ganadores sin sexo determinado. El nombre queda solo en la memoria de la pestaña para revisarlo a mano.
  const dudosos = () => E.filas.filter(f => f.n > 0 && (!f.nom || sexo(f.nom) === '?'))
    .map(f => ({u: f.u, e: f.e, lugar: `${E.nombres[f.u.slice(0, 2) + '0000']}/${E.nombres[f.u]}`, nom: f.nom, org: f.org}));

  // Gobernador regional (idEleccion=1): las dos fórmulas más votadas, su sexo y su porcentaje de votos válidos. La
  // segunda elección regional se hace donde ninguna fórmula llega al 30 % de los votos válidos.
  async function regional() {
    const ub = await fetch('/assets/ubig/v1.json').then(r => r.json());
    E.regional = [];
    for (const x of ub.filter(y => /^\d\d0000$/.test(y.c_ubigeo) && y.c_ubigeo !== '000000')) {
      const r = await fetch('/presentacion-backend/resumen-general/participantes?idEleccion=1&tipoFiltro=ubigeo_nivel_01' +
                            `&idAmbitoGeografico=1&idUbigeoDepartamento=${x.c_ubigeo}`);
      if (r.status !== 200) { E.regional.push({dep: x.c_nombre, st: r.status}); continue; }
      const d = ((await r.json()).data || []).sort((a, b) => b.totalVotosValidos - a.totalVotosValidos);
      const f = i => d[i] ? {s: d[i].nombreCandidato ? sexo(d[i].nombreCandidato) : 'S', pv: d[i].porcentajeVotosValidos,
                             nom: d[i].nombreCandidato} : {s: '', pv: null, nom: ''};
      E.regional.push({dep: norm(x.c_nombre), n: d.length, p1: f(0), p2: f(1)});
      await new Promise(res => setTimeout(res, 150));
    }
    return E.regional.length;
  }

  // Porcentaje de actas contabilizadas por elección (1 regional, 3 provincial, 4 distrital), del resumen del portal
  async function actas() {
    const o = {};
    for (const e of [1, 3, 4]) {
      const r = await fetch(`/presentacion-backend/resumen-general/totales?idEleccion=${e}&tipoFiltro=ambito_geografico&idAmbitoGeografico=1`);
      const t = r.status === 200 ? (await r.json()).data : null;
      o[e] = t ? {actas: t.actasContabilizadas, fecha: t.fechaActualizacion, total: t.totalActas} : r.status;
    }
    return o;
  }

  return {REGLAS, sexo, preparar, lote, conteo, correr, avance, dudosos, regional, actas, E};
})();
