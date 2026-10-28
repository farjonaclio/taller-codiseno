/**
 * Taller de codiseño – Modelo de análisis de información
 * Servidor de la herramienta web del taller (versión 6: un solo enlace, trabajo en plenario).
 * La página vive en GitHub Pages y llama a este servidor (doPost, JSON). Las respuestas quedan en la hoja.
 * Un solo enlace para celular y computador: 1 Para qué · 2 Preguntas · Votar · 3 Qué tenemos · 4 Cómo se ve.
 * ?vista=resultados es la vista del proyector (con modo facilitador).
 *
 * Instalación: Apps Script con Code.gs y Datos.gs. Implementar > Aplicación web
 *   Ejecutar como: Yo   ·   Quién tiene acceso: Cualquier persona
 * (los invitados no tienen cuenta de CliO ni Gmail). Copiar la URL /exec en config.js del repositorio.
 *
 * No se piden nombres ni correos: solo área y, si el facilitador los activa, grupo
 * (minimización de datos personales, Ley 1581 de 2012). No escribir datos de personas atendidas.
 */

const MAX_TEXTO = 500;
const MAX_VOTOS = 5;
const MAX_GRUPOS = 6;
const PRIORIZADAS_SUGERIDAS = 4;

// TALLER (ideas de para qué, ejemplos y destinatarios), AREAS, FUENTES, CLAVE_FACILITADOR y HOJA_ID están en Datos.gs,
// que no se publica en GitHub porque contiene información del cliente.
const FRECUENCIAS = ['Diaria', 'Semanal', 'Mensual', 'Trimestral', 'Semestral', 'Anual', 'Por demanda', 'No sé'];
const TENEMOS = ['Sí', 'Parcial', 'No', 'No sé'];
const LOGRABLE = ['Sí', 'Con ajustes', 'No por ahora'];
const HORIZONTE = ['Prioritaria', 'Para más adelante'];

function doGet(e) {
  if (e && e.parameter && e.parameter.accion) return api_(e.parameter.accion, null);
  try {
    return HtmlService.createHtmlOutputFromFile('Index').setTitle('Taller de codiseño').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    return ContentService.createTextOutput('Servidor del taller activo.');
  }
}

/* ---------- API para la página en GitHub Pages ---------- */
function doPost(e) {
  let b = {};
  try { b = JSON.parse(e.postData.contents); } catch (err) { return salida_({ ok: false, error: 'Solicitud no válida' }); }
  return api_(b.accion, b.arg);
}
function api_(accion, arg) {
  const ACCIONES = { obtenerConfig: obtenerConfig, obtenerResultados: obtenerResultados, guardarParaQue: guardarParaQue, guardarPregunta: guardarPregunta,
    guardarVotos: guardarVotos, guardarInsumo: guardarInsumo, borrarInsumo: borrarInsumo, guardarEvaluacion: guardarEvaluacion, agregarPriorizada: agregarPriorizada,
    fijarDimensiones: fijarDimensiones, ocultarPregunta: ocultarPregunta, fijarActores: fijarActores, fijarGrupos: fijarGrupos, fijarPriorizadas: fijarPriorizadas,
    cargarSimulacion: cargarSimulacion, borrarSimulacion: borrarSimulacion, limpiarTodo: limpiarTodo };
  try {
    if (!ACCIONES[accion]) throw new Error('Acción desconocida');
    return salida_({ ok: true, data: ACCIONES[accion](arg) });
  } catch (err) {
    return salida_({ ok: false, error: err.message });
  }
}
function salida_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function validarClave_(r) { if (!r || String(r.clave || '') !== String(CLAVE_FACILITADOR)) throw new Error('Clave de facilitador incorrecta'); }

function obtenerConfig() {
  return { dimensiones: TALLER.dimensiones, ejemplos: TALLER.ejemplos, destinos: destinos_(), activas: activas_(), grupos: grupos_(),
    areas: AREAS, fuentes: FUENTES, frecuencias: FRECUENCIAS, tenemos: TENEMOS, lograble: LOGRABLE, horizonte: HORIZONTE,
    responsables: responsables_(), maxVotos: MAX_VOTOS };
}
// Sugerencias para "quién produce" y "responsable": áreas del taller y actores. Se puede escribir otro.
function responsables_() {
  const out = AREAS.filter(function (a) { return a !== 'Otra'; });
  nombresDestinos_().forEach(function (d) { if (out.indexOf(d) < 0) out.push(d); });
  return out.concat(['Otra entidad', 'Por definir']);
}

/* ---------- Hojas ---------- */
const HOJAS = {
  Config: ['Clave', 'Valor'],
  ParaQue: ['Fecha', 'Envío', 'Área', 'Grupo', 'Para qué', 'Ejemplo de decisión'],
  Preguntas: ['Fecha', 'ID', 'Área', 'Grupo', 'Pregunta', 'Para qué (dimensiones)', 'A quién se reporta', 'Decisión', 'Oculta', 'Frecuencia', 'Qué se mide', 'Dispositivo'],
  Votos: ['Fecha', 'Envío', 'Área', 'Grupo', 'ID'],
  Insumos: ['Fecha', 'Insumo', 'Pregunta', 'Fuente', 'Quién la produce', 'Frecuencia del dato', '¿La tenemos?', 'Comentario', 'Grupo', 'Dispositivo'],
  Evaluacion: ['Fecha', 'Pregunta', '¿Es lograble?', 'Horizonte', 'Qué falta', 'Cómo se presenta', 'Responsable']
};
function libro_() { return (typeof HOJA_ID !== 'undefined' && HOJA_ID) ? SpreadsheetApp.openById(HOJA_ID) : SpreadsheetApp.getActiveSpreadsheet(); }
function hoja_(nombre) {
  const ss = libro_(), cols = HOJAS[nombre];
  let sh = ss.getSheetByName(nombre);
  if (!sh) sh = ss.insertSheet(nombre);
  if (sh.getLastRow() === 0 || sh.getRange(1, 1, 1, cols.length).getValues()[0].join('|') !== cols.join('|')) {
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function filas_(nombre) {
  const sh = hoja_(nombre);
  const n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, HOJAS[nombre].length).getValues() : [];
}
function limpiar_(v) { return String(v == null ? '' : v).replace(/[\u0000-\u001F]/g, ' ').trim().slice(0, MAX_TEXTO); }
function enLista_(v, lista) { return lista.indexOf(v) >= 0 ? v : ''; }
function lista_(arr, validos, max) {
  return (Array.isArray(arr) ? arr : []).map(limpiar_).filter(function (x, i, a) { return x && a.indexOf(x) === i && (!validos || validos.indexOf(x) >= 0); }).slice(0, max || 12);
}
function corto_(v) { return limpiar_(v).slice(0, 80); }
// Identificador anónimo del dispositivo (lo genera la página; no identifica a la persona). Permite editar y evita duplicados.
function token_(r) { const v = String(r && r.dispositivo || ''); return /^[a-z0-9]{8,24}$/.test(v) ? v : ''; }
function quitarFilas_(nombre, prueba) {
  const sh = hoja_(nombre), datos = filas_(nombre);
  for (let i = datos.length - 1; i >= 0; i--) if (prueba(datos[i])) sh.deleteRow(i + 2);
}
function conLock_(fn) { const lock = LockService.getScriptLock(); lock.waitLock(20000); try { return fn(); } finally { lock.releaseLock(); } }
function grupo_(r) { const n = grupos_(); const g = String(r && r.grupo || ''); return n && Number(g) >= 1 && Number(g) <= n ? g : ''; }
function validarPersona_(r) {
  if (!r || AREAS.indexOf(r.area) < 0) throw new Error('Seleccione su área.');
  return { area: r.area, grupo: grupo_(r) };
}
function config_(clave) {
  const f = filas_('Config').filter(function (x) { return x[0] === clave; })[0];
  try { return f ? JSON.parse(f[1]) : null; } catch (e) { return null; }
}
function fijarConfig_(clave, valor) { upsert_('Config', function (f) { return f[0] === clave; }, [clave, JSON.stringify(valor)]); }
function upsert_(nombre, coincide, fila) {
  const sh = hoja_(nombre), datos = filas_(nombre);
  for (let i = 0; i < datos.length; i++) if (coincide(datos[i])) { sh.getRange(i + 2, 1, 1, fila.length).setValues([fila]); return; }
  sh.getRange(sh.getLastRow() + 1, 1, 1, fila.length).setValues([fila]);
}
function escribir_(nombre, filas) {
  if (!filas.length) return;
  const sh = hoja_(nombre);
  sh.getRange(sh.getLastRow() + 1, 1, filas.length, HOJAS[nombre].length).setValues(filas);
}

/* ---------- Configuración que fija el facilitador ---------- */
function grupos_() { const v = Number(config_('grupos')); return v >= 1 && v <= MAX_GRUPOS ? v : 0; }
function fijarGrupos(r) {
  validarClave_(r);
  const n = Math.max(0, Math.min(MAX_GRUPOS, Math.floor(Number(r.grupos) || 0)));
  conLock_(function () { fijarConfig_('grupos', n); });
  return { ok: true, grupos: n };
}
function activas_() { const v = config_('dimensiones'); return Array.isArray(v) && v.length ? v : null; }
function fijarDimensiones(r) {
  validarClave_(r);
  const v = lista_(r.lista, null, 8).map(function (x) { return x.slice(0, 60); });
  if (!v.length) throw new Error('Elija al menos una dimensión.');
  conLock_(function () { fijarConfig_('dimensiones', v); });
  return { ok: true, activas: v };
}
// Actores a quienes se reporta: lista propuesta (Datos.gs) + los que agrega el facilitador en el momento 1.
function extras_() { const v = config_('actores'); return Array.isArray(v) ? v : []; }
function destinos_() {
  const base = TALLER.destinos.map(function (d) { return d.id; });
  return TALLER.destinos.concat(extras_().filter(function (x) { return base.indexOf(x) < 0; }).map(function (x) { return { id: x, marco: 'Agregado en el taller' }; }));
}
function nombresDestinos_() { return destinos_().map(function (d) { return d.id; }); }
function fijarActores(r) {
  validarClave_(r);
  const v = lista_(r.lista, null, 12).map(function (x) { return x.slice(0, 60); });
  conLock_(function () { fijarConfig_('actores', v); });
  return { ok: true, destinos: destinos_() };
}
// Preguntas que se trabajan en los momentos 3 y 4. Si el facilitador no las fija, se toman las 4 más votadas.
function priorizadasFijadas_() { const v = config_('priorizadas'); return Array.isArray(v) ? v : null; }
function fijarPriorizadas(r) {
  validarClave_(r);
  const ids = preguntas_().map(function (q) { return q.id; });
  const v = lista_(r.ids, ids, 30);
  conLock_(function () { fijarConfig_('priorizadas', v); });
  return { ok: true, priorizadas: v };
}
// Cualquier participante puede sumar una pregunta más a las que se trabajan (no quita ninguna).
function agregarPriorizada(r) {
  const id = limpiar_(r && r.id);
  if (!preguntas_().some(function (q) { return q.id === id; })) throw new Error('Pregunta no encontrada.');
  return conLock_(function () {
    const v = priorizadasFijadas_() || priorizadasEfectivas_(null);
    if (v.indexOf(id) < 0) v.push(id);
    fijarConfig_('priorizadas', v);
    return { ok: true, priorizadas: v };
  });
}

/* ---------- Momento 1 · Para qué (los actores se revisan aparte y se vinculan a cada pregunta en el momento 2) ---------- */
function guardarParaQue(r) {
  const p = validarPersona_(r);
  const base = TALLER.dimensiones.map(function (d) { return d.nombre; });
  const envio = token_(r) || Utilities.getUuid().slice(0, 8), ahora = new Date();
  const filas = [];
  (r.items || []).slice(0, 12).forEach(function (it) {
    const nombre = enLista_(it.nombre, base);
    if (nombre) filas.push([ahora, envio, p.area, p.grupo, nombre, limpiar_(it.ejemplo)]);
  });
  const otro = corto_(r.otro);
  if (otro) filas.push([ahora, envio, p.area, p.grupo, 'Otro: ' + otro, limpiar_(r.otroEjemplo)]);
  // Actores que faltan en la lista (separados por coma o punto y coma).
  limpiar_(r.actores).split(/[;,]/).map(corto_).filter(String).slice(0, 5)
    .forEach(function (a) { filas.push([ahora, envio, p.area, p.grupo, 'Actor: ' + a, '']); });
  // Actores de la lista a quienes reporta esta persona.
  lista_(r.reporta, nombresDestinos_(), 12).forEach(function (a) { filas.push([ahora, envio, p.area, p.grupo, 'Reporta: ' + a, '']); });
  if (!filas.length) throw new Error('Marque al menos un para qué.');
  // Si el mismo dispositivo envía de nuevo, su respuesta anterior se reemplaza (no se duplica).
  conLock_(function () { quitarFilas_('ParaQue', function (f) { return f[1] === envio; }); escribir_('ParaQue', filas); });
  return { ok: true, reemplazada: !!token_(r) };
}

/* ---------- Momento 2 · Preguntas y votos ---------- */
function guardarPregunta(r) {
  const p = validarPersona_(r);
  const texto = limpiar_(r.texto);
  if (texto.length < 10) throw new Error('Escriba la pregunta.');
  const dims = lista_(r.dims, null, 8);
  if (!dims.length) throw new Error('Marque al menos un para qué.');
  const destinos = lista_(r.destinos, null, 10).map(corto_);
  return conLock_(function () {
    const tk = token_(r), datos = filas_('Preguntas');
    // Editar: misma pregunta del mismo dispositivo. Duplicado: mismo texto del mismo dispositivo.
    let i = tk ? datos.findIndex(function (f) { return f[11] === tk && ((r.id && f[1] === r.id) || String(f[4]).trim().toLowerCase() === texto.toLowerCase()); }) : -1;
    const id = i >= 0 ? datos[i][1] : 'Q' + ('00' + (datos.reduce(function (m, f) { const n = /^Q(\d+)$/.exec(f[1]); return n ? Math.max(m, Number(n[1])) : m; }, 0) + 1)).slice(-3);
    const fila = [new Date(), id, p.area, p.grupo, texto, dims.join('; '), destinos.join('; '), limpiar_(r.decision), i >= 0 ? datos[i][8] : '', enLista_(r.frecuencia, FRECUENCIAS), corto_(r.mide), tk];
    if (i >= 0) hoja_('Preguntas').getRange(i + 2, 1, 1, fila.length).setValues([fila]); else escribir_('Preguntas', [fila]);
    return { ok: true, id: id, editada: i >= 0 };
  });
}
function guardarVotos(r) {
  const p = validarPersona_(r);
  const visibles = preguntas_().filter(function (q) { return !q.oculta; }).map(function (q) { return q.id; });
  const ids = lista_(r.ids, visibles, MAX_VOTOS);
  if (!ids.length) throw new Error('Marque al menos una estrella.');
  const envio = token_(r) || Utilities.getUuid().slice(0, 8), ahora = new Date();
  conLock_(function () { quitarFilas_('Votos', function (f) { return f[1] === envio; }); escribir_('Votos', ids.map(function (id) { return [ahora, envio, p.area, p.grupo, id]; })); });
  return { ok: true, votos: ids.length };
}
function ocultarPregunta(r) {
  validarClave_(r);
  const id = r.id, oculta = !!r.oculta;
  return conLock_(function () {
    const sh = hoja_('Preguntas'), datos = filas_('Preguntas');
    for (let i = 0; i < datos.length; i++) if (datos[i][1] === id) { sh.getRange(i + 2, 9).setValue(oculta ? 'Sí' : ''); return { ok: true }; }
    throw new Error('Pregunta no encontrada.');
  });
}
function preguntas_() {
  return filas_('Preguntas').map(function (f) {
    return { id: f[1], area: f[2], grupo: String(f[3] || ''), texto: f[4], dims: String(f[5] || '').split('; ').filter(String),
      destinos: String(f[6] || '').split('; ').filter(String), decision: f[7], oculta: f[8] === 'Sí', frecuencia: f[9], mide: f[10] };
  });
}

/* ---------- Momento 3 · Qué tenemos (insumos por pregunta, en plenario) ---------- */
function guardarInsumo(r) {
  const qid = limpiar_(r && r.pregunta);
  if (!preguntas_().some(function (q) { return q.id === qid; })) throw new Error('Pregunta no encontrada.');
  const fuente = limpiar_(r.fuente).slice(0, 200);
  if (!fuente) throw new Error('Indique la fuente.');
  const tk = token_(r);
  return conLock_(function () {
    const datos = filas_('Insumos');
    // Editar por ID; o, si el mismo dispositivo repite la misma fuente en la misma pregunta, se actualiza en vez de duplicar.
    const i = datos.findIndex(function (f) { return (r.id && f[1] === r.id) || (tk && f[9] === tk && f[2] === qid && f[3] === fuente); });
    const id = i >= 0 ? datos[i][1] : 'I' + Utilities.getUuid().slice(0, 6);
    const fila = [new Date(), id, qid, fuente, corto_(r.produce), enLista_(r.frecuencia, FRECUENCIAS), enLista_(r.tenemos, TENEMOS), limpiar_(r.comentario), grupo_(r), i >= 0 && !tk ? datos[i][9] : tk];
    if (i >= 0) hoja_('Insumos').getRange(i + 2, 1, 1, fila.length).setValues([fila]); else escribir_('Insumos', [fila]);
    return { ok: true, id: id, editado: i >= 0 };
  });
}
function borrarInsumo(r) {
  const id = limpiar_(r && r.id);
  return conLock_(function () {
    const sh = hoja_('Insumos'), datos = filas_('Insumos');
    for (let i = datos.length - 1; i >= 0; i--) if (datos[i][1] === id) { sh.deleteRow(i + 2); return { ok: true }; }
    throw new Error('Insumo no encontrado.');
  });
}
/* ---------- Momentos 3 y 4 · Conclusión por pregunta: lograble, horizonte, qué falta, cómo se presenta y responsable ---------- */
const CAMPOS_EVAL = ['lograble', 'horizonte', 'falta', 'presenta', 'responsable'];
function guardarEvaluacion(r) {
  const qid = limpiar_(r && r.pregunta);
  if (!preguntas_().some(function (q) { return q.id === qid; })) throw new Error('Pregunta no encontrada.');
  return conLock_(function () {
    const actual = evaluacion_()[qid] || {};
    const v = {};
    CAMPOS_EVAL.forEach(function (k) { v[k] = actual[k] || ''; });
    if ('lograble' in r) v.lograble = enLista_(r.lograble, LOGRABLE);
    if ('horizonte' in r) v.horizonte = enLista_(r.horizonte, HORIZONTE);
    if ('falta' in r) v.falta = limpiar_(r.falta);
    if ('presenta' in r) v.presenta = limpiar_(r.presenta);
    if ('responsable' in r) v.responsable = corto_(r.responsable);
    upsert_('Evaluacion', function (f) { return f[1] === qid; }, [new Date(), qid, v.lograble, v.horizonte, v.falta, v.presenta, v.responsable]);
    return { ok: true, evaluacion: v };
  });
}
function evaluacion_() {
  const out = {};
  filas_('Evaluacion').forEach(function (f) { out[f[1]] = { lograble: f[2], horizonte: f[3], falta: f[4], presenta: f[5], responsable: f[6] }; });
  return out;
}

/* ---------- Resultados ---------- */
function priorizadasEfectivas_(ranking) {
  const v = priorizadasFijadas_();
  if (v) return v;
  const R = ranking || ranking_();
  return R.filter(function (q) { return q.votos > 0; }).slice(0, PRIORIZADAS_SUGERIDAS).map(function (q) { return q.id; });
}
function ranking_() {
  const votos = {};
  filas_('Votos').forEach(function (f) { votos[f[4]] = (votos[f[4]] || 0) + 1; });
  return preguntas_().map(function (q) { q.votos = votos[q.id] || 0; return q; })
    .filter(function (q) { return !q.oculta; }).sort(function (a, b) { return (b.votos - a.votos) || (a.id < b.id ? -1 : 1); });
}
function obtenerResultados() {
  const pq = {}, envPQ = {}, actores = {}, reporta = {};
  filas_('ParaQue').forEach(function (f) {
    envPQ[f[1]] = true;
    const k = String(f[4]);
    if (k.indexOf('Actor: ') === 0) { const a = k.slice(7); actores[a] = (actores[a] || 0) + 1; return; }
    if (k.indexOf('Reporta: ') === 0) { const a = k.slice(9); reporta[a] = (reporta[a] || 0) + 1; return; }
    if (!pq[k]) pq[k] = { nombre: k, votos: 0, ejemplos: [] };
    pq[k].votos++;
    if (f[5]) pq[k].ejemplos.push(f[5]);
  });
  const ordenar = function (o) { return Object.keys(o).map(function (a) { return { nombre: a, votos: o[a] }; }).sort(function (a, b) { return b.votos - a.votos; }); };
  const votos = {}, votantes = {};
  filas_('Votos').forEach(function (f) { votos[f[4]] = (votos[f[4]] || 0) + 1; votantes[f[1]] = true; });
  const preguntas = preguntas_().map(function (q) { q.votos = votos[q.id] || 0; return q; });
  const ranking = preguntas.filter(function (q) { return !q.oculta; }).sort(function (a, b) { return (b.votos - a.votos) || (a.id < b.id ? -1 : 1); });
  const insumos = filas_('Insumos').map(function (f) {
    return { id: f[1], pregunta: f[2], fuente: f[3], produce: f[4], frecuencia: f[5], tenemos: f[6], comentario: f[7], grupo: String(f[8] || ''), mio: '' };
  });
  return { activas: activas_(), destinos: destinos_(), grupos: grupos_(), actores: ordenar(actores), reporta: ordenar(reporta),
    simulacion: config_('simulacion') === true, paraque: Object.keys(pq).map(function (k) { return pq[k]; }).sort(function (a, b) { return b.votos - a.votos; }),
    respuestasParaQue: Object.keys(envPQ).length, preguntas: preguntas, ranking: ranking, votantes: Object.keys(votantes).length,
    priorizadas: priorizadasEfectivas_(ranking), priorizadasFijadas: !!priorizadasFijadas_(), insumos: insumos, evaluacion: evaluacion_() };
}

/* ---------- Simulación (solo para probar la herramienta antes del taller) ---------- */
// Los datos simulados llevan Envío "SIM-…", ID de pregunta "S…", insumos "SI…" y textos marcados [Simulado]. borrarSimulacion los quita todos.
function cargarSimulacion(r) {
  validarClave_(r);
  borrarSimulacion(r);
  const ahora = new Date();
  const areas = AREAS.filter(function (a) { return a !== 'Otra'; });
  const N = TALLER.dimensiones.map(function (d) { return d.nombre; });
  const dest = TALLER.destinos.map(function (d) { return d.id; });
  const elecciones = [[0, 2], [0, 3], [1, 0], [2, 5], [0, 1, 3], [3], [2, 4], [0, 2], [4, 1], [5, 2], [0], [3, 2], [1, 4], [2, 0]];
  const pq = [];
  elecciones.forEach(function (sel, i) {
    const env = 'SIM-' + ('0' + (i + 1)).slice(-2), area = areas[i % areas.length];
    sel.forEach(function (k, j) { pq.push([ahora, env, area, '', N[k], j === 0 && i % 3 === 0 ? '[Simulado] ' + TALLER.dimensiones[k].ejemplo : '']); });
  });
  pq.push([ahora, 'SIM-03', areas[2], '', 'Otro: Proteger a NNA en tránsito [simulado]', '']);
  pq.push([ahora, 'SIM-05', areas[4], '', 'Actor: Personería Distrital [simulado]', '']);
  for (let i = 0; i < 14; i++) [2, 1, 3, (i % 4) + 4].slice(0, 2 + (i % 3)).forEach(function (k) { pq.push([ahora, 'SIM-' + ('0' + (i + 1)).slice(-2), areas[i % areas.length], '', 'Reporta: ' + dest[k], '']); });
  pq.push([ahora, 'SIM-09', areas[8 % areas.length], '', 'Actor: Personería Distrital [simulado]', '']);
  pq.push([ahora, 'SIM-11', areas[10 % areas.length], '', 'Actor: Instituto Colombiano de Bienestar Familiar [simulado]', '']);
  const frec = ['Mensual', 'Trimestral', 'Mensual', 'Semestral', 'Mensual', 'Anual', 'Trimestral', 'Mensual', 'Mensual', 'Trimestral'];
  const mide = ['Personas distintas atendidas'];
  const preg = TALLER.ejemplos.map(function (x, i) {
    return [ahora, 'S' + ('0' + (i + 1)).slice(-2), areas[i % areas.length], '', '[Simulado] ' + x.texto,
      x.dims.map(function (id) { return TALLER.dimensiones.filter(function (d) { return d.id === id; })[0].nombre; }).join('; '),
      [dest[(i * 3 + 2) % dest.length], dest[(i * 5 + 1) % dest.length]].filter(function (v, k, a) { return a.indexOf(v) === k; }).join('; '),
      i % 3 === 0 ? '[Simulado] Ajustar la oferta del trimestre' : '', '', frec[i % frec.length], i === 0 ? '[Simulado] ' + mide[0] : '', ''];
  });
  preg.push([ahora, 'S11', areas[1], '', '[Simulado] ¿Cuántas personas atendidas cada mes son de primera vez?', N[0] + '; ' + N[2], dest[2], '[Simulado] Planear cupos del mes', '', 'Mensual', '[Simulado] Personas nuevas por mes', '']);
  preg.push([ahora, 'S12', areas[3], '', '[Simulado] ¿Cuántas personas atendemos al mes?', N[0], dest[2], '', 'Sí', 'Mensual', '', '']);
  const pesos = [9, 3, 6, 4, 8, 7, 5, 2, 6, 1, 4, 0];
  const votos = [];
  for (let v = 0; v < 14; v++) {
    const orden = preg.map(function (q, k) { return { id: q[1], s: pesos[k] + ((v * 7 + k * 3) % 5) }; }).filter(function (x) { return x.id !== 'S12'; })
      .sort(function (a, b) { return b.s - a.s; }).slice(0, MAX_VOTOS);
    orden.forEach(function (x) { votos.push([ahora, 'SIM-' + ('0' + (v + 1)).slice(-2), areas[v % areas.length], '', x.id]); });
  }
  conLock_(function () {
    escribir_('ParaQue', pq); escribir_('Preguntas', preg); escribir_('Votos', votos);
    fijarConfig_('dimensiones', [N[0], N[2], N[3], N[4]]);
    fijarConfig_('actores', ['Personería Distrital [simulado]']);
    fijarConfig_('simulacion', true);
  });
  // Momentos 3 y 4: las 4 más votadas, con 1 a 3 insumos cada una y su conclusión.
  const R = ranking_().slice(0, 4);
  conLock_(function () { fijarConfig_('priorizadas', R.map(function (q) { return q.id; })); });
  const eq = 'Subsecretaría – Equipo de Migrantes';
  const ins = [
    [['Registros de la atención · Archivo 1S (indicadores)', eq, 'Mensual', 'Parcial', '[Simulado] Falta un ID único para no contar dos veces'],
      ['Registros de la atención · Ficha de caracterización (Forms)', eq, 'Diaria', 'Sí', '']],
    [['Registros de la atención · Registro de remisiones', eq, 'Mensual', 'Sí', ''],
      ['Datos administrativos · Otra secretaría o entidad', 'Secretarías sectoriales', 'Por demanda', 'No', '[Simulado] No hay retorno de información de la entidad receptora']],
    [['Registros de la atención · Matriz de atenciones E.A', eq, 'Mensual', 'Sí', ''],
      ['Datos administrativos · SISBEN', 'Planeación Distrital', 'Anual', 'Parcial', '[Simulado] Acceso por convenio'],
      ['Información secundaria · Reportes de cooperación (GIFMM, R4V, agencias)', 'Cooperación', 'Trimestral', 'Sí', '[Simulado] Solo a nivel departamental']],
    [['Información nueva · Hay que levantarla', 'Por definir', 'No sé', 'No', '[Simulado] Requiere una pregunta nueva en la ficha']]];
  const evs = [['Con ajustes', 'Prioritaria', '[Simulado] Regla de conteo de personas distintas', '[Simulado] Tablero mensual con cifra clave y tendencia por punto', eq],
    ['Con ajustes', 'Prioritaria', '[Simulado] Acuerdo de retorno de información con las secretarías', '[Simulado] Informe trimestral con remisiones efectivas por sector', eq],
    ['Sí', 'Prioritaria', '', '[Simulado] Ficha de una página para el Despacho', 'Planeación Distrital'],
    ['No por ahora', 'Para más adelante', '[Simulado] Agregar la pregunta a la ficha de caracterización', '', '']];
  const filasIns = [], filasEv = [];
  R.forEach(function (q, i) {
    ins[i].forEach(function (x, j) { filasIns.push([ahora, 'SI' + i + j, q.id, x[0], x[1], x[2], x[3], x[4], '', '']); });
    filasEv.push([ahora, q.id].concat(evs[i]));
  });
  conLock_(function () { escribir_('Insumos', filasIns); escribir_('Evaluacion', filasEv); });
  return { ok: true };
}
function borrarSimulacion(r) {
  validarClave_(r);
  conLock_(function () {
    const quitar = function (sh, col, prueba) {
      if (!sh || sh.getLastRow() < 2) return;
      const datos = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
      for (let i = datos.length - 1; i >= 0; i--) if (prueba(String(datos[i][col]))) sh.deleteRow(i + 2);
    };
    const S = function (v) { return /^S\d+$/.test(v); };
    const sim = function (v) { return v.indexOf('SIM-') === 0; };
    quitar(hoja_('ParaQue'), 1, sim);
    quitar(hoja_('Votos'), 1, sim);
    quitar(hoja_('Preguntas'), 1, S);
    quitar(hoja_('Insumos'), 2, S);
    quitar(hoja_('Evaluacion'), 1, S);
    quitar(libro_().getSheetByName('Matriz'), 2, S); // hoja de la versión anterior, si existe
    if (config_('simulacion') === true) quitar(hoja_('Config'), 0, function (v) { return ['dimensiones', 'actores', 'simulacion', 'priorizadas'].indexOf(v) >= 0; });
  });
  return { ok: true };
}

/* ---------- Dejar la hoja en blanco para el taller (borra todas las respuestas, también las de prueba) ---------- */
function limpiarTodo(r) {
  validarClave_(r);
  if (!r || r.confirmar !== 'BORRAR') throw new Error('Falta confirmar.');
  conLock_(function () {
    ['ParaQue', 'Preguntas', 'Votos', 'Insumos', 'Evaluacion'].forEach(function (n) { const sh = hoja_(n); if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1); });
    const old = libro_().getSheetByName('Matriz'); if (old && old.getLastRow() > 1) old.deleteRows(2, old.getLastRow() - 1);
    quitarFilas_('Config', function (f) { return ['dimensiones', 'actores', 'simulacion', 'priorizadas'].indexOf(f[0]) >= 0; });
  });
  return { ok: true };
}
