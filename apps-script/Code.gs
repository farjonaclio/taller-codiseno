/**
 * Taller de codiseño – Modelo de análisis de información
 * Servidor de la herramienta web del taller (versión 4: cuatro momentos).
 * La página vive en GitHub Pages y llama a este servidor (doPost, JSON). Las respuestas quedan en la hoja.
 * Vistas de la página: celular (momentos 1 y 2) · ?vista=mesa (momentos 3 y 4) · ?vista=resultados (proyector).
 *
 * Instalación: en la hoja de respuestas > Extensiones > Apps Script. Pegar Code.gs y Datos.gs
 * (botón + > Secuencia de comandos). Implementar > Nueva implementación > Aplicación web
 *   Ejecutar como: Yo   ·   Quién tiene acceso: Cualquier persona
 * (no "Cualquier persona con cuenta de Google": los invitados no tienen cuenta de CliO ni Gmail).
 * Copiar la URL /exec en config.js del repositorio. Opcional: crear un archivo HTML "Index" con Index.html
 * para usar la misma página directo desde Apps Script.
 *
 * No se piden nombres ni correos: solo área y mesa (minimización de datos personales,
 * Ley 1581 de 2012). No escribir datos de personas atendidas.
 */

const MAX_TEXTO = 500;
const MAX_VOTOS = 5;
const MESAS = ['1', '2', '3'];

// TALLER (ideas de para qué, ejemplos y destinatarios) y CLAVE_FACILITADOR están en Datos.gs,
// que no se publica en GitHub porque contiene información del cliente.

// AREAS y FUENTES también están en Datos.gs (nombres internos del cliente).
const FRECUENCIAS = ['Diaria', 'Semanal', 'Mensual', 'Trimestral', 'Semestral', 'Anual', 'Por demanda', 'No sé'];
const ACCESO = { 'verde': '🟢 Sí, lo tenemos', 'amarillo': '🟡 Parcial o hay que pedirlo', 'rojo': '🔴 No lo tenemos' };
const RUTAS = { 'construir': '🛠️ Lo construimos en la consultoría', 'gestionar': '🤝 Lo gestiona otra entidad', 'fuera': '⛔ Fuera de alcance' };
const CUANDO = ['Diagnóstico (nov–dic 2026)', 'Primer trimestre 2027', 'Después de la consultoría', 'Por definir'];
const FORMATOS = ['Tablero', 'Informe', 'Ficha o boletín', 'Alerta', 'Por definir'];
function responsables_() { return AREAS.filter(function (a) { return a !== 'Otra'; }).concat(['Otra entidad', 'Por definir']); }

function doGet(e) {
  if (e && e.parameter && e.parameter.accion) return api_(e.parameter.accion, null);
  try {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Taller de codiseño')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
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
    guardarVotos: guardarVotos, guardarMesa: guardarMesa, fijarDimensiones: fijarDimensiones, ocultarPregunta: ocultarPregunta,
    fijarActores: fijarActores, cargarSimulacion: cargarSimulacion, borrarSimulacion: borrarSimulacion };
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
  return { dimensiones: TALLER.dimensiones, ejemplos: TALLER.ejemplos, destinos: destinos_(), activas: activas_(),
    areas: AREAS, fuentes: FUENTES, frecuencias: FRECUENCIAS, acceso: ACCESO, rutas: RUTAS, cuando: CUANDO,
    formatos: FORMATOS, responsables: responsables_(), maxVotos: MAX_VOTOS };
}

/* ---------- Hojas ---------- */
const HOJAS = {
  Config: ['Clave', 'Valor'],
  ParaQue: ['Fecha', 'Envío', 'Área', 'Mesa', 'Para qué', 'Ejemplo de decisión'],
  Preguntas: ['Fecha', 'ID', 'Área', 'Mesa', 'Pregunta', 'Para qué (dimensiones)', 'A quién se reporta', 'Decisión', 'Oculta'],
  Votos: ['Fecha', 'Envío', 'Área', 'Mesa', 'ID'],
  Matriz: ['Fecha', 'Mesa', 'ID', 'Pregunta', 'Dimensiones', 'Fuente', 'Produce', 'Frecuencia del dato', 'Lo tenemos', 'Qué falta', 'Ruta',
    'Siguiente paso', 'Responsable', 'Cuándo', 'Cómo se usa', 'Frecuencia de uso', 'Formato', 'Frecuencia de reporte', 'Valida', 'Publica', 'Recibe', 'Notas del boceto']
};

function hoja_(nombre) {
  // HOJA_ID (en Datos.gs) apunta a la hoja de respuestas; si no está, usa la hoja a la que está vinculado el script.
  const ss = (typeof HOJA_ID !== 'undefined' && HOJA_ID) ? SpreadsheetApp.openById(HOJA_ID) : SpreadsheetApp.getActiveSpreadsheet();
  const cols = HOJAS[nombre];
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
  return (Array.isArray(arr) ? arr : []).map(limpiar_).filter(function (x) { return x && (!validos || validos.indexOf(x) >= 0); }).slice(0, max || 12);
}
function conLock_(fn) { const lock = LockService.getScriptLock(); lock.waitLock(20000); try { return fn(); } finally { lock.releaseLock(); } }
function validarPersona_(r) {
  if (!r || AREAS.indexOf(r.area) < 0) throw new Error('Seleccione su área.');
  return { area: r.area, mesa: enLista_(String(r.mesa || ''), MESAS) };
}
function config_(clave) {
  const f = filas_('Config').filter(function (x) { return x[0] === clave; })[0];
  try { return f ? JSON.parse(f[1]) : null; } catch (e) { return null; }
}
function fijarConfig_(clave, valor) { upsert_('Config', function (f) { return f[0] === clave; }, [clave, JSON.stringify(valor)]); }
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

/* ---------- Config: dimensiones acordadas ---------- */
function activas_() { const v = config_('dimensiones'); return Array.isArray(v) && v.length ? v : null; }
function fijarDimensiones(r) {
  validarClave_(r);
  const v = lista_(r.lista, null, 8).map(function (x) { return x.slice(0, 60); });
  if (!v.length) throw new Error('Elija al menos una dimensión.');
  conLock_(function () { upsert_('Config', function (f) { return f[0] === 'dimensiones'; }, ['dimensiones', JSON.stringify(v)]); });
  return { ok: true, activas: v };
}

/* ---------- Momento 1 · Para qué ---------- */
function guardarParaQue(r) {
  const p = validarPersona_(r);
  const base = TALLER.dimensiones.map(function (d) { return d.nombre; });
  const envio = Utilities.getUuid().slice(0, 8), ahora = new Date();
  const filas = [];
  (r.items || []).slice(0, 12).forEach(function (it) {
    const nombre = enLista_(it.nombre, base);
    if (nombre) filas.push([ahora, envio, p.area, p.mesa, nombre, limpiar_(it.ejemplo)]);
  });
  const otro = limpiar_(r.otro).slice(0, 60);
  if (otro) filas.push([ahora, envio, p.area, p.mesa, 'Otro: ' + otro, limpiar_(r.otroEjemplo)]);
  // Actores que faltan en la lista de destinatarios (separados por coma o punto y coma).
  limpiar_(r.actores).split(/[;,]/).map(function (x) { return x.trim().slice(0, 60); }).filter(String).slice(0, 5)
    .forEach(function (a) { filas.push([ahora, envio, p.area, p.mesa, 'Actor: ' + a, '']); });
  if (!filas.length) throw new Error('Marque al menos un para qué.');
  conLock_(function () { const sh = hoja_('ParaQue'); sh.getRange(sh.getLastRow() + 1, 1, filas.length, HOJAS.ParaQue.length).setValues(filas); });
  return { ok: true };
}

/* ---------- Momento 2 · Preguntas y votos ---------- */
function guardarPregunta(r) {
  const p = validarPersona_(r);
  const texto = limpiar_(r.texto);
  if (texto.length < 10) throw new Error('Escriba la pregunta.');
  const dims = lista_(r.dims, null, 8);
  if (!dims.length) throw new Error('Marque al menos un para qué.');
  return conLock_(function () {
    const sh = hoja_('Preguntas');
    const max = filas_('Preguntas').reduce(function (m, f) { const n = /^Q(\d+)$/.exec(f[1]); return n ? Math.max(m, Number(n[1])) : m; }, 0);
    const id = 'Q' + ('00' + (max + 1)).slice(-3);
    sh.appendRow([new Date(), id, p.area, p.mesa, texto, dims.join('; '), lista_(r.destinos, nombresDestinos_()).join('; '), limpiar_(r.decision), '']);
    return { ok: true, id: id };
  });
}
function guardarVotos(r) {
  const p = validarPersona_(r);
  const visibles = preguntas_().filter(function (q) { return !q.oculta; }).map(function (q) { return q.id; });
  const ids = lista_(r.ids, visibles, MAX_VOTOS).filter(function (x, i, a) { return a.indexOf(x) === i; });
  if (!ids.length) throw new Error('Marque al menos una estrella.');
  const envio = Utilities.getUuid().slice(0, 8), ahora = new Date();
  conLock_(function () {
    const sh = hoja_('Votos');
    sh.getRange(sh.getLastRow() + 1, 1, ids.length, HOJAS.Votos.length).setValues(ids.map(function (id) { return [ahora, envio, p.area, p.mesa, id]; }));
  });
  return { ok: true, votos: ids.length };
}
function ocultarPregunta(r) {
  validarClave_(r);
  const id = r.id, oculta = !!r.oculta;
  return conLock_(function () {
    const sh = hoja_('Preguntas');
    const datos = filas_('Preguntas');
    for (let i = 0; i < datos.length; i++) if (datos[i][1] === id) { sh.getRange(i + 2, 9).setValue(oculta ? 'Sí' : ''); return { ok: true }; }
    throw new Error('Pregunta no encontrada.');
  });
}
function preguntas_() {
  return filas_('Preguntas').map(function (f) {
    return { id: f[1], area: f[2], mesa: String(f[3]), texto: f[4], dims: String(f[5] || '').split('; ').filter(String),
      destinos: String(f[6] || '').split('; ').filter(String), decision: f[7], oculta: f[8] === 'Sí' };
  });
}

/* ---------- Momentos 3 y 4 · Computador de la mesa ---------- */
function guardarMesa(r) {
  const mesa = enLista_(String(r && r.mesa), MESAS);
  if (!mesa) throw new Error('Seleccione la mesa.');
  const porId = {};
  preguntas_().forEach(function (q) { porId[q.id] = q; });
  const ahora = new Date(), RESPONSABLES = responsables_();
  return conLock_(function () {
    let n = 0;
    (r.filas || []).slice(0, 12).forEach(function (x) {
      const q = porId[x.id];
      if (!q) return;
      upsert_('Matriz', function (f) { return String(f[1]) === mesa && f[2] === q.id; },
        [ahora, mesa, q.id, q.texto, q.dims.join('; '), limpiar_(x.fuente), enLista_(x.produce, RESPONSABLES), enLista_(x.frecuencia, FRECUENCIAS),
          ACCESO[x.semaforo] || '', limpiar_(x.falta), RUTAS[x.ruta] || '', limpiar_(x.paso), enLista_(x.responsable, RESPONSABLES), enLista_(x.cuando, CUANDO),
          limpiar_(x.uso), enLista_(x.frecUso, FRECUENCIAS), enLista_(x.formato, FORMATOS), enLista_(x.frecReporte, FRECUENCIAS),
          enLista_(x.valida, RESPONSABLES), enLista_(x.publica, RESPONSABLES), lista_(x.recibe, nombresDestinos_()).join('; '), limpiar_(x.notas)]);
      n++;
    });
    return { ok: true, filas: n };
  });
}
function upsert_(nombre, coincide, fila) {
  const sh = hoja_(nombre);
  const datos = filas_(nombre);
  for (let i = 0; i < datos.length; i++) if (coincide(datos[i])) { sh.getRange(i + 2, 1, 1, fila.length).setValues([fila]); return; }
  sh.getRange(sh.getLastRow() + 1, 1, 1, fila.length).setValues([fila]);
}

/* ---------- Resultados ---------- */
function obtenerResultados() {
  const pq = {}, envPQ = {};
  const actores = {};
  filas_('ParaQue').forEach(function (f) {
    envPQ[f[1]] = true;
    const k = String(f[4]);
    if (k.indexOf('Actor: ') === 0) { const a = k.slice(7); actores[a] = (actores[a] || 0) + 1; return; }
    if (!pq[k]) pq[k] = { nombre: k, votos: 0, ejemplos: [] };
    pq[k].votos++;
    if (f[5]) pq[k].ejemplos.push(f[5]);
  });
  const votos = {}, votantes = {};
  filas_('Votos').forEach(function (f) { votos[f[4]] = (votos[f[4]] || 0) + 1; votantes[f[1]] = true; });
  const preguntas = preguntas_().map(function (q) { q.votos = votos[q.id] || 0; return q; });
  const ranking = preguntas.filter(function (q) { return !q.oculta; }).sort(function (a, b) { return (b.votos - a.votos) || (a.id < b.id ? -1 : 1); });
  const matriz = filas_('Matriz').map(function (f) {
    return { mesa: String(f[1]), id: f[2], pregunta: f[3], dims: f[4], fuente: f[5], produce: f[6], frecuencia: f[7], semaforo: f[8], falta: f[9], ruta: f[10],
      paso: f[11], responsable: f[12], cuando: f[13], uso: f[14], frecUso: f[15], formato: f[16], frecReporte: f[17], valida: f[18], publica: f[19], recibe: f[20], notas: f[21] };
  });
  return { activas: activas_(), destinos: destinos_(), actores: Object.keys(actores).map(function (a) { return { nombre: a, votos: actores[a] }; }),
    simulacion: config_('simulacion') === true, paraque: Object.keys(pq).map(function (k) { return pq[k]; }).sort(function (a, b) { return b.votos - a.votos; }),
    respuestasParaQue: Object.keys(envPQ).length, preguntas: preguntas, ranking: ranking, votantes: Object.keys(votantes).length, matriz: matriz };
}

/* ---------- Simulación (solo para probar la herramienta antes del taller) ---------- */
// Los datos simulados llevan Envío "SIM-…" e ID "S…", y textos marcados [Simulado]. borrarSimulacion los quita todos.
function cargarSimulacion(r) {
  validarClave_(r);
  borrarSimulacion(r);
  const ahora = new Date();
  const areas = AREAS.filter(function (a) { return a !== 'Otra'; });
  const N = TALLER.dimensiones.map(function (d) { return d.nombre; });
  const elecciones = [[0, 2], [0, 3], [1, 0], [2, 5], [0, 1, 3], [3], [2, 4], [0, 2], [4, 1], [5, 2], [0], [3, 2], [1, 4], [2, 0]];
  const pq = [];
  elecciones.forEach(function (sel, i) {
    const env = 'SIM-' + ('0' + (i + 1)).slice(-2), area = areas[i % areas.length], mesa = String(i % 3 + 1);
    sel.forEach(function (k, j) { pq.push([ahora, env, area, mesa, N[k], j === 0 && i % 3 === 0 ? '[Simulado] ' + TALLER.dimensiones[k].ejemplo : '']); });
  });
  pq.push([ahora, 'SIM-03', areas[2], '3', 'Otro: Proteger a NNA en tránsito [simulado]', '']);
  pq.push([ahora, 'SIM-05', areas[4], '2', 'Actor: Personería Distrital [simulado]', '']);
  pq.push([ahora, 'SIM-09', areas[8 % areas.length], '3', 'Actor: Personería Distrital [simulado]', '']);
  pq.push([ahora, 'SIM-11', areas[10 % areas.length], '2', 'Actor: Instituto Colombiano de Bienestar Familiar [simulado]', '']);
  const dest = TALLER.destinos.map(function (d) { return d.id; });
  const preg = TALLER.ejemplos.map(function (x, i) {
    return [ahora, 'S' + ('0' + (i + 1)).slice(-2), areas[i % areas.length], String(i % 3 + 1), '[Simulado] ' + x.texto,
      x.dims.map(function (id) { return TALLER.dimensiones.filter(function (d) { return d.id === id; })[0].nombre; }).join('; '),
      [dest[(i * 3 + 2) % dest.length], dest[(i * 5 + 1) % dest.length]].filter(function (v, k, a) { return a.indexOf(v) === k; }).join('; '), '', ''];
  });
  preg.push([ahora, 'S11', areas[1], '2', '[Simulado] ¿Cuántas personas atendidas cada mes son de primera vez?', N[0] + '; ' + N[2], dest[2], 'Planear cupos del mes', '']);
  preg.push([ahora, 'S12', areas[3], '1', '[Simulado] ¿Cuántas personas atendemos al mes?', N[0], dest[2], '', 'Sí']);
  const pesos = [9, 3, 6, 4, 8, 7, 5, 2, 6, 1, 4, 0];
  const votos = [];
  for (let v = 0; v < 14; v++) {
    const orden = preg.map(function (q, k) { return { id: q[1], s: pesos[k] + ((v * 7 + k * 3) % 5) }; }).filter(function (x) { return x.id !== 'S12'; })
      .sort(function (a, b) { return b.s - a.s; }).slice(0, MAX_VOTOS);
    orden.forEach(function (x) { votos.push([ahora, 'SIM-' + ('0' + (v + 1)).slice(-2), areas[v % areas.length], String(v % 3 + 1), x.id]); });
  }
  conLock_(function () {
    escribir_('ParaQue', pq); escribir_('Preguntas', preg); escribir_('Votos', votos);
    fijarConfig_('dimensiones', [N[0], N[2], N[3], N[4]]);
    fijarConfig_('actores', ['Personería Distrital [simulado]']);
    fijarConfig_('simulacion', true);
  });
  // Matriz: las 6 más votadas, repartidas por turnos entre las mesas.
  const R = obtenerResultados().ranking.slice(0, 6);
  const RS = responsables_();
  const ej = [
    { fuente: 'Registros de la atención · Archivo 1S (indicadores)', semaforo: 'amarillo', falta: '[Simulado] ID único para no contar dos veces', ruta: 'construir', paso: '[Simulado] Regla de conteo de personas distintas', cuando: CUANDO[0], uso: '[Simulado] Reporte mensual del Plan', formato: 'Tablero' },
    { fuente: 'Registros de la atención · Registro de remisiones', semaforo: 'rojo', falta: '[Simulado] No se sabe si la persona recibió el servicio', ruta: 'gestionar', paso: '[Simulado] Acuerdo de retorno de información con entidades receptoras', cuando: CUANDO[1], uso: '[Simulado] Ajustar la ruta con cada secretaría', formato: 'Informe' },
    { fuente: 'Registros de la atención · Ficha de caracterización (Forms)', semaforo: 'rojo', falta: '[Simulado] La ficha no registra comuna de residencia', ruta: 'construir', paso: '[Simulado] Agregar comuna a la ficha', cuando: CUANDO[0], uso: '[Simulado] Elegir sedes de jornadas', formato: 'Tablero' },
    { fuente: 'Datos administrativos · SISBEN', semaforo: 'amarillo', falta: '[Simulado] Acceso por convenio', ruta: 'gestionar', paso: '[Simulado] Solicitud formal a Planeación', cuando: CUANDO[1], uso: '[Simulado] Cruzar vulnerabilidad', formato: 'Informe' },
    { fuente: 'Registros de la atención · Matriz de atenciones E.A', semaforo: 'verde', uso: '[Simulado] Seguimiento del equipo', formato: 'Ficha o boletín' },
    { fuente: 'Información secundaria · Reportes de cooperación (GIFMM, R4V, agencias)', semaforo: 'amarillo', falta: '[Simulado] Desagregación por ciudad', ruta: 'fuera', paso: '[Simulado] Dejar como recomendación', cuando: CUANDO[3], uso: '[Simulado] Contexto para escenarios', formato: 'Informe' }];
  R.forEach(function (q, i) {
    const x = ej[i];
    guardarMesa({ mesa: String(i % 3 + 1), filas: [{ id: q.id, fuente: x.fuente, produce: RS[i % 4], frecuencia: 'Mensual', semaforo: x.semaforo, falta: x.falta || '', ruta: x.ruta || '',
      paso: x.paso || '', responsable: x.ruta ? RS[0] : '', cuando: x.cuando || '', uso: x.uso, frecUso: 'Mensual', formato: x.formato, frecReporte: i % 2 ? 'Trimestral' : 'Mensual',
      valida: i < 4 ? RS[5 % RS.length] : '', publica: i < 3 ? RS[0] : '', recibe: q.destinos, notas: '' }] });
  });
  return { ok: true };
}
function borrarSimulacion(r) {
  validarClave_(r);
  conLock_(function () {
    const quitar = function (nombre, col, prueba) {
      const sh = hoja_(nombre), datos = filas_(nombre);
      for (let i = datos.length - 1; i >= 0; i--) if (prueba(String(datos[i][col]))) sh.deleteRow(i + 2);
    };
    quitar('ParaQue', 1, function (v) { return v.indexOf('SIM-') === 0; });
    quitar('Votos', 1, function (v) { return v.indexOf('SIM-') === 0; });
    quitar('Preguntas', 1, function (v) { return /^S\d+$/.test(v); });
    quitar('Matriz', 2, function (v) { return /^S\d+$/.test(v); });
    if (config_('simulacion') === true) quitar('Config', 0, function (v) { return ['dimensiones', 'actores', 'simulacion'].indexOf(v) >= 0; });
  });
  return { ok: true };
}
function escribir_(nombre, filas) {
  if (!filas.length) return;
  const sh = hoja_(nombre);
  sh.getRange(sh.getLastRow() + 1, 1, filas.length, HOJAS[nombre].length).setValues(filas);
}
