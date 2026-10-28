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
    guardarVotos: guardarVotos, guardarMesa: guardarMesa, fijarDimensiones: fijarDimensiones, ocultarPregunta: ocultarPregunta };
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
  return { dimensiones: TALLER.dimensiones, ejemplos: TALLER.ejemplos, destinos: TALLER.destinos, activas: activas_(),
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
function nombresDestinos_() { return TALLER.destinos.map(function (d) { return d.id; }); }

/* ---------- Config: dimensiones acordadas ---------- */
function activas_() {
  const f = filas_('Config').filter(function (x) { return x[0] === 'dimensiones'; })[0];
  try { const v = f ? JSON.parse(f[1]) : null; return Array.isArray(v) && v.length ? v : null; } catch (e) { return null; }
}
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
    const id = 'Q' + ('00' + sh.getLastRow()).slice(-3);
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
  filas_('ParaQue').forEach(function (f) {
    envPQ[f[1]] = true;
    const k = f[4];
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
  return { activas: activas_(), paraque: Object.keys(pq).map(function (k) { return pq[k]; }).sort(function (a, b) { return b.votos - a.votos; }),
    respuestasParaQue: Object.keys(envPQ).length, preguntas: preguntas, ranking: ranking, votantes: Object.keys(votantes).length, matriz: matriz };
}
