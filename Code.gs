/**
 * SOCIEDAD MUZKI · back de la app (Google Apps Script)
 * Pega este código en Extensiones > Apps Script de la hoja "Muzki_datos_app".
 * La hoja debe tener una pestaña "Movimientos" con la fila 1: id, fecha, concepto, importe, nota, sobre, creado, modificado.
 *
 * Primera vez:
 *   1. Escribe tu clave en CLAVE_APP (aquí abajo), entre las comillas. Solo letras y números.
 *   2. Implementar > Nueva implementación > Aplicación web
 *      (Ejecutar como: yo · Quién tiene acceso: cualquier usuario). Copia la URL que acaba en /exec.
 *   3. Opcional: ejecuta una vez `activarCopias` para guardar una copia de la hoja cada lunes.
 *   4. Opcional: rellena CORREO_EXCEL y URL_PLANTILLA y ejecuta una vez `activarEnvioMensual`:
 *      el día 1 de cada mes llega por correo la Excel del año (marzo a febrero) con el mes recién cerrado,
 *      y se guarda una copia en la carpeta «Muzki · Excel» de Drive. `probarEnvio` la manda ahora mismo.
 */

// ⬇⬇ ENVÍO DE LA EXCEL CADA MES ⬇⬇
const CORREO_EXCEL = '';      // uno o varios correos separados por comas
const URL_PLANTILLA = '';     // la plantilla de la app: https://<usuario>.github.io/<repo>/plantilla.xlsx

// ⬇⬇ ESCRIBE AQUÍ LA CLAVE DE LA APP (la misma que pondrás en el enlace &k=…) ⬇⬇
const CLAVE_APP = '';

const HOJA_GASTOS = 'Movimientos';
const CAMPOS = ['id', 'fecha', 'concepto', 'importe', 'nota', 'sobre', 'creado', 'modificado'];
// mismas columnas que la Excel de la soci (gastos e ingresos)
const CATEGORIAS = ['IBERDROLA', 'LEÑA', 'INTERNET', 'URTASUN', 'BEBIDA', 'MANCOMUNIDAD', 'VARIOS', 'SEGURO', 'EXTRA',
  'SOBRE CENA', 'SOBRE BEBIDA', 'INGRESO CAJA', 'SOBRE LEÑA', 'COBRO INTERNET', 'INGRESO EXTRA',
  'FIESTAS GASTO', 'FIESTAS INGRESO', 'FIESTAS AYTO'];   // fiestas: cuenta aparte (hoja FIESTAS)

/** Muestra la clave que está usando el script (o genera una si no hay ninguna). */
function configurar() {
  if (CLAVE_APP) { Logger.log('Clave de la app: ' + CLAVE_APP); return CLAVE_APP; }
  const props = PropertiesService.getScriptProperties();
  let clave = props.getProperty('CLAVE');
  if (!clave) {
    clave = Utilities.getUuid().replace(/-/g, '').slice(0, 20);
    props.setProperty('CLAVE', clave);
  }
  Logger.log('Clave de la app: ' + clave);
  return clave;
}

function doGet(e) {
  return responder_(() => {
    // sin clave: dice qué script es, para comprobar que la URL apunta al sitio bueno
    if (!e.parameter.k) return { ok: false, error: 'FALTA_CLAVE', app: 'Sociedad Muzki', hoja: SpreadsheetApp.getActive().getName(), claveConfigurada: !!claveActual_() };
    comprobarClave_(e.parameter.k);
    return { ok: true, movimientos: leerGastos_(), hora: new Date().toISOString() };
  });
}

function doPost(e) {
  return responder_(() => {
    const p = JSON.parse(e.postData.contents);
    comprobarClave_(p.k);
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const hechos = [];
      (p.ops || [p]).forEach(op => hechos.push(aplicar_(op)));
      return { ok: true, hechos: hechos };
    } finally {
      lock.releaseLock();
    }
  });
}

function aplicar_(op) {
  const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_GASTOS);
  const ahora = new Date();
  if (op.accion === 'crear') {
    const g = limpiar_(op.mov);
    // si ya existe (reintento sin conexión), no se duplica
    if (g.id && buscarFila_(hoja, g.id) > 0) return { accion: 'crear', id: g.id, repetido: true };
    g.id = g.id || 'M' + Utilities.formatDate(ahora, 'UTC', 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
    g.creado = ahora; g.modificado = ahora;
    hoja.appendRow(aFila_(g));
    return { accion: 'crear', id: g.id };
  }
  if (op.accion === 'editar') {
    const fila = buscarFila_(hoja, op.mov.id);
    if (fila < 0) throw new Error('No encuentro el apunte ' + op.mov.id);
    const actual = hoja.getRange(fila, 1, 1, CAMPOS.length).getValues()[0];
    const g = limpiar_(op.mov);
    g.creado = actual[CAMPOS.indexOf('creado')] || '';
    g.modificado = ahora;
    hoja.getRange(fila, 1, 1, CAMPOS.length).setValues([aFila_(g)]);
    return { accion: 'editar', id: g.id };
  }
  if (op.accion === 'borrar') {
    const fila = buscarFila_(hoja, op.id);
    if (fila > 0) {
      // copia de seguridad antes de borrar
      const papelera = SpreadsheetApp.getActive().getSheetByName('Papelera') ||
        SpreadsheetApp.getActive().insertSheet('Papelera');
      papelera.appendRow(hoja.getRange(fila, 1, 1, CAMPOS.length).getValues()[0].concat([ahora]));
      hoja.deleteRow(fila);
    }
    return { accion: 'borrar', id: op.id };
  }
  throw new Error('Acción desconocida');
}

function limpiar_(g) {
  const cat = String(g.concepto || '').toUpperCase().trim();
  if (CATEGORIAS.indexOf(cat) < 0) throw new Error('Concepto desconocido: ' + g.concepto);
  const importe = Math.round(Number(g.importe) * 100) / 100;
  if (!isFinite(importe) || importe <= 0) throw new Error('Importe no válido');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(g.fecha || ''))) throw new Error('Fecha no válida');
  return { id: g.id || '', fecha: g.fecha, concepto: cat, importe: importe, nota: String(g.nota || '').slice(0, 200), sobre: String(g.sobre || '').slice(0, 40) };
}

function leerGastos_() {
  const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_GASTOS);
  const n = hoja.getLastRow() - 1;
  if (n < 1) return [];
  const tz = Session.getScriptTimeZone();
  return hoja.getRange(2, 1, n, 6).getValues()
    .filter(v => v[0] !== '' && v[2] !== '')
    .map(v => [String(v[0]), v[1] instanceof Date ? Utilities.formatDate(v[1], tz, 'yyyy-MM-dd') : String(v[1]).slice(0, 10),
      String(v[2]), Number(v[3]) || 0, String(v[4] || ''), String(v[5] || '')]);
}

function aFila_(g) {
  return CAMPOS.map(c => {
    const x = g[c];
    if (x === undefined || x === null) return '';
    if (c === 'fecha') { const p = String(x).split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); }
    return x;
  });
}

function buscarFila_(hoja, id) {
  const n = hoja.getLastRow() - 1;
  if (n < 1) return -1;
  const ids = hoja.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function claveActual_() {
  return String(CLAVE_APP || PropertiesService.getScriptProperties().getProperty('CLAVE') || '').trim();
}
function comprobarClave_(k) {
  const clave = claveActual_();
  if (!clave || String(k || '').trim() !== clave) throw new Error('CLAVE');
}

function responder_(fn) {
  let out;
  try { out = fn(); } catch (err) { out = { ok: false, error: String(err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ====================== copia de seguridad semanal ====================== */
/** Ejecútala una vez: guarda una copia de la hoja cada lunes en la carpeta "Muzki copias" (las 8 últimas). */
function activarCopias() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'copiaSemanal').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('copiaSemanal').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(4).create();
  copiaSemanal();
}
function copiaSemanal() {
  const ss = SpreadsheetApp.getActive();
  const madre = DriveApp.getFileById(ss.getId()).getParents();
  const base = madre.hasNext() ? madre.next() : DriveApp.getRootFolder();
  const it = base.getFoldersByName('Muzki copias');
  const carpeta = it.hasNext() ? it.next() : base.createFolder('Muzki copias');
  DriveApp.getFileById(ss.getId()).makeCopy('Muzki · copia ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'), carpeta);
  const copias = []; const f = carpeta.getFiles(); while (f.hasNext()) copias.push(f.next());
  copias.sort((x, y) => y.getDateCreated() - x.getDateCreated()).slice(8).forEach(x => x.setTrashed(true));
}

/* ====================== la Excel de la soci, cada mes por correo ====================== */
// Rellena la misma plantilla que la app (mismas columnas, un sobre por fila) y la envía.
const MESES_N = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
const COL_DE = { 'IBERDROLA': 'C', 'LEÑA': 'E', 'INTERNET': 'G', 'URTASUN': 'I', 'BEBIDA': 'K', 'MANCOMUNIDAD': 'M', 'VARIOS': 'O',
  'SEGURO': 'Q', 'EXTRA': 'Q', 'SOBRE CENA': 'S', 'SOBRE BEBIDA': 'U', 'INGRESO CAJA': 'W', 'SOBRE LEÑA': 'Y', 'COBRO INTERNET': 'AA', 'INGRESO EXTRA': 'AC' };
const PREFIJO_XL = { 'SEGURO': 'Seguro' };
const PARTES_SOBRE = ['SOBRE CENA', 'SOBRE BEBIDA', 'SOBRE LEÑA'];
const INGRESOS_XL = ['SOBRE CENA', 'SOBRE BEBIDA', 'INGRESO CAJA', 'SOBRE LEÑA', 'COBRO INTERNET', 'INGRESO EXTRA'];
const COLS_XL = 'C D E F G H I J K L M N O P Q R S T U V W X Y Z AA AB AC AD'.split(' ');

/** Ejecútala una vez: deja programado el envío el día 1 de cada mes a las 8. */
function activarEnvioMensual() {
  if (!CORREO_EXCEL || !URL_PLANTILLA) throw new Error('Rellena CORREO_EXCEL y URL_PLANTILLA arriba del todo.');
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'enviarExcelMensual').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('enviarExcelMensual').timeBased().onMonthDay(1).atHour(8).create();
  Logger.log('Listo: el día 1 de cada mes se enviará la Excel a ' + CORREO_EXCEL);
}
/** Para probar: envía ya la Excel con el mes pasado. */
function probarEnvio() { enviarExcelMensual(); }

function enviarExcelMensual() {
  const hoy = new Date();
  const cerrado = new Date(hoy.getFullYear(), hoy.getMonth() - 1, 1);           // el mes que acaba de terminar
  const ym = cerrado.getFullYear() + '-' + ('0' + (cerrado.getMonth() + 1)).slice(-2);
  const anio = cerrado.getMonth() + 1 >= 3 ? cerrado.getFullYear() : cerrado.getFullYear() - 1;   // de marzo a febrero
  const movs = leerGastos_().map(r => ({ id: r[0], fecha: r[1], concepto: r[2], importe: r[3], nota: r[4], sobre: r[5] }));
  const blob = excelDelAnio_(anio, movs);
  const delMes = movs.filter(m => m.fecha.slice(0, 7) === ym && m.concepto.indexOf('FIESTAS') !== 0);
  const entra = delMes.filter(m => INGRESOS_XL.indexOf(m.concepto) >= 0).reduce((a, m) => a + m.importe, 0);
  const sale = delMes.filter(m => INGRESOS_XL.indexOf(m.concepto) < 0).reduce((a, m) => a + m.importe, 0);
  const sobres = {}; delMes.forEach(m => { if (m.sobre) sobres[m.sobre] = 1; });
  const e = n => n.toLocaleString('es-ES', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' €';
  const mes = MESES_N[cerrado.getMonth()] + ' ' + cerrado.getFullYear();
  // copia en Drive
  const it = DriveApp.getFoldersByName('Muzki · Excel');
  (it.hasNext() ? it.next() : DriveApp.createFolder('Muzki · Excel')).createFile(blob.copyBlob().setName('Gastos_Muzqui_' + anio + '-' + (anio + 1) + ' (' + mes + ').xlsx'));
  if (!CORREO_EXCEL) return;
  MailApp.sendEmail({
    to: CORREO_EXCEL,
    subject: 'Sociedad Muzki · cuentas de ' + mes,
    htmlBody: '<p>Cuentas de <b>' + mes + '</b>:</p><ul><li>Entra: <b>' + e(entra) + '</b> (' + Object.keys(sobres).length + ' sobres)</li>' +
      '<li>Sale: <b>' + e(sale) + '</b></li><li>Queda: <b>' + e(entra - sale) + '</b></li></ul>' +
      '<p>Va adjunta la Excel del año ' + anio + '-' + (anio + 1) + ' con todo lo apuntado hasta el ' + ('0' + new Date(hoy.getFullYear(), hoy.getMonth(), 0).getDate()).slice(-2) + ' de ' + MESES_N[cerrado.getMonth()] + '.</p>',
    attachments: [blob]
  });
}

function excelDelAnio_(anio, movs) {
  const r = UrlFetchApp.fetch(URL_PLANTILLA, { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('No encuentro la plantilla en ' + URL_PLANTILLA);
  const partes = Utilities.unzip(r.getBlob().setContentType('application/zip'));
  const salida = partes.map(b => {
    const m = b.getName().match(/^xl\/worksheets\/sheet(\d+)\.xml$/);
    const n = m ? +m[1] : 0;
    if (n === 14) return Utilities.newBlob(hojaFiestasXml(b.getDataAsString('UTF-8'), lineasFiestas_(movs, anio)), 'application/xml', b.getName());
    if (n < 2 || n > 13) return b;                       // sheet2 = Marzo … sheet13 = Febrero
    const mes = (n - 2 + 2) % 12 + 1, y = mes >= 3 ? anio : anio + 1;
    const ym = y + '-' + ('0' + mes).slice(-2);
    const xml = rellenarMesXl_(b.getDataAsString('UTF-8'), movs.filter(x => x.fecha.slice(0, 7) === ym && x.concepto.indexOf('FIESTAS') !== 0));
    return Utilities.newBlob(xml, 'application/xml', b.getName());
  });
  return Utilities.zip(salida, 'Gastos_Muzqui_' + anio + '-' + (anio + 1) + '.xlsx')
    .setContentType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
}

function filasXl_(movs) {
  const lineas = [], grupos = {};
  movs.forEach(m => {
    if (m.sobre && PARTES_SOBRE.indexOf(m.concepto) >= 0) {
      let g = grupos[m.sobre];
      if (!g) { g = grupos[m.sobre] = { sobre: true, fecha: m.fecha, quien: m.nota, partes: {} }; lineas.push(g); }
      g.partes[m.concepto] = m.importe; if (m.fecha > g.fecha) g.fecha = m.fecha;
    } else lineas.push(m);
  });
  const ddmm = f => (+f.slice(8, 10)) + '.' + (+f.slice(5, 7));
  const obs = c => COLS_XL[COLS_XL.indexOf(c) + 1];
  return lineas.sort((a, b) => a.fecha < b.fecha ? -1 : a.fecha > b.fecha ? 1 : 0).map(a => {
    const v = {};
    if (a.sobre) PARTES_SOBRE.forEach(p => { if (a.partes[p]) { v[COL_DE[p]] = a.partes[p]; v[obs(COL_DE[p])] = ('sobre ' + String(a.quien || '').toLowerCase() + ' ' + ddmm(a.fecha)).replace(/\s+/g, ' '); } });
    else {
      const col = COL_DE[a.concepto]; const nota = [PREFIJO_XL[a.concepto], a.nota].filter(Boolean).join(' · ');
      v[col] = a.importe; v[obs(col)] = /^X/.test(a.id) ? nota : ddmm(a.fecha) + (nota ? ' ' + nota : '');
    }
    return v;
  });
}
function celdaXl_(fila, col, n, valor) {
  const ref = col + n;
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const re = new RegExp('<c r="' + ref + '"((?:\\s+[a-zA-Z:]+="[^"]*")*)\\s*(?:\\/>|>[\\s\\S]*?<\\/c>)');
  const prev = fila.match(re);
  const sm = prev && prev[1].match(/ s="(\d+)"/);
  const s = sm ? ' s="' + sm[1] + '"' : '';
  const celda = typeof valor === 'number' ? '<c r="' + ref + '"' + s + '><v>' + valor + '</v></c>'
    : '<c r="' + ref + '"' + s + ' t="inlineStr"><is><t xml:space="preserve">' + esc(valor) + '</t></is></c>';
  if (prev) return fila.replace(re, celda);
  const num = c => { let k = 0; for (let i = 0; i < c.length; i++) k = k * 26 + c.charCodeAt(i) - 64; return k; };
  const k = num(col), cells = /<c r="([A-Z]+)\d+"/g; let m;
  while ((m = cells.exec(fila))) { if (num(m[1]) > k) return fila.slice(0, m.index) + celda + fila.slice(m.index); }
  return fila.replace(/<\/row>$/, celda + '</row>');
}
function rellenarMesXl_(xml, movs) {
  const filas = filasXl_(movs), hechas = {};
  xml = xml.replace(/<row r="(\d+)"[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g, (fila, nn) => {
    const n = +nn, v = filas[n - 6]; if (n < 6 || n > 488 || !v) return fila;
    if (/\/>$/.test(fila) && !/<\/row>$/.test(fila)) fila = fila.replace(/\/>$/, '></row>');
    hechas[n] = 1;
    COLS_XL.forEach(c => { if (v[c] !== undefined) fila = celdaXl_(fila, c, n, v[c]); });
    return fila;
  });
  filas.forEach((v, i) => {
    const n = i + 6; if (hechas[n] || n > 488) return;
    let r = '<row r="' + n + '"></row>';
    COLS_XL.forEach(c => { if (v[c] !== undefined) r = celdaXl_(r, c, n, v[c]); });
    const re = /<row r="(\d+)"/g; let m, pos = -1;
    while ((m = re.exec(xml))) { if (+m[1] > n) { pos = m.index; break; } }
    xml = pos >= 0 ? xml.slice(0, pos) + r + xml.slice(pos) : xml.replace('</sheetData>', r + '</sheetData>');
  });
  return xml;
}

function lineasFiestas_(movs, anio) {
  const m = {}, orden = [];
  movs.filter(g => g.concepto.indexOf('FIESTAS') === 0 && +g.fecha.slice(0, 4) === anio).forEach(g => {
    const k = g.sobre || g.id;
    if (!m[k]) { m[k] = { id: k, fecha: g.fecha, concepto: g.nota, gasto: 0, ingreso: 0, ayto: false }; orden.push(m[k]); }
    if (g.concepto === 'FIESTAS GASTO') m[k].gasto += g.importe; else m[k].ingreso += g.importe;
    if (g.concepto === 'FIESTAS AYTO') m[k].ayto = true;
  });
  return orden.sort((a, b) => a.fecha < b.fecha ? -1 : a.fecha > b.fecha ? 1 : a.id < b.id ? -1 : 1);
}
/* hoja FIESTAS de la Excel: una línea por concepto (gasto, ingreso, y «ayunta» si lo paga el ayuntamiento),
   el total con su beneficio, y debajo lo que paga el ayuntamiento */
function hojaFiestasXml(xml, lineas) {
  const e = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const t = (ref, v, s) => '<c r="' + ref + '"' + (s ? ' s="' + s + '"' : '') + ' t="inlineStr"><is><t xml:space="preserve">' + e(v) + '</t></is></c>';
  const n = (ref, v, s) => '<c r="' + ref + '"' + (s ? ' s="' + s + '"' : '') + (v === '' ? '/>' : '><v>' + v + '</v></c>');
  const f = (ref, fo, s) => '<c r="' + ref + '"' + (s ? ' s="' + s + '"' : '') + '><f>' + fo + '</f></c>';
  const filas = ['<row r="1"/>', '<row r="2">' + t('A2', 'CONCEPTO', 99) + t('B2', 'GASTOS', 99) + t('C2', 'INGRESOS', 100) + t('D2', 'OBSERVACIONES', 99) + '</row>'];
  let r = 3;
  lineas.forEach(l => {
    filas.push('<row r="' + r + '">' + t('A' + r, l.concepto, 76) + n('B' + r, l.gasto || '', 76) + n('C' + r, l.ingreso || '', 76) + (l.ayto ? t('D' + r, 'ayunta') : '') + '</row>');
    r++;
  });
  const fin = Math.max(r - 1, 3), tot = Math.max(r, 4);
  filas.push('<row r="' + tot + '">' + t('A' + tot, 'TOTAL', 99) + f('B' + tot, 'SUM(B3:B' + fin + ')', 105) + f('C' + tot, 'SUM(C3:C' + fin + ')', 105) + f('D' + tot, 'C' + tot + '-B' + tot, 106) + t('E' + tot, '← beneficio') + '</row>');
  const ay = lineas.filter(l => l.ayto);
  let a = tot + 3;
  filas.push('<row r="' + a + '">' + t('A' + a, 'Lo paga el ayuntamiento') + '</row>');
  const ini = a + 1;
  ay.forEach(l => { a++; filas.push('<row r="' + a + '">' + t('B' + a, l.concepto) + n('C' + a, l.ingreso || l.gasto) + '</row>'); });
  a++; filas.push('<row r="' + a + '">' + t('B' + a, 'TOTAL') + f('C' + a, 'SUM(C' + ini + ':C' + Math.max(a - 1, ini) + ')') + '</row>');
  return xml.replace(/<dimension ref="[^"]*"\/>/, '<dimension ref="A1:E' + a + '"/>')
    .replace(/<sheetData>[\s\S]*<\/sheetData>|<sheetData\/>/, '<sheetData>' + filas.join('') + '</sheetData>');
}
