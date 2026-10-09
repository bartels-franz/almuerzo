/**
 * Hoja de pedidos — Almuerzo del sábado 10-10-26
 *
 * 1) Pega este código en Extensiones > Apps Script de una Hoja de cálculo de Google nueva.
 * 2) Ejecuta la función `setup` una vez (crea y da formato a las dos pestañas).
 * 3) Implementar > Nueva implementación > Aplicación web
 *    Ejecutar como: Yo · Quién tiene acceso: Cualquier persona.
 * 4) Copia la URL que termina en /exec y pégala en config.js de la página.
 *
 * Volver a ejecutar `setup` restaura el formato sin borrar los pedidos.
 * Los pedidos se guardan buscando el # del invitado y se comprueba que el nombre coincida.
 * Al ejecutar `setup`, cada pedido se reubica por nombre y orden en la lista
 * (1ª Sofía, 2ª Sofía…), así que renumerar invitados no mezcla pedidos.
 */

// ---------- Datos del Excel ----------
var GUESTS = [
  [1,'Myriam',0],[2,'Sebastían',0],[3,'Andrea',0],[4,'Sofía',0],[5,'Daniel',0],
  [6,'Johana',0],[7,'Nicolas',0],[8,'Sofía',0],[9,'Mariana',0],[10,'Arturo',0],
  [11,'Nicoll',1],[12,'Paula',0],[13,'Richard',0],[14,'Martina',1],[15,'Pilar',0],
  [16,'Oscar',0],[17,'Matías',1],[18,'Ana María',0],[19,'Javier',0],[20,'Cristina',0],
  [21,'Hanna',1],[22,'Alejandra',0],[23,'Andrés',0],[24,'Thiago',1],[25,'Franz',0],
  [26,'Juan',0],[27,'Camila',0],[28,'Julián',0],[29,'Juan Pablo',0]
];
var LABELS = {
  ajiaco_normal: 'Ajiaco plato normal', ajiaco_pequeno: 'Ajiaco plato pequeño',
  frijol: 'Frijol', verdura: 'Verdura con carne molida', pasta: 'Pasta con carne molida',
  lomo: 'Lomo de cerdo a la plancha', bistec: 'Carne en bisteck', pechuga: 'Pechuga a la plancha', pollo: 'Pollo al horno',
  mora: 'Mora', mango: 'Mango'
};
var PENDIENTE = 'Pendiente';
var KID = 'NIÑO/A';

// ---------- Estilo de la página web ----------
var C = {
  bg: '#F5F6F0', surface: '#FFFFFF', surface2: '#ECEFE6', ink: '#17211B', muted: '#5B675F',
  line: '#D9DED3', green: '#173A2A', green2: '#24573F', corn: '#E9B42F', cornInk: '#2A1F00', ok: '#1F7A4D'
};
var F_DISPLAY = 'Young Serif';
var F_BODY = 'Figtree';

var TAB_RESUMEN = 'Resumen para la cocina';
var TAB_PEDIDOS = 'Pedido por invitado';

// Pestaña "Pedido por invitado": encabezado en fila 5, invitados desde la fila 6 (B..J)
var HEAD = 5, FIRST = 6, LAST = FIRST + GUESTS.length - 1;
var COLS = ['#', 'Invitado', 'Niño/a', 'Sopita', 'Principio', 'Proteína', 'Jugo', 'Observación', 'Última actualización'];
var COL = { num: 2, nombre: 3, nino: 4, sopa: 5, principio: 6, proteina: 7, jugo: 8, obs: 9, fecha: 10 };
function colL_(n) { return String.fromCharCode(64 + n); }
function rng_(key) { return "'" + TAB_PEDIDOS + "'!" + colL_(COL[key]) + FIRST + ':' + colL_(COL[key]) + LAST; }

// =====================================================================
function setup() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { setup_(); } finally { lock.releaseLock(); }
}

function setup_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ped = buildPedidos_(ss);
  var res = buildResumen_(ss);
  ss.setActiveSheet(res); ss.moveActiveSheet(1);
  ss.setActiveSheet(ped); ss.moveActiveSheet(2);
  // Quitar la hoja vacía por defecto
  ss.getSheets().forEach(function (s) {
    var n = s.getName();
    if (n !== TAB_RESUMEN && n !== TAB_PEDIDOS && s.getLastRow() === 0) ss.deleteSheet(s);
  });
  ss.setActiveSheet(res);
}

function baseSheet_(ss, name, maxCol, maxRow) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.setFrozenRows(0);
  sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) { p.remove(); });
  sh.setHiddenGridlines(true);
  sh.setTabColor(C.green);
  if (sh.getMaxColumns() > maxCol) sh.deleteColumns(maxCol + 1, sh.getMaxColumns() - maxCol);
  if (sh.getMaxRows() < maxRow) sh.insertRowsAfter(sh.getMaxRows(), maxRow - sh.getMaxRows());
  if (sh.getMaxRows() > maxRow) sh.deleteRows(maxRow + 1, sh.getMaxRows() - maxRow);
  sh.getRange(1, 1, maxRow, maxCol).setBackground(C.bg).setFontFamily(F_BODY).setFontColor(C.ink)
    .setFontSize(11).setVerticalAlignment('middle');
  sh.setColumnWidth(1, 28);
  // Franja superior verde con la etiqueta del evento (como el encabezado de la página)
  sh.getRange(1, 1, 1, maxCol).setBackground(C.green);
  sh.setRowHeight(1, 34);
  sh.getRange(1, 2).setValue('PEDIDO · SÁBADO 10-10-26').setFontColor(C.corn).setFontWeight('bold').setFontSize(9);
  sh.setRowHeight(2, 14);
  sh.getRange(2, 1, 1, maxCol).setBackground(C.bg);
  return sh;
}

function title_(sh, text, subFormula) {
  sh.getRange(3, 2).setValue(text).setFontFamily(F_DISPLAY).setFontSize(22).setFontColor(C.ink).setFontWeight('normal');
  sh.setRowHeight(3, 42);
  sh.getRange(4, 2).setFormula(subFormula).setFontColor(C.muted).setFontSize(11);
  sh.setRowHeight(4, 26);
}

var SUB_FORMULA = (function () {
  var E = rng_('sopa'), D = rng_('nino');
  var kids = 'COUNTIFS(' + D + ',"' + KID + '",' + E + ',"<>' + PENDIENTE + '")';
  return '=COUNTIF(' + E + ',"<>' + PENDIENTE + '")&" de ' + GUESTS.length + ' invitados han registrado su pedido"&IF(' +
    kids + '>0," ("&' + kids + '&" niños)",".")';
})();

// Clave estable de un invitado: nombre + número de aparición de ese nombre en la lista
// (Sofía#1, Sofía#2…). No depende del # de la fila, así que sobrevive a una renumeración.
function keyFor_(names) {
  var seen = {};
  return names.map(function (n) { n = String(n); seen[n] = (seen[n] || 0) + 1; return n + '#' + seen[n]; });
}

// Lee los pedidos actuales indexados por esa clave, en el orden en que están en la hoja
function existingByKey_(sh) {
  var map = {};
  var last = sh.getLastRow();
  if (last < FIRST) return map;
  var rows = sh.getRange(FIRST, 2, last - FIRST + 1, COLS.length).getValues()
    .filter(function (r) { return r[1] !== '' && r[1] !== null; })
    .sort(function (a, b) { return Number(a[0]) - Number(b[0]); });
  var keys = keyFor_(rows.map(function (r) { return r[1]; }));
  rows.forEach(function (r, i) { if (r[3] && r[3] !== PENDIENTE) map[keys[i]] = r; });
  return map;
}

// Fila de un invitado según su # en la columna B (null si no está)
function rowOf_(sh, id) {
  var last = sh.getLastRow();
  if (last < FIRST) return null;
  var ids = sh.getRange(FIRST, 2, last - FIRST + 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (Number(ids[i][0]) === id) return FIRST + i;
  return null;
}

// ---------- Pestaña: Pedido por invitado ----------
function buildPedidos_(ss) {
  var prev = ss.getSheetByName(TAB_PEDIDOS);
  var saved = prev ? existingByKey_(prev) : {};  // pedidos guardados, antes de tocar nada
  if (prev && prev.getMaxRows() >= FIRST) {         // limpiar la tabla vieja (se reescribe abajo)
    prev.getRange(FIRST, 1, prev.getMaxRows() - FIRST + 1, prev.getMaxColumns()).clear();
  }
  var sh = baseSheet_(ss, TAB_PEDIDOS, 11, LAST + 2);
  title_(sh, 'Pedido por invitado', SUB_FORMULA);

  var widths = [28, 44, 130, 78, 170, 230, 200, 80, 220, 150, 28];
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });

  // Encabezado de la tabla (como el thead de la página)
  var head = sh.getRange(HEAD, 2, 1, COLS.length);
  head.setValues([COLS.map(function (c) { return c.toUpperCase(); })])
    .setBackground(C.surface2).setFontColor(C.muted).setFontWeight('bold').setFontSize(9);
  sh.setRowHeight(HEAD, 34);

  // Filas de invitados: cada pedido vuelve a la fila de la misma persona
  var body = sh.getRange(FIRST, 2, GUESTS.length, COLS.length);
  var gkeys = keyFor_(GUESTS.map(function (g) { return g[1]; }));
  var vals = GUESTS.map(function (g, i) {
    var r = saved[gkeys[i]];
    return [g[0], g[1], g[2] ? KID : '',
      r ? r[3] : PENDIENTE, r ? r[4] : '', r ? r[5] : '', r ? r[6] : '', r ? r[7] : '', r ? r[8] : ''];
  });
  body.setValues(vals).setBackground(C.surface).setFontColor(C.ink).setFontSize(11).setWrap(true).setVerticalAlignment('top');
  sh.getRange(FIRST, 2, GUESTS.length, 1).setFontColor(C.muted).setHorizontalAlignment('left');
  sh.getRange(FIRST, COL.nino, GUESTS.length, 1).setFontSize(8).setFontWeight('bold').setHorizontalAlignment('center').setFontColor(C.cornInk);
  sh.getRange(FIRST, COL.fecha, GUESTS.length, 1).setNumberFormat('dd/mm/yyyy hh:mm').setFontColor(C.muted).setFontSize(10);
  for (var r = FIRST; r <= LAST; r++) sh.setRowHeight(r, 36);

  // Bordes: contorno y líneas entre filas (#D9DED3), como la tabla de la página
  var table = sh.getRange(HEAD, 2, GUESTS.length + 1, COLS.length);
  table.setBorder(true, true, true, true, false, true, C.line, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(FIRST, 2, GUESTS.length, COLS.length).setVerticalAlignment('middle');

  // Formato condicional: pendientes en gris, invitados con pedido en negrita, chip NIÑO/A
  var rules = [];
  var all = sh.getRange(FIRST, 3, GUESTS.length, 1);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$E' + FIRST + '="' + PENDIENTE + '"').setFontColor(C.muted).setRanges([all, sh.getRange(FIRST, COL.sopa, GUESTS.length, 1)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$E' + FIRST + '<>"' + PENDIENTE + '"').setBold(true).setFontColor(C.ink).setRanges([all]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(KID).setBackground(C.corn).setFontColor(C.cornInk).setRanges([sh.getRange(FIRST, COL.nino, GUESTS.length, 1)]).build());
  sh.setConditionalFormatRules(rules);

  sh.setFrozenRows(HEAD);
  sh.getRange(HEAD, 2, GUESTS.length + 1, COLS.length).protect().setWarningOnly(true)
    .setDescription('Se llena desde la página web. Edita con cuidado.');
  return sh;
}

// ---------- Pestaña: Resumen para la cocina ----------
var CARDS = [
  { title: 'Sopita',    col: 2, row: 6,  col2: 'sopa',      items: [['ajiaco_normal', '='], ['ajiaco_pequeno', '=']] },
  { title: 'Principio', col: 5, row: 6,  col2: 'principio', items: [['frijol', '*'], ['verdura', '*'], ['pasta', '*']] },
  { title: 'Proteína',  col: 2, row: 12, col2: 'proteina',  items: [['lomo', '='], ['bistec', '='], ['pechuga', '='], ['pollo', '=']] },
  { title: 'Jugo',      col: 5, row: 12, col2: 'jugo',      items: [['mora', '='], ['mango', '=']] }
];

function buildResumen_(ss) {
  var sh = baseSheet_(ss, TAB_RESUMEN, 8, 20);
  title_(sh, 'Resumen para la cocina', SUB_FORMULA);
  [28, 230, 70, 20, 230, 70, 28, 28].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setRowHeight(5, 16);
  sh.setRowHeight(11, 16);

  CARDS.forEach(function (card) {
    var n = card.items.length;
    // tarjeta blanca con borde
    var box = sh.getRange(card.row, card.col, n + 1, 2);
    box.clearContent().setBackground(C.surface)
      .setBorder(true, true, true, true, false, false, C.line, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(card.row, card.col).setValue(card.title.toUpperCase())
      .setFontColor(C.muted).setFontWeight('bold').setFontSize(9);
    sh.setRowHeight(card.row, 32);
    card.items.forEach(function (it, i) {
      var r = card.row + 1 + i;
      var label = LABELS[it[0]];
      var crit = it[1] === '*' ? '"*' + label + '*"' : '"' + label + '"';
      sh.getRange(r, card.col).setValue(label).setFontColor(C.ink).setFontSize(11);
      sh.getRange(r, card.col + 1).setFormula('=COUNTIF(' + rng_(card.col2) + ',' + crit + ')')
        .setFontColor(C.green2).setFontWeight('bold').setFontSize(14).setHorizontalAlignment('right');
      if (sh.getRowHeight(r) < 34) sh.setRowHeight(r, 34);
      if (i > 0) sh.getRange(r, card.col, 1, 2).setBorder(true, null, null, null, null, null, C.line, SpreadsheetApp.BorderStyle.SOLID);
    });
  });
  sh.getRange(18, 2).setValue('Se actualiza solo cada vez que alguien envía o cambia su pedido.')
    .setFontColor(C.muted).setFontSize(9).setFontStyle('italic');
  return sh;
}

// =====================================================================
// API para la página web
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function keyOf_(label, keys) {
  for (var i = 0; i < keys.length; i++) if (LABELS[keys[i]] === label) return keys[i];
  return '';
}
function pedidos_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(TAB_PEDIDOS);
  if (!sh) { setup_(); sh = ss.getSheetByName(TAB_PEDIDOS); }
  return sh;
}

function readAll_() {
  var sh = pedidos_();
  var last = sh.getLastRow();
  if (last < FIRST) return [];
  var valid = {};
  GUESTS.forEach(function (g) { valid[g[0]] = true; });
  var vals = sh.getRange(FIRST, 2, last - FIRST + 1, COLS.length).getValues();
  var out = [];
  vals.forEach(function (r) {
    if (!valid[Number(r[0])] || !r[3] || r[3] === PENDIENTE) return;
    var prin = String(r[4] || '');
    out.push({
      id: Number(r[0]), nombre: r[1], nino: r[2] === KID,
      sopa: keyOf_(r[3], ['ajiaco_normal', 'ajiaco_pequeno']),
      frijol: prin.indexOf('Frijol') >= 0,
      acomp: prin.indexOf(LABELS.verdura) >= 0 ? 'verdura' : (prin.indexOf(LABELS.pasta) >= 0 ? 'pasta' : ''),
      proteina: keyOf_(r[5], ['lomo', 'bistec', 'pechuga', 'pollo']),
      jugo: keyOf_(r[6], ['mora', 'mango']),
      obs: r[7] || '',
      fecha: r[8] instanceof Date ? r[8].toISOString() : ''
    });
  });
  return out;
}

function doGet() {
  try { return json_({ ok: true, orders: readAll_() }); }
  catch (e) { return json_({ ok: false, error: String(e) }); }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var o = body.order || {};
    var id = Number(o.id), idx = -1;
    for (var i = 0; i < GUESTS.length; i++) if (GUESTS[i][0] === id) idx = i;
    if (idx < 0) return json_({ ok: false, error: 'Invitado no válido' });

    var acomp = o.acomp || '';
    var valid = ['ajiaco_normal', 'ajiaco_pequeno'].indexOf(o.sopa) >= 0 &&
      (acomp === '' || ['verdura', 'pasta'].indexOf(acomp) >= 0) &&
      (!!o.frijol || acomp !== '') &&
      ['lomo', 'bistec', 'pechuga', 'pollo'].indexOf(o.proteina) >= 0 &&
      ['mora', 'mango'].indexOf(o.jugo) >= 0;
    if (!valid) return json_({ ok: false, error: 'Pedido incompleto' });

    var prin = [];
    if (o.frijol) prin.push(LABELS.frijol);
    if (acomp) prin.push(LABELS[acomp]);
    var obs = String(o.obs || '').slice(0, 300);
    if (/^[=+\-@]/.test(obs)) obs = "'" + obs; // evita que se interprete como fórmula
    var now = new Date();

    var sh = pedidos_();
    var row = rowOf_(sh, id);
    if (!row) return json_({ ok: false, error: 'Invitado no encontrado en la hoja' });
    // Seguridad: el nombre de esa fila debe ser el del invitado (evita escribir en la fila de otro)
    if (String(sh.getRange(row, COL.nombre).getValue()) !== GUESTS[idx][1])
      return json_({ ok: false, error: 'La hoja no está actualizada: ejecuta setup' });
    sh.getRange(row, COL.sopa, 1, 6)
      .setValues([[LABELS[o.sopa], prin.join(' + '), LABELS[o.proteina], LABELS[o.jugo], obs, now]]);
    return json_({ ok: true, order: { id: id, nombre: GUESTS[idx][1], sopa: o.sopa, frijol: !!o.frijol, acomp: acomp,
      proteina: o.proteina, jugo: o.jugo, obs: String(o.obs || '').slice(0, 300), fecha: now.toISOString() } });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}
