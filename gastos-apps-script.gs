/**************************************************************************
 *  MÓDULO DE GASTOS — para el Apps Script de Moctezuma (Google Sheets)
 *  --------------------------------------------------------------------
 *  La app (index.html v34l+) ya llama estas 3 acciones vía JSONP:
 *     ?action=getGastos
 *     ?action=saveGasto&data=<json>
 *     ?action=deleteGasto&data=<json>&user=<correo>
 *  Sin este snippet, los gastos solo se guardan en el equipo (localStorage).
 *
 *  CÓMO INSTALAR:
 *  1) Abre tu Google Sheet → Extensiones → Apps Script.
 *  2) Pega TODAS las funciones de abajo al final del archivo (Code.gs).
 *  3) En tu función doGet(e), donde ya manejas las acciones de PAGOS
 *     (getPagos / savePago / deletePago), agrega estas 3 ramas usando
 *     EXACTAMENTE el mismo mecanismo de respuesta JSONP que ya usas.
 *     Si tu doGet usa un switch(action) o if/else, quedaría así:
 *
 *        else if (action === 'getGastos')   return _salidaJSONP(e, getGastos());
 *        else if (action === 'saveGasto')   return _salidaJSONP(e, saveGasto(e.parameter.data));
 *        else if (action === 'deleteGasto') return _salidaJSONP(e, deleteGasto(e.parameter.data));
 *
 *     Nota: reemplaza `_salidaJSONP` por el nombre de TU helper de salida
 *     JSONP existente (el que usas para getPagos). Si no tienes uno, usa
 *     el helper `_salidaJSONP` incluido al final de este archivo.
 **************************************************************************/

// Hoja de gastos (se crea con encabezados si no existe)
function _hojaGastos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Gastos');
  if (!sh) {
    sh = ss.insertSheet('Gastos');
    sh.appendRow(['id', 'ub', 'u', 'concepto', 'monto', 'fecha', 'mes', 'usuario', 'timestamp']);
    // Fecha y mes como texto plano para que Sheets no los convierta a número/fecha
    sh.getRange('F:G').setNumberFormat('@');
  }
  return sh;
}

function _fechaTxt(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return v == null ? '' : String(v);
}

// GET: devuelve el arreglo de gastos
function getGastos() {
  var sh = _hojaGastos();
  var v = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < v.length; i++) {
    if (v[i][0] === '' || v[i][0] == null) continue;
    out.push({
      id: String(v[i][0]),
      ub: v[i][1],
      u: v[i][2],
      concepto: v[i][3],
      monto: Number(v[i][4]) || 0,
      fecha: _fechaTxt(v[i][5]),
      mes: v[i][6] == null ? '' : String(v[i][6]),
      usuario: v[i][7],
      timestamp: v[i][8] == null ? '' : String(v[i][8])
    });
  }
  return out;
}

// SAVE: agrega un gasto (data = JSON del objeto gasto)
function saveGasto(data) {
  var g = JSON.parse(data);
  var sh = _hojaGastos();
  sh.appendRow([
    g.id, g.ub, g.u || '', g.concepto, Number(g.monto) || 0,
    _fechaTxt(g.fecha), g.mes || '', g.usuario || '', g.timestamp || new Date().toISOString()
  ]);
  return { ok: true };
}

// DELETE: elimina el/los gasto(s) con ese id (data = {"id": "..."})
function deleteGasto(data) {
  var d = JSON.parse(data);
  var sh = _hojaGastos();
  var v = sh.getDataRange().getValues();
  for (var i = v.length - 1; i >= 1; i--) {
    if (String(v[i][0]) === String(d.id)) sh.deleteRow(i + 1);
  }
  return { ok: true };
}

// Helper de salida JSONP — ÚSALO SOLO si tu Apps Script no tiene ya uno propio.
// Envuelve el resultado en callback(...) igual que lo espera jsonpFetch().
function _salidaJSONP(e, obj) {
  var cb = (e && e.parameter && e.parameter.callback) ? e.parameter.callback : 'callback';
  return ContentService
    .createTextOutput(cb + '(' + JSON.stringify(obj) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
