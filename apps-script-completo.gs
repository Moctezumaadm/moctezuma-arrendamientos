// ═══════════════════════════════════════════════════════
// MOCTEZUMA ADMINISTRADORA INTELIGENTE
// Google Apps Script — v34p (Web App + Permisos + Auditoría + Pagos + Gastos)
// ═══════════════════════════════════════════════════════

const SCRIPT_VERSION = 'v34p';

const SHEET_NAME   = 'Contratos';
const ACCESO_SHEET = 'ACCESO';
const AUDIT_SHEET  = 'AUDITORIA';
const PAGOS_SHEET  = 'PAGOS';
const GASTOS_SHEET = 'GASTOS';

const ADMINS = ['abogado.flores.moctezuma@gmail.com'];

function doGet(e) {
  const action = e.parameter.action;

  if (!action) {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Moctezuma — Control de Arrendamientos')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  const callback = e.parameter.callback || '';
  const data = e.parameter.data ? JSON.parse(decodeURIComponent(e.parameter.data)) : null;
  const user = e.parameter.user ? decodeURIComponent(e.parameter.user) : '';

  let result;

  try {
    // Esquema mínimo (pestaña GASTOS + columnas modoAdmin/pctComision).
    // Idempotente; si falla no debe bloquear la acción solicitada.
    let schemaInfo = null;
    try { schemaInfo = ensureSchema_(); } catch(e) { schemaInfo = { error: e.message }; }

    switch(action) {
      case 'ping':
        result = {
          ok: true,
          version: SCRIPT_VERSION,
          pestanas: SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName()),
          timestamp: new Date().toISOString()
        };
        break;
      case 'setup': {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const contratos = ss.getSheetByName(SHEET_NAME);
        result = {
          ok: true,
          version: SCRIPT_VERSION,
          schema: schemaInfo,
          pestanas: ss.getSheets().map(s => s.getName()),
          encabezadosContratos: contratos && contratos.getLastColumn() > 0
            ? contratos.getRange(1, 1, 1, contratos.getLastColumn()).getValues()[0]
            : []
        };
        break;
      }
      case 'getAll':
        result = getAllContratos();
        break;
      case 'save':
        result = saveContrato(data);
        break;
      case 'saveAll':
        result = clearAll();
        break;
      case 'update':
        result = updateContrato(data);
        break;
      case 'delete':
        if (!user || !ADMINS.includes(user.toLowerCase())) {
          registrarAudit(user || 'desconocido', '—', 'ELIMINACION_BLOQUEADA_SERVIDOR', data, 'Intento sin permisos');
          result = { ok: false, error: 'No autorizado para eliminar' };
          break;
        }
        result = deleteContrato(data);
        registrarAudit(user, 'Administrador', 'ELIMINACION_SERVIDOR', data, 'Eliminado exitosamente');
        break;
      case 'getAcceso':
        result = getAcceso(data);
        break;
      case 'audit':
        if (data) {
          registrarAudit(data.usuario||'', data.rol||'', data.accion||'', data.contratoId||0, data.detalles||'');
        }
        result = { ok: true };
        break;
      case 'getPagos':
        result = getPagos();
        break;
      case 'savePago':
        result = savePago(data);
        break;
      case 'deletePago':
        if (!user || !ADMINS.includes(user.toLowerCase())) {
          registrarAudit(user || 'desconocido', '—', 'PAGO_ELIMINACION_BLOQUEADA_SERVIDOR', data ? data.cid : 0, 'Intento sin permisos');
          result = { ok: false, error: 'No autorizado para eliminar pagos' };
          break;
        }
        result = deletePago(data);
        break;
      case 'migrar':
        result = migrarColumnas();
        break;
      case 'getGastos':
        result = getGastos();
        break;
      case 'saveGasto':
        result = saveGasto(data);
        break;
      case 'deleteGasto':
        if (!user || !ADMINS.includes(user.toLowerCase())) {
          registrarAudit(user || 'desconocido', '—', 'GASTO_ELIMINACION_BLOQUEADA_SERVIDOR', 0, 'Intento sin permisos');
          result = { ok: false, error: 'No autorizado para eliminar gastos' };
          break;
        }
        result = deleteGasto(data, user);
        break;
      default:
        result = { ok: false, error: 'Acción desconocida: ' + action };
    }
  } catch(err) {
    result = { ok: false, error: err.message };
  }

  const jsonStr = JSON.stringify(result);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + jsonStr + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(jsonStr)
    .setMimeType(ContentService.MimeType.JSON);
}

function getAllContratos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && !row[1]) continue;
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      let val = row[j];
      obj[headers[j]] = (val === '' || val === null || val === undefined) ? '' : val;
    }
    if (obj.id) obj.id = Number(obj.id);
    if (obj.r) obj.r = Number(obj.r);
    if (obj.dep) obj.dep = Number(obj.dep);
    if (obj.cpol) obj.cpol = Number(obj.cpol);
    if (obj.dp) obj.dp = Number(obj.dp);
    if (obj.gr) obj.gr = Number(obj.gr);
    if (obj.mor) obj.mor = Number(obj.mor);
    if (obj.pag) obj.pag = Number(obj.pag);
    if (obj.adeudoIni) obj.adeudoIni = Number(obj.adeudoIni);
    if (obj.fechaCorte instanceof Date) obj.fechaCorte = Utilities.formatDate(obj.fechaCorte, 'America/Mexico_City', 'yyyy-MM-dd');
    if (obj.ini instanceof Date) obj.ini = Utilities.formatDate(obj.ini, 'America/Mexico_City', 'yyyy-MM-dd');
    if (obj.fin instanceof Date) obj.fin = Utilities.formatDate(obj.fin, 'America/Mexico_City', 'yyyy-MM-dd');
    if (obj.pol === '' || obj.pol === 'null') obj.pol = null;
    if (obj.mant === true || obj.mant === 'true' || obj.mant === 'TRUE') obj.mant = true;
    else if (obj.mant === false || obj.mant === 'false' || obj.mant === 'FALSE' || obj.mant === '') obj.mant = false;
    result.push(obj);
  }
  return result;
}

function saveContrato(contrato) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = ['id','u','ub','al','dir','inq','obl','prop','pol','cpol','r','dep','caj','ini','fin','dp','gr','mor','bco','cta','tit','clabe','em','pag','mant','tipo','giro','clausulas_add','nt','adeudoIni','fechaCorte'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => contrato[h] !== undefined ? contrato[h] : '');
  sheet.appendRow(row);
  return { ok: true, id: contrato.id };
}

function updateContrato(contrato) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return { ok: false, error: 'Hoja no encontrada' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][idCol]) === Number(contrato.id)) {
      const row = headers.map(h => contrato[h] !== undefined ? contrato[h] : '');
      sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return { ok: true, id: contrato.id };
    }
  }
  return { ok: false, error: 'Contrato no encontrado: ' + contrato.id };
}

function deleteContrato(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return { ok: false, error: 'Hoja no encontrada' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { ok: true, deleted: id };
    }
  }
  return { ok: false, error: 'ID no encontrado: ' + id };
}

function clearAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return { ok: true };
  if (sheet.getLastRow() > 1) {
    sheet.deleteRows(2, sheet.getLastRow() - 1);
  }
  return { ok: true };
}

function getPagos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PAGOS_SHEET);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && !row[2]) continue;
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      let val = row[j];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, 'America/Mexico_City', 'yyyy-MM-dd');
      }
      obj[headers[j]] = (val === null || val === undefined) ? '' : val;
    }
    if (obj.cid) obj.cid = Number(obj.cid);
    if (obj.monto) obj.monto = Number(obj.monto);
    result.push(obj);
  }
  return result;
}

function savePago(pago) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PAGOS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PAGOS_SHEET);
    const headers = ['cid','periodo','fecha','monto','comprobante','referencia','nota','usuario','timestamp'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#0F2027').setFontColor('#7BADA8').setFontWeight('bold');
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => pago[h] !== undefined ? pago[h] : '');
  sheet.appendRow(row);
  registrarAudit(pago.usuario || '', '', 'PAGO_REGISTRADO', pago.cid, (pago.periodo || '') + ' — $' + (pago.monto || 0));
  return { ok: true };
}

function deletePago(d) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PAGOS_SHEET);
  if (!sheet) return { ok: false, error: 'Hoja PAGOS no encontrada' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const cidCol = headers.indexOf('cid');
  const tsCol = headers.indexOf('timestamp');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][cidCol]) === String(d.cid) && String(data[i][tsCol]) === String(d.timestamp)) {
      sheet.deleteRow(i + 1);
      registrarAudit('', 'Administrador', 'PAGO_ELIMINADO', d.cid, 'timestamp: ' + d.timestamp);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Pago no encontrado' };
}


// Garantiza el esquema que necesita v34p: pestaña GASTOS con su fila 1, y
// encabezados modoAdmin/pctComision en la fila 1 de Contratos (buscados por
// nombre, no por letra de columna). Idempotente: se ejecuta en cada doGet y
// no toca ninguna otra pestaña ni filas de datos.
function ensureSchema_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const info = { gastosCreada: false, columnasAgregadas: [] };

  if (!ss.getSheetByName(GASTOS_SHEET)) {
    const sheet = ss.insertSheet(GASTOS_SHEET);
    const headers = ['id','fecha','ubicacion','contratoId','categoria','concepto','monto','comprobante','nota','usuario','timestamp'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#0F2027').setFontColor('#7BADA8').setFontWeight('bold');
    sheet.getRange('B:B').setNumberFormat('@'); // fechas como texto plano
    info.gastosCreada = true;
  }

  const contratos = ss.getSheetByName(SHEET_NAME);
  if (contratos && contratos.getLastColumn() > 0) {
    const fila1 = contratos.getRange(1, 1, 1, contratos.getLastColumn()).getValues()[0];
    ['modoAdmin','pctComision'].forEach(h => {
      if (fila1.indexOf(h) >= 0) return;
      const vacia = fila1.indexOf('');
      const col = (vacia >= 0) ? vacia + 1 : fila1.length + 1; // primera columna vacía de la fila 1
      contratos.getRange(1, col).setValue(h);
      fila1[col - 1] = h;
      info.columnasAgregadas.push(h);
    });
  }
  return info;
}

// Agrega a la hoja Contratos los encabezados que falten (idempotente: si ya
// existen no hace nada). Se invoca con ?action=migrar tras actualizar la app.
function migrarColumnas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return { ok: false, error: 'Hoja no encontrada' };
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const faltan = ['estatus','fechaDesocupacion','modoAdmin','pctComision'].filter(h => headers.indexOf(h) < 0);
  faltan.forEach((h, i) => sheet.getRange(1, headers.length + 1 + i).setValue(h));
  return { ok: true, agregadas: faltan, encabezados: headers.concat(faltan) };
}

// ═══════════ GASTOS (administración integral) ═══════════
// Pestaña GASTOS: id, fecha, ubicacion, contratoId, categoria, concepto,
// monto, comprobante, nota, usuario, timestamp. contratoId vacío = gasto
// general del edificio. Se crea sola si no existe.
function _hojaGastos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(GASTOS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(GASTOS_SHEET);
    const headers = ['id','fecha','ubicacion','contratoId','categoria','concepto','monto','comprobante','nota','usuario','timestamp'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#0F2027').setFontColor('#7BADA8').setFontWeight('bold');
    sheet.getRange('B:B').setNumberFormat('@'); // fechas como texto plano
  }
  return sheet;
}

function getGastos() {
  const sheet = _hojaGastos();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      let val = row[j];
      if (val instanceof Date) val = Utilities.formatDate(val, 'America/Mexico_City', 'yyyy-MM-dd');
      obj[headers[j]] = (val === null || val === undefined) ? '' : val;
    }
    obj.id = String(obj.id);
    if (obj.monto) obj.monto = Number(obj.monto);
    result.push(obj);
  }
  return result;
}

function saveGasto(gasto) {
  const sheet = _hojaGastos();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => gasto[h] !== undefined ? gasto[h] : '');
  sheet.appendRow(row);
  registrarAudit(gasto.usuario || '', '', 'GASTO_REGISTRADO', gasto.contratoId || 0,
    (gasto.ubicacion || '') + ' — ' + (gasto.categoria || '') + ' — ' + (gasto.concepto || '') + ' — $' + (gasto.monto || 0));
  return { ok: true, id: gasto.id };
}

function deleteGasto(d, user) {
  const sheet = _hojaGastos();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][idCol]) === String(d.id)) {
      sheet.deleteRow(i + 1);
      registrarAudit(user || '', 'Administrador', 'GASTO_ELIMINADO', 0, 'id: ' + d.id);
      return { ok: true, deleted: d.id };
    }
  }
  return { ok: false, error: 'Gasto no encontrado: ' + d.id };
}

function getAcceso(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ACCESO_SHEET);
  if (!sheet) return { ok: false, error: 'Hoja ACCESO no encontrada' };
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase().trim() === String(email).toLowerCase().trim()) {
      return { ok: true, email: data[i][0], rol: data[i][1], nombre: data[i][2], notas: data[i][3] || '' };
    }
  }
  return { ok: false, error: 'Usuario no registrado' };
}

function registrarAudit(usuario, rol, accion, contratoId, detalles) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(AUDIT_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(AUDIT_SHEET);
      sheet.getRange(1, 1, 1, 6).setValues([['Fecha', 'Usuario', 'Rol', 'Acción', 'Contrato ID', 'Detalles']]);
      sheet.getRange(1, 1, 1, 6).setBackground('#0F2027').setFontColor('#7BADA8').setFontWeight('bold');
    }
    const timestamp = Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM-dd HH:mm:ss');
    sheet.appendRow([timestamp, usuario, rol, accion, contratoId, detalles]);
  } catch(e) {
    console.error('Error en auditoría:', e.message);
  }
}