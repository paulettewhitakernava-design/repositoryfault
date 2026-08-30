function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'list';
  let result;
  try {
    if (action === 'list') {
      result = { ok: true, data: obtenerMovimientos() };
    } else if (action === 'config') {
      result = { ok: true, data: leerConfig() };
    } else {
      result = { ok: false, error: 'Acción no soportada: ' + action };
    }
  } catch (err) {
    result = { ok: false, error: String(err) };
  }
  return responderJSON(result);
}

function doPost(e) {
  let result;
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'add') {
      guardarMovimiento(body.mov);
      result = { ok: true };
    } else if (body.action === 'update') {
      actualizarMovimiento(body.mov);
      result = { ok: true };
    } else if (body.action === 'delete') {
      eliminarMovimiento(body.id);
      result = { ok: true };
    } else if (body.action === 'set-config') {
      guardarConfigValor(body.key, body.value);
      result = { ok: true };
    } else {
      result = { ok: false, error: 'Acción no soportada: ' + body.action };
    }
  } catch (err) {
    result = { ok: false, error: String(err) };
  }
  return responderJSON(result);
}

function responderJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Movimientos');
  if (!sheet) {
    sheet = ss.insertSheet('Movimientos');
    sheet.appendRow(['id', 'monto', 'fecha', 'nota', 'tipo', 'categoria']);
  }
  return sheet;
}

// Google Sheets auto-convierte texto tipo "2026-08-29" a un valor Date interno.
// Esa mezcla de tipos (string vs Date) es lo que rompe el render en el cliente
// al recargar. Aquí siempre se devuelve una fecha como texto plano.
function normalizarFecha(valor) {
  if (valor instanceof Date) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone() || 'America/Mexico_City', 'yyyy-MM-dd');
  }
  return String(valor || '');
}

function obtenerMovimientos() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1); // sin encabezado
  return rows
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(r => ({
      id: String(r[0]),
      monto: Number(r[1]) || 0,
      fecha: normalizarFecha(r[2]),
      nota: String(r[3] || ''),
      tipo: String(r[4] || '').trim().toLowerCase(),
      categoria: String(r[5] || '').trim().toLowerCase()
    }))
    .reverse();
}

function guardarMovimiento(mov) {
  const sheet = getSheet();
  sheet.appendRow([String(mov.id), Number(mov.monto), mov.fecha, mov.nota || '', mov.tipo, mov.categoria || '']);
  return true;
}

function actualizarMovimiento(mov) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(mov.id)) {
      sheet.getRange(i + 1, 1, 1, 6).setValues([[
        String(mov.id), Number(mov.monto), mov.fecha, mov.nota || '', mov.tipo, mov.categoria || ''
      ]]);
      return true;
    }
  }
  throw new Error('Movimiento no encontrado');
}

function eliminarMovimiento(id) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return true;
}

// ---- Config: meta de ahorro, cierres de mes ya preguntados, presupuestos por
// categoría editados a mano, y datos de perfil. Todo en pares clave/valor. ----

function getConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Config');
  if (!sheet) {
    sheet = ss.insertSheet('Config');
    sheet.appendRow(['clave', 'valor']);
  }
  return sheet;
}

function leerConfig() {
  const sheet = getConfigSheet();
  const data = sheet.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < data.length; i++) {
    const clave = data[i][0];
    let valor = data[i][1];
    if (!clave) continue;
    try { valor = JSON.parse(valor); } catch (e) { /* se queda como texto/numero */ }
    config[clave] = valor;
  }
  return config;
}

function guardarConfigValor(clave, valor) {
  const sheet = getConfigSheet();
  const data = sheet.getDataRange().getValues();
  const texto = (typeof valor === 'object') ? JSON.stringify(valor) : valor;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === clave) {
      sheet.getRange(i + 1, 2).setValue(texto);
      return true;
    }
  }
  sheet.appendRow([clave, texto]);
  return true;
}
