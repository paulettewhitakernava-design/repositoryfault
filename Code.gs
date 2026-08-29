function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Registro de gastos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, viewport-fit=cover');
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Movimientos');
  if (!sheet) {
    sheet = ss.insertSheet('Movimientos');
    sheet.appendRow(['id', 'monto', 'fecha', 'nota', 'tipo']);
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
      tipo: String(r[4] || '').trim().toLowerCase()
    }))
    .reverse();
}

function guardarMovimiento(mov) {
  const sheet = getSheet();
  sheet.appendRow([String(mov.id), Number(mov.monto), mov.fecha, mov.nota || '', mov.tipo]);
  return true;
}

function actualizarMovimiento(mov) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(mov.id)) {
      sheet.getRange(i + 1, 1, 1, 5).setValues([[
        String(mov.id), Number(mov.monto), mov.fecha, mov.nota || '', mov.tipo
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
