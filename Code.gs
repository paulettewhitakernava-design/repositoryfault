function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Registro de gastos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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

function obtenerMovimientos() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1); // sin encabezado
  return rows
    .filter(r => r[0] !== '')
    .map(r => ({
      id: r[0],
      monto: r[1],
      fecha: Utilities.formatDate(new Date(r[2]), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      nota: r[3],
      tipo: r[4]
    }))
    .reverse();
}

function guardarMovimiento(mov) {
  const sheet = getSheet();
  sheet.appendRow([mov.id, mov.monto, mov.fecha, mov.nota, mov.tipo]);
  return true;
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
