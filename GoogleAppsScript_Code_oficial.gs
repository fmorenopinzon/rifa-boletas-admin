// Google Apps Script backend for admin-boletas-comite.html
// Deploy as Web App:
//   Execute as: Me
//   Who has access: Anyone

const BOLETAS_SHEET = 'Boletas';
const NUMBERS_SHEET = 'Números_000-999';
const REG_X_FECHAS_SHEET = 'REGISTRO_X_FECHAS';

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  try {
    if (action === 'associates') {
      return json({ associates: getAssociates() });
    }
    if (action === 'search') {
      const number = Number(e.parameter.number);
      if (!Number.isFinite(number)) return json({ ok: false, error: 'Número inválido.' });
      return json({ ok: true, result: searchBoleta(number) });
    }
    if (action === 'associateBoletas') {
      const associate = String(e.parameter.associate || '').trim();
      return json({ ok: true, boletas: getAssociateBoletas(associate) });
    }
    if (action === 'totals') {
      return json(getTotals());
    }
    if (action === 'dailyTotals') {
      return json({ dailyTotals: getDailyTotals() });
    }
    return json({ ok: false, error: 'Acción no reconocida.' });
  } catch (error) {
    return json({ ok: false, error: String(error && error.message ? error.message : error) });
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || '{}');
    if (payload.action === 'update') {
      return json(markBoleta(payload.row));
    }
    if (payload.action === 'updateBulk') {
      const rows = Array.isArray(payload.rows) ? payload.rows : [];
      return json(markBoletas(rows));
    }
    return json({ ok: false, error: 'Acción no reconocida.' });
  } catch (error) {
    return json({ ok: false, error: String(error && error.message ? error.message : error) });
  }
}

function json(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getBoletasSheet() {
  const sheet = getSpreadsheet().getSheetByName(BOLETAS_SHEET);
  if (!sheet) throw new Error('No existe la hoja Boletas.');
  return sheet;
}

function getRegistryByDateSheet() {
  const sheetByDate = getSpreadsheet().getSheetByName(REG_X_FECHAS_SHEET);
  if (!sheetByDate) throw new Error('No existe la hoja de Registro Por Fecha.');
  return sheetByDate;
}


function getAssociates() {
  const sheet = getBoletasSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const names = new Set();
  values.forEach((row) => {
    const name = String(row[0] || '').trim();
    if (name) names.add(name);
  });
  return Array.from(names).sort((a, b) => a.localeCompare(b, 'es'));
}

function getBoletaNumbers(row) {
  return [row[1], row[2], row[3], row[4]].map((value) => Number(value));
}

function searchBoleta(number) {
  const sheet = getBoletasSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const values = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const numbers = getBoletaNumbers(row);
    if (numbers.includes(number)) {
      return {
        row: i + 2,
        associate: String(row[0] || '').trim(),
        numbers: numbers.map((n) => String(n).padStart(3, '0')),
        sold: row[6],
        deliveryDate: row[8],
        valuePaid: row[9],
        boletaValue: getBoletaValue(sheet)
      };
    }
  }
  return null;
}

function getAssociateBoletas(associate) {
  const sheet = getBoletasSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
  const result = [];
  values.forEach((row, index) => {
    const name = String(row[0] || '').trim();
    if (name !== associate) return;
    result.push({
      row: index + 2,
      associate: name,
      numbers: getBoletaNumbers(row).map((n) => String(n).padStart(3, '0')),
      sold: row[6],
      deliveryDate: row[8],
      valuePaid: row[9]
    });
  });
  return result;
}

function markBoleta(rowNumber) {
  const sheet = getBoletasSheet();
  const row = Number(rowNumber);
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (!Number.isInteger(row) || row < 2 || row > sheet.getLastRow()) {
    return { ok: false, error: 'Fila inválida.' };
  }

  const alreadyMarked = String(sheet.getRange(row, 7).getValue() || '').trim() !== '';
  if (!alreadyMarked) {
    sheet.getRange(row, 7).setValue('x');      // G: Vendida/Pagada
    sheet.getRange(row, 9).setValue(today);   // I: Fecha Entrega (Actual)
  }

  //const sheetByDate = getRegistryByDateSheet();
  //ensureDailyTotalRow(sheetByDate.getRange(row, 2).getValue() || today);
 
  ensureDailyTotalRow(today);
  return { ok: true, marked: true, alreadyMarked };
}

function markBoletas(rows) {
  const sheet = getBoletasSheet();
  const lastRow = sheet.getLastRow();
  const markedRows = [];
  const skippedRows = [];
  const today = new Date();

  rows.forEach((rowNumber) => {
    const row = Number(rowNumber);
    if (!Number.isInteger(row) || row < 2 || row > lastRow) {
      skippedRows.push(rowNumber);
      return;
    }
    const alreadyMarked = String(sheet.getRange(row, 7).getValue() || '').trim() !== '';
    if (alreadyMarked) {
      skippedRows.push(row);
      return;
    }
    sheet.getRange(row, 7).setValue('x');
    sheet.getRange(row, 9).setValue(today);
    markedRows.push(row);
  });

  if (markedRows.length > 0) {
    ensureDailyTotalRow(today);
  }
  return { ok: true, marked: markedRows.length, markedRows, skippedRows };
}

function getTotals() {
  const sheet = getBoletasSheet();
  const lastRow = sheet.getLastRow();
  const todayKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  let todayCount = 0;

  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
    values.forEach((row) => {
      const sold = String(row[6] || '').trim();
      const delivery = row[8];
      if (sold && isSameDate(delivery, todayKey)) todayCount += 1;
    });
  }

  const boletaValue = getBoletaValue(sheet);
  const totalPaid = Number(sheet.getRange('G255').getValue() || 0);

  return {
    todayCount,
    todayValue: todayCount * boletaValue,
    totalPaid,
    boletaValue
  };
}

function getBoletaValue(sheet) {
  const n1 = Number(sheet.getRange('N1').getValue());
  if (Number.isFinite(n1) && n1 > 0) return n1;
  const o1 = Number(sheet.getRange('O1').getValue());
  if (Number.isFinite(o1) && o1 > 0) return o1;
  return 50000;
}

function ensureDailyTotalRow(date) {
  const sheetByDate = getRegistryByDateSheet();
  const startRow = 2;
  let row = startRow;
 
  while (row <= sheetByDate.getMaxRows()) {
    const currentValue = sheetByDate.getRange(row, 1).getValue();
    if (currentValue === '' || currentValue === null) break;
    if (isSameDate(currentValue, Utilities.formatDate(date instanceof Date ? date : new Date(date), Session.getScriptTimeZone(), 'yyyy-MM-dd'))) {
      sheetByDate.getRange(row, 1).setNumberFormat('yyyy-MM-dd');
      sheetByDate.getRange(row, 2).setFormula('=COUNTIF(Boletas!$I$2:$I$252,B' + row + ')');
      sheetByDate.getRange(row, 3).setFormula('=B' + row + ' * Boletas!$O$1');
      return row;
    }
    row += 1;
  }

  if (row > sheetByDate.getMaxRows()) {
    sheetByDate.insertRowsAfter(sheet.getMaxRows(), 100);
  }

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  sheetByDate.getRange(row, 1).setValue(today);
  sheetByDate.getRange(row, 1).setNumberFormat('yyyy-MM-dd');
  sheetByDate.getRange(row, 2).setFormula('=COUNTIF(Boletas!$I$2:$I$252,B' + row + ')');
  sheetByDate.getRange(row, 3).setFormula('=B' + row + ' * Boletas!$O$1');

  return row;
}

function getDailyTotals() {
  const sheet = getRegistryByDateSheet();
  const result = [];
  let row = 260;

  while (row <= sheet.getMaxRows()) {
    const dateValue = sheet.getRange(row, 1).getValue();
    if (dateValue === '' || dateValue === null) break;
    const totalValue = sheet.getRange(row, 3).getValue();
    result.push({
      date: dateValue,
      total: (typeof totalValue === 'number' && Number.isFinite(totalValue)) ? totalValue : Number(totalValue) || 0
    });
    row += 1;
  }

  return result;
}

function isSameDate(value, targetKey) {
  if (!value) return false;
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) return false;
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd') === targetKey;
}