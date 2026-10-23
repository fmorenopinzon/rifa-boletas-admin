// Google Apps Script backend for Administrador_Boletas_GoogleSheets.html
// Paste this code into a Google Apps Script project bound to the Google Sheet.
// Deploy it as a Web App:
//   Execute as: Me
//   Who has access: Anyone

const BOLETAS_SHEET = 'Boletas';
const NUMBERS_SHEET = 'Números_000-999';
const EXCLUDED_SHEETS = new Set([BOLETAS_SHEET, NUMBERS_SHEET, 'FONDO']);

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  try {
    if (action === 'associates') {
      return json({ associates: getAssociates() });
    }
    if (action === 'search') {
      const number = Number(e.parameter.number);
      if (!Number.isFinite(number)) {
        return json({ ok: false, error: 'Número inválido.' });
      }
      return json({ ok: true, result: searchBoleta(number) });
    }
    return json({ ok: false, error: 'Acción no reconocida.' });
  } catch (error) {
    return json({ ok: false, error: String(error && error.message ? error.message : error) });
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || '{}');
    return json(updateBoleta(payload));
  } catch (error) {
    return json({ ok: false, error: String(error && error.message ? error.message : error) });
  }
}

function json(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
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

function searchBoleta(number) {
  const sheet = getBoletasSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const values = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const numbers = [row[1], row[2], row[3], row[4]].map((value) => Number(value));
    if (numbers.includes(number)) {
      const valueCell = sheet.getRange(1, 14).getValue();
      return {
        row: i + 2,
        associate: String(row[0] || '').trim(),
        numbers: numbers.map((n) => String(n).padStart(3, '0')),
        sold: row[6],
        paid: row[7],
        buyer: row[8],
        cedula: row[9],
        phone: row[10],
        value: valueCell || 50000
      };
    }
  }
  return null;
}

function updateBoleta(payload) {
  if (!payload || !payload.row) {
    return { ok: false, error: 'Falta la fila de la boleta.' };
  }

  const sheet = getBoletasSheet();
  const row = Number(payload.row);
  const associate = String(sheet.getRange(row, 1).getValue() || '').trim();

  if (payload.selectedAssociate && associate !== payload.selectedAssociate) {
    return { ok: false, error: 'La boleta no pertenece al asociado seleccionado.' };
  }

  const hasBuyerData = Boolean(payload.buyer || payload.cedula || payload.phone);
  const soldValue = hasBuyerData ? 'X' : (payload.sold ? 'X' : '');
  const paidValue = payload.paid ? 'X' : '';

  sheet.getRange(row, 7).setValue(soldValue);   // G: Vendida
  sheet.getRange(row, 8).setValue(paidValue);   // H: Pagada
  sheet.getRange(row, 9).setValue(payload.buyer || '');   // I: Nombre Comprador
  sheet.getRange(row, 10).setValue(payload.cedula || '');  // J: Cédula
  sheet.getRange(row, 11).setValue(payload.phone || '');   // K: Teléfono

  updateAssociateSheet(associate, payload.numbers || [], soldValue, paidValue, payload.buyer || '');
  return { ok: true };
}

function updateAssociateSheet(associate, numbers, soldValue, paidValue, buyer) {
  if (!associate) return;
  const spreadsheet = getSpreadsheet();
  const sheet = spreadsheet.getSheetByName(associate);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const values = sheet.getRange(2, 2, lastRow - 1, 4).getValues();
  for (let i = 0; i < values.length; i++) {
    const rowNumbers = values[i].map((value) => Number(value));
    const requested = (numbers || []).map((value) => Number(value));
    if (rowNumbers.length === 4 && rowNumbers.every((value, index) => value === requested[index])) {
      const targetRow = i + 2;
      sheet.getRange(targetRow, 7).setValue(soldValue); // G: Vendida
      sheet.getRange(targetRow, 8).setValue(paidValue); // H: Pagada
      sheet.getRange(targetRow, 9).setValue(buyer);     // I: Nombre Comprador
    }
  }
}