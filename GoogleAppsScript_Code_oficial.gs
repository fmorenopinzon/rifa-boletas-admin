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
    if (action === 'paidReportHtml') {
      return HtmlService.createHtmlOutput(generatePaidReportHtml()).setTitle('Reporte de boletas pagadas por asociado');
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

function isMarkedValue(value) {
  const text = String(value === null || value === undefined ? '' : value).trim();
  if (!text) return false;
  const lower = text.toLowerCase();
  if (['x', 'si', 'sí', 's', 'yes', 'y', 'true', '1'].includes(lower)) return true;
  const numeric = Number(text);
  return Number.isFinite(numeric) && numeric !== 0;
}

function getBoletasSheet() {
  const sheet = getSpreadsheet().getSheetByName(BOLETAS_SHEET);
  if (!sheet) throw new Error('No existe la hoja Boletas.');
  return sheet;
}

function generatePaidReportHtml() {
  const grouped = {};
  const sheet = getBoletasSheet();
  const lastRow = sheet.getLastRow();
  const boletaValue = getBoletaValue(sheet);

  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
    values.forEach((row, index) => {
      const associate = String(row[0] || '').trim();
      if (!associate || !isMarkedValue(row[6])) return;

      const numbers = getBoletaNumbers(row)
        .filter((n) => Number.isFinite(n) && n !== 0)
        .map((n) => String(n).padStart(3, '0'));

      if (!grouped[associate]) grouped[associate] = [];
      grouped[associate].push({
        row: index + 2,
        numbers,
        deliveryDate: row[8],
        valuePaid: row[9]
      });
    });
  }

  const associateNames = Object.keys(grouped).sort((a, b) => a.localeCompare(b, 'es'));
  const totalBoletas = associateNames.reduce((sum, associate) => sum + grouped[associate].length, 0);
  let totalRecaudo = 0;
  associateNames.forEach((associate) => {
    grouped[associate].forEach((item) => {
      const val = Number(item.valuePaid);
      if (Number.isFinite(val) && val > 0) {
        totalRecaudo += val;
      } else {
        totalRecaudo += boletaValue;
      }
    });
  });
  const generatedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  const sections = associateNames.length
    ? associateNames.map((associate) => {
        const items = grouped[associate];
        let subtotal = 0;
        items.forEach((item) => {
          const val = Number(item.valuePaid);
          if (Number.isFinite(val) && val > 0) {
            subtotal += val;
          } else {
            subtotal += boletaValue;
          }
        });
        const rows = items.map((item) => {
          const numbers = item.numbers.length ? item.numbers.join(' · ') : '—';
          let dateText = '—';
          if (item.deliveryDate) {
            try {
              dateText = Utilities.formatDate(new Date(item.deliveryDate), Session.getScriptTimeZone(), 'yyyy-MM-dd');
            } catch (e) {
              dateText = String(item.deliveryDate);
            }
          }
          const val = Number(item.valuePaid);
          const valueText = (Number.isFinite(val) && val > 0) ? val : boletaValue;
          return `
            <tr>
              <td>${htmlEscape(numbers)}</td>
              <td>${htmlEscape(dateText)}</td>
              <td class="num">${htmlEscape(String(valueText))}</td>
            </tr>
          `;
        }).join('');

        return `
          <section class="associate-block">
            <div class="associate-header">
              <h2>${htmlEscape(associate)}</h2>
              <div class="associate-meta">${items.length} boleta(s) · Subtotal: $ ${Number(subtotal).toLocaleString('es-CO')}</div>
            </div>
            <table>
              <thead>
                <tr>
                  <th>Números</th>
                  <th>Fecha de entrega</th>
                  <th>Valor</th>
                </tr>
              </thead>
              <tbody>${rows}</tbody>
              <tfoot>
                <tr>
                  <td colspan="2" class="tfoot-label">Subtotal por asociado</td>
                  <td class="tfoot-value num">$ ${Number(subtotal).toLocaleString('es-CO')}</td>
                </tr>
              </tfoot>
            </table>
          </section>
        `;
      }).join('')
    : '<p class="empty">No hay boletas marcadas como pagadas.</p>';

  return `
    <!doctype html>
    <html lang="es">
      <head>
        <meta charset="utf-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>Reporte de boletas pagadas por asociado</title>
        <style>
          :root {
            --green: #1f7c3f;
            --green-light: #edf8f1;
            --border: #dfece4;
            --text: #173425;
            --muted: #486b5c;
          }
          * { box-sizing: border-box; }
          body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            margin: 0;
            background: linear-gradient(180deg, #f7faf7 0%, #f3f8f5 100%);
            color: var(--text);
            line-height: 1.6;
          }
          header {
            background: linear-gradient(90deg, #0f3d24 0%, #1f7c3f 100%);
            color: #fff;
            padding: 28px 24px;
          }
          header .inner {
            max-width: 1100px;
            margin: 0 auto;
          }
          header h1 {
            margin: 0;
            font-size: clamp(24px, 4vw, 36px);
            letter-spacing: -0.02em;
          }
          header p {
            margin: 6px 0 0;
            opacity: 0.9;
          }
          main {
            max-width: 1100px;
            margin: 0 auto;
            padding: 24px;
          }
          .summary-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
            gap: 16px;
            margin-bottom: 26px;
          }
          .summary-card {
            background: #ffffff;
            border: 1px solid var(--border);
            border-radius: 12px;
            padding: 18px 22px;
            box-shadow: 0 10px 40px rgba(15, 61, 36, 0.06);
          }
          .summary-card .label {
            text-transform: uppercase;
            font-size: 11px;
            letter-spacing: 0.12em;
            color: var(--muted);
          }
          .summary-card .value {
            margin-top: 8px;
            font-size: clamp(26px, 4vw, 38px);
            font-weight: 600;
          }
          .associate-block {
            background: #ffffff;
            border: 1px solid var(--border);
            border-radius: 12px;
            margin-bottom: 24px;
            box-shadow: 0 12px 50px rgba(15, 61, 36, 0.05);
            overflow: hidden;
          }
          .associate-header {
            display: flex;
            flex-wrap: wrap;
            align-items: baseline;
            justify-content: space-between;
            gap: 12px;
            padding: 18px 26px 12px;
            border-bottom: 1px solid var(--border);
            background: radial-gradient(circle at top right, rgba(31, 124, 63, 0.08), transparent 55%);
          }
          .associate-header h2 {
            margin: 0;
            font-size: clamp(20px, 3vw, 26px);
          }
          .associate-meta {
            color: var(--muted);
            font-weight: 500;
          }
          table {
            width: 100%;
            border-collapse: collapse;
          }
          thead th {
            background: var(--green-light);
            color: #1a4534;
            text-align: left;
            font-size: 12px;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            padding: 12px 26px;
          }
          tbody td {
            padding: 12px 26px;
            border-bottom: 1px solid #f0f6f2;
            vertical-align: top;
          }
          tbody tr:last-child td {
            border-bottom: none;
          }
          tfoot td {
            padding: 14px 26px;
            background: #f9fcfa;
            border-top: 1px solid var(--border);
            font-weight: 600;
          }
          .num {
            text-align: right;
            font-variant-numeric: tabular-nums;
          }
          .tfoot-label {
            text-align: right;
          }
          .empty {
            text-align: center;
            padding: 40px 20px;
            color: var(--muted);
            background: #ffffff;
            border: 1px dashed var(--border);
            border-radius: 12px;
          }
          footer {
            text-align: center;
            color: var(--muted);
            padding: 18px 24px 40px;
            font-size: 13px;
          }
          @media (max-width: 720px) {
            thead { display: none; }
            table, tbody, tr, td { display: block; width: 100%; }
            tbody tr {
              border-bottom: 1px solid var(--border);
              padding: 12px 18px;
            }
            tbody tr:last-child { border-bottom: none; }
            tbody td {
              border-bottom: none;
              padding: 8px 0;
            }
            tbody td::before {
              content: attr(data-label);
              display: block;
              font-size: 11px;
              text-transform: uppercase;
              letter-spacing: 0.08em;
              color: var(--muted);
              margin-bottom: 4px;
            }
            tfoot { display: block; }
            tfoot tr { display: flex; justify-content: space-between; padding: 14px 18px; }
            tfoot td { border: none; background: transparent; padding: 0; }
            .tfoot-label { text-align: left; }
          }
        </style>
      </head>
      <body>
        <header>
          <div class="inner">
            <h1>Reporte de boletas pagadas y entregadas</h1>
            <p>Organizado por asociado · Fondo Socios Madelena · Rifa 2026</p>
          </div>
        </header>
        <main>
          <div class="summary-grid">
            <article class="summary-card">
              <div class="label">Total de asociados con boletas pagadas</div>
              <div class="value">${associateNames.length}</div>
            </article>
            <article class="summary-card">
              <div class="label">Total de boletas pagadas</div>
              <div class="value">${totalBoletas}</div>
            </article>
            <article class="summary-card">
              <div class="label">Total recaudado</div>
              <div class="value">$ ${Number(totalRecaudo).toLocaleString('es-CO')}</div>
            </article>
          </div>
          <div class="meta" style="margin-bottom: 18px; color: var(--muted);">Generado el: ${htmlEscape(generatedAt)}</div>
          ${sections}
        </main>
        <footer>
          Generado desde el panel de administración de boletas · Comité RIFA 2026
        </footer>
      </body>
    </html>
  `;
}

function htmlEscape(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\"/g, '&quot;')
    .replace(/'/g, '&#39;');
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
 
  //ensureDailyTotalRow(today);
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
  const totalCount = Number(sheet.getRange('G253').getValue() || 0);
  const totalPaid = Number(sheet.getRange('G255').getValue() || 0);

  return {
    todayCount,
    todayValue: todayCount * boletaValue,
    totalPaid,
    boletaValue,
    totalCount
  };
}

function getBoletaValue(sheet) {
  const bolValue = Number(sheet.getRange('O1').getValue());
  if (Number.isFinite(bolValue) && bolValue > 0) return bolValue;
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
  let row = 1;

  while (row <= sheet.getMaxRows()) {
    const dateValue = sheet.getRange(row, 1).getValue();
    if (dateValue === '' || dateValue === null) break;
    const todayCount = sheet.getRange(row, 2).getValue();
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