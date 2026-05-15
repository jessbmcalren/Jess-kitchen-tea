/**
 * Google Apps Script for Jess's Kitchen Tea Registry.
 *
 * Setup:
 *   1. Open the Google Sheet:
 *      https://docs.google.com/spreadsheets/d/1Kp8fy4GbZfoJwgWJ52sWBUwazIVYyIRkiXcAsXlg1eo/edit
 *   2. Extensions → Apps Script.
 *   3. Replace the default code with this file's contents. Save.
 *   4. Deploy → New deployment → type: Web app.
 *      • Execute as: Me (your Google account)
 *      • Who has access: Anyone
 *   5. Copy the Web App URL it gives you.
 *   6. Paste that URL into index.html (APPS_SCRIPT_URL).
 *
 * doGet  → returns the registry as JSON (used to render the page).
 * doPost → marks an item as taken by a guest's name.
 */

const SHEET_ID = '1Kp8fy4GbZfoJwgWJ52sWBUwazIVYyIRkiXcAsXlg1eo';

// Locate the registry sheet by finding the tab whose column A contains a "BOUGHT" header.
function getRegistrySheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheets = ss.getSheets();
  for (let s = 0; s < sheets.length; s++) {
    const lastRow = sheets[s].getLastRow();
    if (lastRow === 0) continue;
    const probe = sheets[s].getRange(1, 1, Math.min(lastRow, 20), 1).getValues();
    for (let i = 0; i < probe.length; i++) {
      if (String(probe[i][0] || '').trim().toUpperCase() === 'BOUGHT') {
        return sheets[s];
      }
    }
  }
  return sheets[0];
}

function doGet() {
  try {
    const sheet = getRegistrySheet();
    const lastRow = sheet.getLastRow();
    const lastCol = Math.max(sheet.getLastColumn(), 7);
    const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const richTexts = sheet.getRange(1, 1, lastRow, lastCol).getRichTextValues();

    // Find header row (row that starts with "BOUGHT")
    let headerRow = -1;
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0]).trim().toUpperCase() === 'BOUGHT') {
        headerRow = i;
        break;
      }
    }
    if (headerRow === -1) {
      return jsonResp({ items: [], error: 'No header row found' });
    }

    // Columns: BOUGHT=0, #=1, LINK=2, ITEM=3, NOTES=4, STORE=5, PRICE=6
    const items = [];
    let section = 'items';
    for (let i = headerRow + 1; i < values.length; i++) {
      const row = values[i];
      const first = String(row[0] || '').trim();
      const upper = first.toUpperCase();

      // Section divider
      if (upper === 'VOUCHERS') { section = 'vouchers'; continue; }

      const itemName = String(row[3] || '').trim();
      if (!itemName) continue;

      // Resolve link URL from rich text (CSV export strips these)
      const linkRich = richTexts[i] && richTexts[i][2];
      let url = linkRich ? linkRich.getLinkUrl() : null;
      // Sheets occasionally puts the link on a sub-run rather than the whole cell
      if (!url && linkRich && linkRich.getRuns) {
        const runs = linkRich.getRuns();
        for (let r = 0; r < runs.length; r++) {
          const u = runs[r].getLinkUrl();
          if (u) { url = u; break; }
        }
      }

      // Taken status: anything other than the placeholder "Write taken if bought" or "Voucher" counts as taken
      const placeholder = upper === 'WRITE TAKEN IF BOUGHT' || upper === 'VOUCHER' || first === '';
      items.push({
        rowIndex: i + 1, // 1-based sheet row for write-back
        taken: !placeholder,
        takenBy: placeholder ? '' : first,
        number: row[1],
        url: url || '',
        name: itemName,
        notes: String(row[4] || '').trim(),
        store: String(row[5] || '').trim(),
        price: String(row[6] || '').trim(),
        section: section,
      });
    }

    return jsonResp({ items: items });
  } catch (err) {
    return jsonResp({ items: [], error: err.toString() });
  }
}

function doPost(e) {
  try {
    const sheet = getRegistrySheet();
    const data = JSON.parse(e.postData.contents);
    const rowIndex = parseInt(data.rowIndex, 10);
    const name = String(data.name || '').trim();
    if (!rowIndex || !name) throw new Error('Missing rowIndex or name');

    // Safety: don't overwrite if already taken (race condition between two guests)
    const current = String(sheet.getRange(rowIndex, 1).getValue() || '').trim().toUpperCase();
    if (current && current !== 'WRITE TAKEN IF BOUGHT' && current !== 'VOUCHER') {
      return jsonResp({ result: 'already_taken', takenBy: sheet.getRange(rowIndex, 1).getValue() });
    }

    sheet.getRange(rowIndex, 1).setValue(name);
    return jsonResp({ result: 'success' });
  } catch (err) {
    return jsonResp({ result: 'error', error: err.toString() });
  }
}

function jsonResp(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
