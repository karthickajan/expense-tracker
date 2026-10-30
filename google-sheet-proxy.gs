/**
 * Google Apps Script proxy for the Mess Expense Splitter.
 * 1) Open the Google Sheet you want to use.
 * 2) Extensions -> Apps Script
 * 3) Paste this file into the script editor.
 * 4) Deploy -> New deployment -> Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5) Copy the Web App URL and paste it into the app Settings panel.
 * 6) Optionally set a passphrase in the script below to protect the endpoint.
 */

const PASS = 'mess-expense';
const SS_ID = '1prXNK5p8C21QpnZGd1gKLDyLlGVD7pPc';

function normalizedSheetName(name) {
  return String(name || '').trim();
}

function setTokenFromRequest(req) {
  return req && req.parameter && req.parameter.token ? String(req.parameter.token) : '';
}

function verifyToken(token) {
  return token === PASS || token === '';
}

function exportWorkbookToJson() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const settings = ss.getSheetByName('Settings');
  const result = { settings: { fee: 0, lw: 1, dw: 1, friends: [] }, months: {} };

  if (!settings) {
    return result;
  }

  const sValues = settings.getDataRange().getValues();
  if (sValues[0] && sValues[0][1]) result.settings.fee = Number(sValues[0][1]) || 0;
  if (sValues[0] && sValues[0][3]) result.settings.lw = Number(sValues[0][3]) || 1;
  if (sValues[0] && sValues[0][5]) result.settings.dw = Number(sValues[0][5]) || 1;

  result.settings.friends = [];
  for (let i = 3; i < sValues.length; i++) {
    const row = sValues[i];
    const name = row[0];
    if (!name || !String(name).trim()) continue;
    result.settings.friends.push({ name: String(name).trim(), active: Number(row[1] || 0) !== 0 });
  }

  const monthSheets = ss.getSheets().filter((sheet) => /^\d{4}-\d{2}$/.test(sheet.getName()));
  for (const sheet of monthSheets) {
    const name = sheet.getName();
    const rows = sheet.getDataRange().getValues();
    if (!rows.length) continue;
    const month = {
      fee: Number(rows[0][1]) || 0,
      lw: Number(rows[0][3]) || 1,
      dw: Number(rows[0][5]) || 1,
      friends: [],
      rows: []
    };

    const header = rows[1] || [];
    for (let c = 2; c < header.length && header[c] !== 'Notes'; c += 2) {
      if (header[c]) month.friends.push(String(header[c]).trim());
    }

    for (let r = 3; r < rows.length; r++) {
      const row = rows[r];
      const date = row[0];
      if (!date || !String(date).trim()) continue;
      const grocery = Number(row[1]) || 0;
      const v = [];
      for (let i = 0; i < month.friends.length; i++) {
        const lunch = Number(row[2 + (i * 2)]) || 0;
        const dinner = Number(row[3 + (i * 2)]) || 0;
        v.push([lunch, dinner]);
      }
      const note = row[2 + (month.friends.length * 2)] || '';

      month.rows.push({
        d: String(date),
        g: grocery,
        v,
        note: String(note)
      });
    }

    result.months[name] = month;
  }

  return result;
}

function writeWorkbookFromJson(payload) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const obj = payload || {};
  const sheets = obj.sheets || {};

  if (sheets.Settings) {
    const settings = ss.getSheetByName('Settings') || ss.insertSheet('Settings');
    settings.clear();
    settings.getRange(1, 1, sheets.Settings.length, sheets.Settings[0].length).setValues(sheets.Settings);
  }

  Object.keys(sheets).forEach((name) => {
    if (name === 'Settings') return;
    const values = sheets[name];
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    sheet.clear();
    if (values && values.length) {
      sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
    }
  });

  return { ok: true };
}

function doGet(e) {
  const token = setTokenFromRequest(e);
  if (!verifyToken(token)) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'invalid token' })).setMimeType(ContentService.MimeType.JSON);
  }

  const payload = exportWorkbookToJson();
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
  let payload = {};
  try {
    payload = JSON.parse(raw);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'invalid json' })).setMimeType(ContentService.MimeType.JSON);
  }

  const token = payload.token || '';
  if (!verifyToken(token)) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'invalid token' })).setMimeType(ContentService.MimeType.JSON);
  }

  const sheetPayload = payload.sheets || {};
  const result = writeWorkbookFromJson({ sheets: sheetPayload });
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}
