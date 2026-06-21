/**
 * Employee Directory — Web App backend
 *
 * Sheet columns:  A Employee ID | B Employee Name | C Department | D Region
 *
 * Deploy:  Apps Script editor → Deploy → New deployment → type "Web app"
 *          Execute as: Me   |   Who has access: (your choice)
 */

const SHEET_NAME = 'Sheet1';          // change if your data tab has another name
const DEFAULT_DEPARTMENTS = ['Sales', 'Marketing', 'Finance', 'HR', 'Operations'];
const DEFAULT_REGIONS = ['North', 'South', 'East', 'West'];

/** Serves the web app page. */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Form')
    .setTitle('Employee Directory')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Returns the data sheet (falls back to the first sheet if the name is missing). */
function getDataSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
}

/** Suggests the next Employee ID by incrementing the highest numeric suffix found. */
function suggestNextId_(sheet) {
  const lastRow = sheet.getLastRow();
  let max = 100;
  if (lastRow >= 2) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      const m = String(ids[i][0]).match(/(\d+)/);
      if (m) max = Math.max(max, parseInt(m[1], 10));
    }
  }
  return 'EMP' + (max + 1);
}

/** Data sent to the page on load: dropdown options, employee count, next ID. */
function getInitialData() {
  const sheet = getDataSheet_();
  const lastRow = sheet.getLastRow();

  const departments = new Set(DEFAULT_DEPARTMENTS);
  const regions = new Set(DEFAULT_REGIONS);
  let count = 0;

  if (lastRow >= 2) {
    const rows = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    count = rows.length;
    rows.forEach(function (r) {
      if (r[2]) departments.add(String(r[2]).trim());
      if (r[3]) regions.add(String(r[3]).trim());
    });
  }

  return {
    departments: Array.from(departments),
    regions: Array.from(regions),
    count: count,
    nextId: suggestNextId_(sheet)
  };
}

/**
 * Validates and appends one employee row.
 * Throws on missing fields or a duplicate Employee ID (caught by the UI).
 */
function addEmployee(data) {
  const sheet = getDataSheet_();

  const id = (data.id || '').toString().trim();
  const name = (data.name || '').toString().trim();
  const department = (data.department || '').toString().trim();
  const region = (data.region || '').toString().trim();

  if (!id || !name || !department || !region) {
    throw new Error('All fields are required.');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim().toLowerCase() === id.toLowerCase()) {
        throw new Error('Employee ID "' + id + '" already exists.');
      }
    }
  }

  sheet.appendRow([id, name, department, region]);

  return {
    message: name + ' added successfully',
    count: sheet.getLastRow() - 1,
    nextId: suggestNextId_(sheet)
  };
}
