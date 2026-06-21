/**
 * Sheet setup template — adapt the column map, validation, and formulas to the
 * target header, then delete any section the user did not request.
 *
 * Column map (EDIT to match the real header; index is 1-based, A=1):
 *   A .. | B Date | ... | K Qty | L UnitPrice | M Amount | N StatusTag
 */
function setupSheet() {
  const SHEET_NAME = 'Sheet1';                       // EDIT: tab name
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" not found.');

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {                                 // universal rule 1: guard empty
    ss.toast('No data rows found below the header.');
    return;
  }
  const numRows = lastRow - 1;
  const lastCol = sheet.getLastColumn();

  // 1) TEXT DATES -> REAL DATES (run before date validation). EDIT the columns.
  convertColumnToDates_(sheet, 2, numRows);          // B
  // convertColumnToDates_(sheet, 8, numRows);        // H (example second date col)

  // 2) INJECT CALC FORMULA, blank-safe (EDIT columns/expression). Amount = Qty*Price
  setColumnFormulas_(sheet, 13, lastRow, function (r) {
    return '=IF(OR(K' + r + '="",L' + r + '=""),"",K' + r + '*L' + r + ')';
  });

  // 3) DERIVED LABEL COLUMN (formula, stays blank when source blank). EDIT thresholds.
  sheet.getRange('N1').setValue('StatusTag');
  setColumnFormulas_(sheet, 14, lastRow, function (r) {
    return '=IF(M' + r + '="","",IF(M' + r + '>=50000,"High Value",' +
           'IF(M' + r + '>=10000,"Medium Value","Low Value")))';
  });

  // 4) DATA VALIDATION & DROPDOWNS (EDIT columns/lists)
  const dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();
  sheet.getRange(2, 2, numRows, 1).setDataValidation(dateRule);                 // B

  const listRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Electronics', 'Furniture'], true).setAllowInvalid(false).build();
  sheet.getRange(2, 10, numRows, 1).setDataValidation(listRule);                // J

  const positiveRule = SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThan(0).setAllowInvalid(false).build();
  sheet.getRange(2, 11, numRows, 1).setDataValidation(positiveRule);            // K
  sheet.getRange(2, 12, numRows, 1).setDataValidation(positiveRule);            // L

  // 5) CONDITIONAL FORMATTING (EDIT rules/colors). Replaces existing rules.
  const rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$N2="High Value"').setBackground('#d9ead3')
    .setRanges([sheet.getRange(2, 1, numRows, lastCol)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(2).setBackground('#f4cccc')
    .setRanges([sheet.getRange(2, 11, numRows, 1)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=COUNTIF($D$2:$D$' + lastRow + ',$D2)>1').setBackground('#cfe2f3')
    .setRanges([sheet.getRange(2, 4, numRows, 1)]).build());
  sheet.setConditionalFormatRules(rules);

  // 6) AUTO WIDTH + CENTER + BOLD HEADER
  sheet.autoResizeColumns(1, lastCol);
  const full = sheet.getRange(1, 1, lastRow, lastCol);
  full.setHorizontalAlignment('center');
  full.setVerticalAlignment('middle');
  sheet.getRange(1, 1, 1, lastCol).setFontWeight('bold');

  SpreadsheetApp.flush();
}

/** Writes one formula per data row (row 2..lastRow) into a column. */
function setColumnFormulas_(sheet, col, lastRow, builder) {
  const out = [];
  for (let r = 2; r <= lastRow; r++) out.push([builder(r)]);
  sheet.getRange(2, col, out.length, 1).setFormulas(out);
}

/** Converts a column of text dates to real Date values so date validation passes. */
function convertColumnToDates_(sheet, col, numRows) {
  const range = sheet.getRange(2, col, numRows, 1);
  const out = range.getValues().map(function (row) {
    const v = row[0];
    if (v instanceof Date) return [v];
    if (v === '' || v == null) return [''];
    const d = new Date(v);
    return [isNaN(d.getTime()) ? v : d];
  });
  range.setValues(out);
  range.setNumberFormat('dd-mmm-yyyy');
}
