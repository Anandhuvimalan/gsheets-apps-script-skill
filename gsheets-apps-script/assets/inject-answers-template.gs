/**
 * Answer-injection template — finds question text in a sheet and writes the
 * answer formula into the adjacent cell. Survives row shifts and source typos.
 *
 * To adapt:
 *  - set TARGET_SHEET to the tab holding the questions
 *  - for each question, add an entry: normalized question text -> answer formula
 *    (lowercase + trim the key; build the formula from references/formulas.md)
 *  - change the answer cell offset if answers should not go to the RIGHT
 */
function injectAnswers() {
  const TARGET_SHEET = 'CustomerMaster';             // EDIT

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TARGET_SHEET);
  if (!sheet) throw new Error('Sheet "' + TARGET_SHEET + '" not found.');

  // EDIT: normalized question -> formula written to the cell on its right.
  const answerMap = {
    'count orders from kochi':        '=COUNTIF(SalesData!D:D,"Kochi")',
    'total sales from kochi':         '=SUMIF(SalesData!D:D,"Kochi",SalesData!H:H)',
    'membership of customer id c003': '=VLOOKUP("C003",CustomerMaster!A:C,3,FALSE)',
    'memership of customer id c003':  '=VLOOKUP("C003",CustomerMaster!A:C,3,FALSE)'  // typo-safe
  };

  const values = sheet.getDataRange().getValues();
  let placed = 0;
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const key = String(values[r][c]).trim().toLowerCase();
      if (answerMap.hasOwnProperty(key)) {
        sheet.getRange(r + 1, c + 2).setFormula(answerMap[key]);   // right of question
        placed++;
      }
    }
  }

  SpreadsheetApp.flush();
  ss.toast(placed + ' answer formula(s) injected.', 'Done', 4);
}
