# Sheet Formatting & Validation

Covers: data validation & dropdowns, calculation-formula injection, conditional
formatting/coloring, auto column width, and center alignment. Build one
`setupSheet()` function that does all requested parts in this order. The complete
working template is `assets/setup-sheet-template.gs` — adapt it.

## Contents
- [Order of operations](#order-of-operations)
- [Text dates → real dates](#text-dates--real-dates)
- [Formula injection (blank-safe)](#formula-injection-blank-safe)
- [Data validation & dropdowns](#data-validation--dropdowns)
- [Derived label column](#derived-label-column)
- [Conditional formatting / coloring](#conditional-formatting--coloring)
- [Auto width & center alignment](#auto-width--center-alignment)

## Order of operations

Sequence matters — later steps read what earlier steps wrote:

1. Convert text dates → real dates (so date validation accepts them).
2. Inject calculation formulas (e.g. `Amount = Qty * Price`), blank-safe.
3. Inject derived label columns (read the calculated values).
4. Apply data validation / dropdowns.
5. Apply conditional formatting.
6. Auto-resize columns + center-align + bold header.
7. `SpreadsheetApp.flush();`

## Text dates → real dates

`requireDate().setAllowInvalid(false)` marks **text** dates invalid. Convert first:

```js
function convertColumnToDates(sheet, col, numRows) {
  const range = sheet.getRange(2, col, numRows, 1);
  const values = range.getValues();
  const out = values.map(function (row) {
    const v = row[0];
    if (v instanceof Date) return [v];
    if (v === '' || v == null) return [''];
    const d = new Date(v);
    return [isNaN(d.getTime()) ? v : d];   // leave unparseable values untouched
  });
  range.setValues(out);
  range.setNumberFormat('dd-mmm-yyyy');
}
```

## Formula injection (blank-safe)

Build one formula per data row; keep empty input rows blank (see universal rule 4):

```js
const formulas = [];
for (let r = 2; r <= lastRow; r++) {
  formulas.push(['=IF(OR(K' + r + '="",L' + r + '=""),"",K' + r + '*L' + r + ')']);
}
sheet.getRange(2, 13, numRows, 1).setFormulas(formulas);   // numRows === formulas.length
```

## Data validation & dropdowns

```js
// Date
const dateRule = SpreadsheetApp.newDataValidation()
  .requireDate().setAllowInvalid(false).build();
sheet.getRange(2, 2, numRows, 1).setDataValidation(dateRule);

// Dropdown from a fixed list (true = show dropdown arrow)
const listRule = SpreadsheetApp.newDataValidation()
  .requireValueInList(['Electronics', 'Furniture'], true)
  .setAllowInvalid(false).build();
sheet.getRange(2, 10, numRows, 1).setDataValidation(listRule);

// Number constraints
const positive = SpreadsheetApp.newDataValidation()
  .requireNumberGreaterThan(0).setAllowInvalid(false).build();
sheet.getRange(2, 11, numRows, 1).setDataValidation(positive);
```

Common builders: `requireNumberBetween(a,b)`, `requireNumberGreaterThanOrEqualTo(n)`,
`requireTextIsEmail()`, `requireValueInRange(range)` (dropdown sourced from cells),
`requireCheckbox()`.

## Derived label column

Use a formula so it stays blank and auto-updates — do **not** write static text:

```js
const tagFormulas = [];
for (let r = 2; r <= lastRow; r++) {
  tagFormulas.push([
    '=IF(M' + r + '="","",' +
    'IF(M' + r + '>=50000,"High Value",' +
    'IF(M' + r + '>=10000,"Medium Value","Low Value")))'
  ]);
}
sheet.getRange(2, 14, numRows, 1).setFormulas(tagFormulas);
```

## Conditional formatting / coloring

Build an array of rules, then set them all at once (`setConditionalFormatRules`
**replaces** existing rules — read existing ones first if you must preserve them).

```js
const rules = [];

// Whole-row highlight driven by a column value. Anchor the column with $,
// leave the row unanchored so it advances down the range. Top-left row = 2.
rules.push(SpreadsheetApp.newConditionalFormatRule()
  .whenFormulaSatisfied('=$N2="High Value"')
  .setBackground('#d9ead3')
  .setRanges([sheet.getRange(2, 1, numRows, 14)])   // A2:N
  .build());

// Simple numeric thresholds
rules.push(SpreadsheetApp.newConditionalFormatRule()
  .whenNumberGreaterThan(2)
  .setBackground('#f4cccc')
  .setRanges([sheet.getRange(2, 11, numRows, 1)])    // K2:K
  .build());

// Duplicate detection (e.g. repeat customer) — bound the COUNTIF range
rules.push(SpreadsheetApp.newConditionalFormatRule()
  .whenFormulaSatisfied('=COUNTIF($D$2:$D$' + lastRow + ',$D2)>1')
  .setBackground('#cfe2f3')
  .setRanges([sheet.getRange(2, 4, numRows, 1)])     // D2:D
  .build());

sheet.setConditionalFormatRules(rules);
```

Tested pastel palette: green `#d9ead3`, yellow `#fff2cc`, red `#f4cccc`,
blue `#cfe2f3`, purple `#ead1dc`.

**Formula-rule gotcha:** the relative reference (`$N2`) is anchored to the
**top-left cell of the range**. If the range starts at row 2, reference row 2; the
rule shifts automatically for every other row.

## Auto width & center alignment

```js
const lastCol = sheet.getLastColumn();
sheet.autoResizeColumns(1, lastCol);

const full = sheet.getRange(1, 1, lastRow, lastCol);
full.setHorizontalAlignment('center');
full.setVerticalAlignment('middle');
sheet.getRange(1, 1, 1, lastCol).setFontWeight('bold');   // header
```
