# Formulas: Lookups, Aggregation & Question-Answering

Two related tasks: (a) inject a formula to compute a value, and (b) answer
free-text questions placed in sheet cells by writing the answer formula into the
adjacent cell. Template: `assets/inject-answers-template.gs`.

## Contents
- [Formula cookbook](#formula-cookbook)
- [Cross-sheet references](#cross-sheet-references)
- [Answering questions written in cells](#answering-questions-written-in-cells)

## Formula cookbook

| Need | Formula |
|---|---|
| Count rows matching a value | `=COUNTIF(D:D,"Kochi")` |
| Count with multiple criteria | `=COUNTIFS(D:D,"Kochi",F:F,"Electronics")` |
| Sum a column where criteria met | `=SUMIF(D:D,"Kochi",H:H)` |
| Sum with multiple criteria | `=SUMIFS(H:H,D:D,"Kochi",F:F,"Electronics")` |
| Average where criteria met | `=AVERAGEIF(D:D,"Kochi",H:H)` |
| Look up a value by key | `=VLOOKUP("C003",A:C,3,FALSE)` |
| Robust lookup (key not leftmost / left of result) | `=INDEX(C:C,MATCH("C003",A:A,0))` |
| Conditional value | `=IF(M2>=50000,"High","Low")` |
| Unique list | `=UNIQUE(D2:D)` |
| Running/grouped totals | `=QUERY(A:H,"select D, sum(H) group by D",1)` |

Always pass `FALSE` (exact match) as the 4th `VLOOKUP` argument unless an
approximate match is genuinely intended.

## Cross-sheet references

Prefix the range with the **exact tab name** and `!`:

```
=SUMIF(SalesData!D:D,"Kochi",SalesData!H:H)
=VLOOKUP("C003",CustomerMaster!A:C,3,FALSE)
```

If the tab name contains spaces, single-quote it: `='Sales Data'!D:D`.

## Answering questions written in cells

Pattern observed: questions sit in cells (e.g. `H11 "Count orders from Kochi"`,
`E18 "membership of customer id C003"`) and the answer belongs in the **adjacent
cell**. Scan for the question text rather than hardcoding addresses, so it survives
row shifts and you handle source typos (e.g. `memership`).

```js
function injectAnswers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('CustomerMaster');
  if (!sheet) throw new Error('Sheet "CustomerMaster" not found.');

  // Normalized question text -> formula written to the cell on its RIGHT.
  const answerMap = {
    'count orders from kochi':           '=COUNTIF(SalesData!D:D,"Kochi")',
    'total sales from kochi':            '=SUMIF(SalesData!D:D,"Kochi",SalesData!H:H)',
    'membership of customer id c003':    '=VLOOKUP("C003",CustomerMaster!A:C,3,FALSE)',
    'memership of customer id c003':     '=VLOOKUP("C003",CustomerMaster!A:C,3,FALSE)'
  };

  const values = sheet.getDataRange().getValues();
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const key = String(values[r][c]).trim().toLowerCase();
      if (answerMap.hasOwnProperty(key)) {
        sheet.getRange(r + 1, c + 2).setFormula(answerMap[key]);  // cell to the right
      }
    }
  }
  SpreadsheetApp.flush();
}
```

**Building the `answerMap`:** read each question, translate it to the right
formula using the cookbook, and key it by the question text **lowercased and
trimmed**. To place answers below or above instead of to the right, change the
target to `(r + 2, c + 1)` / `(r, c + 1)`. The answers are live formulas — they
recompute when the source data changes.
