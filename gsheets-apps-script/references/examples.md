# Worked Examples

Request → correct output, with the common wrong turns called out. Use these as
adherence anchors: match the *approach*, not just the syntax.

## Contents
- [Example 1 — Validation + auto-ID (formula injection)](#example-1--validation--auto-id-formula-injection)
- [Example 2 — Full sheet setup](#example-2--full-sheet-setup)
- [Example 3 — Answer questions in a sheet](#example-3--answer-questions-in-a-sheet)

## Example 1 — Validation + auto-ID (formula injection)

**Request:** "Provide data validation and auto-ID generation. Sheet name is
`Sheet2`."

```
CustomerID  CustomerName  Membership
C001        Rahul         Gold
C002        Priya         Silver
...
```

**Layout:** `A CustomerID | B CustomerName | C Membership`

**Correct approach:**
- Auto-ID = **injected formula** in column A (universal rule 8), not an `onEdit`
  trigger.
- Membership = `requireValueInList(['Gold','Silver','Bronze'], true)`.
- Name length = `requireFormulaSatisfied('=LEN(B2)<=25')` — there is **no**
  `requireTextLength*` method (universal rule 9).

```js
function setupSheet2() {
  const SHEET_NAME = 'Sheet2';
  const ROWS = 1000;                 // buffer so new rows auto-fill
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" not found.');

  // Auto-ID: injected, blank-safe, gap-safe.
  const idFormulas = [];
  for (let r = 2; r <= ROWS + 1; r++) {
    idFormulas.push(['=IF(B' + r + '="","","C"&TEXT(COUNTA($B$2:B' + r + '),"000"))']);
  }
  sheet.getRange(2, 1, idFormulas.length, 1).setFormulas(idFormulas);

  const membershipRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Gold', 'Silver', 'Bronze'], true)
    .setAllowInvalid(false).build();
  sheet.getRange(2, 3, ROWS, 1).setDataValidation(membershipRule);

  const nameRule = SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied('=LEN(B2)<=25')
    .setAllowInvalid(false).build();
  sheet.getRange(2, 2, ROWS, 1).setDataValidation(nameRule);

  sheet.getRange('A1:C1').setFontWeight('bold').setBackground('#d9ead3')
    .setHorizontalAlignment('center');
  sheet.autoResizeColumns(1, 3);
  SpreadsheetApp.flush();
}
```

Result: rows with names show `C001, C002, …`; a new name in the next empty row
auto-gets the next ID; empty rows stay blank.

**❌ Wrong turns to avoid (all seen in the wild):**

| Wrong | Why it fails | Right |
|---|---|---|
| `.requireTextLengthLessThanOrEqualTo(25)` | Method doesn't exist → *"is not a function"* | `.requireFormulaSatisfied('=LEN(B2)<=25')` |
| `.requireTextStartsWith('C')` | Method doesn't exist (fails silently if dropped) | `.requireFormulaSatisfied('=LEFT(A2,1)="C"')` |
| Auto-ID via `onEdit(e)` trigger | Fragile, per-keystroke, needs guards/permissions, not what the skill teaches | Inject the `=IF(B2="",...)` ID formula |
| Plain `="C"&ROW()` with no blank check | Fills every empty row with `C1`, `C2`… | Wrap in `IF(B2="","",...)` |

## Example 2 — Full sheet setup

**Request:** "Add validation, calculate the order amount, tag high/medium/low,
color the rows, auto-fit and center."

Use `assets/setup-sheet-template.gs`. Key points: convert text dates first, inject
`Amount` as `=IF(OR(K2="",L2=""),"",K2*L2)`, make the status tag a **formula** (not
written text), build conditional-format rules in one array, then auto-resize +
center. Full detail in `sheet-formatting.md`.

## Example 3 — Answer questions in a sheet

**Request:** "There are questions in the CustomerMaster sheet — fill the answers."

Questions sit in cells (e.g. `H11 "Count orders from Kochi"`); answers go in the
**adjacent** cell as **live formulas**. Scan for question text, map each to a
formula (`COUNTIF`/`SUMIF`/`VLOOKUP` across `SalesData!`/`CustomerMaster!`), write
to `(r, c+1)`. Use `assets/inject-answers-template.gs`; detail in `formulas.md`.
