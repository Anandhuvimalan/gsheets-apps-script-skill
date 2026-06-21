---
name: gsheets-apps-script
description: >-
  Build error-free Google Apps Script (.gs) and HTML Service solutions for Google
  Sheets from a user's data. Use when a user provides spreadsheet data (CSV/sheet)
  and wants to: (1) clean/format a sheet — data validation & dropdowns, inject
  calculation formulas, conditional formatting/coloring, auto column width,
  center alignment; (2) build a designed data-entry web app / form that writes
  submissions back into the sheet; or (3) answer questions or compute values by
  injecting cross-sheet lookup/aggregation formulas (VLOOKUP, COUNTIF, SUMIF,
  INDEX/MATCH). Triggers include "Google Sheet", "Apps Script", ".gs script",
  "data validation", "inject formula", "conditional formatting", "web form for my
  sheet", "answer the questions in the sheet", or any task that maps a column
  layout to Apps Script automation.
---

# Google Sheets Apps Script Builder

Generate reliable Apps Script for Google Sheets that runs the first time. The
recurring failures in this domain are mechanical (empty-range errors, text vs.
real dates, formulas filling blank rows with `0`, deploy/naming mismatches), so
the rules below exist to avoid them — follow them on every task.

## Step 1 — Establish the layout before writing code

1. **Identify the sheet name.** When data comes from a downloaded Google Sheet CSV,
   the file is named `<workbook> - <SheetName>.csv`. The sheet/tab name is the text
   **after the last ` - ` and before `.csv`**.
   - `customer - SalesData.csv` → sheet `SalesData`
   - `ddd - Sheet1.csv` → sheet `Sheet1`
   If multiple CSVs share a workbook prefix, each is a separate tab in the **same**
   spreadsheet — formulas can reference across them by name (`SalesData!A:A`).
2. **Read the header row and map columns to letters/indexes.** Write the map in a
   comment (`A OrderID | B OrderDate | ...`). Column index is 1-based: A=1, B=2, …
   Confirm which column each validation/formula/format applies to from the header;
   never assume.
3. **Confirm the goal**, then pick the workflow:

| User wants… | Read this reference |
|---|---|
| Validation, dropdowns, calc formulas, conditional formatting, coloring, auto-width, centering | `references/sheet-formatting.md` |
| A designed form / web app that writes entries into the sheet | `references/webapp-form.md` |
| Answers to questions in the sheet, or computed values via formulas (lookup/aggregate) | `references/formulas.md` |

When anything breaks, consult `references/troubleshooting.md` — it catalogs the
exact Apps Script errors seen in practice and their fixes.

## Universal rules (apply to every workflow)

These prevent the most common runtime errors. Do not skip them.

1. **Guard empty sheets.** `getLastRow()` returns 1 when only a header exists.
   Building a range with `lastRow - 1` rows then throws
   *"The number of rows in the range must be at least 1."* Always:
   ```js
   const lastRow = sheet.getLastRow();
   if (lastRow < 2) return;            // nothing to process
   const numRows = lastRow - 1;        // data rows, excluding header
   ```
2. **Prefer numeric ranges for dynamic sizes.** Use
   `sheet.getRange(2, col, numRows, 1)` over A1 strings like `"B2:B" + lastRow`.
   It is harder to get wrong and matches `setFormulas` array dimensions exactly.
3. **`setFormulas` (plural) needs a 2D array matching the range.** Mismatched
   dimensions throw. Build `[["=..."],["=..."]]` with exactly `numRows` entries.
   Use `setFormula` (singular) only for a single cell.
4. **Leave blank input rows blank.** A plain `=K2*L2` shows `0` on empty rows and
   any label column then mislabels them. Always wrap:
   ```js
   '=IF(OR(K' + r + '="",L' + r + '=""),"",K' + r + '*L' + r + ')'
   ```
   Derived label columns must guard too: `=IF(M2="","",IF(M2>=50000,"High",...))`.
   Prefer **formula-driven label columns over static written values** so they stay
   blank and self-update.
5. **Convert text dates before date validation.** CSV dates like `05-Jan-2025` are
   text; `requireDate().setAllowInvalid(false)` flags every such cell invalid.
   Parse to real `Date` (`new Date(str)`), write back, then `setNumberFormat`. See
   `references/sheet-formatting.md`.
6. **End mutations with `SpreadsheetApp.flush();`**
7. **Reference sheets by exact name** in cross-sheet formulas; the tab must exist
   in the same spreadsheet.
8. **Prefer injected formulas over `onEdit` triggers** for any derived, sequential,
   or auto-generated value (auto-IDs, totals, labels, statuses). A formula
   self-updates, needs no trigger setup or permissions, and is blank-safe. Reach
   for `onEdit` only when a formula genuinely cannot express the behavior (e.g.
   writing to an *unrelated* cell, calling an external service, or timestamping an
   immutable value). Example auto-ID (no trigger):
   `=IF(B2="","","C"&TEXT(COUNTA($B$2:B2),"000"))`
9. **Use only real `DataValidationBuilder` methods.** There is **no**
   `requireTextLengthLessThanOrEqualTo`, `requireTextStartsWith`,
   `requireTextMatches`, etc. — inventing them throws *"... is not a function"*.
   For length, prefix, or pattern checks use `requireFormulaSatisfied`. The full
   valid method list is in `references/sheet-formatting.md`.

## Assets — start from working templates

Adapt these rather than writing from scratch; they already encode the rules above.

- `assets/setup-sheet-template.gs` — full sheet setup: text-date conversion,
  formula injection (blank-safe), validation/dropdowns, status-tag formula,
  conditional formatting, auto-width, centering, bold header.
- `assets/inject-answers-template.gs` — scans a sheet for question text and writes
  the answer formula into the adjacent cell (cross-sheet lookup/aggregation).
- `assets/webapp/WebApp.gs` + `assets/webapp/Form.html` — full-screen, designed
  data-entry web app that validates and appends rows, with live stats and an
  auto-suggested next ID.

**Workflow with a template:** copy it into the project, then adapt the column
map, validation rules, formulas, and field list to the user's actual header.
Delete sections the user didn't ask for. Keep the universal-rule scaffolding.

## Delivering the result

Apps Script cannot be run from this environment — output the final `.gs`/`.html`
and give the user precise steps:

- **Bound script:** Extensions → Apps Script → paste into a file → Run the
  function. For HTML, create an *HTML* file whose name **exactly matches** the
  `createHtmlOutputFromFile('<Name>')` call.
- **Web app:** Deploy → New deployment → Web app → Execute as **Me**, set access →
  Deploy → authorize → open the URL. **Re-deploy after every code change** (or use
  Test deployments for the latest-code `/dev` URL).
- State plainly what each function does and any assumptions (sheet name, header
  presence). If a header row is required, say so.
