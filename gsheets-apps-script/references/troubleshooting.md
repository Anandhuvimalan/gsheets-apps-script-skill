# Troubleshooting — Apps Script errors & fixes

Real errors seen building these solutions, with causes and fixes.

| Symptom / error | Cause | Fix |
|---|---|---|
| `Exception: The number of rows in the range must be at least 1` | `getRange(2, c, lastRow-1, 1)` when `lastRow` is 1 (header only), giving 0 rows | Guard `if (lastRow < 2) return;` before computing `numRows = lastRow - 1` |
| `TypeError: <x> is not a function` | Method typo, calling a method on the wrong object type, or **stale/duplicate code** still in the file from a partial paste | Verify the method name; **replace the entire file** rather than pasting alongside old code. The reported line number points at the bad call |
| Every date cell shows a red "invalid" flag after adding date validation | CSV dates are **text**, not real `Date` values | Convert text → `Date` first (`new Date(str)`), then apply `requireDate()` — see `sheet-formatting.md` |
| Calculated column shows `0` and the label column says "Low/None" on empty rows | Plain arithmetic formula evaluates blanks as `0` | Wrap: `=IF(OR(A2="",B2=""),"",A2*B2)`; guard label formulas with `=IF(M2="","",...)` |
| `setFormulas` / `setValues` throws about dimensions | The 2D array's size ≠ the target range size | Make the array exactly `numRows × 1` (or match the range); use `getRange(2, col, arr.length, 1)` |
| Web app shows old behavior after editing code | A deployment pins a code version | **Re-deploy** (New deployment or Manage deployments → edit → new version), or use the Test deployment `/dev` URL |
| `google.script.run` "does nothing" / return value is `undefined` | It's asynchronous and returns no value synchronously | Use `.withSuccessHandler(fn).withFailureHandler(fn)` |
| Blank page / "refused to connect" in the web app | Missing `<base target="_top">`, or returning a string instead of `HtmlOutput` | Add `<base target="_top">`; return `HtmlService.createHtmlOutputFromFile(...)` |
| `createHtmlOutputFromFile` "file not found" | HTML filename ≠ the string argument | Name the HTML file to match exactly (no `.html` in the call) |
| `getUi` / "Cannot call ... from this context" in a web app | `SpreadsheetApp.getUi()` isn't available in web-app execution | Remove it from the web-app path; use `spreadsheet.toast()` or return a message to the page |
| Conditional-format rule highlights the wrong rows | Relative reference not anchored to the range's top-left row | If the range starts at row 2, write the formula for row 2 (`=$N2=...`); Sheets shifts it per row |
| Existing conditional-format rules disappeared | `setConditionalFormatRules(rules)` **replaces** all rules | Read `sheet.getConditionalFormatRules()` first and append, or set everything in one pass |
| Cross-sheet formula returns `#REF!` / `#N/A` | Wrong tab name, tab missing, or no match found | Verify the exact tab name (quote names with spaces); for lookups confirm the key exists and 4th `VLOOKUP` arg is `FALSE` |
| `Cannot read properties of null (reading 'getRange')` | `getSheetByName(name)` returned `null` (name mismatch) | Confirm the tab name (remember: CSV filename's sheet name is after the ` - `); fall back to `getSheets()[0]` |
| Windows: init/package script `UnicodeEncodeError: '\U0001f680'` | Console is cp1252 and the script prints emoji | Run with `PYTHONUTF8=1` (e.g. `export PYTHONUTF8=1` before the command) |
