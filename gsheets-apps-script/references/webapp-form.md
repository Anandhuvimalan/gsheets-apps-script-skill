# Designed Web App Form → writes to the sheet

Build a full-screen, polished data-entry form (HTML Service) whose submissions are
appended as rows. Templates: `assets/webapp/WebApp.gs` (backend) and
`assets/webapp/Form.html` (UI). Adapt the field list to the user's header; the
structure and styling are reusable as-is.

## Contents
- [Architecture](#architecture)
- [Backend pattern (WebApp.gs)](#backend-pattern-webappgs)
- [Frontend pattern (Form.html)](#frontend-pattern-formhtml)
- [Adapting the template](#adapting-the-template)
- [Deploying](#deploying)
- [Critical gotchas](#critical-gotchas)

## Architecture

- `doGet()` serves the HTML page.
- `getInitialData()` returns dropdown options, a live row count, and a suggested
  next ID — called once on page load.
- `addEmployee(data)` (rename per domain) validates and `appendRow`s, returning a
  fresh count + next ID. It **throws** on bad input; the UI catches it.
- The browser calls server functions via `google.script.run` — **asynchronous**,
  so always attach `withSuccessHandler` and `withFailureHandler`.

## Backend pattern (WebApp.gs)

```js
const SHEET_NAME = 'Sheet1';   // set to the real tab name

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Form')   // name must match the HTML file
    .setTitle('My Form')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getDataSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
}

function addEntry(data) {
  const sheet = getDataSheet_();
  const id = (data.id || '').toString().trim();
  if (!id /* || other required fields */) throw new Error('All fields are required.');

  // duplicate-key guard
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < keys.length; i++) {
      if (String(keys[i][0]).trim().toLowerCase() === id.toLowerCase()) {
        throw new Error('ID "' + id + '" already exists.');
      }
    }
  }

  sheet.appendRow([id, /* ...fields in column order... */]);
  return { message: 'Saved', count: sheet.getLastRow() - 1 };
}
```

Useful extras already in the template: `suggestNextId_()` increments the highest
numeric suffix found (`EMP106 → EMP107`); `getInitialData()` merges existing
column values with sensible defaults to populate dropdowns.

## Frontend pattern (Form.html)

- Single self-contained file: inline `<style>` and `<script>` (HTML Service does
  not serve separate static files conveniently).
- `<base target="_top">` in `<head>` so navigation escapes the iframe sandbox.
- Submit handler reads inputs into a plain object and calls the server:

```js
google.script.run
  .withSuccessHandler(function (res) { /* toast, reset form, refresh stats */ })
  .withFailureHandler(function (err) { /* show err.message inline */ })
  .addEntry(data);
```

- Show a loading state on the button during the call; reset the form and refocus
  the first field on success for fast repeat entry.

The template's design: full-viewport two-panel layout (gradient brand panel with
live stats + form card), Google **Inter** font, custom-styled selects, animated
success toast, mobile-responsive stacking. Reuse it; just swap the fields.

## Adapting the template

1. Set `SHEET_NAME` and the column order in `appendRow([...])` to match the header.
2. Rename `addEmployee`/`addEntry` consistently in **both** files.
3. Replace the form fields (`<input>`/`<select>`) and the `data = {...}` object so
   keys line up with the columns.
4. Adjust dropdown sources in `getInitialData()`; remove ID auto-suggest if the
   key isn't a sequential `PREFIX###`.

## Deploying

Deploy → New deployment → **Web app** → Execute as **Me** → set access → Deploy →
authorize → open URL. **Re-deploy after every change**, or use Test deployments
for a `/dev` URL that always runs the latest code.

## Critical gotchas

- **HTML filename must match** `createHtmlOutputFromFile('Form')` exactly (Apps
  Script appends `.html` itself — the file is just named `Form`).
- `google.script.run` returns nothing synchronously — never read its return value
  directly; use the handlers.
- A server function can only return JSON-serializable data (no `Date`/`Range`
  objects — convert to strings/primitives first).
- `SpreadsheetApp.getUi()` is unavailable in the web-app execution context; do not
  call it from `doGet`/server functions invoked by the page. `toast()` on the
  spreadsheet is fine.
- Concurrent submissions can interleave; for high concurrency wrap `appendRow`
  with `LockService.getScriptLock()`.
