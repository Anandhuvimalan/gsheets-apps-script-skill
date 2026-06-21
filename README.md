# gsheets-apps-script — Claude Skill

A [Claude Skill](https://docs.anthropic.com/en/docs/claude-code/skills) that builds
**error-free Google Apps Script** and HTML Service solutions for Google Sheets from
your data. Give it a sheet/CSV and say what you want — it generates `.gs`/`.html`
that runs the first time.

## What it does

- **Format & validate a sheet** — data validation & dropdowns, blank-safe
  calculation formulas, conditional formatting/coloring, auto column width,
  center alignment, bold header.
- **Build a designed data-entry web app** — a full-screen HTML Service form that
  validates input and appends rows to the sheet, with live stats and an
  auto-suggested next ID.
- **Answer questions / compute values** — injects cross-sheet lookup and
  aggregation formulas (`VLOOKUP`, `COUNTIF`, `SUMIF`, `INDEX/MATCH`) into the
  right cells.

It encodes the non-obvious rules that make these tasks work reliably: guarding
empty sheets, converting text dates before date validation, keeping empty rows
blank instead of showing `0`, and matching web-app file names / re-deploy steps.

## Install

**Option A — packaged file:** download `gsheets-apps-script.skill` and place it in
your skills directory, or unzip it there.

**Option B — folder:** copy the `gsheets-apps-script/` folder into your skills
directory:

```bash
# Claude Code (user-level skills)
cp -r gsheets-apps-script ~/.claude/skills/
```

Then ask Claude something like *"format this sheet"*, *"build a form for my
sheet"*, or *"answer the questions in this sheet"* with your data attached.

## Layout

```
gsheets-apps-script/
├── SKILL.md                      # triggers, universal rules, workflow routing
├── references/
│   ├── sheet-formatting.md       # validation, formulas, conditional formatting
│   ├── webapp-form.md            # HTML Service form pattern + deploy
│   ├── formulas.md               # lookup/aggregation cookbook + Q&A
│   └── troubleshooting.md        # common Apps Script errors → fixes
└── assets/
    ├── setup-sheet-template.gs   # full sheet setup (adapt the column map)
    ├── inject-answers-template.gs# question → answer-formula injector
    └── webapp/{WebApp.gs, Form.html}  # the designed form
```
