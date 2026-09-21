---
name: build-model
description: Create a new financial model for a public company, add it to the portfolio index, and link every quarter to its SEC filings. Use when the user asks to build, create, or add a model for a company or ticker (e.g. "build a model for CoreWeave", "add a NVDA model", "new model for Digital Realty").
---

# Build a model

Creates the model **and brings it to the repo's standard** in one command. A
model is not finished when the file exists — it is finished when it spans the
company's public life, its quarters link to the filings, and it appears in the
index.

## Usage

```bash
export SEC_EDGAR_USER_AGENT="you@apexcleanenergy.com"   # SEC 403s without this

python .claude/skills/build-model/scripts/build_model.py CRWV
python .claude/skills/build-model/scripts/build_model.py FSLR --dry-run
python .claude/skills/build-model/scripts/build_model.py FSLR --name "First Solar, Inc."
```

Run from the repo root. Needs `openpyxl`.

## What it does

1. **Creates `companies/<TICKER>.xlsx`** from `base.xlsx`. Never overwrites an
   existing model. Dots in tickers become underscores — `BEP.UN` → `BEP_UN.xlsx`.
2. **Adds the company to `portfolio.xlsx` `Main`** and hyperlinks the Ticker
   cell to the model. Name and industry come from EDGAR, so the row matches how
   the company actually reports. If the ticker is already listed, the existing
   row is linked rather than duplicated.
3. **Sets the header range to the company's public life** — back to its first
   10-Q/10-K, dropping any quarters and years from before it.
4. **Links each quarter to its earnings release**, falling back to the
   10-Q/10-K for quarters that predate the press-release era.
5. **Puts the 10-Q/10-K in row 1**, above each quarter header.
6. **Pins the actuals/forecast arrows** to Q326 and 2026. This runs last,
   because every earlier step can move columns underneath them.

## Resolving the ticker

The filename must be the company's exchange ticker, uppercase. Given a company
name, confirm the ticker before running — the script resolves it against
EDGAR's map and prints the registered name, so check that what comes back is
the company you meant.

**EDGAR ticker collisions are real**: `H` is Hyatt Hotels, not Hydro One; `DIA`
is the SPDR Dow ETF, not Dialight; `SUN` is Sunoco, not Stardust Solar. If the
printed name does not match, stop and check rather than proceeding.

## When something is missing

A company with no EDGAR ticker mapping still gets a model and an index row, and
the run says so plainly instead of linking nothing in silence. Leave those
quarters blank — an empty header honestly means no filing exists, and a wrong
link is worse than a blank.

`refresh-links` carries the detail on how quarters are matched, which workbooks
openpyxl must not write, and what the blanks mean. Read it before changing
anything here.

## Verify

The run prints a summary — quarters, range, release links, filing links. Check
it looks right for the company: a 2024 IPO with 80 quarters means something
resolved to the wrong registrant.

Then open the model in Excel. Reading a workbook back with the library that
wrote it proves nothing; two models were silently corrupted that way before the
Excel check became standard.
