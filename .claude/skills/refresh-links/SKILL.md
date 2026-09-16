---
name: refresh-links
description: Link quarterly earnings press releases from SEC EDGAR into a model's quarter headers. Use when the user asks to add, refresh, or fill in press release links for a model — e.g. "add the press release links to DUK", "refresh links now that Q3 closed", "link the earnings releases for the new models".
---

# Refresh press release links

Hyperlinks each quarter header on a model's `Model` sheet to that quarter's
earnings press release on SEC EDGAR. The target is the exhibit itself — the
release document — not the filing index page.

## Usage

```bash
export SEC_EDGAR_USER_AGENT="you@apexcleanenergy.com"   # SEC 403s without this

python .claude/skills/refresh-links/scripts/refresh_links.py DUK NEE
python .claude/skills/refresh-links/scripts/refresh_links.py --all --dry-run
python .claude/skills/refresh-links/scripts/refresh_links.py NVDA --rebuild
python .claude/skills/refresh-links/scripts/refresh_links.py --all --ignore-cache
```

Run from the repo root. Needs `openpyxl` (`py -3.14 -m pip install openpyxl`).

| Flag | Effect |
|---|---|
| *(none)* | Fills only unlinked quarters. Safe to re-run as quarters close. |
| `--dry-run` | Reports without writing. Use first on anything unfamiliar. |
| `--rebuild` | Re-evaluates quarters that already have links. Where no replacement is found the existing link is **kept**, so a rebuild cannot lose coverage. |
| `--ignore-cache` | Retries quarters previously recorded as empty. Use after changing the matching logic. |

Quarters found to have no release are cached in `.refresh_links_cache.json`
(gitignored) so the companies that yield nothing are not re-searched at full
cost forever. `--ignore-cache` is the only way past it.

## Steps

1. `--dry-run` first and read the output.
2. Check anything flagged `(no prior-year comparison -- verify)`. Open one or
   two and confirm the quarter.
3. Re-run without `--dry-run`.
4. Report which quarters filled and which stayed blank, **with the reason**.
   Never call a quarter linked if the run reported no release found.
5. Verify from the workbooks, not from the run's own summary — see
   **Verify against the files** below.

## How quarters are matched

A filing date alone does **not** identify the quarter. Item 2.02 also carries
guidance updates, preliminary results and one-off charges, so picking the
earliest filing in the window gets the wrong document. Eversource's Q423 has
three candidates in Jan–Feb 2024 and only the February one is the results
release.

Each candidate is fetched and scored on whether its text discusses that
calendar quarter, whether it also cites the prior-year same quarter (the
strongest available signal — a real release nearly always carries the
year-over-year comparison), and how close the filing sits to the normal
reporting lag.

## Fiscal versus calendar quarters

Headers are **calendar** quarters. 24 companies in this repo are off-cycle —
NVDA (January), CSCO (July), EMR/ROK/ATKR (September), ENS/NXT (March).

For those the script ignores the company's own quarter labels, which are
fiscal, and places each release by **nearest calendar quarter end**. Two routes:

- An explicit period end date in the text, mapped to the nearest quarter end.
- Failing that, the fiscal label plus the company's fiscal year end. Cisco's
  releases carry *no dates at all* — only "third quarter of fiscal 2020" — so
  the end is derived: Q4 ends at the fiscal year end, each earlier quarter
  three months before. Cisco Q3 FY20 ends 25 Apr 2020 → Q120.

This was validated against NVDA, where it reproduces links that had been made
by hand, and ROK, where Q120 correctly resolves to fiscal Q2 FY20.

## Gotchas already handled

Each of these caused a real wrong or missing link before it was fixed.

- **Not every earnings 8-K is tagged Item 2.02.** Xcel, MGE and PPL have filed
  under 9.01 only; OGE's Q422 arrived as an 8-K/A. All 8-K forms in the window
  are searched, tagged ones first.
- **The release is not always EX-99.1.** Emera puts the MD&A in EX-99.1 and the
  release in EX-99.4; Bloom's Q320 release is EX-99.2 behind a shareholder
  letter.
- **Filers mistype the exhibit type.** NVIDIA's Q3 FY20 press release is filed
  as `EX-95.1`. Exhibits are matched on description ("Q3FY20 PRESS RELEASE") as
  well as type, so a typo no longer hides the document.
- **Headlines can sit far into the document.** 2000s-era releases pad their
  contact block heavily — NVIDIA's 2007 Q3 release reaches "results" at ~4,850
  characters. The positive scan runs to 8,000; disqualifiers stay at 800 so an
  incidental later mention cannot reject a real release.
- **Item 2.02 documents that are not results.** Segment restatements, call
  advisories, dividend declarations, offering notices. Rockwell's "Provides
  Historical Data for new Operating Segments" was picked for Q320 before this
  filter existed. A *strong* results headline outranks these, because a release
  may report results and announce something else in one breath.
- **Call advisories say "will host a conference call" — so do real releases.**
  They are separated by length: advisories are short and carry no tables.
- **Quarterly results published as a "Shareholder Letter"** — Oklo, Bloom and
  other recent listings. Rejecting that phrasing silently zeroed Oklo.
- **Fiscal labels joined by "and".** "Fourth Quarter and Fiscal 2020" is not
  matched by an of/ended pattern; Q419 came back empty until it was.
- **Q4 phrased as full-year only.** "Vistra Reports Full-Year 2023 Results"
  never says "fourth quarter". The pattern is scoped to the headline verb so
  "reaffirms full-year 2024 guidance" in a Q1 release is not read as Q4.
- **Foreign private issuers file 6-K, not 8-K.** Fortis and Emera are handled.
- **Header formats differ.** Both `Q120` and `Q12024` are parsed, keyed on the
  parsed quarter rather than the label text. RUN uses the four-digit form and
  was silently skipped until this was added. Two-digit years pivot at 70, so
  `Q498` is 1998, not 2098.
- **Reporting windows are generous** (Q4 runs into April) because a
  non-accelerated filer has 90 days for a 10-K. Overlap is harmless — the text
  decides the quarter, not the date.

## Ticker collisions

The resolved EDGAR company name is compared against the company name on the
index workbook's `Main` sheet, and a mismatch **refuses to link** rather than
guessing. This is what catches `H` → Hyatt Hotels (not Hydro One), `DIA` → the
SPDR Dow ETF (not Dialight), `SUN` → Sunoco (not Stardust Solar).

The `SKIP` dict at the top of the script records known-bad tickers with their
reason. Add to it when you confirm one; the name check is the safety net, not
the primary record.

## Running at scale

A full 165-model pass is several hours and tens of thousands of requests.

- **A 403/429 aborts the run.** This is deliberate. A throttle that was
  swallowed would look exactly like "these 90 companies never reported", and at
  scale nobody would notice. Blanks are only trustworthy because transport
  failures cannot produce them.
- Failures are reported as lookup errors and **not** cached as absences, so a
  re-run retries them.
- Links save every few quarters, not once per company, so an interrupted run
  keeps its work. A 165-model run *was* killed mid-way by system memory
  pressure; resuming cost almost nothing.
- One bad company cannot end the run.
- Close Excel first. A `~$TICKER.xlsx` lock makes that workbook unwritable.

## Workbook safety

Never call `wb.save()` directly — use `xlsx_safe.save_workbook()`.

openpyxl rebuilds the file from its own object model and silently drops parts
it does not understand. A plain save deleted `xl/drawings/drawing1.xml` — two
divider arrows on the Model sheet — from 27 models before this was caught. The
helper re-injects dropped parts with their content type, relationship and the
owning sheet's `<drawing>` element.

It deliberately does **not** restore parts openpyxl re-emits under its own
names: `sharedStrings` (strings are inlined) and cell comments, which openpyxl
rewrites as `xl/comments/comment1.xml`. Restoring those leaves two competing
comment representations with the originals orphaned.

Known unavoidable loss: openpyxl converts **threaded comments to classic
ones**. TSLA lost its `threadedComment` part this way. If a model carries
threaded comments and they matter, do not write to it through openpyxl.

## Verify against the files

The run's summary counts *links written this pass*, which is not coverage — a
model already fully linked correctly reports zero. Read the workbooks:

```python
ws = openpyxl.load_workbook(f"{ticker}.xlsx")["Model"]
linked = [c.value for c in ws[2] if c.hyperlink]
```

Spot-check that a sample of links return HTTP 200 and that every link points at
`sec.gov`.

## When there is no link

Leave the header unlinked. An empty header honestly means no release exists; a
wrong link is worse than a blank. Established blanks in this repo:

- **Not SEC-registered** — 45 of the 165 tickers have no EDGAR ticker mapping
  at all: mostly TSX/TSXV/CNSX and OTC Expert Market names (Brookfield
  Renewable, TransAlta, Northland, Capital Power), plus dormant shells whose
  last filings predate 2020.
- **CEG** — no standalone registrant before the February 2022 spinoff.
- **EXC Q421** — no Q4 2021 earnings 8-K exists; results went into the 10-K
  around the Constellation separation.
- **Pre-2003 quarters generally.** EDGAR 8-Ks start around 2000 and Item 2.02
  only exists from August 2004. NVDA's earliest release is Q102, and Q498–Q401
  are blank despite the 1999 IPO because nothing was filed as a release.
- **Quarters not yet reported.** Quarters closing within the last 25 days are
  skipped rather than searched.

Preliminary-results releases are linked when they are the only thing filed —
TXNM Q421, GNE Q425, NVDA Q102. Say so when reporting.
