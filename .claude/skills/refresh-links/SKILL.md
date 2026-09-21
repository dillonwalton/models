---
name: refresh-links
description: Link quarterly earnings press releases from SEC EDGAR into a model's quarter headers. Use when the user asks to add, refresh, or fill in press release links for a model — e.g. "add the press release links to DUK", "refresh links now that Q3 closed", "link the earnings releases for the new models".
---

# Refresh press release links

For a brand new company use **build-model**, which creates the model, adds it to
the index, and then applies everything here. Use this skill directly to refresh
models that already exist.

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

Run from the repo root — models are found under `companies/`, and the index
workbook is `portfolio.xlsx`. Needs `openpyxl` (`py -3.14 -m pip install openpyxl`).

| Flag | Effect |
|---|---|
| *(none)* | Fills only unlinked quarters. Safe to re-run as quarters close. |
| `--dry-run` | Reports without writing. Use first on anything unfamiliar. |
| `--rebuild` | Re-evaluates quarters that already have links. Where no replacement is found the existing link is **kept**, so a rebuild cannot lose coverage. |
| `--ignore-cache` | Retries quarters previously recorded as empty. Use after changing the matching logic. |
| `--fallback-periodic` | Where no earnings release exists, links the 10-Q/10-K covering that period instead. See below. |
| `--no-sync-history` | Skips the header-range check. **On by default**: the header row is made to span exactly the company's public life — extended back to its first report, with quarters and years from before it dropped. |
| `--sync-only` | Fixes the header range and stops without searching for links. |

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
- **Long-form disclosure documents match on sheer size.** A spin-off
  information statement, prospectus or proxy spans years and mentions almost
  any quarter. The Constellation spin-off information statement was picked for
  both EXC and CEG Q421 until those document types were disqualified.
- **Range requests and gzip do not mix.** The fetcher asks for a byte range
  *and* gzip, so bodies are routinely truncated gzip members. `gzip.decompress`
  rejects those outright, which turned readable documents into lookup failures;
  a raw `zlib.decompressobj` returns everything up to the cut.

## A model must span exactly the company's public life

**This is a standing requirement, not an option.** A model should cover every
quarter the company has reported as a public company, and none before it. Both
directions matter:

- Reach back to the **first ever public report**, not to whatever the template
  shell happened to start at.
- **Remove quarters and years from before the company was public.** A header
  for a period when the company did not exist as a filer is not an empty cell
  waiting to be filled — it is wrong, and it invites someone to fill it.

This happens **by default** on every run; `--no-sync-history` opts out, and
`--sync-only` does the structural fix alone (useful because it is cheap and the
re-linking is not). The first public quarter is taken from the earliest 10-Q/10-K on EDGAR: a
company files one only once it reports publicly.

SpaceX is the case that motivated this. Its model shell ran from Q120, but
SpaceX listed on 12 June 2026 and filed its first 10-Q that August — 25 of its
28 quarters described a period when it was a private company.

Two limits to state honestly when reporting:

- **EDGAR bounds the answer.** Electronic filing phased in 1993–1996, so for a
  company listed long before that this is EDGAR's earliest filing, not its true
  first report. AEP reaches Q194 and Alliant Q492; both were public for decades
  before.
- **Models holding data or formulas are refused, not restructured.** Column
  changes would silently misalign them. BE, EQIX, GEV, RUN and TSLA fall in
  this group. The check is `model_is_skeleton()` — anything outside header
  row 2 and `B3`. They were handled through Excel instead; see below.

## Workbooks openpyxl must not write

**openpyxl cannot round-trip every workbook.** Writing `BE.xlsx` or `TSLA.xlsx`
through it produces a file that openpyxl itself reads back perfectly and Excel
refuses to open at all — apparently over its handling of cell comments.
Stripping the threaded comments does not help; the output is rejected either
way.

This went unnoticed for several commits because every check used openpyxl.
**Verifying a workbook by reading it back with the library that wrote it proves
nothing.** Open it in Excel:

```powershell
$xl = New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
Get-ChildItem companies\*.xlsx | Where-Object { $_.Name -notlike '~$*' } | ForEach-Object {
  try { $wb=$xl.Workbooks.Open($_.FullName); $wb.Close($false) } catch { $_.Name }
}
$xl.Quit()
```

For those models, work through Excel instead:

- `compute_links.py TICKER out.json` — resolves the links, writes no workbook.
- `apply_links.ps1 -Ticker TICKER -JsonPath out.json` — applies them in Excel.
- `extend_quarters.ps1 -Ticker TICKER -FirstQuarter N -FirstYear YYYY` — inserts
  historical quarter columns.
- `extend_years.ps1 -Ticker TICKER -FirstYear YYYY` — inserts historical columns
  into the annual block. Run it *after* `extend_quarters.ps1`, since the annual
  block sits to the right of the quarters and shifts when they are inserted.

Excel must do the column insert regardless of the comment problem, because it
re-points every formula: a 35-column shift turned `=861-H3` into `=861-AQ3`.
openpyxl moves the cells and leaves the references behind, silently.

## The actuals/forecast arrows

Two vertical connectors on the Model sheet mark where actuals end: one between
Q226 and Q326, one between 2025 and 2026. They are anchored to absolute column
indices, so **every column insert leaves them pointing at the wrong place** —
after the history sync all 161 were still sitting at the template's original
positions.

```bash
python .claude/skills/refresh-links/scripts/place_arrows.py            # all models
python .claude/skills/refresh-links/scripts/place_arrows.py DUK --dry-run
```

It rewrites `xl/drawings/drawing1.xml` directly, not through openpyxl, which
drops drawings and cannot write some workbooks at all. It also adds the drawing
to models that never had one, and covers `base.xlsx` only when named
explicitly — worth doing, since new models inherit its arrow positions.

Re-run it after anything that inserts columns. Verify in Excel, filtering to
connector shapes: cell comments are also `Shapes` and sit in column C, which
looks like a misplaced arrow if you count them.

## Linking both the release and the full filing

A cell holds one hyperlink, and the release and the periodic report are separate
EDGAR submissions — Duke filed both for Q120 on 12 May 2020 under different
accession numbers, so no single URL reaches both. A filing index page shows
everything in *one* submission (release, CFO commentary, presentation, XBRL),
never the 10-Q.

So `--filing-row` puts the 10-Q/10-K in **row 1**, above the quarter header,
in small grey. Row 1 was empty in 162 of 165 models, so nothing is displaced
and no formulas move.

## Reaching quarters older than EDGAR's press releases

Earnings releases reached EDGAR only because of two rule changes: Regulation FD
(October 2000) pushed companies to file the release rather than just wire it,
and the August 2004 Form 8-K overhaul created Item 2.02. Before that there was
no obligation and most companies filed nothing, which is why NVIDIA's earliest
release is Q102 despite a January 1999 IPO. Nothing is missing — nothing was
required.

The **periodic report was always required**, so `--fallback-periodic` links the
10-Q, or the 10-K for a fiscal year end, covering any quarter with no release.
This is how analysts normally carry a model back past the release era: the
quarterly figures live in the primary filing, and the release is a convenience
layered on top.

The mapping is more reliable than release matching, not less — EDGAR reports
the period end date directly, so the quarter is read off rather than inferred
from text. No scoring is involved.

Periodic links are styled **purple and italic** (`FF7030A0`) against the blue
of a press release, so a model never implies a release where there is only a
filing. Report them as periodic reports when summarising.

Two EDGAR quirks are handled: filings before about 2000 name no
`primaryDocument`, and some 2000-era ones name a file EDGAR does not actually
serve (NVIDIA's 2000 10-Qs claim `0001.txt`, which 404s). Constructed document
URLs are HEAD-checked and fall back to the filing index page, which always
resolves.

Other routes when a link is not enough, roughly in order of usefulness:

- **Selected Quarterly Financial Data**, Item 302(a) of Regulation S-K — older
  10-Ks carry an unaudited eight-quarter table, so one filing backfills two
  years. The SEC removed the requirement in early 2021, so it exists for
  precisely the older years where it is needed.
- **Data vendors** (Capital IQ, FactSet, Bloomberg, LSEG) for normalised
  series. CapIQ calendarises off-cycle filers with the same nearest-quarter-end
  convention used here, so the two line up.
- **News archives** (Factiva, Nexis, ProQuest, PR Newswire and Business Wire)
  when the release text itself matters, or the Wayback Machine against the
  company's old IR page.

Watch comparability, not just availability: restatements, ASC 606 and ASC 842,
segment redefinitions and acquisitions all break continuity. NVIDIA restated
fiscal 2000–2002, which is why those quarters have no clean release at all. A
link to the original filing preserves provenance in a way a vendor's restated
figure does not.

## Ticker collisions

The resolved EDGAR company name is compared against the company name on the
index workbook's `Main` sheet, and a mismatch **refuses to link** rather than
guessing. This is what catches `H` → Hyatt Hotels (not Hydro One), `DIA` → the
SPDR Dow ETF (not Dialight), `SUN` → Sunoco (not Stardust Solar).

The `SKIP` dict at the top of the script records known-bad tickers with their
reason. Add to it when you confirm one; the name check is the safety net, not
the primary record.

**A SKIP entry can go stale.** SPCX was correctly skipped as private until
SpaceX listed in June 2026, at which point the entry was simply false. Re-check
an entry before relying on it, particularly for a company that might have
IPO'd.

`NAME_OVERRIDES` handles the opposite error — a legitimate match the heuristic
rejects because the sheet uses a trade name and EDGAR the registered one
("SpaceX" versus "SPACE EXPLORATION TECHNOLOGIES CORP"). Keep it narrow: each
entry disables the collision guard for that ticker.

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
ws = openpyxl.load_workbook(f"companies/{ticker}.xlsx")["Model"]
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
  only exists from August 2004. NVDA's earliest release is Q102. Those quarters
  are reachable with `--fallback-periodic`, which links the 10-Q/10-K instead —
  NVDA now runs Q498 to Q226 with only unreported quarters blank.
- **Quarters not yet reported.** Quarters closing within the last 25 days are
  skipped rather than searched.

Preliminary-results releases are linked when they are the only thing filed —
TXNM Q421, GNE Q425, NVDA Q102. Say so when reporting.
