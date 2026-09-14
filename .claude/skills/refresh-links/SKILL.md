---
name: refresh-links
description: Link quarterly earnings press releases from SEC EDGAR into a model's quarter headers. Use when the user asks to add, refresh, or fill in press release links for a model — e.g. "add the press release links to DUK", "refresh links now that Q3 closed", "link the earnings releases for the new models".
---

# Refresh press release links

Hyperlinks each quarter header on a model's `Model` sheet to that quarter's
earnings press release on SEC EDGAR. The link target is the `EX-99.x` exhibit —
the release document itself, not the filing index.

## Usage

```bash
export SEC_EDGAR_USER_AGENT="you@apexcleanenergy.com"   # SEC 403s without this

python .claude/skills/refresh-links/scripts/refresh_links.py DUK NEE
python .claude/skills/refresh-links/scripts/refresh_links.py --all --dry-run
python .claude/skills/refresh-links/scripts/refresh_links.py DUK --rebuild
```

Run it from the repo root. Needs `openpyxl` (`py -3.14 -m pip install openpyxl`).

- Default: fills only quarters that are **not already linked**, so it is safe to
  re-run as new quarters close. Existing links are left alone.
- `--dry-run` reports what it would do and writes nothing. **Use this first** on
  anything you haven't run before.
- `--rebuild` replaces links already present. Only for fixing a bad run.
- Quarter range follows the headers in each model, so a model covering Q218
  onward is handled without changing the script.

## Steps

1. Run with `--dry-run` and read the output.
2. Check anything flagged `(no prior-year comparison in text -- verify)` — see
   below. Open one or two and confirm they are the right quarter.
3. Re-run without `--dry-run` to write.
4. Tell the user which quarters were filled and which were left blank, and why.
   Do not describe a quarter as linked if the script reported no release found.

## How quarters are matched

A filing date alone does **not** identify the quarter. Companies file other
Item 2.02 8-Ks — guidance updates, preliminary results, charges — inside the
same reporting window, and picking the earliest one gets the wrong document.
Eversource's Q423, for example, has three candidates in Jan–Feb 2024 and only
the February one is the results release.

So each candidate release is fetched and scored on:

- whether its text actually discusses that calendar quarter, and
- whether it also cites the **prior-year** same quarter — a real earnings
  release nearly always carries the year-over-year comparison, and this is the
  single strongest signal available, and
- how close the filing date sits to the normal reporting lag.

The `(no prior-year comparison ...)` flag means the best match lacked that
comparison. It is usually still correct — some companies phrase Q4 as "2024
year end earnings report" with no quarter language at all — but it is worth a
look, and it is how preliminary-results releases surface.

## Gotchas this already handles

- **Not every earnings 8-K is tagged Item 2.02.** Xcel, MGE and PPL have filed
  them under 9.01 only, and OGE's Q422 arrived as an 8-K/A. All 8-K forms in the
  window are searched, tagged ones first.
- **The release is not always EX-99.1.** Emera puts the MD&A in EX-99.1 and the
  press release in EX-99.4; Bloom's Q320 release is EX-99.2 behind a shareholder
  letter. Every EX-99.x is checked and non-releases (MD&A, interim financials,
  52-109 certifications) are filtered out.
- **Foreign private issuers file 6-K, not 8-K.** Fortis and Emera are handled.
- **Ticker collisions.** EDGAR's `H` is Hyatt Hotels, not Hydro One. The `SKIP`
  dict at the top of the script lists tickers that are not the SEC registrant of
  the same name — add to it rather than letting a wrong match through.

## Fiscal versus calendar quarters

The script maps **calendar** quarters. Every electric utility in this repo has a
December 31 year end, so the two coincide. If a company's fiscal year ends
elsewhere — NVDA (January), CSCO (July) — the script prints a warning and its
output must be reviewed: a release headed "first quarter fiscal 2021" may be a
calendar Q2. Do not trust an unreviewed run for those.

## When there is no link

Leave the header unlinked. An empty header honestly means no release exists;
a wrong link is worse than a blank. Known cases in this repo:

- **Hydro One, Commerce Energy, Alaska Power & Telephone, Summer Energy,
  Dynastar** — not SEC-registered, or dormant for over a decade.
- **CEG** — no standalone registrant before the February 2022 spinoff.
- **EXC Q421** — no Q4 2021 earnings 8-K was ever filed; results went into the
  10-K around the Constellation separation.
- **Quarters that have not closed yet.**
