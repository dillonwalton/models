"""Fill quarter-header links in a model from SEC EDGAR earnings press releases.

Finds each company's quarterly earnings release on EDGAR (the EX-99.x exhibit
on the earnings 8-K, or 6-K for foreign private issuers) and hyperlinks it from
the matching quarter header on the Model sheet.

Quarters are matched against the text of each release -- not the filing date --
because filing dates alone mis-assign releases (a January 8-K is not the Q4
results release just because it lands in the Q4 reporting window).

Usage:
    python refresh_links.py DUK NEE          # specific models
    python refresh_links.py --all            # every model in the repo
    python refresh_links.py --all --dry-run  # report only, write nothing
    python refresh_links.py DUK --rebuild    # redo links already present

Requires: openpyxl.  Set SEC_EDGAR_USER_AGENT to a contact email (SEC rejects
requests without one).
"""
import argparse, datetime, glob, gzip, json, os, re, sys, time, urllib.request

UA = os.environ.get("SEC_EDGAR_USER_AGENT", "").strip()

# Model files that are not the SEC registrant of the same ticker, or not SEC
# registrants at all. Keeps us from linking the wrong company's releases.
SKIP = {
    "H":     "Hydro One (TSX) -- EDGAR ticker 'H' is Hyatt Hotels, a different company",
    "CMNR":  "Commerce Energy Group -- last EDGAR filing 2009",
    "APTL":  "Alaska Power & Telephone -- last EDGAR filing 2005",
    "SUME":  "Summer Energy Holdings -- no EDGAR ticker mapping",
    "DYNA":  "Dynastar Holdings -- no EDGAR ticker mapping",
    "SBGSY": "Schneider Electric ADR -- files in France, not with the SEC",
}
NOT_A_MODEL = {"Utilities.xlsx", "base model.xlsx"}

ORD = {"first": 1, "second": 2, "third": 3, "fourth": 4}
MONTH_END = {"march 31": 1, "june 30": 2, "september 30": 3, "december 31": 4}
QUARTER_END = {1: (3, 31), 2: (6, 30), 3: (9, 30), 4: (12, 31)}
# Exhibits filed alongside the release that are not the release itself. Matched
# as regexes against the start of the document, so keep them specific -- a bare
# "management" would reject real releases that mention a management call.
NOT_A_RELEASE = (r"management.{0,8}s discussion", r"unaudited condensed",
                 r"certification of interim", r"earnings coverage ratio",
                 r"form 52-109", r"consent of independent")


def fetch(url, cap=250000, tries=3):
    for attempt in range(tries):
        try:
            req = urllib.request.Request(url, headers={
                "User-Agent": UA, "Accept-Encoding": "gzip", "Range": "bytes=0-%d" % cap})
            with urllib.request.urlopen(req, timeout=60) as resp:
                raw = resp.read()
                try:
                    raw = gzip.decompress(raw)
                except Exception:
                    pass
                return raw.decode("utf-8", "replace")
        except Exception:
            if attempt == tries - 1:
                return ""
            time.sleep(1.0)
        finally:
            time.sleep(0.11)          # SEC allows 10 req/sec


def plain(html):
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html)).strip()


def periods(text):
    """Calendar (quarter, year) pairs the release claims to report."""
    t = text.lower()
    found = set()
    for m in re.finditer(r"(first|second|third|fourth)[- ]quarter(?:\s+(?:of\s+|ended\s+)?(\d{4}))?", t):
        if m.group(2):
            found.add((ORD[m.group(1)], int(m.group(2))))
    for m in re.finditer(r"(\d{4}) (first|second|third|fourth)[- ]quarter", t):
        found.add((ORD[m.group(2)], int(m.group(1))))
    for m in re.finditer(r"three months ended (march 31|june 30|september 30|december 31),? (\d{4})", t):
        found.add((MONTH_END[m.group(1)], int(m.group(2))))
    for m in re.finditer(r"(?:year|fourth quarter) ended december 31,? (\d{4})", t):
        found.add((4, int(m.group(1))))
    # "fourth quarter and full year 2020" / "... and fiscal year 2020" -- the year
    # does not sit directly after "quarter", so the patterns above miss it.
    for m in re.finditer(r"(first|second|third|fourth)[- ]quarter and (?:the )?"
                         r"(?:full[- ]|fiscal[- ])?year (?:ended [a-z]+ \d{1,2},? )?(\d{4})", t):
        found.add((ORD[m.group(1)], int(m.group(2))))
    for m in re.finditer(r"q([1-4])\s+(\d{4})", t):                # "Q3 2025"
        found.add((int(m.group(1)), int(m.group(2))))
    for m in re.finditer(r"(\d{4}) year[- ]end earnings", t):      # Xcel's phrasing
        found.add((4, int(m.group(1))))
    return found


def looks_like_release(text):
    head = text.lower()[:400]
    if any(re.search(p, head) for p in NOT_A_RELEASE):
        return False
    return bool(re.search(r"\b(reports|announces|releases)\b", head))


def reporting_window(q, year):
    """Calendar window in which quarter (q, year) is normally reported."""
    return {1: ("%d-04-01" % year, "%d-06-30" % year),
            2: ("%d-07-01" % year, "%d-09-30" % year),
            3: ("%d-10-01" % year, "%d-12-31" % year),
            4: ("%d-01-01" % (year + 1), "%d-03-31" % (year + 1))}[q]


def lag_days(q, year, filed):
    month, day = QUARTER_END[q]
    filed_on = datetime.date(int(filed[:4]), int(filed[5:7]), int(filed[8:10]))
    return (filed_on - datetime.date(year, month, day)).days


def resolve_cik(ticker):
    data = json.loads(fetch("https://www.sec.gov/files/company_tickers.json", cap=6000000))
    for row in data.values():
        if row["ticker"].upper() == ticker.upper():
            return row["cik_str"], row["title"]
    return None, None


def all_filings(cik):
    sub = json.loads(fetch("https://data.sec.gov/submissions/CIK%010d.json" % cik, cap=8000000))
    rows = []

    def absorb(block):
        for i in range(len(block["accessionNumber"])):
            rows.append({k: block[k][i] for k in
                         ("accessionNumber", "filingDate", "form", "items")})

    absorb(sub["filings"]["recent"])
    for extra in sub["filings"].get("files", []):
        absorb(json.loads(fetch("https://data.sec.gov/submissions/" + extra["name"], cap=8000000)))
    return sub.get("fiscalYearEnd"), sub.get("name"), rows


def exhibits(cik, accession):
    nodash = accession.replace("-", "")
    html = fetch("https://www.sec.gov/Archives/edgar/data/%d/%s/%s-index.htm" % (cik, nodash, accession))
    out = []
    for row in re.findall(r"<tr[^>]*>(.*?)</tr>", html, re.S | re.I):
        cells = [re.sub(r"<[^>]+>", "", c).replace("&nbsp;", " ").strip()
                 for c in re.findall(r"<td[^>]*>(.*?)</td>", row, re.S | re.I)]
        href = re.search(r'href="([^"]+)"', row)
        if len(cells) >= 4 and href and cells[3].upper().startswith("EX-99"):
            out.append({"type": cells[3],
                        "url": "https://www.sec.gov" + href.group(1).replace("/ix?doc=", "")})
    return out


def best_release(cik, filings, q, year):
    """Best EX-99 release for calendar quarter (q, year), or None."""
    low, high = reporting_window(q, year)
    candidates = [f for f in filings
                  if low <= f["filingDate"] <= high
                  and (f["form"].startswith("8-K") or f["form"].startswith("6-K"))]
    # Earnings 8-Ks are usually tagged 2.02, but not always -- some are filed
    # under 9.01 only, and some arrive as 8-K/A. Try the tagged ones first.
    candidates.sort(key=lambda f: ("2.02" not in (f["items"] or ""), f["filingDate"]))
    scored = []
    for f in candidates:
        for ex in exhibits(cik, f["accessionNumber"]):
            text = plain(fetch(ex["url"]))
            if not looks_like_release(text):
                continue
            claimed = periods(text)
            if (q, year) not in claimed:
                continue
            # A real earnings release nearly always cites the prior-year quarter
            # for comparison; that is the strongest signal available.
            yoy = (q, year - 1) in claimed
            ideal = 55 if q == 4 else 40
            scored.append({
                "score": (100 if yoy else 0) - abs(lag_days(q, year, f["filingDate"]) - ideal) * 0.5,
                "url": ex["url"], "filed": f["filingDate"], "form": f["form"], "yoy": yoy})
            break
    scored.sort(key=lambda c: -c["score"])
    return scored[0] if scored else None


def quarters_through_today(start_year):
    """Every quarter from start_year that has actually closed."""
    today = datetime.date.today()
    out = []
    for year in range(start_year, today.year + 1):
        for q in (1, 2, 3, 4):
            month, day = QUARTER_END[q]
            if datetime.date(year, month, day) < today:
                out.append((q, year))
    return out


def refresh(ticker, dry_run=False, rebuild=False):
    import openpyxl
    from openpyxl.styles import Font

    path = "%s.xlsx" % ticker
    if not os.path.exists(path):
        print("  %s: no model file" % ticker)
        return 0
    if ticker in SKIP:
        print("  %s: skipped -- %s" % (ticker, SKIP[ticker]))
        return 0

    cik, _ = resolve_cik(ticker)
    if not cik:
        print("  %s: no EDGAR ticker match; add to SKIP if not SEC-registered" % ticker)
        return 0

    fye, edgar_name, filings = all_filings(cik)
    if fye and fye != "1231":
        print("  %s: WARNING fiscal year ends %s, not 12-31. This script maps CALENDAR "
              "quarters; review the results before trusting them." % (ticker, fye))

    wb = openpyxl.load_workbook(path)
    ws = wb["Model"]
    headers = {}
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=2, column=col)
        if cell.value and re.fullmatch(r"Q[1-4]\d{2}", str(cell.value).strip()):
            headers[str(cell.value).strip()] = cell

    if not headers:
        print("  %s: no quarter headers found on the Model sheet row 2" % ticker)
        return 0
    # Follow whatever range the model itself covers rather than assuming one.
    start_year = min(2000 + int(lbl[2:]) for lbl in headers)

    written = 0
    for q, year in quarters_through_today(start_year):
        label = "Q%d%s" % (q, str(year)[2:])
        cell = headers.get(label)
        if cell is None or (cell.hyperlink and not rebuild):
            continue
        hit = best_release(cik, filings, q, year)
        if not hit:
            print("    %s %s: no release found" % (ticker, label))
            continue
        flag = "" if hit["yoy"] else "  (no prior-year comparison in text -- verify)"
        print("    %s %s: %s %s %s%s" % (ticker, label, hit["filed"], hit["form"],
                                         hit["url"].split("/")[-1], flag))
        if not dry_run:
            cell.hyperlink = hit["url"]
            f = cell.font
            cell.font = Font(name=f.name, sz=f.sz, b=f.b, i=f.i, u="single", color="FF0563C1")
        written += 1

    if written and not dry_run:
        wb.save(path)
    print("  %s (%s): %d link(s) %s" % (ticker, edgar_name, written,
                                        "found" if dry_run else "written"))
    return written


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("tickers", nargs="*", help="model tickers, e.g. DUK NEE")
    ap.add_argument("--all", action="store_true", help="every model in the repo")
    ap.add_argument("--dry-run", action="store_true", help="report without writing")
    ap.add_argument("--rebuild", action="store_true", help="replace links already present")
    args = ap.parse_args()

    if not UA or "@" not in UA:
        sys.exit("Set SEC_EDGAR_USER_AGENT to a contact email -- SEC returns 403 without one.")

    tickers = args.tickers
    if args.all:
        tickers = sorted(os.path.basename(p)[:-5] for p in glob.glob("*.xlsx")
                         if os.path.basename(p) not in NOT_A_MODEL
                         and not os.path.basename(p).startswith("~$"))
    if not tickers:
        ap.error("give one or more tickers, or --all")

    total = sum(refresh(t, args.dry_run, args.rebuild) for t in tickers)
    print("\n%d link(s) %s across %d model(s)" % (total, "found" if args.dry_run else "written",
                                                  len(tickers)))


if __name__ == "__main__":
    main()
