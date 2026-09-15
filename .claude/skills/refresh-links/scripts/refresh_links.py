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
import argparse, datetime, difflib, glob, gzip, json, os, re, sys, time, urllib.error, urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlsx_safe

UA = os.environ.get("SEC_EDGAR_USER_AGENT", "").strip()
CACHE_FILE = ".refresh_links_cache.json"
INDEX_WORKBOOK = "Utilities.xlsx"

# Tickers whose EDGAR entry is a different company, or that are not SEC
# registrants. Verified against the index workbook at run time as well, but
# these are recorded so the reason survives.
SKIP = {
    "H":     "Hydro One (TSX) -- EDGAR ticker 'H' is Hyatt Hotels",
    "DIA":   "Dialight plc (LSE) -- EDGAR ticker 'DIA' is the SPDR Dow Jones ETF",
    "SUN":   "Stardust Solar (TSXV) -- EDGAR ticker 'SUN' is Sunoco LP",
    "SPCX":  "SpaceX -- private; files no earnings releases",
    "CMNR":  "Commerce Energy Group -- last EDGAR filing 2009",
    "APTL":  "Alaska Power & Telephone -- last EDGAR filing 2005",
    "SUME":  "Summer Energy Holdings -- no EDGAR ticker mapping",
    "DYNA":  "Dynastar Holdings -- no EDGAR ticker mapping",
    "SBGSY": "Schneider Electric ADR -- files in France, not with the SEC",
}
NOT_A_MODEL = {"Utilities.xlsx", "base model.xlsx"}

ORD = {"first": 1, "second": 2, "third": 3, "fourth": 4}
MONTHS = {m: i + 1 for i, m in enumerate(
    ["january", "february", "march", "april", "may", "june", "july",
     "august", "september", "october", "november", "december"])}
MONTH_END = {"march 31": 1, "june 30": 2, "september 30": 3, "december 31": 4}
QUARTER_END = {1: (3, 31), 2: (6, 30), 3: (9, 30), 4: (12, 31)}
# Exhibits filed alongside the release that are not the release itself. Matched
# as regexes against the start of the document, so keep them specific -- a bare
# "management" would reject real releases that mention a management call.
NOT_A_RELEASE = (r"management.{0,8}s discussion", r"unaudited condensed",
                 r"certification of interim", r"earnings coverage ratio",
                 r"form 52-109", r"consent of independent")
# Filed under Item 2.02 but not the quarterly results: segment restatements,
# call advisories, dividend declarations, offering notices. These mention the
# quarter and would otherwise score as a release -- Rockwell's "Provides
# Historical Data for new Operating Segments" is a real example that was
# picked for Q320 before this filter existed.
NOT_QUARTERLY_RESULTS = (r"provides historical data",
                         r"\bto announce\b",
                         r"schedules?\b[^.]{0,40}\b(call|conference)",
                         r"announces change", r"declares[^.]{0,30}dividend",
                         r"pricing of", r"prices[^.]{0,30}offering",
                         r"commences[^.]{0,30}offering")
# Headline language that does mark a results document. "Shareholder letter" is
# included because it is how Oklo, Bloom and other recent listings publish
# quarterly results.
IS_RESULTS = (r"financial results", r"results for", r"\bresults\b", r"earnings",
              r"shareholder letter", r"interim report", r"quarterly report",
              r"half[- ]year", r"year[- ]end")
# An unmistakable results headline. A release may also announce something else
# in the same breath -- Rockwell's Q3 FY20 reports results AND announces new
# operating segments -- so a strong headline outranks the soft disqualifiers.
STRONG_RESULTS = (r"reports[^.]{0,80}results", r"reports[^.]{0,60}earnings",
                  r"announces[^.]{0,60}financial results", r"shareholder letter",
                  r"year[- ]end earnings", r"interim report")


class Throttled(Exception):
    """SEC is refusing requests. Fatal: blanks after this point are meaningless."""


class FetchFailed(Exception):
    """A request failed for a reason that is not a definitive 'absent'."""


def fetch(url, cap=250000, tries=3):
    """Return the body, or raise. Never returns '' to mean failure -- a caller
    that cannot tell a throttle from an empty document will silently record
    'no release exists' for an entire run."""
    last = None
    for attempt in range(tries):
        try:
            req = urllib.request.Request(url, headers={
                "User-Agent": UA, "Accept-Encoding": "gzip", "Range": "bytes=0-%d" % cap})
            with urllib.request.urlopen(req, timeout=60) as resp:
                raw = resp.read()
                if resp.headers.get("Content-Encoding") == "gzip":
                    try:
                        raw = gzip.decompress(raw)
                    except Exception:
                        # A range-truncated gzip member cannot be decoded; decoding
                        # the raw bytes would yield mojibake that fails every regex.
                        raise FetchFailed("truncated gzip from %s" % url)
                return raw.decode("utf-8", "replace")
        except urllib.error.HTTPError as e:
            if e.code in (403, 429):
                raise Throttled("SEC returned %d for %s -- stopping. Check "
                                "SEC_EDGAR_USER_AGENT and back off." % (e.code, url))
            if e.code == 404:
                raise FetchFailed("404 %s" % url)
            last = e
        except Exception as e:
            last = e
        if attempt < tries - 1:
            time.sleep(1.5 * (attempt + 1))          # backoff, not a flat retry
        time.sleep(0.11)                              # SEC allows 10 req/sec
    raise FetchFailed("%s after %d tries: %s" % (url, tries, last))


def plain(html):
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html)).strip()


def nearest_quarter(date):
    """Calendar quarter whose end is nearest `date`.

    This is how an off-cycle fiscal quarter is placed on calendar headers:
    NVDA's quarter ending 3 May 2020 sits nearest 31 Mar, so it is Q1 2020.
    """
    best = None
    for year in (date.year - 1, date.year, date.year + 1):
        for q, (m, d) in QUARTER_END.items():
            delta = abs((datetime.date(year, m, d) - date).days)
            if best is None or delta < best[0]:
                best = (delta, q, year)
    return best[1], best[2]


def _dates_in(text):
    for m in re.finditer(r"(january|february|march|april|may|june|july|august|"
                         r"september|october|november|december)\s+(\d{1,2}),?\s+(\d{4})", text):
        try:
            yield datetime.date(int(m.group(3)), MONTHS[m.group(1)], int(m.group(2)))
        except ValueError:
            continue


def fiscal_period_end(fye, n, fiscal_year):
    """End date of fiscal quarter `n` of `fiscal_year`, from a MMDD year end.

    Cisco's releases carry no period end dates at all -- only "third quarter of
    fiscal 2020" -- so the date has to be derived from the fiscal calendar.
    Q4 ends at the fiscal year end and each earlier quarter three months before.
    """
    import calendar
    month, day = int(fye[:2]), int(fye[2:])
    month -= (4 - n) * 3
    year = fiscal_year
    while month <= 0:
        month += 12
        year -= 1
    return datetime.date(year, month, min(day, calendar.monthrange(year, month)[1]))


def periods(text, calendar_filer=True, fye=None):
    """Calendar (quarter, year) pairs the release reports.

    For a calendar-year filer the company's own quarter labels are the calendar
    quarters. For an off-cycle filer they are not -- "first quarter fiscal 2021"
    may be calendar Q1 2020 -- so only explicit period end dates are trusted,
    mapped to the nearest calendar quarter.
    """
    t = text.lower()
    found = set()

    for m in re.finditer(r"(?:three|thirteen) months ended\s+([a-z]+\s+\d{1,2},?\s+\d{4})", t):
        for d in _dates_in(m.group(1)):
            found.add(nearest_quarter(d))
    for m in re.finditer(r"quarter(?:ly period)? ended\s+([a-z]+\s+\d{1,2},?\s+\d{4})", t):
        for d in _dates_in(m.group(1)):
            found.add(nearest_quarter(d))

    if not calendar_filer:
        # Off-cycle filers often state only a fiscal label. Convert it via the
        # fiscal calendar; the prior-year label these releases also carry is
        # what supplies the year-over-year signal.
        if fye:
            for m in re.finditer(r"(first|second|third|fourth) quarter (?:of |ended )?"
                                 r"fiscal(?: year)? (\d{4})", t):
                found.add(nearest_quarter(fiscal_period_end(fye, ORD[m.group(1)],
                                                            int(m.group(2)))))
            for m in re.finditer(r"\bq([1-4]) fy ?(\d{4})", t):
                found.add(nearest_quarter(fiscal_period_end(fye, int(m.group(1)),
                                                            int(m.group(2)))))
            for m in re.finditer(r"fiscal(?: year)? (\d{4}) (first|second|third|fourth) quarter", t):
                found.add(nearest_quarter(fiscal_period_end(fye, ORD[m.group(2)],
                                                            int(m.group(1)))))
        return found

    for m in re.finditer(r"(first|second|third|fourth)[- ]quarter(?:\s+(?:of\s+|ended\s+)?(\d{4}))?", t):
        if m.group(2):
            found.add((ORD[m.group(1)], int(m.group(2))))
    for m in re.finditer(r"(\d{4}) (first|second|third|fourth)[- ]quarter", t):
        found.add((ORD[m.group(2)], int(m.group(1))))
    for m in re.finditer(r"three months ended (march 31|june 30|september 30|december 31),? (\d{4})", t):
        found.add((MONTH_END[m.group(1)], int(m.group(2))))
    for m in re.finditer(r"(?:year|fourth quarter) ended december 31,? (\d{4})", t):
        found.add((4, int(m.group(1))))
    # "fourth quarter and full year 2020" -- the year does not sit directly
    # after "quarter", so the patterns above miss it.
    for m in re.finditer(r"(first|second|third|fourth)[- ]quarter and (?:the )?"
                         r"(?:full[- ]|fiscal[- ])?year (?:ended [a-z]+ \d{1,2},? )?(\d{4})", t):
        found.add((ORD[m.group(1)], int(m.group(2))))
    for m in re.finditer(r"q([1-4])\s+(\d{4})", t):
        found.add((int(m.group(1)), int(m.group(2))))
    # "Vistra Reports Full-Year 2023 Results" -- a Q4 release that never says
    # "fourth quarter". Scoped to the headline verb so that "full-year 2024
    # guidance" inside a Q1 release does not register as Q4.
    for m in re.finditer(r"(?:reports|announces|delivers)[^.]{0,40}full[- ]year "
                         r"(\d{4})\s+(?:results|earnings|financial)", t):
        found.add((4, int(m.group(1))))
    for m in re.finditer(r"(\d{4}) year[- ]end earnings", t):
        found.add((4, int(m.group(1))))
    return found


def looks_like_release(text):
    """Reject the MD&A, interim financials and certifications filed alongside.

    Deliberately permissive on the positive side: small-cap and foreign
    headlines say "Interim Report" or "Financial Results for...", not
    "Reports". Period matching does the real work of picking the right one.
    """
    # Disqualifiers are judged on the headline region only, so an incidental
    # later mention does not reject a real release. The positive test scans
    # further in: letterhead and address blocks routinely push the headline
    # past 600 characters (Rockwell's releases open with a Milwaukee address).
    head = text.lower()[:800]
    body = text.lower()[:3000]
    if any(re.search(p, head) for p in NOT_A_RELEASE):
        return False
    if any(re.search(p, head) for p in STRONG_RESULTS):
        return True
    if any(re.search(p, head) for p in NOT_QUARTERLY_RESULTS):
        return False
    # A call advisory announces a future release: short, no financial tables.
    # The real release says "will host a conference call" too, so the verb
    # alone cannot separate them -- length does.
    if len(text) < 3000 and re.search(r"will (announce|report)|conference call|webcast", head):
        return False
    return any(re.search(p, body) for p in IS_RESULTS)


def reporting_window(q, year):
    """Calendar window in which quarter (q, year) may be reported.

    Deliberately generous: a non-accelerated filer has 90 days for a 10-K and
    routinely publishes full-year results in April. Overlap between windows is
    harmless because the release text, not the date, decides the quarter.
    """
    return {1: ("%d-04-01" % year, "%d-07-31" % year),
            2: ("%d-07-01" % year, "%d-10-31" % year),
            3: ("%d-10-01" % year, "%d-01-31" % (year + 1)),
            4: ("%d-01-01" % (year + 1), "%d-04-30" % (year + 1))}[q]


def lag_days(q, year, filed):
    month, day = QUARTER_END[q]
    filed_on = datetime.date(int(filed[:4]), int(filed[5:7]), int(filed[8:10]))
    return (filed_on - datetime.date(year, month, day)).days


_TICKER_MAP = None


def ticker_map():
    global _TICKER_MAP
    if _TICKER_MAP is None:                       # ~1.5MB; fetch once per run
        data = json.loads(fetch("https://www.sec.gov/files/company_tickers.json", cap=6000000))
        _TICKER_MAP = {row["ticker"].upper(): (row["cik_str"], row["title"])
                       for row in data.values()}
    return _TICKER_MAP


def normalize_name(s):
    s = re.sub(r"\(.*?\)", "", s).lower()
    s = re.sub(r"\b(inc|corp|corporation|company|co|ltd|limited|plc|holdings|group|"
               r"the|lp|sa|nv|ag|incorporated|technologies|energy|power)\b", "", s)
    return re.sub(r"[^a-z0-9]", "", s)


def names_match(sheet_name, edgar_name):
    a, b = normalize_name(sheet_name), normalize_name(edgar_name)
    if not a or not b:
        return False
    return a in b or b in a or difflib.SequenceMatcher(None, a, b).ratio() >= 0.6


def index_companies():
    """ticker -> company name, from the index workbook's Main sheet."""
    import openpyxl
    if not os.path.exists(INDEX_WORKBOOK):
        return {}
    ws = openpyxl.load_workbook(INDEX_WORKBOOK, read_only=True)["Main"]
    out = {}
    for row in ws.iter_rows(min_col=2, max_col=3, values_only=True):
        company, ticker = row[0], row[1]
        if company and ticker:
            out[str(ticker).strip().upper()] = str(company)
    return out


def all_filings(cik):
    sub = json.loads(fetch("https://data.sec.gov/submissions/CIK%010d.json" % cik, cap=20000000))
    rows = []

    def absorb(block):
        for i in range(len(block["accessionNumber"])):
            rows.append({k: block[k][i] for k in
                         ("accessionNumber", "filingDate", "form", "items")})

    absorb(sub["filings"]["recent"])
    for extra in sub["filings"].get("files", []):
        # Older filings are paginated into separate files; these are already
        # bare {field: [...]} blocks, not wrapped in "filings".
        absorb(json.loads(fetch("https://data.sec.gov/submissions/" + extra["name"],
                                cap=20000000)))
    return sub.get("fiscalYearEnd"), sub.get("name"), rows


def exhibits(cik, accession):
    nodash = accession.replace("-", "")
    html = fetch("https://www.sec.gov/Archives/edgar/data/%d/%s/%s-index.htm"
                 % (cik, nodash, accession))
    out = []
    for row in re.findall(r"<tr[^>]*>(.*?)</tr>", html, re.S | re.I):
        cells = [re.sub(r"<[^>]+>", "", c).replace("&nbsp;", " ").strip()
                 for c in re.findall(r"<td[^>]*>(.*?)</td>", row, re.S | re.I)]
        href = re.search(r'href="([^"]+)"', row)
        if len(cells) >= 4 and href and cells[3].upper().startswith("EX-99"):
            out.append({"type": cells[3],
                        "url": "https://www.sec.gov" + href.group(1).replace("/ix?doc=", "")})
    return out


def best_release(cik, filings, q, year, calendar_filer=True, fye=None):
    """(best match or None, whether any lookup failed).

    The second value matters: without it a network failure is indistinguishable
    from a company that never reported, and gets cached as a permanent absence.
    """
    low, high = reporting_window(q, year)
    candidates = [f for f in filings
                  if low <= f["filingDate"] <= high
                  and (f["form"].startswith("8-K") or f["form"].startswith("6-K"))]
    # Earnings 8-Ks are usually tagged 2.02, but not always -- some are filed
    # under 9.01 only, and some arrive as 8-K/A. Try the tagged ones first.
    candidates.sort(key=lambda f: ("2.02" not in (f["items"] or ""), f["filingDate"]))
    scored, had_error = [], False
    for f in candidates:
        try:
            found = exhibits(cik, f["accessionNumber"])
        except FetchFailed:
            had_error = True
            continue
        for ex in found:
            try:
                text = plain(fetch(ex["url"]))
            except FetchFailed:
                had_error = True
                continue
            if not looks_like_release(text):
                continue
            claimed = periods(text, calendar_filer, fye)
            if (q, year) not in claimed:
                continue
            # A real earnings release nearly always cites the prior-year quarter
            # for comparison; that is the strongest signal available.
            yoy = (q, year - 1) in claimed
            ideal = 55 if q == 4 else 40
            scored.append({
                "score": (100 if yoy else 0) - abs(lag_days(q, year, f["filingDate"]) - ideal) * 0.2,
                "url": ex["url"], "filed": f["filingDate"], "form": f["form"], "yoy": yoy})
            break
    scored.sort(key=lambda c: -c["score"])
    return (scored[0] if scored else None), had_error


def quarters_through_today(start_year):
    """Every quarter from start_year that has closed and had time to be reported."""
    cutoff = datetime.date.today() - datetime.timedelta(days=25)
    out = []
    for year in range(start_year, datetime.date.today().year + 1):
        for q in (1, 2, 3, 4):
            month, day = QUARTER_END[q]
            if datetime.date(year, month, day) < cutoff:
                out.append((q, year))
    return out


def load_cache(ignore=False):
    if ignore or not os.path.exists(CACHE_FILE):
        return {}
    try:
        return json.load(open(CACHE_FILE))
    except Exception:
        return {}


def save_cache(cache):
    tmp = CACHE_FILE + ".tmp"
    json.dump(cache, open(tmp, "w"), indent=0)
    os.replace(tmp, CACHE_FILE)


def file_to_ticker(stem, known):
    """Model filenames sanitize dots: BEP_UN.xlsx is ticker BEP.UN."""
    if stem.upper() in known or "_" not in stem:
        return stem.upper()
    dotted = stem.replace("_", ".").upper()
    return dotted if dotted in known else stem.upper()


def refresh(stem, companies, cache, dry_run=False, rebuild=False):
    import openpyxl
    from openpyxl.styles import Font

    path = "%s.xlsx" % stem
    if not os.path.exists(path):
        print("  %s: no model file" % stem); return 0, 0
    ticker = file_to_ticker(stem, companies)
    if ticker in SKIP:
        print("  %s: skipped -- %s" % (ticker, SKIP[ticker])); return 0, 0

    tmap = ticker_map()
    if ticker not in tmap:
        print("  %s: not SEC-registered (no EDGAR ticker)" % ticker); return 0, 0
    cik, edgar_name = tmap[ticker]

    # Guard against ticker collisions -- EDGAR's 'H' is Hyatt, not Hydro One.
    expected = companies.get(ticker)
    if expected and not names_match(expected, edgar_name):
        print("  %s: REFUSED -- sheet says %r, EDGAR ticker is %r"
              % (ticker, expected[:40], edgar_name[:40]))
        return 0, 0

    fye, _, filings = all_filings(cik)
    calendar_filer = (fye or "1231") == "1231"
    if not calendar_filer:
        print("  %s: fiscal year ends %s -- calendarizing by period end date" % (ticker, fye))

    wb = openpyxl.load_workbook(path)
    ws = wb["Model"]
    # Headers are written either Q120 or Q12024 depending on the model's
    # vintage, so key on the parsed quarter rather than the label text.
    headers = {}
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=2, column=col)
        m = re.fullmatch(r"Q([1-4])(\d{2}|\d{4})", str(cell.value).strip()) if cell.value else None
        if m:
            year = int(m.group(2))
            headers[(int(m.group(1)), year + 2000 if year < 100 else year)] = cell
    if not headers:
        print("  %s: no quarter headers on Model row 2" % ticker); return 0, 0
    start_year = min(y for _, y in headers)

    written = errors = 0
    pending = False
    for q, year in quarters_through_today(start_year):
        label = "Q%d%s" % (q, str(year)[2:])
        cell = headers.get((q, year))
        if cell is None or (cell.hyperlink and not rebuild):
            continue
        key = "%s|%s" % (ticker, label)
        if cache.get(key) == "none" and not rebuild:
            continue
        hit, had_error = best_release(cik, filings, q, year, calendar_filer, fye)
        if not hit:
            if had_error:
                print("    %s %s: LOOKUP ERROR (not recorded as absent)" % (ticker, label))
                errors += 1
            else:
                cache[key] = "none"
            continue
        flag = "" if hit["yoy"] else "  (no prior-year comparison -- verify)"
        print("    %s %s: %s %s %s%s" % (ticker, label, hit["filed"], hit["form"],
                                         hit["url"].split("/")[-1], flag))
        if not dry_run:
            cell.hyperlink = hit["url"]
            f = cell.font
            cell.font = Font(name=f.name, sz=f.sz, b=f.b, i=f.i, u="single", color="FF0563C1")
            pending = True
        written += 1
        # Save as we go: a multi-hour run will be interrupted, and saving only
        # at the end would discard every link found for this company.
        if pending and written % 4 == 0:
            xlsx_safe.save_workbook(wb, path); pending = False

    if pending and not dry_run:
        xlsx_safe.save_workbook(wb, path)
    print("  %s (%s): %d link(s) %s%s"
          % (ticker, edgar_name[:38], written, "found" if dry_run else "written",
             ", %d lookup error(s)" % errors if errors else ""))
    return written, errors


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("tickers", nargs="*", help="model tickers, e.g. DUK NEE")
    ap.add_argument("--all", action="store_true", help="every model in the repo")
    ap.add_argument("--dry-run", action="store_true", help="report without writing")
    ap.add_argument("--rebuild", action="store_true", help="replace links already present")
    ap.add_argument("--ignore-cache", action="store_true", help="re-search known-empty quarters")
    args = ap.parse_args()

    if not UA or "@" not in UA:
        sys.exit("Set SEC_EDGAR_USER_AGENT to a contact email -- SEC returns 403 without one.")

    stems = args.tickers
    if args.all:
        stems = sorted(os.path.basename(p)[:-5] for p in glob.glob("*.xlsx")
                       if os.path.basename(p) not in NOT_A_MODEL
                       and not os.path.basename(p).startswith("~$"))
    if not stems:
        ap.error("give one or more tickers, or --all")

    companies = index_companies()
    cache = load_cache(args.ignore_cache)
    total = total_err = failed = 0
    try:
        for i, stem in enumerate(stems, 1):
            print("[%d/%d]" % (i, len(stems)), flush=True)
            try:
                n, e = refresh(stem, companies, cache, args.dry_run, args.rebuild)
                total += n; total_err += e
            except Throttled:
                raise
            except Exception as exc:              # one bad company must not end the run
                failed += 1
                print("  %s: FAILED -- %s" % (stem, exc))
            if not args.dry_run:
                save_cache(cache)
    except Throttled as t:
        save_cache(cache)
        sys.exit("\nABORTED: %s\nLinks written so far are saved; re-run to resume." % t)

    if not args.dry_run:
        save_cache(cache)
    print("\n%d link(s) %s across %d model(s)" % (total, "found" if args.dry_run else "written",
                                                  len(stems)))
    if total_err or failed:
        print("%d lookup error(s), %d model(s) failed -- those quarters are NOT "
              "recorded as empty; re-run to retry." % (total_err, failed))


if __name__ == "__main__":
    main()
