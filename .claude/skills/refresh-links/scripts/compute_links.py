"""Work out the links for a model and emit them as JSON, writing no workbook.

For models that openpyxl cannot round-trip safely (see apply_links.ps1), the
links are computed here and applied through Excel instead.
"""
import json, os, re, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import refresh_links as R


def main():
    ticker, out = sys.argv[1], sys.argv[2]
    companies = R.index_companies()
    tmap = R.ticker_map()
    t = R.file_to_ticker(ticker, companies)
    cik, edgar_name = tmap[t]
    fye, _, filings = R.all_filings(cik)
    calendar = (fye or "1231") == "1231"

    import openpyxl
    ws = openpyxl.load_workbook("%s.xlsx" % ticker)["Model"]
    headers = {}
    for c in ws[2]:
        m = re.fullmatch(r"Q([1-4])(\d{2}|\d{4})", str(c.value).strip()) if c.value else None
        if m:
            y = int(m.group(2))
            if y < 100:
                y += 2000 if y < 70 else 1900
            headers[(int(m.group(1)), y)] = (str(c.value).strip(), c.hyperlink is not None)

    results = {}
    for (q, y), (label, already) in sorted(headers.items(), key=lambda kv: (kv[0][1], kv[0][0])):
        if already:
            continue
        hit, err = R.best_release(cik, filings, q, y, calendar, fye)
        if not hit and not err:
            hit = R.best_periodic(cik, filings, q, y)
        if hit:
            results[label] = {"url": hit["url"], "periodic": bool(hit.get("periodic")),
                              "filed": hit["filed"], "form": hit.get("form", "8-K")}
            print("  %s %s: %s%s" % (ticker, label, hit["url"].split("/")[-1],
                                     "  [periodic]" if hit.get("periodic") else ""), flush=True)
        elif err:
            print("  %s %s: LOOKUP ERROR" % (ticker, label), flush=True)

    json.dump(results, open(out, "w"), indent=1)
    print("%s: %d link(s) computed -> %s" % (ticker, len(results), out))


if __name__ == "__main__":
    main()
