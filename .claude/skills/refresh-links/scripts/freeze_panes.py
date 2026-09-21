"""Freeze the Model sheet so labels and quarter headers stay put while scrolling.

The split goes one column right of the label column and one row below the
quarter headers -- C3, with rows 1-2 and columns A-B frozen. Every model shares
that geometry: labels in column B (`Revenue` at B3 in a skeleton), the filing
row in row 1 and quarter headers across row 2. The five data-bearing models
were already frozen at C3 by hand, which is where the anchor comes from.

Rewrites xl/worksheets/sheetN.xml directly rather than through openpyxl, which
drops drawings and cannot write some of these workbooks at all. Needs nothing
beyond the standard library -- deliberately, since it must never round-trip a
workbook through a library that rebuilds it.

    python freeze_panes.py              # every model
    python freeze_panes.py DUK NVDA     # named models
    python freeze_panes.py base         # the template, so new models inherit it
    python freeze_panes.py --dry-run

Re-running is a no-op: a sheet already frozen below the headers is left alone,
including the horizontal scroll position it was saved with.
"""
import glob, os, re, shutil, sys, zipfile

MODEL_DIR = "companies"                  # models live here; the template does not
PANE = ('<pane xSplit="2" ySplit="2" topLeftCell="%s" activePane="bottomRight" '
        'state="frozen"/>')
ANCHOR_COL, ANCHOR_ROW = "C", 3          # right of the label column, below row 2


def model_part(z):
    """The sheet part backing the Model tab, via the workbook's relationships."""
    book = z.read("xl/workbook.xml").decode("utf-8", "replace")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", "replace")
    by_id = {}
    for rel in re.findall(r"<Relationship\b[^>]*/>", rels):
        rid = re.search(r'Id="(rId\d+)"', rel)
        tgt = re.search(r'Target="([^"]+)"', rel)
        if rid and tgt:
            by_id[rid.group(1)] = "xl/" + tgt.group(1).lstrip("/").replace("xl/", "", 1)
    for name, rid in re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', book):
        if name == "Model":
            return by_id.get(rid)
    return None


def scroll_anchor(view):
    """Keep the sheet's horizontal scroll position, moved below the frozen rows.

    With frozen panes the <pane> element owns the top-left of the scrollable
    region, so a sheetView@topLeftCell of W1 becomes a pane topLeftCell of W3 --
    the models stay scrolled to their recent quarters instead of jumping back
    to the first one.
    """
    m = re.search(r'\btopLeftCell="([A-Z]+)(\d+)"', view)
    if not m:
        return ANCHOR_COL + str(ANCHOR_ROW)
    col, row = m.group(1), int(m.group(2))
    if len(col) < len(ANCHOR_COL) or (len(col) == len(ANCHOR_COL) and col < ANCHOR_COL):
        col = ANCHOR_COL                      # a frozen column cannot lead the pane
    return "%s%d" % (col, max(row, ANCHOR_ROW))


def freeze(sheet_xml):
    """Return (xml, status). Status describes what the sheet needed."""
    m = re.search(r"<sheetViews>.*?</sheetViews>", sheet_xml, re.S)
    if not m:
        view = ('<sheetViews><sheetView workbookViewId="0">%s</sheetView></sheetViews>'
                % (PANE % (ANCHOR_COL + str(ANCHOR_ROW))))
        # sheetViews follows dimension and precedes sheetFormatPr in CT_Worksheet
        for after in (r"</dimension>", r"<dimension\b[^>]*/>", r"</sheetPr>",
                      r"<sheetPr\b[^>]*/>"):
            am = re.search(after, sheet_xml)
            if am:
                return sheet_xml[:am.end()] + view + sheet_xml[am.end():], "added view"
        return sheet_xml.replace("<sheetData", view + "<sheetData", 1), "added view"

    view = m.group(0)
    existing = re.search(r"<pane\b[^>]*/>", view)
    if existing:
        e = existing.group(0)
        top = re.search(r'\btopLeftCell="[A-Z]+(\d+)"', e)
        if (re.search(r'\bxSplit="2"', e) and re.search(r'\bySplit="2"', e)
                and re.search(r'\bstate="frozen"', e)
                and top and int(top.group(1)) >= ANCHOR_ROW):
            return sheet_xml, "already frozen"
        new_view = view.replace(e, PANE % scroll_anchor(view))
        return sheet_xml.replace(view, new_view, 1), "repointed"

    pane = PANE % scroll_anchor(view)
    new_view = re.sub(r'\s*\btopLeftCell="[A-Z]+\d+"', "", view, count=1)
    # <pane> precedes <selection>; the active cell belongs to the scrollable pane
    sel = re.search(r"<selection\b[^>]*/>", new_view)
    if sel and 'pane="' not in sel.group(0):
        new_view = new_view.replace(
            sel.group(0),
            sel.group(0).replace("<selection", '<selection pane="bottomRight"', 1), 1)
    new_view = re.sub(r"(<sheetView\b[^>]*>)", r"\1" + pane, new_view, count=1)
    return sheet_xml.replace(view, new_view, 1), "frozen"


def apply(path, dry_run=False):
    with zipfile.ZipFile(path) as z:
        part = model_part(z)
        if not part:
            return "no Model sheet"
        infos = z.infolist()
        data = {i.filename: z.read(i.filename) for i in infos}
    xml = data[part].decode("utf-8")
    new, status = freeze(xml)
    if new == xml or dry_run:
        return status
    data[part] = new.encode("utf-8")
    tmp = path + ".freeze"
    with zipfile.ZipFile(tmp, "w") as out:
        for i in infos:                       # every other part byte-for-byte
            zi = zipfile.ZipInfo(i.filename, date_time=i.date_time)
            zi.compress_type = i.compress_type
            zi.external_attr = i.external_attr
            zi.internal_attr = i.internal_attr
            zi.create_system = i.create_system
            out.writestr(zi, data[i.filename])
    shutil.move(tmp, path)
    return status


def resolve(name):
    """Name -> workbook path. Models live in MODEL_DIR, but the template sits at
    the repo root, so `freeze_panes.py base` still finds it."""
    path = "%s/%s.xlsx" % (MODEL_DIR, name)
    return path if os.path.exists(path) else "%s.xlsx" % name


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    dry = "--dry-run" in sys.argv
    targets = ([resolve(a) for a in args] if args else
               sorted(p.replace("\\", "/") for p in glob.glob(MODEL_DIR + "/*.xlsx")
                      if not os.path.basename(p).startswith("~$")))
    counts = {}
    for t in targets:
        if not os.path.exists(t):
            print("  %-8s no workbook at %s" % (os.path.basename(t)[:-5], t))
            counts["missing"] = counts.get("missing", 0) + 1
            continue
        r = apply(t, dry)
        counts[r] = counts.get(r, 0) + 1
        print("  %-8s %s" % (os.path.basename(t)[:-5], r))
    print("\n" + ", ".join("%s: %d" % kv for kv in sorted(counts.items())))


if __name__ == "__main__":
    main()
