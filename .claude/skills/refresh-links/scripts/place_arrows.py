"""Put the actuals/forecast divider arrows where they belong.

Two vertical connectors mark where actuals stop and the forecast begins: one
between Q226 and Q326, one between 2025 and 2026. They are anchored to absolute
column indices, so every column insert leaves them pointing at the wrong place.

Rewrites xl/drawings/drawing1.xml directly rather than through openpyxl, which
drops drawings entirely and cannot write some of these workbooks at all.

    python place_arrows.py            # every model
    python place_arrows.py DUK NVDA   # named models
    python place_arrows.py --dry-run
"""
import glob, os, re, shutil, sys, zipfile
import openpyxl

MODEL_DIR = "companies"                         # models live here; the index does not
QUARTER_ANCHOR = re.compile(r"^Q3(26|2026)$")   # arrow sits at Q326's left edge
YEAR_ANCHOR = "2026"                            # and at 2026's left edge
FIRST_ROW, LAST_ROW = 0, 41

DRAWING_NS = ('xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
              'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"')
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DRAWING_REL = REL_NS + "/drawing"
DRAWING_TYPE = "application/vnd.openxmlformats-officedocument.drawing+xml"


def connector(col, ident):
    """One vertical arrow pinned to the left edge of `col` (0-indexed)."""
    return (
        '<xdr:twoCellAnchor>'
        '<xdr:from><xdr:col>%d</xdr:col><xdr:colOff>25400</xdr:colOff>'
        '<xdr:row>%d</xdr:row><xdr:rowOff>63500</xdr:rowOff></xdr:from>'
        '<xdr:to><xdr:col>%d</xdr:col><xdr:colOff>25400</xdr:colOff>'
        '<xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
        '<xdr:cxnSp macro="">'
        '<xdr:nvCxnSpPr><xdr:cNvPr id="%d" name="Straight Arrow Connector %d"/>'
        '<xdr:cNvCxnSpPr/></xdr:nvCxnSpPr>'
        '<xdr:spPr><a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>'
        '<a:ln><a:tailEnd type="triangle"/></a:ln></xdr:spPr>'
        '<xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>'
        '<a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>'
        '<a:effectRef idx="1"><a:schemeClr val="accent1"/></a:effectRef>'
        '<a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></xdr:style>'
        '</xdr:cxnSp><xdr:clientData/></xdr:twoCellAnchor>'
        % (col, FIRST_ROW, col, LAST_ROW, ident, ident - 1))


def target_columns(path):
    """0-indexed columns for the Q226/Q326 and 2025/2026 boundaries."""
    ws = openpyxl.load_workbook(path)["Model"]
    q = y = None
    for c in ws[2]:
        v = str(c.value).strip() if c.value is not None else ""
        if QUARTER_ANCHOR.fullmatch(v):
            q = c.column - 1
        elif v == YEAR_ANCHOR:
            y = c.column - 1
    return q, y


def model_sheet_part(zf):
    """Which worksheet part is the Model sheet."""
    wb = zf.read("xl/workbook.xml").decode("utf-8", "replace")
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8", "replace")
    m = re.search(r'<sheet[^>]*name="Model"[^>]*?r:id="([^"]+)"', wb)
    if not m:
        return None
    rid = m.group(1)
    # Attribute order varies by writer, and targets may be absolute
    # ("/xl/worksheets/sheet2.xml") or relative ("worksheets/sheet2.xml").
    for rel in re.findall(r"<Relationship\b[^>]*/>", rels):
        if re.search(r'Id="%s"' % re.escape(rid), rel):
            t = re.search(r'Target="([^"]+)"', rel)
            if not t:
                return None
            target = t.group(1).replace("../", "").lstrip("/")
            return target if target.startswith("xl/") else "xl/" + target
    return None


def place(path, dry_run=False):
    q, y = target_columns(path)
    cols = [c for c in (q, y) if c is not None]
    if not cols:
        return "no Q326 or 2026 header"

    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        content = {n: z.read(n) for n in names}
        sheet = model_sheet_part(z)
    if not sheet or sheet not in content:
        return "cannot locate Model sheet part"

    existing = content.get("xl/drawings/drawing1.xml", b"").decode("utf-8", "replace")
    current = [int(v) for v in re.findall(r"<xdr:from><xdr:col>(\d+)</xdr:col>", existing)]
    if current == cols:
        return "already correct"
    if dry_run:
        return "would move %s -> %s" % (current or "none", cols)

    body = "".join(connector(c, i + 3) for i, c in enumerate(cols))
    content["xl/drawings/drawing1.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<xdr:wsDr %s>%s</xdr:wsDr>' % (DRAWING_NS, body)).encode("utf-8")

    # Wire the part up if this workbook never had one.
    ct = content["[Content_Types].xml"].decode("utf-8", "replace")
    if 'PartName="/xl/drawings/drawing1.xml"' not in ct:
        ct = ct.replace("</Types>", '<Override PartName="/xl/drawings/drawing1.xml" '
                                    'ContentType="%s"/></Types>' % DRAWING_TYPE)
        content["[Content_Types].xml"] = ct.encode("utf-8")

    rels_name = sheet.replace("xl/worksheets/", "xl/worksheets/_rels/") + ".rels"
    rels = content.get(rels_name, b"").decode("utf-8", "replace")
    if "drawings/drawing1.xml" not in rels:
        used = [int(n) for n in re.findall(r'Id="rId(\d+)"', rels)]
        rid = "rId%d" % (max(used) + 1 if used else 1)
        entry = ('<Relationship Id="%s" Type="%s" Target="../drawings/drawing1.xml"/>'
                 % (rid, DRAWING_REL))
        if rels.strip():
            rels = rels.replace("</Relationships>", entry + "</Relationships>")
        else:
            rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
                    'relationships">%s</Relationships>' % entry)
        content[rels_name] = rels.encode("utf-8")

        sx = content[sheet].decode("utf-8", "replace")
        if not re.search(r"<drawing\b", sx):
            sx = sx.replace("</worksheet>", '<drawing xmlns:r="%s" r:id="%s"/></worksheet>'
                            % (REL_NS, rid))
            content[sheet] = sx.encode("utf-8")

    tmp = path + ".arrows"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for n, d in content.items():
            out.writestr(n, d)
    shutil.move(tmp, path)
    return "moved %s -> %s" % (current or "none", cols)


def resolve(name):
    """Name -> workbook path. Models live in MODEL_DIR, but the template sits at
    the repo root, so `place_arrows.py base` still finds it."""
    path = "%s/%s.xlsx" % (MODEL_DIR, name)
    return path if os.path.exists(path) else "%s.xlsx" % name


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    dry = "--dry-run" in sys.argv
    targets = [resolve(a) for a in args] if args else [
        "%s/%s" % (MODEL_DIR, os.path.basename(p))
        for p in sorted(glob.glob(MODEL_DIR + "/*.xlsx"))
        if os.path.basename(p) not in ("portfolio.xlsx", "base.xlsx")
        and not os.path.basename(p).startswith("~$")]
    counts = {}
    for t in targets:
        r = place(t, dry)
        counts[r.split(" ")[0]] = counts.get(r.split(" ")[0], 0) + 1
        print("  %-8s %s" % (os.path.basename(t)[:-5], r))
    print("\n" + ", ".join("%s: %d" % kv for kv in sorted(counts.items())))


if __name__ == "__main__":
    main()
