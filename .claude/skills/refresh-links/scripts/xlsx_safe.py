"""Save an openpyxl workbook without losing zip parts openpyxl does not model.

openpyxl rebuilds the xlsx from its own object model on save, so any part it
does not understand is silently dropped. The models in this repo carry
`xl/drawings/drawing1.xml` -- two arrow connectors on the Model sheet -- and a
plain `wb.save()` deletes them along with the sheet's <drawing> reference.

`save_workbook()` saves through openpyxl, then re-injects anything that went
missing and repairs the references that point at it.

Parts that openpyxl legitimately rewrites are not restored: sharedStrings is
dropped because openpyxl inlines strings, and calcChain is a disposable cache
Excel rebuilds.
"""
import os
import re
import shutil
import zipfile

# Dropped on purpose by openpyxl; restoring these would corrupt the file.
EXPECTED_TO_CHANGE = ("xl/sharedStrings.xml", "xl/calcChain.xml")

CONTENT_TYPES = {
    ".xml":  "application/vnd.openxmlformats-officedocument.drawing+xml",
    ".vml":  "application/vnd.openxmlformats-officedocument.vmlDrawing",
    ".png":  "image/png",
    ".jpeg": "image/jpeg",
    ".jpg":  "image/jpeg",
    ".gif":  "image/gif",
    ".emf":  "image/x-emf",
}
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DRAWING_REL = REL_NS + "/drawing"


def _drawing_owners(zf):
    """Map drawing part -> worksheet part that references it, per the original."""
    owners = {}
    for name in zf.namelist():
        m = re.fullmatch(r"xl/worksheets/_rels/(sheet\d+\.xml)\.rels", name)
        if not m:
            continue
        rels = zf.read(name).decode("utf-8", "replace")
        for rel in re.findall(r"<Relationship\b[^>]*/>", rels):
            if "/relationships/drawing" not in rel:
                continue
            target = re.search(r'Target="([^"]+)"', rel)
            if target:
                part = "xl/" + target.group(1).replace("../", "")
                owners[part] = "xl/worksheets/" + m.group(1)
    return owners


def _next_rel_id(rels_xml):
    used = [int(n) for n in re.findall(r'Id="rId(\d+)"', rels_xml)]
    return "rId%d" % (max(used) + 1 if used else 1)


def _add_relationship(rels_xml, rel_id, rel_type, target):
    entry = '<Relationship Id="%s" Type="%s" Target="%s"/>' % (rel_id, rel_type, target)
    if not rels_xml.strip():
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
                'relationships">%s</Relationships>' % entry)
    return rels_xml.replace("</Relationships>", entry + "</Relationships>")


def _add_drawing_element(sheet_xml, rel_id):
    """Insert <drawing> immediately before </worksheet>.

    That is the correct slot in the CT_Worksheet sequence for the parts these
    models use (it follows pageMargins). The namespace is declared inline
    because openpyxl does not always declare r: on the root element.
    """
    if re.search(r"<drawing\b", sheet_xml):
        return sheet_xml
    element = '<drawing xmlns:r="%s" r:id="%s"/>' % (REL_NS, rel_id)
    return sheet_xml.replace("</worksheet>", element + "</worksheet>")


def _add_content_type(ct_xml, part):
    if 'PartName="/%s"' % part in ct_xml:
        return ct_xml
    ctype = CONTENT_TYPES.get(os.path.splitext(part)[1].lower())
    if not ctype:
        return ct_xml
    override = '<Override PartName="/%s" ContentType="%s"/>' % (part, ctype)
    return ct_xml.replace("</Types>", override + "</Types>")


def restore_parts(original, rebuilt):
    """Re-inject parts of `original` that are absent from `rebuilt`, in place.

    Returns the list of restored part names.
    """
    with zipfile.ZipFile(original) as src:
        src_names = set(src.namelist())
        owners = _drawing_owners(src)
        with zipfile.ZipFile(rebuilt) as dst:
            dst_names = set(dst.namelist())
        missing = sorted(n for n in src_names - dst_names
                         if n not in EXPECTED_TO_CHANGE and not n.endswith("/"))
        if not missing:
            return []

        with zipfile.ZipFile(rebuilt) as dst:
            content = {n: dst.read(n) for n in dst.namelist()}

        for part in missing:
            content[part] = src.read(part)
            ct = content.get("[Content_Types].xml", b"").decode("utf-8", "replace")
            content["[Content_Types].xml"] = _add_content_type(ct, part).encode("utf-8")

            sheet = owners.get(part)
            if not sheet or sheet not in content:
                continue
            rels_name = sheet.replace("xl/worksheets/", "xl/worksheets/_rels/") + ".rels"
            rels = content.get(rels_name, b"").decode("utf-8", "replace")
            rel_id = _next_rel_id(rels)
            target = "../" + part.replace("xl/", "")
            content[rels_name] = _add_relationship(rels, rel_id, DRAWING_REL, target).encode("utf-8")
            sheet_xml = content[sheet].decode("utf-8", "replace")
            content[sheet] = _add_drawing_element(sheet_xml, rel_id).encode("utf-8")

    tmp = rebuilt + ".parts"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for name, data in content.items():
            out.writestr(name, data)
    shutil.move(tmp, rebuilt)
    return missing


def save_workbook(wb, path):
    """openpyxl save that preserves parts openpyxl would drop.

    Refuses to write if the workbook is open in Excel, which would otherwise
    raise PermissionError partway through a long batch.
    """
    lock = os.path.join(os.path.dirname(os.path.abspath(path)),
                        "~$" + os.path.basename(path))
    if os.path.exists(lock):
        raise PermissionError("%s appears to be open in Excel (%s)" %
                              (path, os.path.basename(lock)))

    backup = path + ".orig"
    shutil.copy2(path, backup)
    try:
        wb.save(path)
        restored = restore_parts(backup, path)
    except Exception:
        shutil.move(backup, path)        # leave the file exactly as it was
        raise
    os.remove(backup)
    return restored
