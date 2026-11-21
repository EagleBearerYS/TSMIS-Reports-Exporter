"""Lite copies of comparison workbooks — the deliverable version (v0.55.0).

The owner's 10.30 deliverables were the app's values workbooks lightened by
hand: the Spot Check, Notes and Provenance sheets deleted, and the Summary cut
to its title ("<A> vs <B> — <report>"), the ROW COUNTS, MATCH STATUS and ROUTE
COVERAGE blocks and the DIFFERENCES BY FIELD table. `write_lite_copy` makes
exactly that from a finished values workbook:

  * every part it keeps is copied from the values workbook byte for byte (its
    compressed bytes, never re-encoded), so a lite copy cannot disagree with the
    workbook it came from, and writing one takes seconds even for the biggest
    comparison;
  * it rewrites only four small parts: the Summary sheet, the workbook part
    (its sheet list and the sheet positions of the filter names), the
    workbook's relationships and the content-types list.

It reads compare_core's own Summary layout and refuses anything else with a
ValueError, so a Summary that changed shape can never produce a wrong lite
copy quietly. The matrices write one beside each values workbook when their
"Also write a lite copy" option is on (`matrix_build._settle_lite_copy`).
"""
import os
import posixpath
import re
import shutil
import struct
import zipfile
import zlib
from pathlib import Path
from xml.dom import minidom

LITE_SUFFIX = " (lite)"
SUMMARY_SHEET = "Summary"
DROPPED_SHEETS = ("Spot Check", "Notes", "Provenance")
TITLE_CUT = " — Discrepancy Report"

# compare_core._write_summary's banners (column B). The lite Summary keeps the
# title, every row from ROW COUNTS up to the FIELD-LEVEL DISCREPANCIES banner,
# one blank row, then the DIFFERENCES BY FIELD table to its last field row.
KEEP_FROM = "ROW COUNTS"
KEEP_UNTIL = "FIELD-LEVEL DISCREPANCIES"
BY_FIELD = "DIFFERENCES BY FIELD"

_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_CALC_CHAIN = _DOC_REL + "/calcChain"
_SHARED_STRINGS = _DOC_REL + "/sharedStrings"
_OFFICE_DOCUMENT = _DOC_REL + "/officeDocument"
_CONTENT_TYPES = "[Content_Types].xml"

# Summary parts a row renumbering would have to move; compare_core writes none
# of them, so meeting one means the layout is not the one this module knows.
_UNHANDLED_SUMMARY_PARTS = ("mergeCells", "hyperlinks", "dataValidations",
                            "drawing", "legacyDrawing", "tableParts",
                            "rowBreaks")

_LOCAL_HEADER = struct.Struct("<4s2B4HL2L2H")
_CENTRAL_HEADER = struct.Struct("<4s4B4HL2L5H2L")
_END_RECORD = struct.Struct("<4s4H2LH")
_LOCAL_SIGNATURE = b"PK\x03\x04"
_CENTRAL_SIGNATURE = b"PK\x01\x02"
_END_SIGNATURE = b"PK\x05\x06"
_MAX32 = 0xFFFFFFFF
_MAX16 = 0xFFFF
_ENCRYPTED = 0x01
_DATA_DESCRIPTOR = 0x08
_UTF8_NAME = 0x800
_CHUNK = 1 << 20


def lite_sibling(values_path):
    """'<name> (lite).xlsx' beside a values workbook."""
    p = Path(values_path)
    return p.with_name(f"{p.stem}{LITE_SUFFIX}{p.suffix}")


def lite_title(title):
    """The Summary title without its 'Discrepancy Report (…)' tail."""
    if TITLE_CUT in title:
        return title.split(TITLE_CUT, 1)[0].rstrip()
    return title


# --------------------------------------------------------------------------- #
# small DOM helpers
# --------------------------------------------------------------------------- #
def _children(node, local):
    return [n for n in node.childNodes
            if n.nodeType == n.ELEMENT_NODE and n.localName == local]


def _child(node, local):
    found = _children(node, local)
    return found[0] if found else None


def _text(node):
    return "".join(n.data for n in node.childNodes
                   if n.nodeType in (n.TEXT_NODE, n.CDATA_SECTION_NODE))


def _qualified(like, local):
    """`local` with the same prefix as the element `like`."""
    return f"{like.prefix}:{local}" if like.prefix else local


def _column(ref):
    m = re.match(r"([A-Z]+)(\d+)$", ref or "")
    if not m:
        raise ValueError(f"unexpected cell reference {ref!r} on the Summary sheet")
    return m.group(1)


def _cell_text(cell, shared):
    kind = cell.getAttribute("t")
    if kind == "inlineStr":
        return "".join(_text(t) for t in cell.getElementsByTagNameNS(_MAIN, "t"))
    v = _child(cell, "v")
    value = _text(v) if v is not None else ""
    if kind == "s":
        if shared is None:
            raise ValueError("the Summary uses shared strings but the workbook has none")
        return shared[int(value)]
    return value


def _row_number(row):
    try:
        return int(row.getAttribute("r"))
    except ValueError:
        raise ValueError("a Summary row has no row number") from None


def _row_has_content(row, shared):
    return any(_cell_text(c, shared) for c in _children(row, "c"))


def _column_b_text(row, shared):
    for cell in _children(row, "c"):
        if _column(cell.getAttribute("r")) == "B":
            return _cell_text(cell, shared)
    return ""


# --------------------------------------------------------------------------- #
# the four rewritten parts
# --------------------------------------------------------------------------- #
def lite_summary(xml_bytes, shared=None):
    """The Summary sheet XML cut to the lite layout. Raises ValueError for a
    Summary that is not compare_core's values layout."""
    doc = minidom.parseString(xml_bytes)
    ws = doc.documentElement
    for local in _UNHANDLED_SUMMARY_PARTS:
        if _child(ws, local) is not None:
            raise ValueError(f"the Summary sheet has {local}, which a lite copy "
                             "does not handle")
    data = _child(ws, "sheetData")
    if data is None:
        raise ValueError("the Summary sheet has no rows")
    rows = _children(data, "row")
    texts = [_column_b_text(r, shared) for r in rows]

    def find(label, test):
        for i, t in enumerate(texts):
            if t and test(t):
                return i
        raise ValueError(f"the Summary has no {label} block")

    title_i = find("title", lambda t: True)
    counts_i = find(KEEP_FROM, lambda t: t == KEEP_FROM)
    until_i = find(KEEP_UNTIL, lambda t: t.startswith(KEEP_UNTIL))
    by_field_i = find(BY_FIELD, lambda t: t == BY_FIELD)
    if not title_i < counts_i < until_i < by_field_i:
        raise ValueError("the Summary blocks are not in the expected order")
    end_i = by_field_i
    while (end_i + 1 < len(rows)
           and _row_number(rows[end_i + 1]) == _row_number(rows[end_i]) + 1
           and _row_has_content(rows[end_i + 1], shared)):
        end_i += 1
    if end_i < by_field_i + 2:
        raise ValueError("the Summary's DIFFERENCES BY FIELD table is empty")

    numbering, n = {}, 1
    for i in [title_i, *range(counts_i, until_i)]:
        numbering[i], n = n, n + 1
    n += 1                                   # the blank row before the field table
    for i in range(by_field_i, end_i + 1):
        numbering[i], n = n, n + 1
    last = n - 1

    for i, row in enumerate(rows):
        if i not in numbering:
            data.removeChild(row)
            continue
        new = numbering[i]
        row.setAttribute("r", str(new))
        for cell in _children(row, "c"):
            cell.setAttribute("r", f"{_column(cell.getAttribute('r'))}{new}")

    title_cell = next(c for c in _children(rows[title_i], "c")
                      if _column(c.getAttribute("r")) == "B")
    while title_cell.firstChild is not None:
        title_cell.removeChild(title_cell.firstChild)
    title_cell.setAttribute("t", "inlineStr")
    inline = doc.createElementNS(_MAIN, _qualified(title_cell, "is"))
    t = doc.createElementNS(_MAIN, _qualified(title_cell, "t"))
    t.appendChild(doc.createTextNode(lite_title(texts[title_i])))
    inline.appendChild(t)
    title_cell.appendChild(inline)

    # The only conditional format compare_core puts on a Summary colours the
    # verdict line, which the lite copy leaves out.
    for cf in _children(ws, "conditionalFormatting"):
        ws.removeChild(cf)
    dimension = _child(ws, "dimension")
    if dimension is not None:
        m = re.match(r"([A-Z]+\d+):([A-Z]+)\d+$", dimension.getAttribute("ref"))
        if m:
            dimension.setAttribute("ref", f"{m.group(1)}:{m.group(2)}{last}")
    for selection in ws.getElementsByTagNameNS(_MAIN, "selection"):
        selection.setAttribute("activeCell", "A1")
        selection.setAttribute("sqref", "A1")
    return doc.toxml(encoding="UTF-8")


def _names_sheet(formula, sheet):
    """Does a defined name's formula point at `sheet`?"""
    quoted = "'" + sheet.replace("'", "''") + "'!"
    return quoted in formula or re.search(
        r"(?<![\w.'])" + re.escape(sheet) + "!", formula) is not None


def lite_workbook(xml_bytes, dropped):
    """The workbook part without the `dropped` sheets: their <sheet> entries go,
    the filter names keep pointing at their own sheets, and the active tab stays
    on a sheet that is still there. Returns (xml, the dropped sheets' r:ids)."""
    doc = minidom.parseString(xml_bytes)
    wb = doc.documentElement
    sheets_el = _child(wb, "sheets")
    position, kept, rids = {}, 0, []
    for i, sheet in enumerate(_children(sheets_el, "sheet")):
        if sheet.getAttribute("name") in dropped:
            rids.append(sheet.getAttributeNS(_DOC_REL, "id"))
            sheets_el.removeChild(sheet)
        else:
            position[i] = kept
            kept += 1
    names = _child(wb, "definedNames")
    if names is not None:
        for name in _children(names, "definedName"):
            local = name.getAttribute("localSheetId")
            if local != "":
                if int(local) not in position:
                    names.removeChild(name)
                    continue
                name.setAttribute("localSheetId", str(position[int(local)]))
            if any(_names_sheet(_text(name), s) for s in dropped):
                names.removeChild(name)
        if not _children(names, "definedName"):
            wb.removeChild(names)
    views = _child(wb, "bookViews")
    for view in (_children(views, "workbookView") if views is not None else ()):
        for attr in ("activeTab", "firstSheet"):
            if view.hasAttribute(attr):
                view.setAttribute(attr, str(position.get(int(view.getAttribute(attr)), 0)))
    return doc.toxml(encoding="UTF-8"), rids


def _relationships(xml_bytes, base):
    """[(Id, Type, part name or None for an external target)] of a .rels part."""
    doc = minidom.parseString(xml_bytes)
    out = []
    for rel in _children(doc.documentElement, "Relationship"):
        target = rel.getAttribute("Target")
        if rel.getAttribute("TargetMode") == "External":
            part = None
        elif target.startswith("/"):
            part = target.lstrip("/")
        else:
            part = posixpath.normpath(posixpath.join(base, target))
        out.append((rel.getAttribute("Id"), rel.getAttribute("Type"), part))
    return out


def lite_rels(xml_bytes, dropped_ids):
    doc = minidom.parseString(xml_bytes)
    for rel in _children(doc.documentElement, "Relationship"):
        if rel.getAttribute("Id") in dropped_ids:
            doc.documentElement.removeChild(rel)
    return doc.toxml(encoding="UTF-8")


def lite_content_types(xml_bytes, dropped_parts):
    doc = minidom.parseString(xml_bytes)
    for override in _children(doc.documentElement, "Override"):
        if override.getAttribute("PartName").lstrip("/") in dropped_parts:
            doc.documentElement.removeChild(override)
    return doc.toxml(encoding="UTF-8")


def _rels_part(part):
    return posixpath.join(posixpath.dirname(part), "_rels",
                          posixpath.basename(part) + ".rels")


def _shared_strings(z, part):
    if part is None or part not in z.namelist():
        return None
    doc = minidom.parseString(z.read(part))
    return ["".join(_text(t) for t in si.getElementsByTagNameNS(_MAIN, "t"))
            for si in _children(doc.documentElement, "si")]


def plan_lite(z):
    """(dropped part names, {part: new bytes}) for one open values workbook."""
    names = set(z.namelist())
    workbook = "xl/workbook.xml"
    if "_rels/.rels" in names:
        for _id, kind, part in _relationships(z.read("_rels/.rels"), ""):
            if kind == _OFFICE_DOCUMENT and part:
                workbook = part
    base = posixpath.dirname(workbook)
    workbook_rels = _rels_part(workbook)
    if workbook not in names or workbook_rels not in names:
        raise ValueError("not an Excel workbook (no workbook part)")
    rels = _relationships(z.read(workbook_rels), base)
    by_id = {rid: part for rid, _kind, part in rels}
    wb_doc = minidom.parseString(z.read(workbook))
    sheets = {s.getAttribute("name"): by_id.get(s.getAttributeNS(_DOC_REL, "id"))
              for s in _children(_child(wb_doc.documentElement, "sheets"), "sheet")}
    if sheets.get(SUMMARY_SHEET) is None:
        raise ValueError("the workbook has no Summary sheet")
    dropped_sheets = tuple(n for n in DROPPED_SHEETS if n in sheets)

    def linked(part):
        rp = _rels_part(part)
        if rp not in names:
            return set()
        return {p for _i, _k, p in _relationships(z.read(rp), posixpath.dirname(part))
                if p is not None} | {rp}

    dropped = set()
    for sheet in dropped_sheets:
        dropped |= {sheets[sheet]} | linked(sheets[sheet])
    kept_links = set()
    for sheet, part in sheets.items():
        if sheet not in dropped_sheets and part is not None:
            kept_links |= linked(part)
    dropped -= kept_links
    calc_ids = {rid for rid, kind, _p in rels if kind == _CALC_CHAIN}
    dropped |= {by_id[rid] for rid in calc_ids if by_id[rid]}

    summary = sheets[SUMMARY_SHEET]
    if _rels_part(summary) in names:
        raise ValueError("the Summary sheet has linked parts a lite copy does not handle")
    shared = _shared_strings(z, next((p for _i, k, p in rels if k == _SHARED_STRINGS), None))
    workbook_xml, sheet_ids = lite_workbook(z.read(workbook), dropped_sheets)
    replaced = {
        summary: lite_summary(z.read(summary), shared),
        workbook: workbook_xml,
        workbook_rels: lite_rels(z.read(workbook_rels), set(sheet_ids) | calc_ids),
        _CONTENT_TYPES: lite_content_types(z.read(_CONTENT_TYPES), dropped),
    }
    return dropped, replaced


# --------------------------------------------------------------------------- #
# writing the zip: kept members by their compressed bytes
# --------------------------------------------------------------------------- #
class _NeedsZip64(Exception):
    """A size or offset past the plain zip format's 4 GiB / 65,535 limits."""


def _dos_stamp(date_time):
    y, mo, d, h, mi, s = date_time
    return (h << 11) | (mi << 5) | (s // 2), ((y - 1980) << 9) | (mo << 5) | d


class _ZipWriter:
    """A plain (non-zip64) zip writer that can copy a member's compressed bytes
    as they are. `close` writes the central directory."""

    def __init__(self, fp):
        self._fp = fp
        self._entries = []

    def _begin(self, info, method, crc, csize, usize):
        offset = self._fp.tell()
        if max(offset, csize, usize) >= _MAX32:
            raise _NeedsZip64
        name = info.filename.encode("utf-8")
        flags = info.flag_bits & ~(_DATA_DESCRIPTOR | _UTF8_NAME)
        if not name.isascii():
            flags |= _UTF8_NAME
        stamp, date = _dos_stamp(info.date_time)
        self._fp.write(_LOCAL_HEADER.pack(_LOCAL_SIGNATURE, 20, 0, flags, method,
                                          stamp, date, crc, csize, usize,
                                          len(name), 0))
        self._fp.write(name)
        self._entries.append((name, flags, method, stamp, date, crc, csize, usize,
                              info.create_system, info.external_attr, offset))

    def copy(self, src, info):
        if info.flag_bits & _ENCRYPTED:
            raise ValueError(f"{info.filename} is encrypted")
        src.seek(info.header_offset)
        head = _LOCAL_HEADER.unpack(src.read(_LOCAL_HEADER.size))
        if head[0] != _LOCAL_SIGNATURE:
            raise ValueError(f"{info.filename}: damaged zip entry")
        start = info.header_offset + _LOCAL_HEADER.size + head[10] + head[11]
        self._begin(info, info.compress_type, info.CRC, info.compress_size,
                    info.file_size)
        src.seek(start)
        left = info.compress_size
        while left:
            chunk = src.read(min(left, _CHUNK))
            if not chunk:
                raise ValueError(f"{info.filename}: the zip entry is cut short")
            self._fp.write(chunk)
            left -= len(chunk)

    def add(self, info, data):
        packer = zlib.compressobj(6, zlib.DEFLATED, -15)
        packed = packer.compress(data) + packer.flush()
        self._begin(info, zipfile.ZIP_DEFLATED, zlib.crc32(data) & _MAX32,
                    len(packed), len(data))
        self._fp.write(packed)

    def close(self):
        start = self._fp.tell()
        for (name, flags, method, stamp, date, crc, csize, usize, system, attr,
             offset) in self._entries:
            self._fp.write(_CENTRAL_HEADER.pack(
                _CENTRAL_SIGNATURE, 20, system, 20, 0, flags, method, stamp, date,
                crc, csize, usize, len(name), 0, 0, 0, 0, attr, offset))
            self._fp.write(name)
        size = self._fp.tell() - start
        count = len(self._entries)
        if count >= _MAX16 or start >= _MAX32 or size >= _MAX32:
            raise _NeedsZip64
        self._fp.write(_END_RECORD.pack(_END_SIGNATURE, 0, 0, count, count, size,
                                        start, 0))


def _write_raw(src, z, out, dropped, replaced):
    writer = _ZipWriter(out)
    for info in z.infolist():
        if info.filename in dropped:
            continue
        if info.filename in replaced:
            writer.add(info, replaced[info.filename])
        else:
            writer.copy(src, info)
    writer.close()


def _write_recompressed(z, out, dropped, replaced):
    """The zip64-capable path for a workbook past the plain zip limits: every
    kept member is decompressed and compressed again (slow, never needed below
    4 GiB)."""
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED, allowZip64=True) as zout:
        for info in z.infolist():
            if info.filename in dropped:
                continue
            new = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            new.compress_type = zipfile.ZIP_DEFLATED
            new.external_attr = info.external_attr
            if info.filename in replaced:
                zout.writestr(new, replaced[info.filename])
                continue
            with z.open(info) as fin, zout.open(
                    new, "w", force_zip64=info.file_size > zipfile.ZIP64_LIMIT) as fout:
                shutil.copyfileobj(fin, fout, _CHUNK)


def _check_written(path, expected, replaced):
    with zipfile.ZipFile(path) as z:
        if z.namelist() != expected:
            raise ValueError("the lite copy's parts are not the ones planned")
        for name, data in replaced.items():
            if z.read(name) != data:
                raise ValueError(f"the lite copy's {name} did not read back")


def _identity(path):
    st = os.stat(path)
    return st.st_size, st.st_mtime_ns


def write_lite_copy(values_path, lite_path=None, allow=None):
    """Write the lite copy of the values workbook `values_path` to `lite_path`
    ('<name> (lite).xlsx' beside it by default) through a temporary file, and
    return its path. `allow(path)` is the caller's destination guard, asked
    before the write and again before the rename. Raises ValueError when the
    workbook is not a comparison this module knows, changes while it is read,
    or the destination is not allowed; OSError when the file cannot be written
    (for instance, the lite copy is open in Excel)."""
    values_path = Path(values_path)
    lite_path = Path(lite_path) if lite_path is not None else lite_sibling(values_path)
    if lite_path.resolve() == values_path.resolve():
        raise ValueError("the lite copy would overwrite its own workbook")
    tmp = lite_path.with_name(f"{lite_path.name}.{os.getpid()}.tmp")

    def _require(path, action):
        if allow is not None and not allow(path):
            raise ValueError(f"the lite copy's destination changed before the {action}")

    _require(lite_path, "write")
    _require(tmp, "write")
    before = _identity(values_path)
    try:
        with open(values_path, "rb") as src, zipfile.ZipFile(src) as z:
            dropped, replaced = plan_lite(z)
            expected = [i.filename for i in z.infolist() if i.filename not in dropped]
            try:
                with open(tmp, "wb") as out:
                    _write_raw(src, z, out, dropped, replaced)
            except _NeedsZip64:
                with open(tmp, "wb") as out:
                    _write_recompressed(z, out, dropped, replaced)
        if _identity(values_path) != before:
            raise ValueError("the comparison workbook changed while its lite copy "
                             "was being written")
        _check_written(tmp, expected, replaced)
        _require(lite_path, "rename")
        os.replace(tmp, lite_path)
    finally:
        try:
            tmp.unlink()
        except FileNotFoundError:  # silent-ok: the rename consumed it
            pass
    return lite_path
