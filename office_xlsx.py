#!/usr/bin/env python3
"""OOXML cell writer for the office MCP.

Reads a JSON request on stdin and writes a JSON result on stdout:

    {"file": "/abs/path.xlsx", "sheet": "Sheet1",
     "cells": [{"ref": "B7", "value": 42}, {"ref": "C7", "formula": "A1+B1"}]}

Why Python and not the Node server itself: xlsx is a ZIP container, Node has no bundled ZIP
support and the host has no `zip` binary, so the alternative would be hand-rolling ZIP
central-directory writing - exactly the kind of code that silently produces a corrupt
workbook. `zipfile` is stdlib and battle-tested, so the rewrite happens here.

Two correctness details that matter more than they look:

* A formula cell is written WITHOUT a cached `<v>`. A stale cache is worse than no cache: a
  reader that does not recalculate would hand back a confidently wrong number.
* Because of that, `calcChain.xml` is dropped and `fullCalcOnLoad` is set, so a spreadsheet
  application recalculates on open instead of trusting what we left behind.
"""

import json
import re
import shutil
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from office_paths import PathNotAllowed, check_path

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

ET.register_namespace("", MAIN_NS)

CELL_REF = re.compile(r"^([A-Za-z]+)([1-9][0-9]*)$")


def fail(message):
    json.dump({"ok": False, "error": message}, sys.stdout)
    sys.exit(0)


def col_index(letters):
    """'A' -> 1, 'AA' -> 27. Used to keep cells in ascending column order."""
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def q(tag):
    return f"{{{MAIN_NS}}}{tag}"


def worksheet_path(zf, sheet_name):
    """Map a sheet NAME to its part path, via workbook.xml + its rels."""
    try:
        wb = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    except KeyError as exc:
        fail(f"not a valid xlsx: missing {exc}")

    rel_target = {}
    for rel in rels:
        rel_target[rel.get("Id")] = rel.get("Target")

    for sheet in wb.iter(q("sheet")):
        if sheet.get("name") == sheet_name:
            rid = sheet.get(f"{{{DOC_REL_NS}}}id")
            target = rel_target.get(rid)
            if not target:
                fail(f"sheet '{sheet_name}' has no resolvable relationship")
            if target.startswith("/"):
                return target.lstrip("/")
            return f"xl/{target}" if not target.startswith("xl/") else target

    names = [s.get("name") for s in wb.iter(q("sheet"))]
    fail(f"sheet '{sheet_name}' not found; workbook has: {', '.join(n for n in names if n)}")


def get_or_create_row(sheet_data, row_num):
    """Rows must stay in ascending r order, so insert at the right position."""
    insert_at = len(sheet_data)
    for i, row in enumerate(sheet_data.findall(q("row"))):
        r = int(row.get("r", "0"))
        if r == row_num:
            return row
        if r > row_num:
            insert_at = list(sheet_data).index(row)
            break
    row = ET.Element(q("row"), {"r": str(row_num)})
    sheet_data.insert(insert_at, row)
    return row


def get_or_create_cell(row, ref, col_num):
    """Cells must stay in ascending column order within their row."""
    insert_at = len(row)
    for cell in row.findall(q("c")):
        cur = cell.get("r", "")
        m = CELL_REF.match(cur)
        if not m:
            continue
        cur_col = col_index(m.group(1))
        if cur_col == col_num:
            return cell
        if cur_col > col_num:
            insert_at = list(row).index(cell)
            break
    cell = ET.Element(q("c"), {"r": ref})
    row.insert(insert_at, cell)
    return cell


def write_cell(cell, spec):
    # Start from a clean cell: leaving an old <v> next to a new <f> is how stale values
    # survive an edit.
    for child in list(cell):
        cell.remove(child)
    cell.attrib.pop("t", None)

    if "formula" in spec and spec["formula"] is not None:
        formula = str(spec["formula"]).lstrip("=")
        f = ET.SubElement(cell, q("f"))
        f.text = formula
        return "formula"

    value = spec.get("value")
    if isinstance(value, bool):
        cell.set("t", "b")
        v = ET.SubElement(cell, q("v"))
        v.text = "1" if value else "0"
        return "boolean"
    if isinstance(value, (int, float)):
        v = ET.SubElement(cell, q("v"))
        v.text = repr(value) if isinstance(value, float) else str(value)
        return "number"

    # Strings go in as inline strings: no sharedStrings.xml bookkeeping, no index drift.
    cell.set("t", "inlineStr")
    is_el = ET.SubElement(cell, q("is"))
    t_el = ET.SubElement(is_el, q("t"))
    text = "" if value is None else str(value)
    if text != text.strip():
        t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t_el.text = text
    return "text"


def set_full_calc_on_load(workbook_xml):
    """Ask the reader to recalculate on open, since formula cells carry no cached value."""
    wb = ET.fromstring(workbook_xml)
    calc_pr = wb.find(q("calcPr"))
    if calc_pr is None:
        calc_pr = ET.SubElement(wb, q("calcPr"))
    calc_pr.set("fullCalcOnLoad", "1")
    return ET.tostring(wb, encoding="UTF-8", xml_declaration=True)


def main():
    try:
        req = json.load(sys.stdin)
    except Exception as exc:
        fail(f"invalid request JSON: {exc}")

    path = req.get("file")
    sheet_name = req.get("sheet")
    cells = req.get("cells") or []
    if not path or not sheet_name or not cells:
        fail("file, sheet and a non-empty cells list are required")

    # Where the caller may write at all, checked before anything is opened.
    try:
        path = check_path(path)
    except PathNotAllowed as exc:
        fail(str(exc))

    for spec in cells:
        ref = str(spec.get("ref", ""))
        if not CELL_REF.match(ref):
            fail(f"invalid cell reference: '{ref}'")
        if "value" not in spec and "formula" not in spec:
            fail(f"cell {ref}: give either value or formula")

    try:
        zin = zipfile.ZipFile(path, "r")
    except FileNotFoundError:
        fail(f"file not found: {path}")
    except zipfile.BadZipFile:
        fail(f"not a valid xlsx (bad zip): {path}")

    with zin:
        ws_path = worksheet_path(zin, sheet_name)
        try:
            sheet_xml = zin.read(ws_path)
        except KeyError:
            fail(f"worksheet part missing from archive: {ws_path}")

        sheet = ET.fromstring(sheet_xml)
        sheet_data = sheet.find(q("sheetData"))
        if sheet_data is None:
            sheet_data = ET.SubElement(sheet, q("sheetData"))

        applied = []
        for spec in cells:
            ref = spec["ref"].upper()
            m = CELL_REF.match(ref)
            col, row_num = m.group(1), int(m.group(2))
            row = get_or_create_row(sheet_data, row_num)
            cell = get_or_create_cell(row, ref, col_index(col))
            applied.append({"ref": ref, "kind": write_cell(cell, spec)})

        new_sheet = ET.tostring(sheet, encoding="UTF-8", xml_declaration=True)
        workbook_xml = set_full_calc_on_load(zin.read("xl/workbook.xml"))

        # Rewrite the container: copy every part except the ones we replace, and drop
        # calcChain.xml, which is now stale for any formula we touched.
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        tmp.close()
        with zipfile.ZipFile(tmp.name, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "xl/calcChain.xml":
                    continue
                if item.filename == ws_path:
                    zout.writestr(item, new_sheet)
                elif item.filename == "xl/workbook.xml":
                    zout.writestr(item, workbook_xml)
                else:
                    zout.writestr(item, zin.read(item.filename))

    shutil.move(tmp.name, path)
    json.dump({"ok": True, "file": path, "sheet": sheet_name, "applied": applied}, sys.stdout)


if __name__ == "__main__":
    main()
