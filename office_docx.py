"""Text replacement in a docx, at the level a document actually stores text.

The trap this module exists for: a sentence you see as one line is stored as
SEVERAL runs. Word splits them for formatting, spell-check state, revision ids -
so "12 500 000" can be `<w:t>12 5</w:t><w:t>00 000</w:t>` with nothing visibly
wrong. A naive string replace over document.xml then finds nothing and reports
success, which is the worst outcome available: the caller believes the amount was
updated.

The run-spanning replacement itself lives in office_ooxml_text, because
PowerPoint has exactly the same trap in a different dialect (<a:t> instead of
<w:t>) and two copies of that algorithm would drift. This module owns what is
specific to a Word document: which elements are paragraphs, cells, rows and
tables, and how the package is read and written back.

A replacement that spans a bold/plain boundary comes out in the first run's
style - worth knowing rather than discovering.

Everything outside the touched <w:t> elements is left byte-for-byte alone. XML
declarations, namespace prefixes, revision ids and the rest are exactly as Word
wrote them: every extra edit is another chance to hand LibreOffice or Word a file
it silently refuses.
"""

import json
import os
import re
import sys
import zipfile

from office_ooxml_text import (  # noqa: F401 - re-exported for callers and tests
    ENTITIES, escape, replace_in_runs, runs_text, unescape, write_first_run,
)

PARAGRAPH_RE = re.compile(r"<w:p\b.*?</w:p>", re.S)
TEXT_RE = re.compile(r"(<w:t\b[^>]*>)(.*?)(</w:t>)", re.S)
CELL_RE = re.compile(r"<w:tc\b.*?</w:tc>", re.S)
ROW_RE = re.compile(r"<w:tr\b.*?</w:tr>", re.S)
TABLE_RE = re.compile(r"<w:tbl\b.*?</w:tbl>", re.S)


class DocxError(Exception):
    """The document cannot be read, or the requested target does not exist."""


def paragraph_text(paragraph_xml):
    """What a reader sees in this paragraph, with the run boundaries collapsed."""
    return runs_text(paragraph_xml, TEXT_RE)


def replace_in_paragraph(paragraph_xml, find, replace):
    """Replace every occurrence of `find` in this paragraph. Returns (xml, count)."""
    if not find:
        raise DocxError("nothing to find: the search text is empty")
    return replace_in_runs(paragraph_xml, find, replace, TEXT_RE)


def replace_text(document_xml, find, replace):
    """Apply the replacement paragraph by paragraph. Returns (xml, count)."""
    total = 0
    out, last = [], 0
    for match in PARAGRAPH_RE.finditer(document_xml):
        new_paragraph, count = replace_in_paragraph(match.group(0), find, replace)
        total += count
        out.append(document_xml[last:match.start()])
        out.append(new_paragraph)
        last = match.end()
    out.append(document_xml[last:])
    return "".join(out), total


def set_table_cell(document_xml, table, row, column, text):
    """Put `text` in one table cell, addressed by position (all zero-based).

    Tables are the one place in a document with real addressing, so this does not
    need to search for anything. The cell's first run keeps its formatting and
    receives the text; the remaining runs in that cell are emptied, otherwise the
    old value would still be sitting there after the new one.
    """
    tables = list(TABLE_RE.finditer(document_xml))
    if table >= len(tables):
        raise DocxError("the document has %d table(s); no table at index %d" % (len(tables), table))
    table_xml = tables[table].group(0)

    rows = list(ROW_RE.finditer(table_xml))
    if row >= len(rows):
        raise DocxError("table %d has %d row(s); no row at index %d" % (table, len(rows), row))
    row_xml = rows[row].group(0)

    cells = list(CELL_RE.finditer(row_xml))
    if column >= len(cells):
        raise DocxError(
            "row %d of table %d has %d cell(s); no cell at index %d"
            % (row, table, len(cells), column)
        )
    cell_xml = cells[column].group(0)

    new_cell_xml = write_first_run(cell_xml, text, TEXT_RE)
    if new_cell_xml is None:
        raise DocxError(
            "the cell at table %d, row %d, column %d has no text run to write into"
            % (table, row, column)
        )

    new_row = row_xml[:cells[column].start()] + new_cell_xml + row_xml[cells[column].end():]
    new_table = table_xml[:rows[row].start()] + new_row + table_xml[rows[row].end():]
    return (document_xml[:tables[table].start()] + new_table
            + document_xml[tables[table].end():])


def apply_edits(path, replacements=(), cells=()):
    """Apply text replacements and table-cell writes to `path`, in place.

    Returns a report of what was changed. A replacement that matched NOTHING is
    reported with a count of 0 rather than passed over: a template placeholder
    that silently fails to fill is the failure this tool exists to prevent, and
    the caller has to be able to see it.

    The file is only overwritten once every edit has been computed, so a failure
    half-way through leaves the original exactly as it was.
    """
    document = read_document(path)
    applied = []

    for item in replacements:
        find = item.get("find")
        replace = item.get("replace")
        if find is None or replace is None:
            raise DocxError("each replacement needs a find and a replace")
        document, count = replace_text(document, find, replace)
        applied.append({"find": find, "replaced": count})

    for cell in cells:
        try:
            table = int(cell["table"])
            row = int(cell["row"])
            column = int(cell["column"])
            text = cell["text"]
        except (KeyError, TypeError, ValueError):
            raise DocxError("each cell needs table, row, column (numbers) and text")
        document = set_table_cell(document, table, row, column, text)
        applied.append({"cell": [table, row, column]})

    tmp = path + ".tmp-docx"
    write_document(path, tmp, document)
    os.replace(tmp, path)
    return applied


def read_document(path):
    try:
        with zipfile.ZipFile(path) as zf:
            return zf.read("word/document.xml").decode("utf-8")
    except (zipfile.BadZipFile, KeyError, OSError) as exc:
        raise DocxError("cannot read the document (%s): %s" % (exc, path))


def write_document(src, dst, document_xml):
    """Copy the docx, replacing only word/document.xml."""
    with zipfile.ZipFile(src) as zin:
        parts = [(item, zin.read(item.filename)) for item in zin.infolist()]
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as out:
        for item, data in parts:
            if item.filename == "word/document.xml":
                data = document_xml.encode("utf-8")
            out.writestr(item.filename, data)
    return dst


def main():
    """CLI contract, same shape as the other office helpers: JSON in, JSON out."""
    from office_paths import PathNotAllowed, check_path

    try:
        req = json.load(sys.stdin)
    except Exception as exc:  # noqa: BLE001
        json.dump({"ok": False, "error": "invalid request JSON: %s" % exc}, sys.stdout)
        return

    path = req.get("file")
    if not path:
        json.dump({"ok": False, "error": "file is required"}, sys.stdout)
        return

    try:
        path = check_path(path)
    except PathNotAllowed as exc:
        json.dump({"ok": False, "error": str(exc)}, sys.stdout)
        return

    try:
        applied = apply_edits(path,
                              replacements=req.get("replacements") or (),
                              cells=req.get("cells") or ())
    except DocxError as exc:
        json.dump({"ok": False, "error": str(exc)}, sys.stdout)
        return
    except Exception as exc:  # noqa: BLE001 - a bug must not leak a traceback
        import traceback
        traceback.print_exc(file=sys.stderr)
        json.dump({"ok": False, "error": "docx edit failed: %s" % exc}, sys.stdout)
        return

    json.dump({"ok": True, "file": path, "applied": applied}, sys.stdout)


if __name__ == "__main__":
    main()
