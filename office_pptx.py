"""Editing an EXISTING pptx: text replacement and table-cell writes.

The docx analogue (office_docx) and this module solve the same problem in two
dialects, and share the run-spanning replacement itself (office_ooxml_text). What
is genuinely different about a presentation, and is handled here:

  * A deck is MANY parts. Text lives in ppt/slides/slideN.xml, one part per slide,
    so a replacement has to visit all of them and report WHICH slides it changed.
  * SLIDE NUMBER IS NOT FILE NUMBER. What the viewer calls slide 1 is the first
    entry of <p:sldIdLst> in presentation.xml, resolved through the relationships.
    Reorder a deck in PowerPoint and slide1.xml can end up third. Addressing a
    slide by its file name would edit a different slide than the caller sees, and
    the file would look perfectly fine afterwards - so the order is resolved
    properly, and slides are numbered from 1 the way a human counts them.
  * An empty table cell has no run to write into. Word cells practically always
    carry one; PowerPoint cells often do not. Refusing there would make the tool
    useless for filling in a blank template, so a run is INSERTED - with no
    explicit formatting, which means it inherits the table/placeholder style.

What is deliberately NOT touched, and is named rather than silently skipped:
speaker notes (notesSlides), the master and layouts, and text inside charts or
SmartArt. A replacement that reports 0 says so.
"""

import json
import os
import posixpath
import re
import sys
import xml.etree.ElementTree as ET
import zipfile

from office_ooxml_text import escape, replace_in_runs, runs_text, write_first_run

TEXT_RE = re.compile(r"(<a:t\b[^>]*>)(.*?)(</a:t>)", re.S)
PARAGRAPH_RE = re.compile(r"<a:p\b(?:[^>]*)>.*?</a:p>", re.S)
TABLE_RE = re.compile(r"<a:tbl\b.*?</a:tbl>", re.S)
ROW_RE = re.compile(r"<a:tr\b.*?</a:tr>", re.S)
CELL_RE = re.compile(r"<a:tc\b.*?</a:tc>", re.S)
END_PARA_RE = re.compile(r"<a:endParaRPr\b[^>]*/>|<a:endParaRPr\b.*?</a:endParaRPr>", re.S)
EMPTY_PARAGRAPH_RE = re.compile(r"<a:p\b[^>]*/>")

RELS_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
SLIDE_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"


class PptxError(Exception):
    """The deck cannot be read, or the requested target does not exist."""


def slide_text(slide_xml):
    """What a reader sees on this slide, run boundaries collapsed."""
    return runs_text(slide_xml, TEXT_RE)


def replace_in_slide(slide_xml, find, replace):
    """Replace every occurrence on ONE slide, paragraph by paragraph.

    Paragraph by paragraph, not whole-slide, for the same reason as in a document:
    two separate text boxes are two separate texts, and a match must never be
    invented across the gap between them.
    """
    if not find:
        raise PptxError("nothing to find: the search text is empty")
    total = 0
    out, last = [], 0
    for match in PARAGRAPH_RE.finditer(slide_xml):
        new_paragraph, count = replace_in_runs(match.group(0), find, replace, TEXT_RE)
        total += count
        out.append(slide_xml[last:match.start()])
        out.append(new_paragraph)
        last = match.end()
    out.append(slide_xml[last:])
    return "".join(out), total


def insert_run(cell_xml, text):
    """Put `text` into a cell that has no run at all, or return None if impossible.

    The run is inserted into the cell's FIRST paragraph, before <a:endParaRPr> when
    there is one: the schema puts that element last, and a run after it makes the
    cell render empty in PowerPoint while looking fine in the XML.
    """
    paragraph = PARAGRAPH_RE.search(cell_xml)
    if not paragraph:
        # A blank cell is often stored as the self-closing <a:p/>, which has no
        # inside to insert into - it gets one.
        empty = EMPTY_PARAGRAPH_RE.search(cell_xml)
        if not empty:
            return None
        run = '<a:r><a:rPr lang="hu-HU"/><a:t>%s</a:t></a:r>' % escape(text)
        return cell_xml[:empty.start()] + "<a:p>" + run + "</a:p>" + cell_xml[empty.end():]
    body = paragraph.group(0)
    run = '<a:r><a:rPr lang="hu-HU"/><a:t>%s</a:t></a:r>' % escape(text)
    end = END_PARA_RE.search(body)
    if end:
        new_body = body[:end.start()] + run + body[end.start():]
    else:
        new_body = body[:-len("</a:p>")] + run + "</a:p>"
    return cell_xml[:paragraph.start()] + new_body + cell_xml[paragraph.end():]


def set_table_cell(slide_xml, table, row, column, text):
    """Write one table cell of ONE slide, addressed by position (zero-based here).

    Slides are numbered from 1 because that is how a person counts them; table,
    row and column are zero-based to match docx_set_table_cell. The mixture is
    deliberate and is stated in every error message, because a silent off-by-one
    on a slide edits the wrong slide and looks like success.
    """
    tables = list(TABLE_RE.finditer(slide_xml))
    if table >= len(tables):
        raise PptxError("this slide has %d table(s); no table at index %d (tables are "
                        "numbered from 0)" % (len(tables), table))
    table_xml = tables[table].group(0)

    rows = list(ROW_RE.finditer(table_xml))
    if row >= len(rows):
        raise PptxError("table %d has %d row(s); no row at index %d (rows are numbered "
                        "from 0)" % (table, len(rows), row))
    row_xml = rows[row].group(0)

    cells = list(CELL_RE.finditer(row_xml))
    if column >= len(cells):
        raise PptxError("row %d of table %d has %d cell(s); no cell at index %d (cells "
                        "are numbered from 0)" % (row, table, len(cells), column))
    cell_xml = cells[column].group(0)

    new_cell_xml = write_first_run(cell_xml, text, TEXT_RE)
    if new_cell_xml is None:
        # An empty cell: no run to overwrite, so one is created.
        new_cell_xml = insert_run(cell_xml, text)
    if new_cell_xml is None:
        raise PptxError("the cell at table %d, row %d, column %d has no paragraph to "
                        "write into - the deck stores this cell in a shape this tool "
                        "does not handle" % (table, row, column))

    new_row = row_xml[:cells[column].start()] + new_cell_xml + row_xml[cells[column].end():]
    new_table = table_xml[:rows[row].start()] + new_row + table_xml[rows[row].end():]
    return slide_xml[:tables[table].start()] + new_table + slide_xml[tables[table].end():]


def slide_parts(zf):
    """Slide part names in PRESENTATION ORDER - index 0 is what the viewer calls 1."""
    try:
        presentation = zf.read("ppt/presentation.xml")
        rels = zf.read("ppt/_rels/presentation.xml.rels")
    except KeyError as exc:
        raise PptxError("this does not look like a presentation (%s missing)" % exc)

    try:
        targets = {}
        for relationship in ET.fromstring(rels):
            if relationship.get("Type") == SLIDE_TYPE:
                targets[relationship.get("Id")] = relationship.get("Target")
        order = []
        for element in ET.fromstring(presentation).iter():
            if not element.tag.endswith("}sldIdLst"):
                continue
            for slide in element:
                target = targets.get(slide.get(RELS_NS + "id"))
                if not target:
                    continue
                order.append(posixpath.normpath(posixpath.join("ppt", target)).lstrip("/"))
    except ET.ParseError as exc:
        raise PptxError("the presentation part is not readable XML: %s" % exc)

    if not order:
        raise PptxError("the deck lists no slides in presentation.xml")
    return order


def read_slides(path):
    """(ordered part names, {name: xml}) for every slide of the deck."""
    try:
        with zipfile.ZipFile(path) as zf:
            names = slide_parts(zf)
            missing = [n for n in names if n not in zf.namelist()]
            if missing:
                raise PptxError("the deck references slide part(s) that are not in the "
                                "file: %s" % ", ".join(missing))
            return names, {name: zf.read(name).decode("utf-8") for name in names}
    except (zipfile.BadZipFile, OSError) as exc:
        raise PptxError("cannot read the presentation (%s): %s" % (exc, path))


def write_slides(src, dst, slides):
    """Copy the pptx, replacing only the slide parts given in `slides`."""
    with zipfile.ZipFile(src) as zin:
        parts = [(item, zin.read(item.filename)) for item in zin.infolist()]
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as out:
        for item, data in parts:
            if item.filename in slides:
                data = slides[item.filename].encode("utf-8")
            out.writestr(item.filename, data)
    return dst


def _resolve_slide(names, number):
    """1-based slide number -> part name, or a clean refusal."""
    try:
        number = int(number)
    except (TypeError, ValueError):
        raise PptxError("slide must be a number; got %r" % (number,))
    if number < 1:
        raise PptxError("slides are numbered from 1 (the first slide is 1, not 0); got %d"
                        % number)
    if number > len(names):
        raise PptxError("the deck has %d slide(s); no slide %d" % (len(names), number))
    return names[number - 1]


def apply_edits(path, replacements=(), cells=(), slide=None):
    """Apply text replacements and table-cell writes to `path`, in place.

    A replacement that matched NOTHING is reported with a count of 0 rather than
    passed over: a placeholder that silently fails to fill is exactly the failure
    this tool exists to prevent.

    The file is only overwritten once every edit has been computed, so a failure
    half-way through leaves the original exactly as it was.
    """
    names, slides = read_slides(path)
    scope = [_resolve_slide(names, slide)] if slide is not None else names
    applied = []

    for item in replacements:
        find = item.get("find")
        replace = item.get("replace")
        if find is None or replace is None:
            raise PptxError("each replacement needs a find and a replace")
        total, touched = 0, []
        for name in scope:
            new_xml, count = replace_in_slide(slides[name], find, replace)
            if count:
                slides[name] = new_xml
                total += count
                touched.append(names.index(name) + 1)
        applied.append({"find": find, "replaced": total, "slides": touched})

    for cell in cells:
        try:
            number = cell["slide"]
            table = int(cell["table"])
            row = int(cell["row"])
            column = int(cell["column"])
            text = cell["text"]
        except (KeyError, TypeError, ValueError):
            raise PptxError("each cell needs slide (from 1), table, row, column (numbers "
                            "from 0) and text")
        name = _resolve_slide(names, number)
        slides[name] = set_table_cell(slides[name], table, row, column, text)
        applied.append({"cell": [int(number), table, row, column]})

    tmp = path + ".tmp-pptx"
    write_slides(path, tmp, slides)
    os.replace(tmp, path)
    return applied


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
                              cells=req.get("cells") or (),
                              slide=req.get("slide"))
    except PptxError as exc:
        json.dump({"ok": False, "error": str(exc)}, sys.stdout)
        return
    except Exception as exc:  # noqa: BLE001 - a bug must not leak a traceback
        import traceback
        traceback.print_exc(file=sys.stderr)
        json.dump({"ok": False, "error": "pptx edit failed: %s" % exc}, sys.stdout)
        return

    json.dump({"ok": True, "file": path, "applied": applied}, sys.stdout)


if __name__ == "__main__":
    main()
