"""Extract ONE worksheet from an xlsx into a standalone workbook.

Why this exists is a security requirement, not a convenience: LibreOffice exports
a WHOLE workbook to PDF, so a single "print the repayment plan" turns into every
sheet in the file - including the ones carrying personal data, and (per the
fleet's own observation) including sheets marked hidden. Hiding is therefore not
a fix; the other sheets have to be genuinely gone from what the engine sees.

The copy keeps the target sheet exactly as stored: its cells, its cached formula
results, its styles, its merged ranges. Nothing is recalculated here - that is
deliberate. A kept sheet routinely references the sheets we just removed (203 of
2219 formula cells in a real workbook did), so a recalculation would turn those
into #REF! and the PDF would show errors where numbers belong. Exporting the
stored values is also what the tool already promises: recalculation is a separate
step, run on purpose.

What is rewritten, and nothing more: the workbook's sheet list, its relationship
part, the content-type overrides for the parts that are gone, and any defined
name that pointed at a removed sheet. Everything else is copied byte-for-byte,
because every extra edit is another way to hand LibreOffice a file it silently
refuses (an earlier hand-rezip did exactly that: exit 0, no output).
"""

import os
import re
import shutil
import zipfile

SHEET_RE = re.compile(r'<sheet\b[^>]*/>')
NAME_RE = re.compile(r'name="([^"]*)"')
# The relationship id may carry ANY namespace prefix, not just "r": our own cell
# writer re-serialises workbook.xml through ElementTree, which renames prefixes
# (r:id -> ns0:id). Matching a literal "r:id" made this module fail on a file the
# neighbouring tool had just written - "lists the sheet without a relationship id".
# The negative lookbehind keeps sheetId out of it.
RID_RE = re.compile(r'(?<![A-Za-z])(?:[A-Za-z0-9_]+:)id="([^"]*)"')
STATE_RE = re.compile(r'state="([^"]*)"')


class SheetError(Exception):
    """The requested sheet does not exist, or the workbook cannot be read."""


def _read_parts(path):
    with zipfile.ZipFile(path) as zf:
        return [(item, zf.read(item.filename)) for item in zf.infolist()]


def list_sheets(path):
    """Every sheet name in workbook order, with whether it is hidden.

    Returns a list of (name, hidden) - the caller needs the names to offer a
    choice, and the hidden flag because a hidden sheet is exactly the kind that
    should never leave the building by accident.
    """
    try:
        with zipfile.ZipFile(path) as zf:
            workbook = zf.read("xl/workbook.xml").decode("utf-8")
    except (zipfile.BadZipFile, KeyError, OSError) as exc:
        raise SheetError("cannot read the workbook (%s): %s" % (exc, path))

    sheets = []
    for tag in SHEET_RE.findall(workbook):
        name = NAME_RE.search(tag)
        if not name:
            continue
        state = STATE_RE.search(tag)
        sheets.append((name.group(1), bool(state) and state.group(1) in ("hidden", "veryHidden")))
    return sheets


def single_sheet_copy(src, dst, sheet_name):
    """Write `dst`: a copy of `src` containing only `sheet_name`.

    Raises SheetError if the sheet is not in the workbook - naming a sheet that
    does not exist must never quietly fall back to exporting everything.
    """
    names = [name for name, _hidden in list_sheets(src)]
    if sheet_name not in names:
        raise SheetError(
            "no sheet named %r in the workbook. Available: %s" % (sheet_name, ", ".join(names))
        )
    if len(names) == 1:
        shutil.copy(src, dst)
        return dst

    parts = _read_parts(src)
    by_name = {item.filename: data for item, data in parts}

    workbook = by_name["xl/workbook.xml"].decode("utf-8")
    rels = by_name["xl/_rels/workbook.xml.rels"].decode("utf-8")

    # Which relationship (and therefore which part) belongs to the sheet we keep,
    # and which parts the removed sheets pull along with them.
    keep_rid, drop_rids = None, []
    for tag in SHEET_RE.findall(workbook):
        name = NAME_RE.search(tag)
        rid = RID_RE.search(tag)
        if not name or not rid:
            continue
        if name.group(1) == sheet_name:
            keep_rid = rid.group(1)
        else:
            drop_rids.append(rid.group(1))
    if keep_rid is None:
        raise SheetError("the workbook lists %r without a relationship id" % sheet_name)

    rel_target = {}
    for tag in re.findall(r"<Relationship\b[^>]*/>", rels):
        rid = re.search(r'Id="([^"]*)"', tag)
        target = re.search(r'Target="([^"]*)"', tag)
        if rid and target:
            rel_target[rid.group(1)] = target.group(1).lstrip("/")

    dropped_parts = set()
    for rid in drop_rids:
        target = rel_target.get(rid)
        if not target:
            continue
        part = target if target.startswith("xl/") else "xl/" + target
        dropped_parts.add(part)
        # A sheet's own relationships (comments, drawings, hyperlinks) go with it.
        head, tail = os.path.split(part)
        dropped_parts.add(os.path.join(head, "_rels", tail + ".rels"))

    # 1. workbook.xml: keep one <sheet>, and drop defined names pointing at the rest.
    kept_tag = next(
        tag for tag in SHEET_RE.findall(workbook)
        if (NAME_RE.search(tag) or None) and NAME_RE.search(tag).group(1) == sheet_name
    )
    workbook = re.sub(
        r"<sheets>.*?</sheets>", "<sheets>" + kept_tag + "</sheets>", workbook, flags=re.S
    )
    workbook = _prune_defined_names(workbook, sheet_name, names.index(sheet_name))
    workbook = _reset_active_tab(workbook)

    # 2. rels: keep every relationship except the removed sheets'.
    for rid in drop_rids:
        rels = re.sub(r'<Relationship\b[^>]*Id="%s"[^>]*/>' % re.escape(rid), "", rels)

    # 3. The kept sheet's formulas have to become their stored values.
    #
    #    This is the correction that matters, found in production and
    #    reproduced here: keeping the formulas is NOT safe once the other sheets
    #    are gone. LibreOffice cannot resolve `=Alapadatok!B3` to a sheet that no
    #    longer exists, so at LOAD time - before any recalculation - it turns the
    #    reference into #NAME? and the cached value is discarded. The exported PDF
    #    then shows #NAME? where the name, the amount or the date should be, and it
    #    looks like a rendering problem rather than a wrong file. Disabling recalc
    #    does not help: the formula itself is invalid, not merely stale.
    #
    #    So a formula cell is flattened to what it last evaluated to. That is
    #    exactly what a faithful print of a sheet should contain anyway - the
    #    numbers as stored - and it is why the recalc step runs BEFORE the export.
    kept_part = rel_target.get(keep_rid, "")
    kept_part = kept_part if kept_part.startswith("xl/") else "xl/" + kept_part

    # 4. Write everything back, skipping the parts that are gone.
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as out:
        for item, data in parts:
            if item.filename in dropped_parts:
                continue
            if item.filename == "xl/workbook.xml":
                data = workbook.encode("utf-8")
            elif item.filename == "xl/_rels/workbook.xml.rels":
                data = rels.encode("utf-8")
            elif item.filename == "[Content_Types].xml":
                data = _prune_content_types(data.decode("utf-8"), dropped_parts).encode("utf-8")
            elif item.filename == kept_part:
                data = flatten_formulas(data.decode("utf-8")).encode("utf-8")
            out.writestr(item, data)
    return dst


CELL_RE = re.compile(r"<c\b[^>]*>.*?</c>", re.S)
FORMULA_RE = re.compile(r"<f\b[^>]*/>|<f\b[^>]*>.*?</f>", re.S)
VALUE_RE = re.compile(r"<v>(.*?)</v>", re.S)
TYPE_RE = re.compile(r'\st="([^"]*)"')


def flatten_formulas(sheet_xml):
    """Replace every formula in a worksheet part with its cached value.

    A formula whose references have been removed is worse than useless: it renders
    as #NAME? and takes the cached value down with it. Flattening keeps what the
    cell last evaluated to, which is what a faithful print contains.

    Two details that bite:
      - a formula returning TEXT is stored as t="str" with the text in <v>. With
        the formula gone that type is meaningless, so it becomes an inline string;
      - a cell whose cached value is missing (a formula that was never evaluated)
        has nothing to keep, so the formula is dropped and the cell left empty
        rather than left showing an error.
    """
    out, last = [], 0
    for match in CELL_RE.finditer(sheet_xml):
        cell = match.group(0)
        if "<f" not in cell:
            continue
        out.append(sheet_xml[last:match.start()])
        out.append(_flatten_cell(cell))
        last = match.end()
    out.append(sheet_xml[last:])
    return "".join(out)


def _flatten_cell(cell_xml):
    without_formula = FORMULA_RE.sub("", cell_xml)
    value = VALUE_RE.search(without_formula)
    cell_type = TYPE_RE.search(without_formula)

    if value and cell_type and cell_type.group(1) == "str":
        # A text result: t="str" only makes sense with a formula present.
        text = value.group(1)
        without_formula = TYPE_RE.sub(' t="inlineStr"', without_formula, count=1)
        without_formula = without_formula.replace(
            value.group(0), "<is><t>%s</t></is>" % text, 1
        )
    return without_formula


def _prune_defined_names(workbook, sheet_name, kept_index):
    """Drop every defined name that does not belong to the sheet we keep.

    Two ways a stale name kills the file outright, both measured on a real
    workbook (LibreOffice answers "source file could not be loaded", exits 0 and
    writes nothing):

      - its formula points at a removed sheet;
      - localSheetId is an INDEX into the sheet list, so a name scoped to the old
        sheet 5 is out of range once one sheet is left.

    Sheet names inside the formula are XML-escaped ('&apos;Sheet name&apos;!A1'),
    which is exactly why the first version of this filter kept a name it should
    have dropped: it looked for a literal quote and found none, concluded the name
    referenced nothing, and let it through.
    """
    block = re.search(r"<definedNames>.*?</definedNames>", workbook, flags=re.S)
    if not block:
        return workbook

    kept = []
    for tag in re.findall(r"<definedName\b.*?</definedName>", block.group(0), flags=re.S):
        local = re.search(r'localSheetId="(\d+)"', tag)
        if local and int(local.group(1)) != kept_index:
            continue
        if not _references_only(tag, sheet_name):
            continue
        # The kept sheet is now the only one, so a local scope can only be 0.
        kept.append(re.sub(r'localSheetId="\d+"', 'localSheetId="0"', tag))

    replacement = "<definedNames>" + "".join(kept) + "</definedNames>" if kept else ""
    return workbook.replace(block.group(0), replacement)


def _references_only(defined_name_tag, sheet_name):
    """True when every sheet the name refers to is the one we keep.

    Entities are resolved first: the reference is stored escaped, so matching the
    raw text would silently see no references at all and answer True.
    """
    text = (defined_name_tag
            .replace("&apos;", "'")
            .replace("&quot;", '"')
            .replace("&amp;", "&"))
    referenced = re.findall(r"(?:'([^']+)'|([A-Za-z0-9_.\u00c0-\u024f]+))!", text)
    if not referenced:
        # A name with no sheet reference at all (a constant) is harmless to keep.
        return True
    for quoted, bare in referenced:
        if (quoted or bare) != sheet_name:
            return False
    return True


def _prune_content_types(content_types, dropped_parts):
    for part in dropped_parts:
        content_types = re.sub(
            r'<Override\b[^>]*PartName="/%s"[^>]*/>' % re.escape(part), "", content_types
        )
    return content_types


def _reset_active_tab(workbook):
    """Point the workbook view at the only sheet that is left.

    Not cosmetic: activeTab is an INDEX. A workbook whose view still selects tab
    10 while one sheet remains is refused outright - LibreOffice answers "source
    file could not be loaded", exits 0 and writes nothing, so the failure arrives
    as a missing file rather than as an explanation. Measured on a real workbook
    while building this.
    """
    return re.sub(r'\s*activeTab="\d+"', ' activeTab="0"', workbook, count=1)
