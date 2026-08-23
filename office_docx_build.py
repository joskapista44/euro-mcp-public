#!/usr/bin/env python3
"""Build a .docx from a small structured description, with no third-party library.

Why by hand: python-docx is not installed, there is no pip and no root here, so a
proper install is an operator decision. Meanwhile this fleet already produces good
docx output with the standard library (skill build-docx-no-libs, and a
verified two-page document), so this module packages a method that works rather
than betting on one that is not available.

The parity floor is deliberate: whatever an agent can produce by hand today, this
must also produce, or the tool is a downgrade and nobody will use it. That means
inline formatting inside a paragraph, both bullet and numbered lists, a table with
a styled repeating header and banded rows, a quote/callout block, images, page
breaks, and a footer with "page / total".

Anything beyond that list is REFUSED with a message naming what is unsupported.
A builder that silently drops a block is the same failure as a text replace that
matches nothing: the caller believes the document contains something it does not.

Input (JSON on stdin):

    {"file": "/tmp/out.docx",
     "title": "optional document title",
     "header": "optional header text",
     "footer": "optional footer text (a page number is added as 'n / m')",
     "blocks": [
       {"type": "heading", "text": "...", "level": 1},
       {"type": "paragraph", "runs": [{"text": "bold bit", "bold": true},
                                       {"text": " plain bit"}]},
       {"type": "paragraph", "text": "shorthand for a single plain run"},
       {"type": "list", "items": ["a", "b"], "ordered": false},
       {"type": "table", "header": ["A", "B"], "rows": [["1", "2"]], "banded": true},
       {"type": "quote", "text": "...", "author": "optional"},
       {"type": "image", "path": "/tmp/chart.png", "caption": "optional"},
       {"type": "page_break"}
     ]}
"""

import json
import os
import struct
import sys
import xml.etree.ElementTree as ET
import zipfile

from office_paths import PathNotAllowed, check_path

EMU_PER_INCH = 914400
# A4 with 2.5 cm margins leaves about 16 cm of usable width.
USABLE_WIDTH_EMU = 5760720

SUPPORTED = ("heading", "paragraph", "list", "table", "quote", "image", "page_break")


class BuildError(Exception):
    """The description cannot be turned into a document."""


def esc(text):
    return (str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


# --- image ------------------------------------------------------------------

def png_size(path):
    """Width and height from the PNG header - no Pillow needed, and no guessing.

    A wrong size here is not cosmetic: the drawing would be stretched, and that is
    the kind of defect nobody reports as a bug, they just stop using the tool.
    """
    with open(path, "rb") as fh:
        signature = fh.read(8)
        if signature != b"\x89PNG\r\n\x1a\n":
            raise BuildError("not a PNG image: %s" % path)
        fh.read(4)
        if fh.read(4) != b"IHDR":
            raise BuildError("PNG header missing (IHDR): %s" % path)
        width, height = struct.unpack(">II", fh.read(8))
    if not width or not height:
        raise BuildError("PNG reports a zero dimension: %s" % path)
    return width, height


# --- runs and paragraphs ----------------------------------------------------

def run_xml(run):
    """One <w:r>, with the inline formatting the parity floor requires."""
    if not isinstance(run, dict):
        run = {"text": run}
    text = esc(run.get("text", ""))
    props = []
    if run.get("bold"):
        props.append("<w:b/>")
    if run.get("italic"):
        props.append("<w:i/>")
    if run.get("underline"):
        props.append('<w:u w:val="single"/>')
    if run.get("highlight"):
        # Word accepts named highlights; an unknown name would be ignored silently,
        # so the value is passed through and documented rather than validated here.
        props.append('<w:highlight w:val="%s"/>' % esc(run["highlight"]))
    if run.get("color"):
        props.append('<w:color w:val="%s"/>' % esc(run["color"]).lstrip("#"))
    if run.get("size"):
        # half-points, as OOXML counts them
        props.append('<w:sz w:val="%d"/>' % (int(run["size"]) * 2))
    rpr = "<w:rPr>%s</w:rPr>" % "".join(props) if props else ""
    return '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr, text)


def paragraph_xml(runs, style=None, extra_ppr=""):
    props = []
    if style:
        props.append('<w:pStyle w:val="%s"/>' % style)
    ppr = "<w:pPr>%s%s</w:pPr>" % ("".join(props), extra_ppr) if (props or extra_ppr) else ""
    return "<w:p>%s%s</w:p>" % (ppr, "".join(run_xml(r) for r in runs))


def block_xml(block, images):
    kind = block.get("type")
    if kind == "heading":
        level = int(block.get("level", 1))
        if level not in (1, 2, 3):
            raise BuildError("heading level must be 1, 2 or 3 (got %r)" % block.get("level"))
        return paragraph_xml([{"text": block.get("text", "")}], style="Heading%d" % level)

    if kind == "paragraph":
        runs = block.get("runs") or [{"text": block.get("text", "")}]
        return paragraph_xml(runs)

    if kind == "list":
        items = block.get("items") or []
        if not items:
            raise BuildError("a list block needs at least one item")
        num_id = 2 if block.get("ordered") else 1
        out = []
        for item in items:
            runs = item if isinstance(item, list) else [{"text": item}]
            numpr = ('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="%d"/></w:numPr>' % num_id)
            out.append(paragraph_xml(runs, style="ListParagraph", extra_ppr=numpr))
        return "".join(out)

    if kind == "table":
        return table_xml(block)

    if kind == "quote":
        # A left bar plus a tint: the same shape as the hand-built callout, so the
        # output does not regress into an ordinary indented paragraph.
        border = ('<w:pBdr><w:left w:val="single" w:sz="18" w:space="8" w:color="BF9000"/></w:pBdr>'
                  '<w:shd w:val="clear" w:fill="FFF7E6"/>'
                  '<w:ind w:left="360"/><w:spacing w:before="120" w:after="120"/>')
        runs = [{"text": block.get("text", ""), "italic": True}]
        if block.get("author"):
            runs.append({"text": "  - " + str(block["author"]), "italic": True, "color": "7F6000"})
        return paragraph_xml(runs, extra_ppr=border)

    if kind == "image":
        return image_xml(block, images)

    if kind == "page_break":
        return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

    raise BuildError(
        "unsupported block type %r. This builder handles: %s. Anything richer "
        "(columns, footnotes, native charts) is deliberately not supported rather "
        "than half-supported - say what you need and it gets designed."
        % (kind, ", ".join(SUPPORTED))
    )


def table_xml(block):
    header = block.get("header") or []
    rows = block.get("rows") or []
    if not header and not rows:
        raise BuildError("a table block needs a header or at least one row")
    width = len(header) if header else len(rows[0])
    if any(len(row) != width for row in rows):
        raise BuildError(
            "every table row must have %d cell(s) to match the header; got %s"
            % (width, [len(r) for r in rows])
        )
    col_width = int(9360 / max(width, 1))  # twentieths of a point across the text area
    grid = "".join('<w:gridCol w:w="%d"/>' % col_width for _ in range(width))

    def cell(text, fill=None, bold=False):
        shading = '<w:shd w:val="clear" w:fill="%s"/>' % fill if fill else ""
        return ('<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/>%s</w:tcPr>%s</w:tc>'
                % (col_width, shading, paragraph_xml([{"text": text, "bold": bold}])))

    out = ['<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>'
           '<w:top w:val="single" w:sz="4" w:color="AAAAAA"/>'
           '<w:left w:val="single" w:sz="4" w:color="AAAAAA"/>'
           '<w:bottom w:val="single" w:sz="4" w:color="AAAAAA"/>'
           '<w:right w:val="single" w:sz="4" w:color="AAAAAA"/>'
           '<w:insideH w:val="single" w:sz="4" w:color="CCCCCC"/>'
           '<w:insideV w:val="single" w:sz="4" w:color="CCCCCC"/>'
           '</w:tblBorders></w:tblPr><w:tblGrid>%s</w:tblGrid>' % grid]

    if header:
        # tblHeader repeats the row when the table breaks across pages - without it
        # a long table's second page has no column titles.
        out.append('<w:tr><w:trPr><w:tblHeader/></w:trPr>'
                   + "".join(cell(text, fill="D9E2F3", bold=True) for text in header)
                   + "</w:tr>")
    banded = block.get("banded", True)
    for index, row in enumerate(rows):
        fill = "F2F2F2" if (banded and index % 2 == 1) else None
        out.append("<w:tr>" + "".join(cell(text, fill=fill) for text in row) + "</w:tr>")
    out.append("</w:tbl>")
    # A table followed directly by another table would merge visually; an empty
    # paragraph keeps them apart, which is what a hand-built document does too.
    out.append("<w:p/>")
    return "".join(out)


def image_xml(block, images):
    path = block.get("path")
    if not path:
        raise BuildError("an image block needs a path")
    path = check_path(path)
    if not os.path.exists(path):
        raise BuildError("image not found: %s" % path)
    width_px, height_px = png_size(path)
    index = len(images) + 1
    images.append(path)
    cx = USABLE_WIDTH_EMU
    cy = int(cx * height_px / width_px)
    drawing = (
        '<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
        '<wp:extent cx="%d" cy="%d"/><wp:docPr id="%d" name="Picture %d"/>'
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:nvPicPr><pic:cNvPr id="%d" name="image%d.png"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="rIdImg%d"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"
        % (cx, cy, index, index, index, index, index, cx, cy)
    )
    if block.get("caption"):
        drawing += paragraph_xml(
            [{"text": str(block["caption"]), "italic": True, "size": 9}],
            extra_ppr='<w:jc w:val="center"/>',
        )
    return drawing


# --- the package ------------------------------------------------------------

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
<w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/>
<w:pPr><w:spacing w:after="120"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/>
<w:pPr><w:spacing w:after="240"/></w:pPr>
<w:rPr><w:b/><w:sz w:val="52"/><w:color w:val="1F3864"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>
<w:pPr><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr>
<w:rPr><w:b/><w:sz w:val="32"/><w:color w:val="1F3864"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/>
<w:pPr><w:spacing w:before="200" w:after="100"/><w:outlineLvl w:val="1"/></w:pPr>
<w:rPr><w:b/><w:sz w:val="26"/><w:color w:val="2E5395"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/>
<w:pPr><w:spacing w:before="160" w:after="80"/><w:outlineLvl w:val="2"/></w:pPr>
<w:rPr><w:b/><w:i/><w:sz w:val="24"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/>
<w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/>
<w:rPr><w:sz w:val="18"/><w:color w:val="7F7F7F"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/>
<w:rPr><w:sz w:val="18"/><w:color w:val="7F7F7F"/></w:rPr></w:style>
</w:styles>"""

# Both list flavours, because "bullet only" would be a downgrade. The bullet glyph
# MUST be spelled out with the Symbol font: an empty lvlText renders no bullet at
# all, which is a documented trap from the hand-built path.
NUMBERING = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/>
<w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/><w:lvlJc w:val="left"/>
<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl></w:abstractNum>
<w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:start w:val="1"/>
<w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>
<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
<w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>
</w:numbering>"""


def page_field():
    """PAGE / NUMPAGES as a field pair, so the footer reads "2 / 7" and stays right."""
    return (
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        "<w:r><w:t>1</w:t></w:r>"
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        '<w:r><w:t xml:space="preserve"> / </w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        "<w:r><w:t>1</w:t></w:r>"
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    )


def build(spec, out_path):
    blocks = spec.get("blocks") or []
    if not blocks and not spec.get("title"):
        raise BuildError("nothing to build: give a title or at least one block")

    images = []
    body = []
    if spec.get("title"):
        body.append(paragraph_xml([{"text": spec["title"]}], style="Title"))
    for block in blocks:
        if not isinstance(block, dict):
            raise BuildError("every block must be an object, got %r" % type(block).__name__)
        body.append(block_xml(block, images))

    header_text = spec.get("header")
    footer_text = spec.get("footer")

    section = (
        '<w:sectPr>'
        + ('<w:headerReference w:type="default" r:id="rIdHdr"/>' if header_text else "")
        + '<w:footerReference w:type="default" r:id="rIdFtr"/>'
        + '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="709" w:footer="709"/>'
        "</w:sectPr>"
    )

    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        "<w:body>" + "".join(body) + section + "</w:body></w:document>"
    )

    footer = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:p><w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="center"/></w:pPr>'
        + (('<w:r><w:t xml:space="preserve">%s   </w:t></w:r>' % esc(footer_text))
           if footer_text else "")
        + page_field()
        + "</w:p></w:ftr>"
    )

    header_part = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr>'
        '<w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p></w:hdr>' % esc(header_text or "")
    )

    rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rIdSty" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships/styles" Target="styles.xml"/>'
            '<Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships/numbering" Target="numbering.xml"/>'
            '<Relationship Id="rIdFtr" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships/footer" Target="footer1.xml"/>']
    if header_text:
        rels.append('<Relationship Id="rIdHdr" Type="http://schemas.openxmlformats.org/'
                    'officeDocument/2006/relationships/header" Target="header1.xml"/>')
    for index in range(1, len(images) + 1):
        rels.append('<Relationship Id="rIdImg%d" Type="http://schemas.openxmlformats.org/'
                    'officeDocument/2006/relationships/image" Target="media/image%d.png"/>'
                    % (index, index))
    rels.append("</Relationships>")

    overrides = ['<Override PartName="/word/document.xml" ContentType="application/vnd.'
                 'openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
                 '<Override PartName="/word/styles.xml" ContentType="application/vnd.'
                 'openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
                 '<Override PartName="/word/numbering.xml" ContentType="application/vnd.'
                 'openxmlformats-officedocument.wordprocessingml.numbering+xml"/>',
                 '<Override PartName="/word/footer1.xml" ContentType="application/vnd.'
                 'openxmlformats-officedocument.wordprocessingml.footer+xml"/>']
    if header_text:
        overrides.append('<Override PartName="/word/header1.xml" ContentType="application/vnd.'
                         'openxmlformats-officedocument.wordprocessingml.header+xml"/>')

    content_types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                     '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                     '<Default Extension="rels" ContentType="application/vnd.openxmlformats-'
                     'package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
                     '<Default Extension="png" ContentType="image/png"/>'
                     + "".join(overrides) + "</Types>")

    package_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
                    'relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
                    'officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
                    "</Relationships>")

    parts = {
        "[Content_Types].xml": content_types,
        "_rels/.rels": package_rels,
        "word/document.xml": document,
        "word/styles.xml": STYLES,
        "word/numbering.xml": NUMBERING,
        "word/footer1.xml": footer,
        "word/_rels/document.xml.rels": "".join(rels),
    }
    if header_text:
        parts["word/header1.xml"] = header_part

    # Well-formedness is checked BEFORE anything is written: a malformed part means
    # LibreOffice refuses the file with "could not be loaded", exits 0 and produces
    # nothing - a failure that arrives as a missing file rather than a reason.
    for name, xml in parts.items():
        try:
            ET.fromstring(xml)
        except ET.ParseError as exc:
            raise BuildError("generated %s is not well-formed XML: %s" % (name, exc))

    tmp = out_path + ".tmp-build"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for name, xml in parts.items():
            out.writestr(name, xml)
        for index, image in enumerate(images, start=1):
            with open(image, "rb") as fh:
                out.writestr("word/media/image%d.png" % index, fh.read())
    os.replace(tmp, out_path)

    counts = {}
    for block in blocks:
        counts[block.get("type")] = counts.get(block.get("type"), 0) + 1
    return {"blocks": counts, "images": len(images)}


def main():
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
        report = build(req, path)
    except (BuildError, PathNotAllowed) as exc:
        json.dump({"ok": False, "error": str(exc)}, sys.stdout)
        return
    except Exception as exc:  # noqa: BLE001
        import traceback
        traceback.print_exc(file=sys.stderr)
        json.dump({"ok": False, "error": "docx build failed: %s" % exc}, sys.stdout)
        return

    json.dump({"ok": True, "file": path, "built": report}, sys.stdout)


if __name__ == "__main__":
    main()
