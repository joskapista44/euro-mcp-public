#!/usr/bin/env python3
"""Build a .pptx from a small structured description, with no third-party library.

Same reasoning as the docx builder: python-pptx is not installed, there is no pip
and no root, and this fleet already produces good decks by hand (skill
build-pptx-no-libs; an eight-slide deck rendered correctly through our own
engine). This packages that method instead of waiting for a dependency.

Every element is absolutely positioned with its own a:xfrm rather than inheriting
from a placeholder. One master and one blank layout are enough that way, and it is
by far the most stable route for generated decks - placeholder inheritance is
where hand-built files usually start rendering differently in different viewers.

The parity floor from the docx side applies here too: text with inline formatting,
bullet and numbered lists, a REAL table (not a picture of one), a callout box,
images, and a title on every slide. Anything else is refused by name.

Input (JSON on stdin):

    {"file": "/tmp/deck.pptx",
     "title": "optional title slide heading",
     "subtitle": "optional",
     "slides": [
       {"title": "Slide title",
        "blocks": [
          {"type": "text", "runs": [{"text": "bold", "bold": true}, {"text": " rest"}]},
          {"type": "list", "items": ["a", "b"], "ordered": false},
          {"type": "table", "header": ["A", "B"], "rows": [["1", "2"]]},
          {"type": "callout", "text": "the one thing to remember"},
          {"type": "image", "path": "/tmp/chart.png"}
        ]}
     ]}
"""

import json
import os
import struct
import sys
import xml.etree.ElementTree as ET
import zipfile

from office_paths import PathNotAllowed, check_path

SLIDE_W, SLIDE_H = 12192000, 6858000          # 16:9, in EMU
MARGIN = 838200                                # ~0.92", the safe side margin
CONTENT_W = SLIDE_W - 2 * MARGIN
TITLE_TOP, TITLE_H = 520000, 900000
BODY_TOP = TITLE_TOP + TITLE_H + 200000
BODY_H = SLIDE_H - BODY_TOP - 500000

SUPPORTED = ("text", "list", "table", "callout", "image")


class BuildError(Exception):
    """The description cannot be turned into a deck."""


def esc(text):
    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def png_size(path):
    with open(path, "rb") as fh:
        if fh.read(8) != b"\x89PNG\r\n\x1a\n":
            raise BuildError("not a PNG image: %s" % path)
        fh.read(4)
        if fh.read(4) != b"IHDR":
            raise BuildError("PNG header missing (IHDR): %s" % path)
        return struct.unpack(">II", fh.read(8))


def run_xml(run, default_size=1600, default_color="303030"):
    if not isinstance(run, dict):
        run = {"text": run}
    props = ['lang="hu-HU"', 'sz="%d"' % int(run.get("size", default_size / 100) * 100)]
    if run.get("bold"):
        props.append('b="1"')
    if run.get("italic"):
        props.append('i="1"')
    if run.get("underline"):
        props.append('u="sng"')
    color = str(run.get("color", default_color)).lstrip("#")
    return ('<a:r><a:rPr %s><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:rPr>'
            "<a:t>%s</a:t></a:r>" % (" ".join(props), esc(color), esc(run.get("text", ""))))


def textbox(x, y, cx, cy, paragraphs, anchor="t", fill=None, line=None):
    body = "".join(paragraphs)
    fill_xml = ('<a:solidFill><a:srgbClr val="%s"/></a:solidFill>' % fill.lstrip("#")
                if fill else "<a:noFill/>")
    line_xml = ('<a:ln w="28575"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:ln>'
                % line.lstrip("#") if line else "")
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="tb%d"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 8000"/></a:avLst></a:prstGeom>'
        "%s%s</p:spPr>"
        '<p:txBody><a:bodyPr wrap="square" anchor="%s" lIns="137160" rIns="137160"/><a:lstStyle/>%s</p:txBody>'
        "</p:sp>" % (textbox.counter(), textbox.counter(peek=True), x, y, cx, cy,
                     fill_xml, line_xml, anchor, body)
    )


def _counter():
    state = {"n": 1}

    def next_id(peek=False):
        if peek:
            return state["n"]
        state["n"] += 1
        return state["n"]
    return next_id


textbox.counter = _counter()


def paragraph(runs, bullet=None, size=1600, color="303030", align="l", space_after=600):
    if bullet == "bullet":
        marker = '<a:buChar char="&#9679;"/>'
        indent = ' marL="285750" indent="-274320"'
    elif bullet == "number":
        marker = '<a:buAutoNum type="arabicPeriod"/>'
        indent = ' marL="285750" indent="-274320"'
    else:
        marker = "<a:buNone/>"
        indent = ""
    return ('<a:p><a:pPr algn="%s"%s><a:spcAft><a:spcPts val="%d"/></a:spcAft>%s</a:pPr>%s</a:p>'
            % (align, indent, space_after, marker,
               "".join(run_xml(r, size, color) for r in runs)))


def table_xml(block, x, y, cx):
    header = block.get("header") or []
    rows = block.get("rows") or []
    if not header and not rows:
        raise BuildError("a table block needs a header or at least one row")
    width = len(header) if header else len(rows[0])
    if any(len(row) != width for row in rows):
        raise BuildError("every table row must have %d cell(s) to match the header" % width)
    col = int(cx / width)
    row_h = 370000
    total_h = row_h * (len(rows) + (1 if header else 0))

    def cell(text, fill, bold, color):
        return ('<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>%s</a:txBody>'
                '<a:tcPr marL="91440" marR="91440" anchor="ctr">'
                '<a:lnL w="12700"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:lnL>'
                '<a:lnR w="12700"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:lnR>'
                '<a:lnT w="12700"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:lnT>'
                '<a:lnB w="12700"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:lnB>'
                '<a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:tcPr></a:tc>'
                % (paragraph([{"text": text, "bold": bold}], size=1300, color=color,
                             space_after=0), fill))

    trs = []
    if header:
        # Explicit per-cell fills instead of a tableStyleId: a style id without a
        # tableStyles.xml part renders as an unstyled grid in some viewers.
        trs.append('<a:tr h="%d">%s</a:tr>'
                   % (row_h, "".join(cell(t, "1F4E79", True, "FFFFFF") for t in header)))
    for index, row in enumerate(rows):
        fill = "EDF2F9" if index % 2 else "FFFFFF"
        trs.append('<a:tr h="%d">%s</a:tr>'
                   % (row_h, "".join(cell(t, fill, False, "303030") for t in row)))

    return ('<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="%d" name="table"/>'
            "<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>"
            '<p:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></p:xfrm>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
            '<a:tbl><a:tblPr firstRow="1" bandRow="1"/><a:tblGrid>%s</a:tblGrid>%s</a:tbl>'
            "</a:graphicData></a:graphic></p:graphicFrame>"
            % (textbox.counter(), x, y, cx, total_h,
               "".join('<a:gridCol w="%d"/>' % col for _ in range(width)), "".join(trs))), total_h


def picture_xml(rel_id, x, y, cx, cy, index):
    return ('<p:pic><p:nvPicPr><p:cNvPr id="%d" name="image%d"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>'
            '<p:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
            '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
            % (textbox.counter(), index, rel_id, x, y, cx, cy))


def slide_xml(slide, images):
    """One slide. Blocks stack down the body area; the caller keeps them short."""
    shapes = []
    title = slide.get("title")
    if title:
        shapes.append(textbox(MARGIN, TITLE_TOP, CONTENT_W, TITLE_H,
                              [paragraph([{"text": title, "bold": True}], size=3200,
                                         color="1F4E79", space_after=0)],
                              anchor="ctr"))

    y = BODY_TOP
    rels = []
    for block in slide.get("blocks") or []:
        if not isinstance(block, dict):
            raise BuildError("every block must be an object, got %r" % type(block).__name__)
        kind = block.get("type")

        if kind == "text":
            runs = block.get("runs") or [{"text": block.get("text", "")}]
            height = 500000
            shapes.append(textbox(MARGIN, y, CONTENT_W, height, [paragraph(runs)]))
            y += height + 80000

        elif kind == "list":
            items = block.get("items") or []
            if not items:
                raise BuildError("a list block needs at least one item")
            marker = "number" if block.get("ordered") else "bullet"
            paragraphs = [paragraph(item if isinstance(item, list) else [{"text": item}],
                                    bullet=marker) for item in items]
            height = 340000 * len(items) + 120000
            shapes.append(textbox(MARGIN, y, CONTENT_W, height, paragraphs))
            y += height + 80000

        elif kind == "table":
            frame, height = table_xml(block, MARGIN, y, CONTENT_W)
            shapes.append(frame)
            y += height + 140000

        elif kind == "callout":
            height = 620000
            shapes.append(textbox(
                MARGIN, y, CONTENT_W, height,
                [paragraph([{"text": block.get("text", ""), "bold": True}],
                           size=1700, color="7F4F00", align="ctr", space_after=0)],
                anchor="ctr", fill="FFF3D6", line="E0A800"))
            y += height + 100000

        elif kind == "image":
            path = block.get("path")
            if not path:
                raise BuildError("an image block needs a path")
            path = check_path(path)
            if not os.path.exists(path):
                raise BuildError("image not found: %s" % path)
            width_px, height_px = png_size(path)
            cx = min(CONTENT_W, int(CONTENT_W * float(block.get("scale", 1.0))))
            cy = int(cx * height_px / width_px)
            remaining = SLIDE_H - y - 400000
            if cy > remaining > 0:
                # Keep the picture on the slide rather than letting it run off the
                # bottom edge, where nobody would see that it was cropped.
                cy, cx = remaining, int(remaining * width_px / height_px)
            images.append(path)
            index = len(images)
            rels.append(index)
            shapes.append(picture_xml("rIdImg%d" % index, MARGIN, y, cx, cy, index))
            y += cy + 100000

        else:
            raise BuildError(
                "unsupported block type %r. This builder handles: %s. Anything richer "
                "(animation, native charts, speaker notes) is deliberately not supported "
                "rather than half-supported." % (kind, ", ".join(SUPPORTED)))

        if y > SLIDE_H:
            raise BuildError(
                "the blocks on slide %r do not fit: content runs past the bottom edge. "
                "Split them across two slides - silently overflowing would hide half of it."
                % (title or "(untitled)"))

    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
            '<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
            "<a:effectLst/></p:bgPr></p:bg>"
            '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            "<p:grpSpPr/>" + "".join(shapes) + "</p:spTree></p:cSld>"
            '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'), rels


THEME = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
<a:themeElements>
<a:clrScheme name="M"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="1F4E79"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
<a:accent1><a:srgbClr val="1F4E79"/></a:accent1><a:accent2><a:srgbClr val="2E75B6"/></a:accent2>
<a:accent3><a:srgbClr val="E0A800"/></a:accent3><a:accent4><a:srgbClr val="70AD47"/></a:accent4>
<a:accent5><a:srgbClr val="C00000"/></a:accent5><a:accent6><a:srgbClr val="7F7F7F"/></a:accent6>
<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>
<a:fontScheme name="M"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>
<a:fmtScheme name="M">
<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
<a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
<a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle>
<a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
</a:fmtScheme></a:themeElements></a:theme>"""

MASTER = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr/></p:spTree></p:cSld>
<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"
 accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rIdL1"/></p:sldLayoutIdLst>
</p:sldMaster>"""

LAYOUT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">
<p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>"""


def build(spec, out_path):
    slides = list(spec.get("slides") or [])
    if spec.get("title"):
        cover = {"title": spec["title"], "blocks": []}
        if spec.get("subtitle"):
            cover["blocks"].append({"type": "text", "runs": [
                {"text": spec["subtitle"], "size": 20, "color": "5A5A5A"}]})
        slides.insert(0, cover)
    if not slides:
        raise BuildError("nothing to build: give a title or at least one slide")

    images, parts, slide_rels = [], {}, []
    for index, slide in enumerate(slides, start=1):
        before = len(images)
        xml, _ = slide_xml(slide, images)
        parts["ppt/slides/slide%d.xml" % index] = xml
        used = list(range(before + 1, len(images) + 1))
        slide_rels.append(used)

    for index, used in enumerate(slide_rels, start=1):
        rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rIdLay" Type="http://schemas.openxmlformats.org/officeDocument/'
                '2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>']
        for image_index in used:
            # The image relationship belongs to the SLIDE, not to presentation.xml -
            # putting it in the wrong part is the classic way to get a deck that
            # opens with an empty frame where the picture should be.
            rels.append('<Relationship Id="rIdImg%d" Type="http://schemas.openxmlformats.org/'
                        'officeDocument/2006/relationships/image" Target="../media/image%d.png"/>'
                        % (image_index, image_index))
        rels.append("</Relationships>")
        parts["ppt/slides/_rels/slide%d.xml.rels" % index] = "".join(rels)

    slide_ids = "".join('<p:sldId id="%d" r:id="rIdSld%d"/>' % (255 + i, i)
                        for i in range(1, len(slides) + 1))
    parts["ppt/presentation.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMst"/></p:sldMasterIdLst>'
        "<p:sldIdLst>%s</p:sldIdLst>"
        '<p:sldSz cx="%d" cy="%d"/><p:notesSz cx="%d" cy="%d"/></p:presentation>'
        % (slide_ids, SLIDE_W, SLIDE_H, SLIDE_H, SLIDE_W))

    pres_rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rIdMst" Type="http://schemas.openxmlformats.org/officeDocument/'
                 '2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>'
                 '<Relationship Id="rIdThm" Type="http://schemas.openxmlformats.org/officeDocument/'
                 '2006/relationships/theme" Target="theme/theme1.xml"/>']
    for index in range(1, len(slides) + 1):
        pres_rels.append('<Relationship Id="rIdSld%d" Type="http://schemas.openxmlformats.org/'
                         'officeDocument/2006/relationships/slide" Target="slides/slide%d.xml"/>'
                         % (index, index))
    pres_rels.append("</Relationships>")
    parts["ppt/_rels/presentation.xml.rels"] = "".join(pres_rels)

    parts["ppt/slideMasters/slideMaster1.xml"] = MASTER
    parts["ppt/slideMasters/_rels/slideMaster1.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rIdL1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
        '<Relationship Id="rIdThm" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/theme" Target="../theme/theme1.xml"/></Relationships>')
    parts["ppt/slideLayouts/slideLayout1.xml"] = LAYOUT
    parts["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rIdMst" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>')
    parts["ppt/theme/theme1.xml"] = THEME
    parts["_rels/.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>')

    overrides = ['<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.'
                 'openxmlformats-officedocument.presentationml.presentation.main+xml"/>',
                 '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/'
                 'vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>',
                 '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/'
                 'vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>',
                 '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.'
                 'openxmlformats-officedocument.theme+xml"/>']
    for index in range(1, len(slides) + 1):
        overrides.append('<Override PartName="/ppt/slides/slide%d.xml" ContentType="application/'
                         'vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' % index)
    parts["[Content_Types].xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="png" ContentType="image/png"/>' + "".join(overrides) + "</Types>")

    for name, xml in parts.items():
        try:
            ET.fromstring(xml)
        except ET.ParseError as exc:
            raise BuildError("generated %s is not well-formed XML: %s" % (name, exc))

    # The three counts that must agree, checked rather than assumed: a deck whose
    # sldIdLst, slide parts and content-type overrides disagree opens short a slide
    # or not at all, and the file itself looks fine.
    listed = parts["ppt/presentation.xml"].count("<p:sldId ")
    files = len([n for n in parts if n.startswith("ppt/slides/slide")])
    typed = parts["[Content_Types].xml"].count("/ppt/slides/slide")
    if not listed == files == typed == len(slides):
        raise BuildError("internal inconsistency: %d listed, %d slide parts, %d content types"
                         % (listed, files, typed))

    tmp = out_path + ".tmp-build"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for name, xml in parts.items():
            out.writestr(name, xml)
        for index, image in enumerate(images, start=1):
            with open(image, "rb") as fh:
                out.writestr("ppt/media/image%d.png" % index, fh.read())
    os.replace(tmp, out_path)
    return {"slides": len(slides), "images": len(images)}


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
        json.dump({"ok": False, "error": "pptx build failed: %s" % exc}, sys.stdout)
        return

    json.dump({"ok": True, "file": path, "built": report}, sys.stdout)


if __name__ == "__main__":
    main()
