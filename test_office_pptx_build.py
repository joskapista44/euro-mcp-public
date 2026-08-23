#!/usr/bin/env python3
"""Tests for building a .pptx from scratch.

The same two defences as the docx builder (parity floor + named refusals), plus
three failure modes that only exist in PresentationML and that a hand-built deck
usually gets wrong:

  * THE THREE COUNTS. sldIdLst entries, ppt/slides/slideN.xml parts and the
    [Content_Types].xml overrides must all agree. When they do not, the deck opens
    short a slide - or not at all - and the file itself looks perfectly fine.
  * WHERE THE IMAGE RELATIONSHIP LIVES. It belongs to the SLIDE part, not to
    presentation.xml. Put it in the wrong place and the deck opens with an empty
    frame where the picture should be.
  * OVERFLOW. Content that runs past the bottom edge is not visible and nobody can
    tell from the file that half a table is missing, so it must be REFUSED.

Run:  python3 -m unittest test_office_pptx_build
"""

import json
import os
import re
import struct
import subprocess
import sys
import tempfile
import unittest
import zipfile
import zlib

from office_pptx_build import BuildError, SLIDE_H, build

HERE = os.path.dirname(os.path.abspath(__file__))


def make_png(path, width=40, height=20):
    """A real (tiny) PNG; the builder reads the IHDR for the aspect ratio."""
    def chunk(tag, data):
        payload = tag + data
        return (struct.pack(">I", len(data)) + payload
                + struct.pack(">I", zlib.crc32(payload) & 0xFFFFFFFF))

    raw = b"".join(b"\x00" + bytes([30, 90, 160] * width) for _ in range(height))
    with open(path, "wb") as fh:
        fh.write(b"\x89PNG\r\n\x1a\n"
                 + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
                 + chunk(b"IDAT", zlib.compress(raw))
                 + chunk(b"IEND", b""))


class ParityFloorTests(unittest.TestCase):
    """Every element of the floor must be present in the written package."""

    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.mkdtemp(prefix="office-pptx-build-")
        cls.png = os.path.join(cls.tmp, "kep.png")
        make_png(cls.png)
        cls.path = os.path.join(cls.tmp, "full.pptx")
        cls.report = build({
            "file": cls.path,
            "title": "Teljes prezentacio",
            "subtitle": "alcim",
            "slides": [
                {"title": "Szoveg es listak", "blocks": [
                    {"type": "text", "runs": [
                        {"text": "sima "},
                        {"text": "felkover", "bold": True},
                        {"text": " es dolt", "italic": True},
                    ]},
                    {"type": "list", "items": ["elso", "masodik"]},
                    {"type": "list", "items": ["egy", "ketto"], "ordered": True},
                ]},
                {"title": "Tabla es callout", "blocks": [
                    {"type": "table", "header": ["Tetel", "Osszeg"],
                     "rows": [["Alap", "1000"], ["Kamat", "200"]]},
                    {"type": "callout", "text": "Ez az egy dolog, amit vinni kell."},
                ]},
                {"title": "Kep", "blocks": [{"type": "image", "path": cls.png}]},
            ],
        }, cls.path)
        with zipfile.ZipFile(cls.path) as z:
            cls.names = z.namelist()
            cls.presentation = z.read("ppt/presentation.xml").decode()
            cls.pres_rels = z.read("ppt/_rels/presentation.xml.rels").decode()
            cls.types = z.read("[Content_Types].xml").decode()
            cls.slides = [z.read("ppt/slides/slide%d.xml" % i).decode()
                          for i in range(1, cls.report["slides"] + 1)]
            cls.slide_rels = [z.read("ppt/slides/_rels/slide%d.xml.rels" % i).decode()
                              for i in range(1, cls.report["slides"] + 1)]

    def test_the_title_becomes_a_COVER_slide_in_front_of_the_others(self):
        self.assertEqual(self.report["slides"], 4)
        self.assertIn("Teljes prezentacio", self.slides[0])
        self.assertIn("alcim", self.slides[0])

    def test_the_package_has_the_parts_powerpoint_needs(self):
        for name in ("[Content_Types].xml", "_rels/.rels", "ppt/presentation.xml",
                     "ppt/_rels/presentation.xml.rels", "ppt/slideMasters/slideMaster1.xml",
                     "ppt/slideLayouts/slideLayout1.xml", "ppt/theme/theme1.xml"):
            self.assertIn(name, self.names)

    def test_the_THREE_COUNTS_agree(self):
        listed = self.presentation.count("<p:sldId ")
        parts = len([n for n in self.names if re.match(r"ppt/slides/slide\d+\.xml$", n)])
        typed = self.types.count("/ppt/slides/slide")
        self.assertEqual((listed, parts, typed), (4, 4, 4))

    def test_every_slide_is_related_from_presentation_xml_and_relates_to_the_layout(self):
        for index in range(1, 5):
            self.assertIn('Target="slides/slide%d.xml"' % index, self.pres_rels)
            self.assertIn("slideLayout1.xml", self.slide_rels[index - 1])

    def test_every_slide_has_a_title(self):
        for index, expected in enumerate(
                ["Teljes prezentacio", "Szoveg es listak", "Tabla es callout", "Kep"]):
            self.assertIn(expected, self.slides[index])

    def test_inline_formatting_survives_inside_one_text_block(self):
        slide = self.slides[1]
        self.assertIn('b="1"', slide)
        self.assertIn('i="1"', slide)
        self.assertIn("sima ", slide)

    def test_a_bulleted_and_a_numbered_list_use_DIFFERENT_markers(self):
        # buChar is a bullet, buAutoNum is real automatic numbering. If the ordered
        # list came out with buChar, it silently renders as bullets.
        slide = self.slides[1]
        self.assertIn("<a:buChar", slide)
        self.assertIn('<a:buAutoNum type="arabicPeriod"/>', slide)

    def test_the_table_is_a_REAL_table_not_a_picture_of_one(self):
        slide = self.slides[2]
        self.assertIn("<a:tbl>", slide)
        self.assertIn("graphicFrame", slide)
        self.assertEqual(slide.count("<a:gridCol"), 2)
        self.assertEqual(slide.count("<a:tr "), 3)          # header + 2 rows

    def test_the_table_header_is_styled_and_the_rows_are_banded_with_EXPLICIT_fills(self):
        # A tableStyleId without a tableStyles.xml part renders as an unstyled grid
        # in some viewers, so the fills have to be per cell.
        slide = self.slides[2]
        self.assertIn('<a:srgbClr val="1F4E79"/>', slide)   # header fill
        self.assertIn('<a:srgbClr val="EDF2F9"/>', slide)   # banded row
        self.assertIn('firstRow="1"', slide)

    def test_the_callout_is_a_filled_outlined_box(self):
        slide = self.slides[2]
        self.assertIn('<a:srgbClr val="FFF3D6"/>', slide)   # box fill
        self.assertIn('<a:srgbClr val="E0A800"/>', slide)   # box outline
        self.assertIn("Ez az egy dolog", slide)

    def test_the_image_relationship_belongs_to_the_SLIDE_not_to_presentation_xml(self):
        self.assertIn("ppt/media/image1.png", self.names)
        self.assertIn('Target="../media/image1.png"', self.slide_rels[3])
        self.assertIn('r:embed="rIdImg1"', self.slides[3])
        self.assertNotIn("media/image1.png", self.pres_rels)

    def test_the_image_keeps_its_aspect_ratio_from_the_IHDR(self):
        pic = self.slides[3].split("<p:pic>")[1]
        ext = pic.split("<a:ext ")[1].split("/>")[0]
        cx = int(ext.split('cx="')[1].split('"')[0])
        cy = int(ext.split('cy="')[1].split('"')[0])
        self.assertEqual(round(cx / cy, 2), 2.0)

    def test_the_png_content_type_is_declared(self):
        self.assertIn('Extension="png"', self.types)

    def test_the_report_counts_slides_and_images(self):
        self.assertEqual(self.report, {"slides": 4, "images": 1})


class RefusalTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="office-pptx-build-bad-")
        self.path = os.path.join(self.tmp, "out.pptx")

    def test_an_unknown_block_type_is_refused_and_the_supported_ones_are_named(self):
        with self.assertRaises(BuildError) as ctx:
            build({"slides": [{"title": "x", "blocks": [{"type": "animation"}]}]}, self.path)
        message = str(ctx.exception)
        self.assertIn("animation", message)
        for supported in ("text", "list", "table", "callout", "image"):
            self.assertIn(supported, message)

    def test_a_slide_whose_content_runs_off_the_bottom_is_refused_BY_NAME(self):
        # 40 list items cannot fit; the refusal must name the slide so the caller
        # knows which one to split.
        with self.assertRaises(BuildError) as ctx:
            build({"slides": [{"title": "Tulcsordulo dia",
                               "blocks": [{"type": "list", "items": ["sor"] * 40}]}]}, self.path)
        message = str(ctx.exception)
        self.assertIn("Tulcsordulo dia", message)
        self.assertIn("Split", message)
        self.assertFalse(os.path.exists(self.path))

    def test_a_slide_that_only_just_fits_is_NOT_refused(self):
        # The overflow gate must not be a blanket "long content is refused": if it
        # fired too early it would be indistinguishable from a broken builder.
        build({"slides": [{"title": "Eppen elfer",
                           "blocks": [{"type": "list", "items": ["sor"] * 8}]}]}, self.path)
        self.assertTrue(zipfile.is_zipfile(self.path))

    def test_an_empty_description_is_refused(self):
        with self.assertRaises(BuildError):
            build({}, self.path)

    def test_a_row_that_does_not_match_the_header_width_is_refused(self):
        with self.assertRaises(BuildError) as ctx:
            build({"slides": [{"title": "x", "blocks": [
                {"type": "table", "header": ["A", "B"], "rows": [["1"]]}]}]}, self.path)
        self.assertIn("2", str(ctx.exception))

    def test_an_empty_table_is_refused(self):
        with self.assertRaises(BuildError):
            build({"slides": [{"title": "x", "blocks": [{"type": "table"}]}]}, self.path)

    def test_an_empty_list_is_refused(self):
        with self.assertRaises(BuildError):
            build({"slides": [{"title": "x", "blocks": [
                {"type": "list", "items": []}]}]}, self.path)

    def test_a_missing_image_is_refused_by_path(self):
        missing = os.path.join(self.tmp, "nincs.png")
        with self.assertRaises(BuildError) as ctx:
            build({"slides": [{"title": "x", "blocks": [
                {"type": "image", "path": missing}]}]}, self.path)
        self.assertIn(missing, str(ctx.exception))

    def test_a_file_that_is_not_a_png_is_refused_as_not_a_png(self):
        fake = os.path.join(self.tmp, "kep.png")
        with open(fake, "wb") as fh:
            fh.write(b"GIF89a not a png")
        with self.assertRaises(BuildError) as ctx:
            build({"slides": [{"title": "x", "blocks": [
                {"type": "image", "path": fake}]}]}, self.path)
        self.assertIn("PNG", str(ctx.exception))

    def test_a_block_that_is_not_an_object_is_refused_by_type(self):
        with self.assertRaises(BuildError) as ctx:
            build({"slides": [{"title": "x", "blocks": ["sima string"]}]}, self.path)
        self.assertIn("str", str(ctx.exception))

    def test_nothing_is_written_when_a_later_slide_is_refused(self):
        with self.assertRaises(BuildError):
            build({"slides": [{"title": "jo", "blocks": [{"type": "text", "text": "ok"}]},
                              {"title": "nem jo", "blocks": [{"type": "diagram"}]}]}, self.path)
        self.assertFalse(os.path.exists(self.path))
        self.assertFalse(os.path.exists(self.path + ".tmp-build"))


class ImageFittingTests(unittest.TestCase):
    def test_a_tall_image_is_shrunk_to_stay_ON_the_slide(self):
        # A picture taller than the remaining space must be scaled down, not cropped
        # by the slide edge - a crop is invisible in the file and obvious on screen.
        tmp = tempfile.mkdtemp(prefix="office-pptx-build-img-")
        png = os.path.join(tmp, "magas.png")
        make_png(png, width=20, height=200)
        path = os.path.join(tmp, "tall.pptx")
        build({"slides": [{"title": "Magas kep", "blocks": [
            {"type": "image", "path": png}]}]}, path)
        with zipfile.ZipFile(path) as z:
            slide = z.read("ppt/slides/slide1.xml").decode()
        pic = slide.split("<p:pic>")[1]
        off = pic.split("<a:off ")[1].split("/>")[0]
        ext = pic.split("<a:ext ")[1].split("/>")[0]
        y = int(off.split('y="')[1].split('"')[0])
        cy = int(ext.split('cy="')[1].split('"')[0])
        self.assertLessEqual(y + cy, SLIDE_H)


class EscapingTests(unittest.TestCase):
    def test_markup_characters_in_the_text_do_not_break_the_package(self):
        tmp = tempfile.mkdtemp(prefix="office-pptx-build-esc-")
        path = os.path.join(tmp, "esc.pptx")
        build({"slides": [{"title": "A & B <Kft.>", "blocks": [
            {"type": "text", "text": "1 < 2 & mindig"},
            {"type": "table", "header": ["A & B"], "rows": [["<x>"]]},
        ]}]}, path)
        with zipfile.ZipFile(path) as z:
            slide = z.read("ppt/slides/slide1.xml").decode()
        self.assertIn("&amp;", slide)
        self.assertIn("&lt;", slide)
        self.assertNotIn("<Kft.>", slide)


class StdinStdoutContractTests(unittest.TestCase):
    """The contract office-mcp.js drives: one JSON in, one JSON object out."""

    def run_helper(self, request, env=None):
        proc = subprocess.run(
            [sys.executable, os.path.join(HERE, "office_pptx_build.py")],
            input=json.dumps(request), capture_output=True, text=True, cwd=HERE,
            env={**os.environ, **(env or {})},
        )
        self.assertEqual(proc.returncode, 0, proc.stderr)
        return json.loads(proc.stdout), proc.stderr

    def test_a_good_request_answers_ok_with_the_path_and_the_counts(self):
        tmp = tempfile.mkdtemp(prefix="office-pptx-build-cli-")
        path = os.path.join(tmp, "cli.pptx")
        answer, _ = self.run_helper({"file": path, "slides": [
            {"title": "Egy dia", "blocks": [{"type": "text", "text": "szoveg"}]}]})
        self.assertTrue(answer["ok"])
        self.assertEqual(answer["file"], path)
        self.assertEqual(answer["built"]["slides"], 1)
        self.assertTrue(zipfile.is_zipfile(path))

    def test_a_refusal_comes_back_as_ok_false_with_ONE_clean_message(self):
        tmp = tempfile.mkdtemp(prefix="office-pptx-build-cli-")
        answer, stderr = self.run_helper({"file": os.path.join(tmp, "x.pptx"), "slides": [
            {"title": "x", "blocks": [{"type": "smartart"}]}]})
        self.assertFalse(answer["ok"])
        self.assertIn("smartart", answer["error"])
        self.assertNotIn("Traceback", answer["error"])
        self.assertNotIn("Traceback", stderr)

    def test_a_missing_file_argument_is_refused(self):
        answer, _ = self.run_helper({"slides": [{"title": "x"}]})
        self.assertFalse(answer["ok"])
        self.assertIn("file", answer["error"])

    def test_a_path_outside_the_allowed_roots_is_refused_before_anything_is_written(self):
        # Roots pinned for the test, so the result does not depend on where this
        # file happens to live.
        allowed = tempfile.mkdtemp(prefix="office-pptx-build-in-")
        outside = tempfile.mkdtemp(prefix="office-pptx-build-out-")
        target = os.path.join(outside, "must-not-appear.pptx")
        answer, _ = self.run_helper({"file": target, "slides": [{"title": "x"}]},
                                    env={"OFFICE_ALLOWED_ROOTS": allowed})
        self.assertFalse(answer["ok"])
        self.assertIn("allowed roots", answer["error"])
        self.assertFalse(os.path.exists(target))

    def test_an_image_outside_the_allowed_roots_is_refused_too(self):
        # A REAL png outside the roots: the refusal proves the guard, not the format
        # check or a missing file.
        allowed = tempfile.mkdtemp(prefix="office-pptx-build-in-")
        outside = tempfile.mkdtemp(prefix="office-pptx-build-out-")
        png = os.path.join(outside, "kep.png")
        make_png(png)
        answer, _ = self.run_helper({"file": os.path.join(allowed, "x.pptx"), "slides": [
            {"title": "x", "blocks": [{"type": "image", "path": png}]}]},
            env={"OFFICE_ALLOWED_ROOTS": allowed})
        self.assertFalse(answer["ok"])
        self.assertIn("allowed roots", answer["error"])

    def test_unreadable_request_json_is_refused_cleanly(self):
        proc = subprocess.run([sys.executable, os.path.join(HERE, "office_pptx_build.py")],
                              input="{{{", capture_output=True, text=True, cwd=HERE)
        self.assertEqual(proc.returncode, 0)
        answer = json.loads(proc.stdout)
        self.assertFalse(answer["ok"])
        self.assertIn("JSON", answer["error"])


if __name__ == "__main__":
    unittest.main()
