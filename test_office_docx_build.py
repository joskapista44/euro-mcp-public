#!/usr/bin/env python3
"""Tests for building a .docx from scratch.

Two things are being defended here, and they are different:

  1. THE PARITY FLOOR. The builder exists to replace a hand-written document, so
     every element an agent can produce by hand must come out of it: inline
     formatting inside one paragraph, bullet AND numbered lists, a table with a
     repeating styled header and banded rows, a quote/callout, an image, a page
     break, and a footer with "page / total". A builder that quietly renders a
     numbered list as bullets is a downgrade nobody would notice until the
     document is in front of a client.

  2. THE REFUSALS. Every unsupported or impossible input must come back as one
     clean error that NAMES the problem. A builder that silently drops a block is
     the same failure mode as a text replace that matches nothing: the caller
     believes the file contains something it does not.

The assertions look at the XML the builder actually writes into the package, not
at a return value, because the return value is not what Word reads.

Run:  python3 -m unittest test_office_docx_build
"""

import json
import os
import struct
import subprocess
import sys
import tempfile
import unittest
import zipfile
import zlib

from office_docx_build import BuildError, build

HERE = os.path.dirname(os.path.abspath(__file__))


def make_png(path, width=40, height=20):
    """A real (tiny) PNG. The builder reads the IHDR for the aspect ratio, so a
    fake file with a .png name would not exercise the same path."""
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


EVERY_BLOCK = [
    {"type": "heading", "text": "Elso fejezet", "level": 1},
    {"type": "paragraph", "runs": [
        {"text": "Ez a resz "},
        {"text": "felkover", "bold": True},
        {"text": " es ez "},
        {"text": "dolt", "italic": True},
    ]},
    {"type": "list", "items": ["elso", "masodik"], "ordered": False},
    {"type": "list", "items": ["egy", "kettő"], "ordered": True},
    {"type": "table", "header": ["Tetel", "Osszeg"],
     "rows": [["Alap", "1000"], ["Kamat", "200"], ["Dij", "50"]]},
    {"type": "quote", "text": "A lenyeg egy mondatban.", "author": "Owner"},
    {"type": "page_break"},
    {"type": "heading", "text": "Masodik fejezet", "level": 2},
]


class ParityFloorTests(unittest.TestCase):
    """Every element of the floor must be present in the written package."""

    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.mkdtemp(prefix="office-docx-build-")
        cls.png = os.path.join(cls.tmp, "kep.png")
        make_png(cls.png)
        cls.path = os.path.join(cls.tmp, "full.docx")
        blocks = EVERY_BLOCK + [{"type": "image", "path": cls.png, "caption": "1. abra"}]
        cls.report = build({"title": "Teljes dokumentum", "header": "Example Corp",
                            "footer": "belso hasznalatra", "blocks": blocks}, cls.path)
        with zipfile.ZipFile(cls.path) as z:
            cls.names = z.namelist()
            cls.document = z.read("word/document.xml").decode()
            cls.footer = z.read("word/footer1.xml").decode()
            cls.header = z.read("word/header1.xml").decode()
            cls.rels = z.read("word/_rels/document.xml.rels").decode()
            cls.types = z.read("[Content_Types].xml").decode()

    def test_the_package_has_the_parts_word_needs(self):
        for name in ("[Content_Types].xml", "_rels/.rels", "word/document.xml",
                     "word/styles.xml", "word/numbering.xml", "word/footer1.xml",
                     "word/header1.xml", "word/_rels/document.xml.rels"):
            self.assertIn(name, self.names)

    def test_inline_formatting_survives_INSIDE_one_paragraph(self):
        # Not "the document contains bold somewhere": bold and italic must sit in
        # the same paragraph as the plain text around them.
        para = [p for p in self.document.split("<w:p>") if "felkov" in p]
        self.assertEqual(len(para), 1)
        self.assertIn("<w:b/>", para[0])
        self.assertIn("<w:i/>", para[0])
        self.assertIn("Ez a resz", para[0])

    def test_a_bulleted_and_a_numbered_list_use_DIFFERENT_numbering(self):
        # numId 1 is the bullet definition, 2 the decimal one. If both lists ended
        # up on the same numId, one of them silently renders as the other.
        self.assertIn('<w:numId w:val="1"/>', self.document)
        self.assertIn('<w:numId w:val="2"/>', self.document)

    def test_the_numbering_part_defines_both_a_bullet_and_a_decimal_format(self):
        with zipfile.ZipFile(self.path) as z:
            numbering = z.read("word/numbering.xml").decode()
        self.assertIn('w:val="bullet"', numbering)
        self.assertIn('w:val="decimal"', numbering)

    def test_the_table_header_repeats_and_is_styled_and_the_rows_are_banded(self):
        self.assertIn("<w:tblHeader/>", self.document)          # repeats across pages
        self.assertIn('w:fill="D9E2F3"', self.document)         # header tint
        self.assertIn('w:fill="F2F2F2"', self.document)         # banded row
        self.assertEqual(self.document.count('<w:gridCol w:w='), 2)

    def test_the_quote_is_a_callout_not_just_an_indented_paragraph(self):
        self.assertIn("<w:pBdr>", self.document)
        self.assertIn('w:color="BF9000"', self.document)
        self.assertIn("Owner", self.document)

    def test_the_page_break_is_a_real_break(self):
        self.assertIn('<w:br w:type="page"/>', self.document)

    def test_the_image_is_embedded_and_related_and_typed(self):
        self.assertIn("word/media/image1.png", self.names)
        self.assertIn('r:embed="rIdImg1"', self.document)
        self.assertIn('Target="media/image1.png"', self.rels)
        self.assertIn('Extension="png"', self.types)
        self.assertIn("1. abra", self.document)

    def test_the_image_keeps_its_aspect_ratio_from_the_IHDR(self):
        # 40x20 px must come out twice as wide as it is tall, whatever the width is
        # scaled to. A guessed height is how hand-built documents get squashed pictures.
        extent = self.document.split('<wp:extent ')[1].split("/>")[0]
        cx = int(extent.split('cx="')[1].split('"')[0])
        cy = int(extent.split('cy="')[1].split('"')[0])
        self.assertEqual(round(cx / cy, 2), 2.0)

    def test_the_footer_carries_the_page_number_as_a_FIELD(self):
        # Literal text would print "1 / 1" on every page. The field pair is what
        # makes page 2 say 2.
        self.assertIn(" PAGE ", self.footer)
        self.assertIn(" NUMPAGES ", self.footer)
        self.assertIn("belso hasznalatra", self.footer)

    def test_the_header_text_is_in_the_header_part(self):
        self.assertIn("Example Corp", self.header)

    def test_the_title_is_rendered_with_the_title_style(self):
        self.assertIn('<w:pStyle w:val="Title"/>', self.document)
        self.assertIn("Teljes dokumentum", self.document)

    def test_the_report_counts_what_was_built(self):
        self.assertEqual(self.report["images"], 1)
        self.assertEqual(self.report["blocks"]["list"], 2)
        self.assertEqual(self.report["blocks"]["heading"], 2)


class RefusalTests(unittest.TestCase):
    """Bad input must come back named, not half-applied."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="office-docx-build-bad-")
        self.path = os.path.join(self.tmp, "out.docx")

    def test_an_unknown_block_type_is_refused_and_the_supported_ones_are_named(self):
        with self.assertRaises(BuildError) as ctx:
            build({"blocks": [{"type": "chart", "text": "x"}]}, self.path)
        message = str(ctx.exception)
        self.assertIn("chart", message)
        for supported in ("heading", "paragraph", "list", "table", "quote", "image", "page_break"):
            self.assertIn(supported, message)

    def test_nothing_is_written_when_a_later_block_is_refused(self):
        # The refusal must not leave a half-built document behind: a file that
        # exists is taken as a success by anyone who checks for one.
        with self.assertRaises(BuildError):
            build({"blocks": [{"type": "heading", "text": "jo"},
                              {"type": "hasab", "text": "nem jo"}]}, self.path)
        self.assertFalse(os.path.exists(self.path))
        self.assertFalse(os.path.exists(self.path + ".tmp-build"))

    def test_an_empty_description_is_refused(self):
        with self.assertRaises(BuildError):
            build({}, self.path)

    def test_a_heading_level_outside_1_to_3_is_refused_with_the_value(self):
        with self.assertRaises(BuildError) as ctx:
            build({"blocks": [{"type": "heading", "text": "x", "level": 4}]}, self.path)
        self.assertIn("4", str(ctx.exception))

    def test_a_row_that_does_not_match_the_header_width_is_refused_with_both_widths(self):
        with self.assertRaises(BuildError) as ctx:
            build({"blocks": [{"type": "table", "header": ["A", "B"],
                               "rows": [["1", "2"], ["3"]]}]}, self.path)
        message = str(ctx.exception)
        self.assertIn("2", message)
        self.assertIn("1", message)

    def test_an_empty_table_is_refused(self):
        with self.assertRaises(BuildError):
            build({"blocks": [{"type": "table"}]}, self.path)

    def test_an_empty_list_is_refused(self):
        with self.assertRaises(BuildError):
            build({"blocks": [{"type": "list", "items": []}]}, self.path)

    def test_a_missing_image_is_refused_by_path(self):
        missing = os.path.join(self.tmp, "nincs.png")
        with self.assertRaises(BuildError) as ctx:
            build({"blocks": [{"type": "image", "path": missing}]}, self.path)
        self.assertIn(missing, str(ctx.exception))

    def test_a_file_that_is_not_a_png_is_refused_as_not_a_png(self):
        fake = os.path.join(self.tmp, "kep.png")
        with open(fake, "wb") as fh:
            fh.write(b"\xff\xd8\xff\xe0 JPEG really")
        with self.assertRaises(BuildError) as ctx:
            build({"blocks": [{"type": "image", "path": fake}]}, self.path)
        self.assertIn("PNG", str(ctx.exception))

    def test_an_image_block_without_a_path_is_refused(self):
        with self.assertRaises(BuildError):
            build({"blocks": [{"type": "image"}]}, self.path)

    def test_a_block_that_is_not_an_object_is_refused_by_type(self):
        with self.assertRaises(BuildError) as ctx:
            build({"blocks": ["egy sima string"]}, self.path)
        self.assertIn("str", str(ctx.exception))


class EscapingTests(unittest.TestCase):
    def test_markup_characters_in_the_text_do_not_break_the_package(self):
        # An unescaped "&" is the classic way to produce a file LibreOffice refuses
        # to open while the builder reports success.
        tmp = tempfile.mkdtemp(prefix="office-docx-build-esc-")
        path = os.path.join(tmp, "esc.docx")
        build({"blocks": [
            {"type": "paragraph", "text": "Kovacs & Tarsa <Kft.> \"idezet\""},
            {"type": "table", "header": ["A & B"], "rows": [["<x>"]]},
        ]}, path)
        with zipfile.ZipFile(path) as z:
            document = z.read("word/document.xml").decode()
        self.assertIn("&amp;", document)
        self.assertIn("&lt;", document)
        self.assertNotIn("<Kft.>", document)


class StdinStdoutContractTests(unittest.TestCase):
    """The contract office-mcp.js drives: one JSON in, one JSON object out."""

    def run_helper(self, request, env=None):
        proc = subprocess.run(
            [sys.executable, os.path.join(HERE, "office_docx_build.py")],
            input=json.dumps(request), capture_output=True, text=True,
            env={**os.environ, **(env or {})}, cwd=HERE,
        )
        self.assertEqual(proc.returncode, 0, proc.stderr)
        return json.loads(proc.stdout), proc.stderr

    def test_a_good_request_answers_ok_with_the_path_and_the_counts(self):
        tmp = tempfile.mkdtemp(prefix="office-docx-build-cli-")
        path = os.path.join(tmp, "cli.docx")
        answer, _ = self.run_helper({"file": path, "blocks": [
            {"type": "paragraph", "text": "szoveg"}]})
        self.assertTrue(answer["ok"])
        self.assertEqual(answer["file"], path)
        self.assertEqual(answer["built"]["blocks"], {"paragraph": 1})
        self.assertTrue(zipfile.is_zipfile(path))

    def test_a_refusal_comes_back_as_ok_false_with_ONE_clean_message(self):
        tmp = tempfile.mkdtemp(prefix="office-docx-build-cli-")
        answer, stderr = self.run_helper({"file": os.path.join(tmp, "x.docx"),
                                          "blocks": [{"type": "labjegyzet"}]})
        self.assertFalse(answer["ok"])
        self.assertIn("labjegyzet", answer["error"])
        # A caller must never receive a traceback.
        self.assertNotIn("Traceback", answer["error"])
        self.assertNotIn("Traceback", stderr)

    def test_a_missing_file_argument_is_refused(self):
        answer, _ = self.run_helper({"blocks": [{"type": "paragraph", "text": "x"}]})
        self.assertFalse(answer["ok"])
        self.assertIn("file", answer["error"])

    def test_a_path_outside_the_allowed_roots_is_refused_before_anything_is_written(self):
        # The roots are pinned for the test rather than inherited, so the result does
        # not depend on where this file happens to live.
        allowed = tempfile.mkdtemp(prefix="office-docx-build-in-")
        outside = tempfile.mkdtemp(prefix="office-docx-build-out-")
        target = os.path.join(outside, "must-not-appear.docx")
        answer, _ = self.run_helper({"file": target, "blocks": [
            {"type": "paragraph", "text": "x"}]}, env={"OFFICE_ALLOWED_ROOTS": allowed})
        self.assertFalse(answer["ok"])
        self.assertIn("allowed roots", answer["error"])
        self.assertFalse(os.path.exists(target))

    def test_an_image_outside_the_allowed_roots_is_refused_too(self):
        # The guard has to cover the files the builder READS, not only the one it
        # writes - otherwise the document becomes a way to copy any PNG on the box.
        # The image is a REAL png, so the refusal proves the guard and not the
        # format check or a missing file.
        allowed = tempfile.mkdtemp(prefix="office-docx-build-in-")
        outside = tempfile.mkdtemp(prefix="office-docx-build-out-")
        png = os.path.join(outside, "kep.png")
        make_png(png)
        answer, _ = self.run_helper({"file": os.path.join(allowed, "x.docx"), "blocks": [
            {"type": "image", "path": png}]}, env={"OFFICE_ALLOWED_ROOTS": allowed})
        self.assertFalse(answer["ok"])
        self.assertIn("allowed roots", answer["error"])

    def test_unreadable_request_json_is_refused_cleanly(self):
        proc = subprocess.run([sys.executable, os.path.join(HERE, "office_docx_build.py")],
                              input="not json at all", capture_output=True, text=True, cwd=HERE)
        self.assertEqual(proc.returncode, 0)
        answer = json.loads(proc.stdout)
        self.assertFalse(answer["ok"])
        self.assertIn("JSON", answer["error"])


if __name__ == "__main__":
    unittest.main()
