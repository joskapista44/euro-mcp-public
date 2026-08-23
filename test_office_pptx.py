#!/usr/bin/env python3
"""Tests for editing an existing pptx.

Two adversarial fixtures carry most of the weight, and both are checked to BE
adversarial before anything else is tested:

  SPLIT_PARAGRAPH - a value stored across two runs, the way PowerPoint writes it
    when part of it is formatted differently. A plain string replace over the XML
    finds nothing there and reports success.

  REORDERED DECK - a deck whose slide FILE names do not match the order the viewer
    shows: slide1.xml is displayed third. Addressing slides by file name silently
    edits a different slide than the caller meant, and the result looks fine.

Run:  python3 -m unittest test_office_pptx
"""

import json
import os
import subprocess
import sys
import tempfile
import unittest
import zipfile

from office_pptx import (
    PptxError,
    apply_edits,
    insert_run,
    read_slides,
    replace_in_slide,
    set_table_cell,
    slide_parts,
    slide_text,
)

HERE = os.path.dirname(os.path.abspath(__file__))

# "12 500 000" exists only as "12 5" + "00 000": no run contains it whole.
SPLIT_PARAGRAPH = ('<a:p><a:r><a:rPr lang="hu-HU"/><a:t>A hitel osszege 12 5</a:t></a:r>'
                   '<a:r><a:rPr lang="hu-HU" b="1"/><a:t>00 000</a:t></a:r>'
                   '<a:r><a:rPr lang="hu-HU"/><a:t> Ft.</a:t></a:r></a:p>')

WHOLE_PARAGRAPH = '<a:p><a:r><a:t>A hitel osszege 12 500 000 Ft.</a:t></a:r></a:p>'

TABLE = (
    '<a:tbl><a:tblPr firstRow="1"/><a:tblGrid><a:gridCol w="100"/><a:gridCol w="100"/></a:tblGrid>'
    '<a:tr h="100"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>Nev</a:t></a:r></a:p></a:txBody>'
    '<a:tcPr/></a:tc>'
    '<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>Ertek</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>'
    '<a:tr h="100"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>Osszeg</a:t></a:r></a:p></a:txBody>'
    '<a:tcPr/></a:tc>'
    '<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>regi</a:t></a:r><a:r><a:t>-maradek</a:t></a:r>'
    '</a:p></a:txBody><a:tcPr/></a:tc></a:tr>'
    '<a:tr h="100"><a:tc><a:txBody><a:bodyPr/><a:p><a:endParaRPr lang="hu-HU"/></a:p></a:txBody>'
    '<a:tcPr/></a:tc>'
    '<a:tc><a:txBody><a:bodyPr/><a:p/></a:txBody><a:tcPr/></a:tc></a:tr>'
    '</a:tbl>'
)


def slide_xml(*body):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
            '<p:cSld><p:spTree>'
            '<p:sp><p:txBody><a:bodyPr/>' + "".join(body) + '</p:txBody></p:sp>'
            '</p:spTree></p:cSld></p:sld>')


PRESENTATION = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
                'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
                '<p:sldIdLst>%s</p:sldIdLst></p:presentation>')

PRES_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '%s</Relationships>')

SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"


def build_deck(path, slides_in_display_order):
    """Write a deck where the FILE names run opposite to the display order.

    slides_in_display_order[0] is what the viewer calls slide 1, but it is stored
    as the HIGHEST-numbered slide file. Any code that assumes slideN.xml is the
    Nth slide edits the wrong one here.
    """
    count = len(slides_in_display_order)
    files = {}
    ids, rels = [], []
    for display_index, body in enumerate(slides_in_display_order):
        file_number = count - display_index          # reversed on purpose
        name = "ppt/slides/slide%d.xml" % file_number
        files[name] = body
        rid = "rId%d" % (display_index + 1)
        ids.append('<p:sldId id="%d" r:id="%s"/>' % (256 + display_index, rid))
        rels.append('<Relationship Id="%s" Type="%s" Target="slides/slide%d.xml"/>'
                    % (rid, SLIDE_REL_TYPE, file_number))

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", "<Types/>")
        z.writestr("_rels/.rels", "<Relationships/>")
        z.writestr("ppt/presentation.xml", PRESENTATION % "".join(ids))
        z.writestr("ppt/_rels/presentation.xml.rels", PRES_RELS % "".join(rels))
        z.writestr("ppt/theme/theme1.xml", "<a:theme>KEEP-ME</a:theme>")
        for name, body in files.items():
            z.writestr(name, body)
    return path


class TheFixturesAreAdversarialTest(unittest.TestCase):
    def test_a_naive_string_replace_finds_NOTHING_in_the_split_paragraph(self):
        # If this ever fails, the fixture stopped being the hard case and every
        # other test in this file became weaker without anyone noticing.
        self.assertNotIn("12 500 000", SPLIT_PARAGRAPH)
        self.assertEqual(slide_text(SPLIT_PARAGRAPH), "A hitel osszege 12 500 000 Ft.")

    def test_the_deck_fixture_really_has_file_names_against_the_display_order(self):
        tmp = tempfile.mkdtemp(prefix="office-pptx-")
        path = build_deck(os.path.join(tmp, "d.pptx"),
                          [slide_xml('<a:p><a:r><a:t>ELSO</a:t></a:r></a:p>'),
                           slide_xml('<a:p><a:r><a:t>MASODIK</a:t></a:r></a:p>')])
        with zipfile.ZipFile(path) as z:
            self.assertIn("ELSO", z.read("ppt/slides/slide2.xml").decode())
            self.assertIn("MASODIK", z.read("ppt/slides/slide1.xml").decode())
            self.assertEqual(slide_parts(z),
                             ["ppt/slides/slide2.xml", "ppt/slides/slide1.xml"])


class ReplaceInSlideTests(unittest.TestCase):
    def test_a_value_split_across_runs_IS_replaced(self):
        out, count = replace_in_slide(slide_xml(SPLIT_PARAGRAPH), "12 500 000", "9 900 000")
        self.assertEqual(count, 1)
        self.assertIn("9 900 000", slide_text(out))
        self.assertNotIn("12 5", out)

    def test_the_surrounding_formatting_survives(self):
        out, _ = replace_in_slide(slide_xml(SPLIT_PARAGRAPH), "12 500 000", "9 900 000")
        self.assertIn('b="1"', out)
        self.assertEqual(out.count("<a:r>"), slide_xml(SPLIT_PARAGRAPH).count("<a:r>"))

    def test_several_occurrences_on_one_slide_are_all_replaced(self):
        body = slide_xml('<a:p><a:r><a:t>ma 100, holnap 1</a:t></a:r>'
                         '<a:r><a:t>00, aztan 100.</a:t></a:r></a:p>')
        out, count = replace_in_slide(body, "100", "200")
        self.assertEqual(count, 3)
        self.assertEqual(slide_text(out), "ma 200, holnap 200, aztan 200.")

    def test_a_match_does_NOT_span_two_text_boxes(self):
        # Two paragraphs are two texts; joining them would invent a match nobody
        # can see on the slide.
        body = slide_xml('<a:p><a:r><a:t>vege: 12 5</a:t></a:r></a:p>',
                         '<a:p><a:r><a:t>00 000 kezdet</a:t></a:r></a:p>')
        _out, count = replace_in_slide(body, "12 500 000", "9 900 000")
        self.assertEqual(count, 0)

    def test_text_with_no_match_is_returned_untouched(self):
        body = slide_xml(SPLIT_PARAGRAPH)
        out, count = replace_in_slide(body, "nincs ilyen", "x")
        self.assertEqual(count, 0)
        self.assertEqual(out, body)

    def test_entities_are_handled_as_the_reader_sees_them(self):
        body = slide_xml('<a:p><a:r><a:t>Kovacs &amp; Tarsa</a:t></a:r></a:p>')
        out, count = replace_in_slide(body, "Kovacs & Tarsa", "Nagy & Fia")
        self.assertEqual(count, 1)
        self.assertIn("&amp;", out)
        self.assertIn("Nagy & Fia", slide_text(out))

    def test_an_empty_search_is_refused(self):
        with self.assertRaises(PptxError):
            replace_in_slide(slide_xml(WHOLE_PARAGRAPH), "", "x")

    def test_the_slide_structure_outside_the_text_is_untouched(self):
        body = slide_xml(WHOLE_PARAGRAPH)
        out, _ = replace_in_slide(body, "12 500 000", "9 900 000")
        self.assertIn("<p:spTree>", out)
        self.assertIn('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"', out)


class TableCellTests(unittest.TestCase):
    def setUp(self):
        self.slide = slide_xml('<a:p><a:r><a:t>cim</a:t></a:r></a:p>') .replace(
            "</p:txBody></p:sp>", "</p:txBody></p:sp><p:graphicFrame>%s</p:graphicFrame>" % TABLE)

    def test_a_cell_is_addressed_by_position(self):
        out = set_table_cell(self.slide, 0, 1, 1, "uj ertek")
        self.assertIn("uj ertek", out)

    def test_the_old_value_does_not_survive_in_a_second_run(self):
        out = set_table_cell(self.slide, 0, 1, 1, "uj ertek")
        self.assertNotIn("-maradek", out)
        self.assertNotIn(">regi<", out)

    def test_other_cells_are_left_alone(self):
        out = set_table_cell(self.slide, 0, 1, 1, "uj ertek")
        self.assertIn("<a:t>Nev</a:t>", out)
        self.assertIn("<a:t>Osszeg</a:t>", out)

    def test_a_SELF_CLOSING_empty_paragraph_cell_gets_a_run_rather_than_a_refusal(self):
        # A blank cell is often stored as <a:p/>, which has no inside to insert
        # into. Refusing here would make the tool useless for filling a template.
        out = set_table_cell(self.slide, 0, 2, 1, "beirt ertek")
        self.assertIn("<a:t>beirt ertek</a:t>", out)
        self.assertNotIn("<a:p/>", out.split("beirt ertek")[0].rsplit("<a:tc>", 1)[-1])

    def test_a_run_inserted_into_an_empty_cell_lands_BEFORE_endParaRPr(self):
        # a:endParaRPr must stay the last child of a:p; a run after it renders as
        # an empty cell in PowerPoint while the XML looks fine.
        out = set_table_cell(self.slide, 0, 2, 0, "elso oszlop")
        paragraph = out.split("elso oszlop")[0]
        self.assertNotIn("<a:endParaRPr", paragraph.rsplit("<a:p>", 1)[-1])
        self.assertIn("<a:endParaRPr", out.split("elso oszlop")[1])

    def test_an_out_of_range_address_is_refused_with_the_actual_size(self):
        for args in ((1, 0, 0), (0, 9, 0), (0, 0, 9)):
            with self.assertRaises(PptxError) as ctx:
                set_table_cell(self.slide, *args, "x")
            self.assertRegex(str(ctx.exception), r"\d")

    def test_markup_in_the_written_text_is_escaped(self):
        out = set_table_cell(self.slide, 0, 1, 1, "A & B <Kft.>")
        self.assertIn("&amp;", out)
        self.assertNotIn("<Kft.>", out)


class SlideNumberingTests(unittest.TestCase):
    """Slide 1 is the first slide the VIEWER sees, not slide1.xml."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="office-pptx-order-")
        self.path = build_deck(os.path.join(self.tmp, "deck.pptx"), [
            slide_xml('<a:p><a:r><a:t>ELSO dia szovege</a:t></a:r></a:p>'),
            slide_xml('<a:p><a:r><a:t>MASODIK dia szovege</a:t></a:r></a:p>'),
            slide_xml('<a:p><a:r><a:t>HARMADIK dia szovege</a:t></a:r></a:p>'),
        ])

    def test_read_slides_returns_them_in_display_order(self):
        names, slides = read_slides(self.path)
        self.assertEqual([slide_text(slides[n]).split()[0] for n in names],
                         ["ELSO", "MASODIK", "HARMADIK"])

    def test_editing_slide_1_edits_the_slide_the_viewer_sees_first(self):
        applied = apply_edits(self.path, replacements=[{"find": "dia szovege",
                                                        "replace": "MODOSITVA"}], slide=1)
        self.assertEqual(applied[0]["replaced"], 1)
        self.assertEqual(applied[0]["slides"], [1])
        names, slides = read_slides(self.path)
        self.assertIn("ELSO MODOSITVA", slide_text(slides[names[0]]))
        self.assertIn("MASODIK dia szovege", slide_text(slides[names[1]]))

    def test_the_report_names_the_slides_in_DISPLAY_numbers(self):
        applied = apply_edits(self.path, replacements=[{"find": "dia", "replace": "lap"}])
        self.assertEqual(applied[0]["replaced"], 3)
        self.assertEqual(applied[0]["slides"], [1, 2, 3])

    def test_slide_zero_is_refused_with_the_convention_named(self):
        with self.assertRaises(PptxError) as ctx:
            apply_edits(self.path, replacements=[{"find": "x", "replace": "y"}], slide=0)
        self.assertIn("numbered from 1", str(ctx.exception))

    def test_a_slide_past_the_end_is_refused_with_the_actual_count(self):
        with self.assertRaises(PptxError) as ctx:
            apply_edits(self.path, replacements=[{"find": "x", "replace": "y"}], slide=9)
        self.assertIn("3 slide(s)", str(ctx.exception))


class ApplyEditsTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="office-pptx-apply-")
        self.path = build_deck(os.path.join(self.tmp, "deck.pptx"), [
            slide_xml(SPLIT_PARAGRAPH),
            slide_xml('<a:p><a:r><a:t>A hitel osszege 12 500 000 Ft.</a:t></a:r></a:p>'),
        ])

    def test_a_replacement_visits_every_slide_by_default(self):
        applied = apply_edits(self.path, replacements=[{"find": "12 500 000",
                                                        "replace": "9 900 000"}])
        self.assertEqual(applied[0]["replaced"], 2)
        self.assertEqual(applied[0]["slides"], [1, 2])

    def test_a_replacement_that_matches_NOTHING_is_reported_as_zero(self):
        applied = apply_edits(self.path, replacements=[{"find": "nincs ilyen",
                                                        "replace": "x"}])
        self.assertEqual(applied[0]["replaced"], 0)
        self.assertEqual(applied[0]["slides"], [])

    def test_the_parts_nobody_edited_are_copied_through_untouched(self):
        apply_edits(self.path, replacements=[{"find": "12 500 000", "replace": "9 900 000"}])
        with zipfile.ZipFile(self.path) as z:
            self.assertIn("KEEP-ME", z.read("ppt/theme/theme1.xml").decode())
            self.assertIn("ppt/presentation.xml", z.namelist())

    def test_a_refusal_half_way_leaves_the_ORIGINAL_file_untouched(self):
        with open(self.path, "rb") as fh:
            before = fh.read()
        with self.assertRaises(PptxError):
            apply_edits(self.path,
                        replacements=[{"find": "12 500 000", "replace": "9 900 000"}],
                        cells=[{"slide": 1, "table": 7, "row": 0, "column": 0, "text": "x"}])
        with open(self.path, "rb") as fh:
            self.assertEqual(fh.read(), before)

    def test_a_replacement_without_find_or_replace_is_refused(self):
        with self.assertRaises(PptxError):
            apply_edits(self.path, replacements=[{"find": "csak ez van"}])

    def test_a_cell_edit_needs_the_slide_number(self):
        with self.assertRaises(PptxError) as ctx:
            apply_edits(self.path, cells=[{"table": 0, "row": 0, "column": 0, "text": "x"}])
        self.assertIn("slide", str(ctx.exception))

    def test_a_file_that_is_not_a_deck_is_refused_cleanly(self):
        broken = os.path.join(self.tmp, "broken.pptx")
        with open(broken, "w") as fh:
            fh.write("not a zip")
        with self.assertRaises(PptxError):
            apply_edits(broken, replacements=[{"find": "x", "replace": "y"}])

    def test_a_zip_without_a_presentation_part_is_refused_by_name(self):
        empty = os.path.join(self.tmp, "empty.pptx")
        with zipfile.ZipFile(empty, "w") as z:
            z.writestr("hello.txt", "no office parts here")
        with self.assertRaises(PptxError) as ctx:
            apply_edits(empty, replacements=[{"find": "x", "replace": "y"}])
        self.assertIn("presentation", str(ctx.exception))


class StdinStdoutContractTests(unittest.TestCase):
    """The contract office-mcp.js drives: one JSON in, one JSON object out."""

    def run_helper(self, request, env=None):
        proc = subprocess.run(
            [sys.executable, os.path.join(HERE, "office_pptx.py")],
            input=json.dumps(request), capture_output=True, text=True, cwd=HERE,
            env={**os.environ, **(env or {})},
        )
        self.assertEqual(proc.returncode, 0, proc.stderr)
        return json.loads(proc.stdout), proc.stderr

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="office-pptx-cli-")
        self.path = build_deck(os.path.join(self.tmp, "deck.pptx"),
                               [slide_xml(SPLIT_PARAGRAPH)])

    def test_a_good_request_answers_ok_with_the_counts(self):
        answer, _ = self.run_helper({"file": self.path, "replacements": [
            {"find": "12 500 000", "replace": "9 900 000"}]})
        self.assertTrue(answer["ok"])
        self.assertEqual(answer["applied"][0]["replaced"], 1)
        self.assertEqual(answer["applied"][0]["slides"], [1])

    def test_a_refusal_comes_back_as_ok_false_with_ONE_clean_message(self):
        answer, stderr = self.run_helper({"file": self.path, "cells": [
            {"slide": 1, "table": 0, "row": 0, "column": 0, "text": "x"}]})
        self.assertFalse(answer["ok"])
        self.assertIn("table", answer["error"])
        self.assertNotIn("Traceback", answer["error"])
        self.assertNotIn("Traceback", stderr)

    def test_a_missing_file_argument_is_refused(self):
        answer, _ = self.run_helper({"replacements": [{"find": "x", "replace": "y"}]})
        self.assertFalse(answer["ok"])
        self.assertIn("file", answer["error"])

    def test_a_path_outside_the_allowed_roots_is_refused(self):
        allowed = tempfile.mkdtemp(prefix="office-pptx-in-")
        outside = tempfile.mkdtemp(prefix="office-pptx-out-")
        target = build_deck(os.path.join(outside, "deck.pptx"), [slide_xml(WHOLE_PARAGRAPH)])
        answer, _ = self.run_helper({"file": target, "replacements": [
            {"find": "12 500 000", "replace": "9 900 000"}]},
            env={"OFFICE_ALLOWED_ROOTS": allowed})
        self.assertFalse(answer["ok"])
        self.assertIn("allowed roots", answer["error"])

    def test_unreadable_request_json_is_refused_cleanly(self):
        proc = subprocess.run([sys.executable, os.path.join(HERE, "office_pptx.py")],
                              input="}{", capture_output=True, text=True, cwd=HERE)
        self.assertEqual(proc.returncode, 0)
        answer = json.loads(proc.stdout)
        self.assertFalse(answer["ok"])
        self.assertIn("JSON", answer["error"])


if __name__ == "__main__":
    unittest.main()
