#!/usr/bin/env python3
"""Tests for docx text replacement.

The fixture that matters is SPLIT_PARAGRAPH: a value stored across two runs, the
way Word writes it when part of it is formatted differently. The first test does
not check our code at all - it checks the FIXTURE, by showing that a plain string
replace over the XML finds nothing there. Without that, a green suite would only
prove we handled the easy case, and the whole point of this module is the case
that looks fine and silently does nothing.

Run:  python3 -m unittest test_office_docx
"""

import os
import tempfile
import unittest
import zipfile

from office_docx import (
    DocxError,
    paragraph_text,
    read_document,
    replace_in_paragraph,
    replace_text,
    set_table_cell,
    write_document,
)

# "12 500 000" exists only as "12 5" + "00 000": no run contains it whole.
SPLIT_PARAGRAPH = (
    '<w:p><w:r><w:t>A hitel osszege 12 5</w:t></w:r>'
    '<w:r><w:rPr><w:b/></w:rPr><w:t>00 000</w:t></w:r>'
    '<w:r><w:t> Ft.</w:t></w:r></w:p>'
)

WHOLE_PARAGRAPH = '<w:p><w:r><w:t>A hitel osszege 12 500 000 Ft.</w:t></w:r></w:p>'

TABLE = (
    '<w:tbl>'
    '<w:tr><w:tc><w:p><w:r><w:t>Nev</w:t></w:r></w:p></w:tc>'
    '<w:tc><w:p><w:r><w:t>Ertek</w:t></w:r></w:p></w:tc></w:tr>'
    '<w:tr><w:tc><w:p><w:r><w:t>Osszeg</w:t></w:r></w:p></w:tc>'
    '<w:tc><w:p><w:r><w:t>regi</w:t></w:r><w:r><w:t>-maradek</w:t></w:r></w:p></w:tc></w:tr>'
    '</w:tbl>'
)


class TheFixtureIsAdversarialTest(unittest.TestCase):
    def test_a_naive_string_replace_finds_NOTHING_in_the_split_paragraph(self):
        # If this ever fails, the fixture stopped being the hard case and every
        # other test in this file became weaker without anyone noticing.
        self.assertNotIn("12 500 000", SPLIT_PARAGRAPH)
        self.assertEqual(paragraph_text(SPLIT_PARAGRAPH), "A hitel osszege 12 500 000 Ft.")


class ReplaceInParagraphTests(unittest.TestCase):
    def test_a_value_split_across_runs_IS_replaced(self):
        out, count = replace_in_paragraph(SPLIT_PARAGRAPH, "12 500 000", "9 900 000")
        self.assertEqual(count, 1)
        self.assertEqual(paragraph_text(out), "A hitel osszege 9 900 000 Ft.")

    def test_the_surrounding_formatting_survives(self):
        out, _ = replace_in_paragraph(SPLIT_PARAGRAPH, "12 500 000", "9 900 000")
        self.assertIn("<w:b/>", out)          # the bold run is still there
        self.assertEqual(out.count("<w:r>"), SPLIT_PARAGRAPH.count("<w:r>"))

    def test_a_whole_value_in_one_run_works_too(self):
        out, count = replace_in_paragraph(WHOLE_PARAGRAPH, "12 500 000", "9 900 000")
        self.assertEqual(count, 1)
        self.assertEqual(paragraph_text(out), "A hitel osszege 9 900 000 Ft.")

    def test_several_occurrences_in_one_paragraph_are_all_replaced(self):
        paragraph = ('<w:p><w:r><w:t>ma 100, holnap 1</w:t></w:r>'
                     '<w:r><w:t>00, aztan 100.</w:t></w:r></w:p>')
        out, count = replace_in_paragraph(paragraph, "100", "200")
        self.assertEqual(count, 3)
        self.assertEqual(paragraph_text(out), "ma 200, holnap 200, aztan 200.")

    def test_text_with_no_match_is_returned_untouched(self):
        out, count = replace_in_paragraph(SPLIT_PARAGRAPH, "nincs ilyen", "x")
        self.assertEqual(count, 0)
        self.assertEqual(out, SPLIT_PARAGRAPH)

    def test_entities_are_handled_as_the_reader_sees_them(self):
        paragraph = '<w:p><w:r><w:t>Kovacs &amp; Tarsa</w:t></w:r></w:p>'
        self.assertEqual(paragraph_text(paragraph), "Kovacs & Tarsa")
        out, count = replace_in_paragraph(paragraph, "Kovacs & Tarsa", "Nagy & Fia")
        self.assertEqual(count, 1)
        self.assertIn("&amp;", out)           # written back escaped
        self.assertEqual(paragraph_text(out), "Nagy & Fia")

    def test_a_replacement_containing_a_special_character_is_escaped(self):
        out, _ = replace_in_paragraph(WHOLE_PARAGRAPH, "12 500 000", "10 < 20 & tobb")
        self.assertIn("&lt;", out)
        self.assertIn("&amp;", out)
        self.assertEqual(paragraph_text(out), "A hitel osszege 10 < 20 & tobb Ft.")

    def test_an_empty_search_is_refused(self):
        with self.assertRaises(DocxError):
            replace_in_paragraph(WHOLE_PARAGRAPH, "", "x")

    def test_xml_space_preserve_attributes_survive(self):
        paragraph = '<w:p><w:r><w:t xml:space="preserve">ar: 100 </w:t></w:r></w:p>'
        out, _ = replace_in_paragraph(paragraph, "100", "200")
        self.assertIn('xml:space="preserve"', out)


class ReplaceTextTests(unittest.TestCase):
    def test_every_paragraph_is_visited(self):
        document = "<w:body>" + SPLIT_PARAGRAPH + WHOLE_PARAGRAPH + "</w:body>"
        out, count = replace_text(document, "12 500 000", "9 900 000")
        self.assertEqual(count, 2)
        self.assertNotIn("12 5", out)

    def test_a_match_does_NOT_span_two_paragraphs(self):
        # Paragraphs are separate texts; joining them would invent a match that a
        # reader never sees.
        document = ("<w:body><w:p><w:r><w:t>vege: 12 5</w:t></w:r></w:p>"
                    "<w:p><w:r><w:t>00 000 kezdet</w:t></w:r></w:p></w:body>")
        _out, count = replace_text(document, "12 500 000", "9 900 000")
        self.assertEqual(count, 0)

    def test_content_outside_the_text_runs_is_untouched(self):
        document = '<w:body><w:sectPr w:rsidR="00AB12"/>' + WHOLE_PARAGRAPH + "</w:body>"
        out, _ = replace_text(document, "12 500 000", "9 900 000")
        self.assertIn('<w:sectPr w:rsidR="00AB12"/>', out)


class TableCellTests(unittest.TestCase):
    def test_a_cell_is_addressed_by_position(self):
        out = set_table_cell(TABLE, 0, 1, 1, "uj ertek")
        self.assertIn("uj ertek", out)

    def test_the_old_value_does_not_survive_in_a_second_run(self):
        # The cell's second run held "-maradek"; leaving it would append the new
        # value to the old one on the page.
        out = set_table_cell(TABLE, 0, 1, 1, "uj ertek")
        self.assertNotIn("-maradek", out)
        self.assertNotIn(">regi<", out)

    def test_other_cells_are_left_alone(self):
        out = set_table_cell(TABLE, 0, 1, 1, "uj ertek")
        self.assertIn("<w:t>Nev</w:t>", out)
        self.assertIn("<w:t>Osszeg</w:t>", out)

    def test_an_out_of_range_address_is_refused_with_the_actual_size(self):
        for args in ((1, 0, 0), (0, 9, 0), (0, 0, 9)):
            with self.assertRaises(DocxError) as ctx:
                set_table_cell(TABLE, *args, "x")
            self.assertRegex(str(ctx.exception), r"\d")


class DocumentIoTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="office-docx-")
        self.src = os.path.join(self.tmp, "in.docx")
        with zipfile.ZipFile(self.src, "w") as z:
            z.writestr("[Content_Types].xml", "<Types/>")
            z.writestr("word/document.xml", "<w:document><w:body>%s</w:body></w:document>"
                       % SPLIT_PARAGRAPH)
            z.writestr("word/styles.xml", "<w:styles>KEEP-ME</w:styles>")

    def test_reading_a_non_docx_raises_DocxError(self):
        broken = os.path.join(self.tmp, "broken.docx")
        with open(broken, "w") as fh:
            fh.write("not a zip")
        with self.assertRaises(DocxError):
            read_document(broken)

    def test_only_the_document_part_is_replaced(self):
        document = read_document(self.src)
        new_document, count = replace_text(document, "12 500 000", "9 900 000")
        self.assertEqual(count, 1)
        dst = os.path.join(self.tmp, "out.docx")
        write_document(self.src, dst, new_document)
        with zipfile.ZipFile(dst) as z:
            self.assertIn("KEEP-ME", z.read("word/styles.xml").decode())
            self.assertIn("9 900 000", z.read("word/document.xml").decode())
            self.assertEqual(sorted(z.namelist()),
                             sorted(["[Content_Types].xml", "word/document.xml", "word/styles.xml"]))

    def test_the_source_is_not_modified(self):
        with open(self.src, "rb") as fh:
            before = fh.read()
        document = read_document(self.src)
        new_document, _ = replace_text(document, "12 500 000", "9 900 000")
        write_document(self.src, os.path.join(self.tmp, "out.docx"), new_document)
        with open(self.src, "rb") as fh:
            self.assertEqual(fh.read(), before)


if __name__ == "__main__":
    unittest.main()
