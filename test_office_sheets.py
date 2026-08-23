#!/usr/bin/env python3
"""Tests for single-sheet extraction.

Three of these are regressions for traps found on a real workbook, and all three
failed the same silent way: LibreOffice answered "source file could not be
loaded", exited 0 and wrote nothing, so the symptom was a missing PDF rather than
an explanation. They are worth keeping precisely because nothing about the
extracted file LOOKS wrong - Python opens it happily.

  * activeTab is an INDEX into the sheet list, so it has to be reset;
  * a defined name scoped (localSheetId) or pointing at a removed sheet kills the
    file - and its sheet reference is XML-escaped, which is how the first version
    of the filter let one through;
  * the relationship id may carry any namespace prefix, because our own cell
    writer re-serialises workbook.xml and renames prefixes (r:id -> ns0:id).

Run:  python3 -m unittest test_office_sheets
"""

import os
import tempfile
import unittest
import zipfile

from office_sheets import SheetError, flatten_formulas, list_sheets, single_sheet_copy

NS = ('xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')


def build_workbook(path, sheets, active_tab=0, defined_names="", rid_prefix="r"):
    """A workbook with one part per sheet, shaped like the real thing."""
    listing = "".join(
        '<sheet name="%s" sheetId="%d"%s %s:id="rId%d"/>'
        % (name, i + 1, ' state="hidden"' if hidden else "", rid_prefix, i + 1)
        for i, (name, hidden) in enumerate(sheets)
    )
    rels = "".join(
        '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/worksheet" Target="worksheets/sheet%d.xml"/>' % (i + 1, i + 1)
        for i in range(len(sheets))
    )
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", "<Types>%s</Types>" % "".join(
            '<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="sheet"/>' % (i + 1)
            for i in range(len(sheets))))
        z.writestr("xl/workbook.xml",
                   '<?xml version="1.0"?><workbook %s><bookViews><workbookView activeTab="%d"/>'
                   "</bookViews><sheets>%s</sheets>%s</workbook>"
                   % (NS, active_tab, listing, defined_names))
        z.writestr("xl/_rels/workbook.xml.rels",
                   '<?xml version="1.0"?><Relationships>%s</Relationships>' % rels)
        for i, (name, _hidden) in enumerate(sheets):
            z.writestr("xl/worksheets/sheet%d.xml" % (i + 1),
                       "<worksheet><sheetData><row r=\"1\">"
                       "<c r=\"A1\" t=\"inlineStr\"><is><t>CONTENT-OF-%s</t></is></c>"
                       "<c r=\"B1\"><f>Secret!B3</f><v>4242</v></c>"
                       "</row></sheetData></worksheet>" % name)
    return path


class ListSheetsTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="office-sheets-")

    def test_names_and_hidden_flags_come_back_in_order(self):
        src = build_workbook(os.path.join(self.tmp, "b.xlsx"),
                             [("Data", False), ("Secret", True), ("Plan", False)])
        self.assertEqual(list_sheets(src), [("Data", False), ("Secret", True), ("Plan", False)])

    def test_an_unreadable_file_raises_SheetError_not_a_zip_error(self):
        broken = os.path.join(self.tmp, "broken.xlsx")
        with open(broken, "w") as fh:
            fh.write("not a zip")
        with self.assertRaises(SheetError):
            list_sheets(broken)


class SingleSheetCopyTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="office-sheets-")
        self.dst = os.path.join(self.tmp, "out.xlsx")

    def _parts(self, path):
        with zipfile.ZipFile(path) as z:
            return {n: z.read(n).decode("utf-8", "replace") for n in z.namelist()}

    def test_only_the_named_sheet_survives(self):
        src = build_workbook(os.path.join(self.tmp, "b.xlsx"),
                             [("Data", False), ("Secret", True), ("Plan", False)])
        single_sheet_copy(src, self.dst, "Plan")
        self.assertEqual([n for n, _h in list_sheets(self.dst)], ["Plan"])
        parts = self._parts(self.dst)
        worksheets = [n for n in parts if n.startswith("xl/worksheets/sheet")]
        self.assertEqual(len(worksheets), 1)

    def test_the_OTHER_sheets_content_is_GONE_from_the_file(self):
        # The security claim, checked on the bytes: not merely delisted, absent.
        src = build_workbook(os.path.join(self.tmp, "b.xlsx"),
                             [("Data", False), ("Secret", True), ("Plan", False)])
        single_sheet_copy(src, self.dst, "Plan")
        blob = "".join(self._parts(self.dst).values())
        self.assertIn("CONTENT-OF-Plan", blob)
        self.assertNotIn("CONTENT-OF-Secret", blob)
        self.assertNotIn("CONTENT-OF-Data", blob)

    def test_a_HIDDEN_sheet_can_be_the_one_kept_and_the_rest_still_go(self):
        src = build_workbook(os.path.join(self.tmp, "b.xlsx"),
                             [("Data", False), ("Secret", True)])
        single_sheet_copy(src, self.dst, "Secret")
        blob = "".join(self._parts(self.dst).values())
        self.assertIn("CONTENT-OF-Secret", blob)
        self.assertNotIn("CONTENT-OF-Data", blob)

    def test_activeTab_is_reset_to_the_surviving_sheet(self):
        # Left at 2, LibreOffice refuses the file outright.
        src = build_workbook(os.path.join(self.tmp, "b.xlsx"),
                             [("A", False), ("B", False), ("C", False)], active_tab=2)
        single_sheet_copy(src, self.dst, "A")
        self.assertIn('activeTab="0"', self._parts(self.dst)["xl/workbook.xml"])

    def test_a_defined_name_pointing_at_a_removed_sheet_is_dropped(self):
        names = ('<definedNames><definedName name="_xlnm.Print_Area" localSheetId="1">'
                 "&apos;Secret&apos;!$A$1:$E$9</definedName></definedNames>")
        src = build_workbook(os.path.join(self.tmp, "b.xlsx"),
                             [("Plan", False), ("Secret", False)], defined_names=names)
        single_sheet_copy(src, self.dst, "Plan")
        workbook = self._parts(self.dst)["xl/workbook.xml"]
        self.assertNotIn("Secret", workbook)

    def test_a_defined_name_of_the_KEPT_sheet_survives_and_is_rescoped(self):
        names = ('<definedNames><definedName name="_xlnm.Print_Area" localSheetId="1">'
                 "&apos;Plan&apos;!$A$1:$E$9</definedName></definedNames>")
        src = build_workbook(os.path.join(self.tmp, "b.xlsx"),
                             [("Secret", False), ("Plan", False)], defined_names=names)
        single_sheet_copy(src, self.dst, "Plan")
        workbook = self._parts(self.dst)["xl/workbook.xml"]
        self.assertIn("Print_Area", workbook)
        self.assertIn('localSheetId="0"', workbook)

    def test_a_renamed_namespace_prefix_on_the_relationship_id_still_resolves(self):
        # Our own cell writer produces ns0:id instead of r:id.
        src = build_workbook(os.path.join(self.tmp, "b.xlsx"),
                             [("Data", False), ("Plan", False)], rid_prefix="ns0")
        single_sheet_copy(src, self.dst, "Plan")
        self.assertEqual([n for n, _h in list_sheets(self.dst)], ["Plan"])

    def test_naming_a_sheet_that_does_not_exist_REFUSES(self):
        # It must never fall back to exporting everything.
        src = build_workbook(os.path.join(self.tmp, "b.xlsx"), [("Data", False), ("Plan", False)])
        with self.assertRaises(SheetError) as ctx:
            single_sheet_copy(src, self.dst, "Nope")
        self.assertIn("Data", str(ctx.exception))  # lists what IS available
        self.assertFalse(os.path.exists(self.dst))

    def test_a_single_sheet_workbook_is_copied_untouched(self):
        src = build_workbook(os.path.join(self.tmp, "b.xlsx"), [("Only", False)])
        single_sheet_copy(src, self.dst, "Only")
        with open(src, "rb") as a, open(self.dst, "rb") as b:
            self.assertEqual(a.read(), b.read())


class FlattenFormulasTests(unittest.TestCase):
    """A formula pointing at a removed sheet renders as #NAME? and takes its cached
    value down with it - found in production (on a real customer workbook) after
    the first version of the extraction shipped keeping formulas. Flattening to the
    stored value is what makes a single-sheet export faithful.
    """

    def test_a_numeric_formula_keeps_its_value_and_loses_the_formula(self):
        cell = '<c r="B1"><f>Alapadatok!B3</f><v>4242</v></c>'
        out = flatten_formulas("<sheetData><row>%s</row></sheetData>" % cell)
        self.assertNotIn("<f>", out)
        self.assertIn("<v>4242</v>", out)

    def test_a_TEXT_formula_becomes_an_inline_string(self):
        # t="str" only means anything while a formula is present.
        cell = '<c r="B2" t="str"><f>Alapadatok!A1</f><v>Kovacs Bt</v></c>'
        out = flatten_formulas("<sheetData><row>%s</row></sheetData>" % cell)
        self.assertNotIn("<f>", out)
        self.assertIn('t="inlineStr"', out)
        self.assertIn("<is><t>Kovacs Bt</t></is>", out)
        self.assertNotIn("<v>", out)

    def test_a_formula_with_no_cached_value_leaves_an_empty_cell(self):
        # Nothing to keep: better empty than showing an error.
        cell = '<c r="B3"><f>Alapadatok!B9</f></c>'
        out = flatten_formulas("<sheetData><row>%s</row></sheetData>" % cell)
        self.assertNotIn("<f>", out)
        self.assertNotIn("<v>", out)

    def test_a_shared_formula_is_flattened_too(self):
        cell = '<c r="B4"><f t="shared" si="3"/><v>7</v></c>'
        out = flatten_formulas("<sheetData><row>%s</row></sheetData>" % cell)
        self.assertNotIn("<f", out)
        self.assertIn("<v>7</v>", out)

    def test_cells_without_formulas_are_untouched(self):
        body = '<sheetData><row><c r="A1" t="inlineStr"><is><t>Nev</t></is></c></row></sheetData>'
        self.assertEqual(flatten_formulas(body), body)

    def test_the_extracted_sheet_contains_NO_formula_at_all(self):
        # The end-to-end property: whatever the source had, the copy is value-only,
        # so no reference can dangle into a sheet that is gone.
        tmp = tempfile.mkdtemp(prefix="office-sheets-flat-")
        src = build_workbook(os.path.join(tmp, "b.xlsx"), [("Plan", False), ("Secret", False)])
        dst = os.path.join(tmp, "out.xlsx")
        single_sheet_copy(src, dst, "Plan")
        with zipfile.ZipFile(dst) as z:
            sheet = [n for n in z.namelist() if n.startswith("xl/worksheets/sheet")][0]
            body = z.read(sheet).decode()
        self.assertNotIn("<f>", body)
        self.assertNotIn("Secret!", body)     # the dangling reference is gone
        self.assertIn("<v>4242</v>", body)    # its value survived


if __name__ == "__main__":
    unittest.main()
