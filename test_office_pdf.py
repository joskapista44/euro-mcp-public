#!/usr/bin/env python3
"""Tests for the PDF export helper.

The engine itself is already covered by the recalc suite - this module only adds
a target format, so what is worth testing here is what is NEW: choosing the right
LibreOffice filter per source type, refusing what cannot be exported, and never
leaving a half-written PDF behind when the run fails.

The filter matters more than it looks: `--convert-to docx` (the reverse
direction) failed with "no export filter" on this very engine, so a bare "pdf"
is not something to assume works for every source.

Run:  python3 -m unittest test_office_pdf
"""

import json
import os
import tempfile
import time
import unittest
import unittest.mock
import zipfile

import contextlib
import sys

import office_pdf
from office_pdf import PDF_FILTERS, export_pdf, pdf_filter_for
from office_recalc import ExternalToolError

HERE = os.path.dirname(os.path.abspath(__file__))
FAKE = os.path.join(HERE, "fake_soffice_recalc.py")


@contextlib.contextmanager
def set_env(**pairs):
    saved = {k: os.environ.get(k) for k in pairs}
    try:
        for k, v in pairs.items():
            if v is None:
                os.environ.pop(k, None)
            else:
                os.environ[k] = v
        yield
    finally:
        for k, v in saved.items():
            if v is None:
                os.environ.pop(k, None)
            else:
                os.environ[k] = v


def make_xlsx(path, sheets=("Sheet1",)):
    """A minimal but STRUCTURALLY HONEST workbook.

    It has to list its sheets: the export refuses a multi-sheet workbook that does
    not say which sheet to print, so a fixture without a workbook part would test
    a code path no real file takes.
    """
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", "<Types/>")
        listing = "".join(
            '<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (name, i + 1, i + 1)
            for i, name in enumerate(sheets)
        )
        z.writestr("xl/workbook.xml",
                   '<?xml version="1.0"?><workbook '
                   'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                   'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                   "<sheets>%s</sheets></workbook>" % listing)


class FilterChoiceTests(unittest.TestCase):
    def test_each_family_gets_its_own_filter(self):
        self.assertEqual(pdf_filter_for("/tmp/a.xlsx"), "pdf:calc_pdf_Export")
        self.assertEqual(pdf_filter_for("/tmp/a.docx"), "pdf:writer_pdf_Export")
        self.assertEqual(pdf_filter_for("/tmp/a.pptx"), "pdf:impress_pdf_Export")

    def test_the_extension_is_matched_case_insensitively(self):
        self.assertEqual(pdf_filter_for("/tmp/REPORT.XLSX"), "pdf:calc_pdf_Export")

    def test_an_unsupported_source_is_refused_with_the_list(self):
        with self.assertRaises(ExternalToolError) as ctx:
            pdf_filter_for("/tmp/notes.txt")
        self.assertIn(".txt", str(ctx.exception))
        self.assertIn(".docx", str(ctx.exception))

    def test_a_file_without_an_extension_is_refused_readably(self):
        with self.assertRaises(ExternalToolError) as ctx:
            pdf_filter_for("/tmp/README")
        self.assertIn("without an extension", str(ctx.exception))

    def test_every_listed_filter_actually_names_a_pdf_export(self):
        for ext, flt in PDF_FILTERS.items():
            self.assertTrue(flt.startswith("pdf"), "%s -> %s" % (ext, flt))


class ExportGuardTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="office-pdf-test-")

    def test_a_missing_source_is_reported_as_such(self):
        with self.assertRaises(ExternalToolError) as ctx:
            export_pdf(os.path.join(self.tmp, "nope.xlsx"), os.path.join(self.tmp, "o.pdf"), 5)
        self.assertIn("file not found", str(ctx.exception))

    def test_a_directory_is_not_mistaken_for_a_document(self):
        with self.assertRaises(ExternalToolError) as ctx:
            export_pdf(self.tmp, os.path.join(self.tmp, "o.pdf"), 5)
        self.assertIn("not a file", str(ctx.exception))

    def test_a_corrupt_office_file_fails_before_the_engine_is_started(self):
        broken = os.path.join(self.tmp, "broken.xlsx")
        with open(broken, "w") as fh:
            fh.write("this is not a zip")
        with unittest.mock.patch.object(office_pdf, "build_runner") as runner:
            with self.assertRaises(ExternalToolError) as ctx:
                export_pdf(broken, os.path.join(self.tmp, "o.pdf"), 5)
        self.assertIn("not a valid office file", str(ctx.exception))
        # Validation is intentionally engine-independent: a bad input must not even
        # require runner discovery, let alone start a container or host process.
        runner.assert_not_called()

    def test_no_pdf_is_left_behind_when_the_engine_fails(self):
        src = os.path.join(self.tmp, "book.xlsx")
        make_xlsx(src)
        out = os.path.join(self.tmp, "book.pdf")
        with unittest.mock.patch.object(office_pdf, "prepare_profile"), \
             unittest.mock.patch.object(office_pdf, "run_producing_file",
                                        side_effect=ExternalToolError("engine died")):
            with self.assertRaises(ExternalToolError):
                export_pdf(src, out, 5)
        self.assertFalse(os.path.exists(out), "a failed export must not leave a PDF")

    def test_the_source_is_not_modified(self):
        src = os.path.join(self.tmp, "book.xlsx")
        make_xlsx(src)
        with open(src, "rb") as fh:
            before = fh.read()
        with unittest.mock.patch.object(office_pdf, "prepare_profile"), \
             unittest.mock.patch.object(office_pdf, "run_producing_file",
                                        side_effect=ExternalToolError("engine died")):
            with self.assertRaises(ExternalToolError):
                export_pdf(src, os.path.join(self.tmp, "o.pdf"), 5)
        with open(src, "rb") as fh:
            self.assertEqual(fh.read(), before)


class EngineFailureModesTests(unittest.TestCase):
    """The four ways an external tool fails, re-proven THROUGH the PDF entry point."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="office-pdf-modes-")
        self.src = os.path.join(self.tmp, "book.xlsx")
        make_xlsx(self.src)
        self.out = os.path.join(self.tmp, "book.pdf")

    def _export(self, behavior, timeout=5):
        with set_env(OFFICE_ENGINE="host", OFFICE_SOFFICE_BIN=FAKE,
                     FAKE_SOFFICE_BEHAVIOR=behavior):
            return export_pdf(self.src, self.out, timeout)

    def test_exit_zero_without_an_output_file_is_still_a_failure(self):
        with self.assertRaises(ExternalToolError) as ctx:
            self._export("no-output")
        self.assertFalse(os.path.exists(self.out))
        self.assertNotIn("Traceback", str(ctx.exception))

    def test_a_non_zero_exit_surfaces_its_message_once(self):
        with self.assertRaises(ExternalToolError) as ctx:
            self._export("error")
        self.assertFalse(os.path.exists(self.out))
        self.assertNotIn("Traceback", str(ctx.exception))

    def test_a_hanging_engine_is_killed_by_the_timeout(self):
        started = time.time()
        with self.assertRaises(ExternalToolError):
            self._export("hang", timeout=2)
        self.assertLess(time.time() - started, 30, "the timeout did not stop the run")
        self.assertFalse(os.path.exists(self.out))

    def test_a_missing_engine_binary_is_a_clean_error(self):
        with set_env(OFFICE_ENGINE="host",
                     OFFICE_SOFFICE_BIN=os.path.join(HERE, "no-such-soffice-xyz"),
                     FAKE_SOFFICE_BEHAVIOR=None):
            with self.assertRaises(ExternalToolError) as ctx:
                export_pdf(self.src, self.out, 5)
        self.assertFalse(os.path.exists(self.out))
        self.assertNotIn("Traceback", str(ctx.exception))


class ContractTests(unittest.TestCase):
    """main() must always answer with one clean JSON line - the MCP tool parses it."""

    def _run(self, request):
        import contextlib, io, sys
        out = io.StringIO()
        with unittest.mock.patch("sys.stdin", io.StringIO(json.dumps(request))), \
             contextlib.redirect_stdout(out), contextlib.redirect_stderr(io.StringIO()):
            office_pdf.main()
        return json.loads(out.getvalue())

    def test_a_path_outside_the_allowed_roots_is_refused(self):
        result = self._run({"file": "/etc/passwd.xlsx"})
        self.assertFalse(result["ok"])
        self.assertIn("outside the allowed roots", result["error"])

    def test_an_allowed_source_with_a_forbidden_DESTINATION_is_refused(self):
        src = os.path.join(tempfile.gettempdir(), "ok.xlsx")
        make_xlsx(src)
        result = self._run({"file": src, "out_file": "/etc/out.pdf"})
        self.assertFalse(result["ok"])
        self.assertIn("outside the allowed roots", result["error"])

    def test_a_missing_file_field_is_a_clean_error(self):
        self.assertIn("file is required", self._run({})["error"])


class SheetSelectionTests(unittest.TestCase):
    """The workbook gate: what ends up in the PDF, and what must not."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="office-pdf-sheets-")

    def _run(self, request):
        import contextlib, io
        out = io.StringIO()
        with unittest.mock.patch("sys.stdin", io.StringIO(json.dumps(request))), \
             contextlib.redirect_stdout(out), contextlib.redirect_stderr(io.StringIO()):
            office_pdf.main()
        return json.loads(out.getvalue())

    def _run_reaching_engine(self, request):
        # These tests prove that the workbook privacy gate permits the request to
        # proceed. Runner discovery is not part of that contract, so replace it
        # with a sentinel failure after all engine-independent validation passed.
        with unittest.mock.patch.object(
                office_pdf, "build_runner",
                side_effect=ExternalToolError("engine not run here")):
            return self._run(request)

    def test_a_multi_sheet_workbook_without_a_sheet_is_REFUSED(self):
        src = os.path.join(self.tmp, "book.xlsx")
        make_xlsx(src, sheets=("Plan", "Decision"))
        result = self._run({"file": src, "out_file": os.path.join(self.tmp, "o.pdf")})
        self.assertFalse(result["ok"])
        self.assertIn("sheet=", result["error"])
        self.assertIn("Decision", result["error"])

    def test_a_single_sheet_workbook_needs_no_choice(self):
        src = os.path.join(self.tmp, "one.xlsx")
        make_xlsx(src, sheets=("Only",))
        result = self._run_reaching_engine(
            {"file": src, "out_file": os.path.join(self.tmp, "o.pdf")})
        self.assertIn("engine not run here", result["error"])

    def test_all_sheets_is_an_explicit_opt_in(self):
        src = os.path.join(self.tmp, "book.xlsx")
        make_xlsx(src, sheets=("Plan", "Decision"))
        result = self._run_reaching_engine(
            {"file": src, "all_sheets": True,
             "out_file": os.path.join(self.tmp, "o.pdf")})
        self.assertIn("engine not run here", result["error"])

    def test_sheet_selection_is_refused_for_a_document(self):
        src = os.path.join(self.tmp, "doc.docx")
        with zipfile.ZipFile(src, "w") as z:
            z.writestr("[Content_Types].xml", "<Types/>")
        result = self._run({"file": src, "sheet": "Plan",
                            "out_file": os.path.join(self.tmp, "o.pdf")})
        self.assertFalse(result["ok"])
        self.assertIn("only applies to spreadsheets", result["error"])


if __name__ == "__main__":
    unittest.main()
