#!/usr/bin/env python3
"""PDF export for the office MCP: xlsx / docx / pptx -> pdf.

Deliberately thin. Everything that was hard about driving LibreOffice was already
solved for recalc and lives in office_recalc/office_runner: engine selection
(container by default, visible host fallback), the per-job profile that keeps a
headless run from touching a shared LibreOffice profile, the job workspace, and
the four failure modes of an external tool. This module reuses all of it and only
changes the target format - writing a second, parallel invocation would mean two
places to fix the next LibreOffice trap.
"""

import json
import os
import shutil
import sys
import zipfile

from office_paths import PathNotAllowed, check_path
from office_sheets import SheetError, list_sheets, single_sheet_copy
from office_recalc import ExternalToolError, build_runner, job_workspace, prepare_profile
from office_runner import run_producing_file

PDF_FILTERS = {
    ".xlsx": "pdf:calc_pdf_Export", ".xls": "pdf:calc_pdf_Export",
    ".docx": "pdf:writer_pdf_Export", ".doc": "pdf:writer_pdf_Export",
    ".odt": "pdf:writer_pdf_Export", ".ods": "pdf:calc_pdf_Export",
    ".pptx": "pdf:impress_pdf_Export", ".ppt": "pdf:impress_pdf_Export",
    ".odp": "pdf:impress_pdf_Export",
}


def pdf_filter_for(path):
    ext = os.path.splitext(path)[1].lower()
    try:
        return PDF_FILTERS[ext]
    except KeyError:
        raise ExternalToolError(
            "cannot export %s to PDF: supported sources are %s"
            % (ext or "a file without an extension", ", ".join(sorted(PDF_FILTERS)))
        )


def export_pdf(path, out_path, timeout, report=None, sheet=None, all_sheets=False):
    """Write `path` as a PDF at `out_path`. Returns the engine description."""
    # Validate everything that does not require LibreOffice before selecting an
    # engine. Besides producing better errors, this keeps validation deterministic
    # on clean machines where neither the container image nor soffice is installed.
    if not os.path.exists(path):
        raise ExternalToolError("file not found: " + path)
    if os.path.isdir(path):
        raise ExternalToolError("not a file: " + path)
    if os.path.splitext(path)[1].lower() in (".xlsx", ".docx", ".pptx"):
        try:
            with zipfile.ZipFile(path):
                pass
        except (zipfile.BadZipFile, OSError) as exc:
            raise ExternalToolError("not a valid office file (%s): %s" % (exc, path))

    convert_to = pdf_filter_for(path)
    name = os.path.basename(path)
    produced = os.path.splitext(name)[0] + ".pdf"
    is_workbook = os.path.splitext(path)[1].lower() in (".xlsx", ".xls", ".ods")

    if sheet and not is_workbook:
        raise ExternalToolError("sheet selection only applies to spreadsheets, not %s" % name)

    if is_workbook and not sheet and not all_sheets:
        try:
            sheets = list_sheets(path)
        except SheetError as exc:
            raise ExternalToolError(str(exc))
        if len(sheets) > 1:
            listing = ", ".join(
                "%s%s" % (n, " (hidden)" if hidden else "") for n, hidden in sheets
            )
            raise ExternalToolError(
                "this workbook has %d sheets and all of them would end up in the PDF, "
                "including hidden ones. Name the sheet to export (sheet=...), or ask for "
                "the whole workbook explicitly (all_sheets=true). Sheets: %s"
                % (len(sheets), listing)
            )

    runner = build_runner(report=report)

    with job_workspace(prefix="office-pdf-") as jobdir:
        runner.prepare(jobdir)
        if sheet:
            try:
                single_sheet_copy(path, os.path.join(jobdir, name), sheet)
            except SheetError as exc:
                raise ExternalToolError(str(exc))
        else:
            shutil.copy(path, os.path.join(jobdir, name))
        os.makedirs(os.path.join(jobdir, "out"))
        out_local = os.path.join(jobdir, "out", produced)

        prepare_profile(runner, jobdir, timeout, force_recalc=False)
        run_producing_file(
            runner.command(jobdir, [
                "--headless", "--norestore",
                "-env:UserInstallation=file://" + runner.seen(jobdir, "profile"),
                "--convert-to", convert_to,
                "--outdir", runner.seen(jobdir, "out"),
                runner.seen(jobdir, name)]),
            out_local, timeout,
        )
        shutil.move(out_local, out_path)
    return runner.describe()


def main():
    try:
        req = json.load(sys.stdin)
    except Exception as exc:
        json.dump({"ok": False, "error": "invalid request JSON: %s" % exc}, sys.stdout)
        return

    path = req.get("file")
    if not path:
        json.dump({"ok": False, "error": "file is required"}, sys.stdout)
        return
    timeout = req.get("timeout", 180)

    try:
        path = check_path(path)
        out_path = req.get("out_file") or os.path.splitext(path)[0] + ".pdf"
        out_path = check_path(out_path)
    except PathNotAllowed as exc:
        json.dump({"ok": False, "error": str(exc)}, sys.stdout)
        return

    try:
        engine = export_pdf(path, out_path, timeout,
                            report=lambda m: sys.stderr.write(m + "\n"),
                            sheet=req.get("sheet"), all_sheets=bool(req.get("all_sheets")))
    except ExternalToolError as exc:
        json.dump({"ok": False, "error": str(exc)}, sys.stdout)
        return
    except Exception as exc:
        import traceback
        traceback.print_exc(file=sys.stderr)
        json.dump({"ok": False, "error": "pdf export failed: %s" % exc}, sys.stdout)
        return

    json.dump({"ok": True, "file": out_path, "engine": engine}, sys.stdout)


if __name__ == "__main__":
    main()
