#!/usr/bin/env python3
"""PDF export for the office MCP: xlsx / docx / pptx -> pdf.

Deliberately thin. Everything that was hard about driving LibreOffice was already
solved for recalc and lives in office_recalc/office_runner: engine selection
(container by default, visible host fallback), the per-job profile that keeps a
headless run from touching a shared LibreOffice profile, the job workspace, and
the four failure modes of an external tool. This module reuses all of it and only
changes the target format - writing a second, parallel invocation would mean two
places to fix the next LibreOffice trap.

One such trap is already in the record: without
`-env:UserInstallation=file:///work/profile` inside the mount, soffice in the
container dies with "User installation could not be completed" and produces
nothing. Measured today - which is exactly why this path goes through
prepare_profile like recalc does, instead of assembling its own command.

The other is the filter name: `--convert-to pdf` works for xlsx and pptx, but a
docx source needs `pdf:writer_pdf_Export`. Plain `--convert-to docx` (the reverse
direction) already failed today with "no export filter", so the filter is chosen
per source type here rather than hoped for.
"""

import json
import os
import shutil
import sys
import zipfile

from office_paths import PathNotAllowed, check_path
from office_sheets import SheetError, list_sheets, single_sheet_copy
from office_recalc import (
    ExternalToolError,
    build_runner,
    job_workspace,
    prepare_profile,
)
from office_runner import run_producing_file

# Source extension -> the filter that produces a PDF from it. The value is what
# `--convert-to` receives, so it carries the LibreOffice filter name where the
# bare "pdf" is ambiguous.
PDF_FILTERS = {
    ".xlsx": "pdf:calc_pdf_Export",
    ".xls": "pdf:calc_pdf_Export",
    ".docx": "pdf:writer_pdf_Export",
    ".doc": "pdf:writer_pdf_Export",
    ".odt": "pdf:writer_pdf_Export",
    ".ods": "pdf:calc_pdf_Export",
    ".pptx": "pdf:impress_pdf_Export",
    ".ppt": "pdf:impress_pdf_Export",
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
    """Write `path` as a PDF at `out_path`. Returns the engine description.

    The output is only put in place once a real PDF exists, so a failed export
    never leaves a truncated file behind - and never silently overwrites a good
    PDF from an earlier run with nothing.
    """
    runner = build_runner(report=report)

    if not os.path.exists(path):
        raise ExternalToolError("file not found: " + path)
    if os.path.isdir(path):
        raise ExternalToolError("not a file: " + path)
    # Office formats are zips; catching it here gives a clear message instead of a
    # LibreOffice error twenty seconds later.
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
            # An unreadable workbook must not surface as a stray traceback just
            # because we asked it how many sheets it has.
            raise ExternalToolError(str(exc))
        if len(sheets) > 1:
            # Fail closed. LibreOffice exports EVERY sheet, hidden ones included, so
            # a whole-workbook PDF of a customer file can carry personal data from
            # sheets nobody asked for. Exporting all of them has to be asked for.
            listing = ", ".join(
                "%s%s" % (n, " (hidden)" if hidden else "") for n, hidden in sheets
            )
            raise ExternalToolError(
                "this workbook has %d sheets and all of them would end up in the PDF, "
                "including hidden ones. Name the sheet to export (sheet=...), or ask for "
                "the whole workbook explicitly (all_sheets=true). Sheets: %s"
                % (len(sheets), listing)
            )

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

        # No recalculation: the PDF must show what the file stores. With a single
        # extracted sheet a recalc would resolve references to the removed sheets as
        # #REF!, and even on a whole workbook it would silently change numbers the
        # caller never asked to change.
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
    except Exception as exc:  # noqa: BLE001 - any parse failure becomes a clean result
        json.dump({"ok": False, "error": "invalid request JSON: %s" % exc}, sys.stdout)
        return

    path = req.get("file")
    if not path:
        json.dump({"ok": False, "error": "file is required"}, sys.stdout)
        return
    timeout = req.get("timeout", 180)

    # Both ends are guarded: reading a file the tools may not read, and writing a
    # PDF somewhere they may not write, are two separate ways to leave the sandbox.
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
    except Exception as exc:  # noqa: BLE001 - a bug here must not leak a traceback
        import traceback
        traceback.print_exc(file=sys.stderr)
        json.dump({"ok": False, "error": "pdf export failed: %s" % exc}, sys.stdout)
        return

    json.dump({"ok": True, "file": out_path, "engine": engine}, sys.stdout)


if __name__ == "__main__":
    main()
