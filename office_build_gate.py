#!/usr/bin/env python3
"""Acceptance gate for docx_build / pptx_build, driven through the MCP tool interface.

The unit tests prove the generated XML says the right things. That is not the same
as the file OPENING and RENDERING: a package can be well-formed, pass every string
assertion, and still be refused by LibreOffice or come out with an empty frame
where a picture should be. So this gate goes the whole way - tools/call through
office-mcp.js, then office_to_pdf through the real engine, then read the PDF back
and look for the words that must be on the page.

Four cases, each proving something different, reported separately:

  CASE 1  docx_build -> PDF: every parity-floor element is BUILT, the file renders,
          the page break really produced a second page, and no error marker
          (#NAME?, Error, undefined, a leftover field code) reached the paper.
  CASE 2  pptx_build -> PDF: the deck renders with one page per slide and the text
          of every block type is on it.
  CASE 3  an unsupported block type is REFUSED through the tool: isError, one clean
          message naming what is supported, no stack trace, and NO file left behind.
  CASE 4  a slide that overflows is REFUSED by name - the failure mode that is
          invisible in the file and obvious on screen.
  CASE 5  pptx_replace_text / pptx_set_table_cell EDIT a built deck and the change
          reaches the paper: the new text is in the PDF and the OLD text is gone.
          The searched value is deliberately stored SPLIT ACROSS TWO RUNS, which is
          what a plain string replace over the XML silently fails to find.
  CASE 6  the edit tools REFUSE what they cannot address (a slide past the end, a
          table that is not on the slide) and leave the file byte-for-byte alone.

Verdicts, worst wins: HARNESS > RED > PENDING > GREEN.
  0 GREEN    every case passed.
  1 RED      a real defect: a wrong/missing render, a leaked stack trace, a refusal
             that did not refuse, or a file written by a refused call.
  2 HARNESS  the test could not run (no node, handshake failed, no pdftotext).
  3 PENDING  a tool is not listed yet, or the render engine is not available here.

stdlib only, no network. Writes under /tmp, which is where the tools are allowed
to work by default (OFFICE_ALLOWED_ROOTS).

  python3 office_build_gate.py [/path/to/office-mcp.js]
"""

import os
import struct
import subprocess
import sys
import tempfile
import zlib

from office_e2e import (  # the MCP stdio client, already used by the recalc gate
    GREEN, HARNESS, PENDING, RED, McpClient, McpError, PROTOCOL_VERSION,
    _NAME, aggregate, hr, tool_result_text,
)

HERE = os.path.dirname(os.path.abspath(__file__))

# Anything that means "the document rendered, but wrong". Checked on the PDF text,
# because a positive assertion alone would pass a page full of #NAME?.
ERROR_MARKERS = ("#NAME?", "#REF!", "#VALUE!", "Error:", "undefined", "NUMPAGES", "PAGE \\")


def make_png(path, width=120, height=60):
    def chunk(tag, data):
        payload = tag + data
        return (struct.pack(">I", len(data)) + payload
                + struct.pack(">I", zlib.crc32(payload) & 0xFFFFFFFF))

    raw = b"".join(b"\x00" + bytes([31, 78, 121] * width) for _ in range(height))
    with open(path, "wb") as fh:
        fh.write(b"\x89PNG\r\n\x1a\n"
                 + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
                 + chunk(b"IDAT", zlib.compress(raw))
                 + chunk(b"IEND", b""))


def pdf_text_and_pages(path):
    """(text, page_count) via poppler. Raises RuntimeError if the tools are missing."""
    try:
        text = subprocess.run(["pdftotext", "-layout", path, "-"],
                              capture_output=True, text=True, timeout=120)
        info = subprocess.run(["pdfinfo", path], capture_output=True, text=True, timeout=60)
    except FileNotFoundError as exc:
        raise RuntimeError("pdftotext/pdfinfo not available: %s" % exc)
    if text.returncode != 0:
        raise RuntimeError("pdftotext failed: %s" % text.stderr.strip()[:200])
    pages = None
    for line in info.stdout.splitlines():
        if line.startswith("Pages:"):
            pages = int(line.split()[1])
    return text.stdout, pages


def call(client, name, arguments, timeout=240):
    """tools/call -> ('harness'|'proto_error'|'error_trace'|'error_clean'|'ok', text)."""
    try:
        answer = client.request("tools/call", {"name": name, "arguments": arguments},
                                timeout=timeout)
    except McpError as exc:
        return "harness", str(exc)
    if "error" in answer:
        return "proto_error", str(answer["error"])
    result = answer.get("result") or {}
    text = tool_result_text(result)
    if result.get("isError"):
        if "Traceback" in text or "most recent call last" in text:
            return "error_trace", text
        return "error_clean", text
    return "ok", text


def render(client, path, out):
    """office_to_pdf -> (verdict_or_None, text, pdf_text, pages)."""
    kind, text = call(client, "office_to_pdf", {"file": path, "out_file": out})
    print("    office_to_pdf said: %s" % (text or "(no text)"))
    if kind == "harness":
        return HARNESS, text, None, None
    if kind == "error_trace":
        print("    RED: the export leaked a stack trace - a caller must get one clean message.")
        return RED, text, None, None
    if kind in ("proto_error", "error_clean"):
        print("    PENDING: the render engine is not available here; the build itself is "
              "not disproven, only unverified on paper.")
        return PENDING, text, None, None
    if not os.path.exists(out):
        print("    RED: office_to_pdf reported success but wrote no file.")
        return RED, text, None, None
    try:
        pdf_text, pages = pdf_text_and_pages(out)
    except RuntimeError as exc:
        print("    HARNESS: %s" % exc)
        return HARNESS, text, None, None
    return None, text, pdf_text, pages


def check_words(pdf_text, expected, tag):
    missing = [w for w in expected if w not in pdf_text]
    dirty = [m for m in ERROR_MARKERS if m in pdf_text]
    if missing:
        print("    %s RED: missing from the rendered page: %s" % (tag, ", ".join(missing)))
    if dirty:
        print("    %s RED: error marker(s) on the rendered page: %s" % (tag, ", ".join(dirty)))
    return RED if (missing or dirty) else GREEN


def case_1_docx(client, work):
    print("CASE 1 - docx_build renders every parity-floor element onto paper")
    png = os.path.join(work, "kep.png")
    make_png(png)
    path = os.path.join(work, "gate.docx")
    kind, text = call(client, "docx_build", {
        "file": path,
        "title": "Elfogadasi dokumentum",
        "header": "Example office toolkit",
        "footer": "belso hasznalatra",
        "blocks": [
            {"type": "heading", "text": "Bevezeto fejezet", "level": 1},
            {"type": "paragraph", "runs": [
                {"text": "Ez a bekezdes "},
                {"text": "felkover reszt", "bold": True},
                {"text": " es dolt reszt is tartalmaz.", "italic": True},
            ]},
            {"type": "list", "items": ["Pontozott elem", "Masodik pont"]},
            {"type": "list", "items": ["Szamozott elem", "Masodik szam"], "ordered": True},
            {"type": "table", "header": ["Tetel", "Osszeg"],
             "rows": [["Alapdij", "1000"], ["Kamat", "200"], ["Kezelesi", "50"]]},
            {"type": "quote", "text": "Egy kiemelt mondat.", "author": "Owner"},
            {"type": "image", "path": png, "caption": "Abra a keprol"},
            {"type": "page_break"},
            {"type": "heading", "text": "Masodik oldal fejezete", "level": 2},
            {"type": "paragraph", "text": "Ez mar a masodik oldalon van."},
        ],
    })
    print("    docx_build said: %s" % (text or "(no text)"))
    if kind == "harness":
        return HARNESS
    if kind == "error_trace":
        print("    RED: docx_build leaked a stack trace.")
        return RED
    if kind != "ok":
        print("    RED: docx_build refused a request it must handle: %s" % text)
        return RED
    if not os.path.exists(path):
        print("    RED: docx_build reported success but wrote no file.")
        return RED

    out = os.path.join(work, "gate-docx.pdf")
    verdict, _, pdf_text, pages = render(client, path, out)
    if verdict is not None:
        return verdict
    print("    rendered: %s page(s)" % pages)
    if pages != 2:
        print("    CASE 1 RED: expected 2 pages (the page break must start a new one), got %s"
              % pages)
        return RED
    verdict = check_words(pdf_text, [
        "Elfogadasi dokumentum", "Bevezeto fejezet", "felkover reszt", "Pontozott elem",
        "Szamozott elem", "Alapdij", "Kezelesi", "Egy kiemelt mondat", "Owner",
        "Abra a keprol", "Masodik oldal fejezete", "belso hasznalatra",
        "Example office toolkit",
    ], "CASE 1")
    if verdict == GREEN and "1. Szamozott elem" not in pdf_text.replace("  ", " "):
        # The numbered list must render WITH numbers; a numbering part that never
        # gets applied looks identical in the XML string check.
        print("    CASE 1 RED: the numbered list rendered without its numbers.")
        return RED
    if verdict == GREEN:
        print("    CASE 1: 2 pages, every element on the paper, no error marker.")
    return verdict


def case_2_pptx(client, work):
    print("CASE 2 - pptx_build renders every parity-floor element onto slides")
    png = os.path.join(work, "kep.png")
    make_png(png)
    path = os.path.join(work, "gate.pptx")
    kind, text = call(client, "pptx_build", {
        "file": path,
        "title": "Elfogadasi prezentacio",
        "subtitle": "office toolkit",
        "slides": [
            {"title": "Szoveg es listak", "blocks": [
                {"type": "text", "runs": [
                    {"text": "Sima szoveg "},
                    {"text": "felkover resszel", "bold": True},
                ]},
                {"type": "list", "items": ["Pontozott dia elem", "Masodik pont"]},
                {"type": "list", "items": ["Szamozott dia elem"], "ordered": True},
            ]},
            {"title": "Tabla es kiemeles", "blocks": [
                {"type": "table", "header": ["Tetel", "Osszeg"],
                 "rows": [["Alapdij", "1000"], ["Kamat", "200"]]},
                {"type": "callout", "text": "Ez az egy mondat, ami szamit."},
            ]},
            {"title": "Kep a dian", "blocks": [{"type": "image", "path": png, "scale": 0.5}]},
        ],
    })
    print("    pptx_build said: %s" % (text or "(no text)"))
    if kind == "harness":
        return HARNESS
    if kind == "error_trace":
        print("    RED: pptx_build leaked a stack trace.")
        return RED
    if kind != "ok":
        print("    RED: pptx_build refused a request it must handle: %s" % text)
        return RED
    if not os.path.exists(path):
        print("    RED: pptx_build reported success but wrote no file.")
        return RED

    out = os.path.join(work, "gate-pptx.pdf")
    verdict, _, pdf_text, pages = render(client, path, out)
    if verdict is not None:
        return verdict
    print("    rendered: %s page(s)" % pages)
    if pages != 4:
        print("    CASE 2 RED: expected 4 slides (cover + 3), got %s pages. A deck whose "
              "sldIdLst and parts disagree opens short a slide." % pages)
        return RED
    verdict = check_words(pdf_text, [
        "Elfogadasi prezentacio", "office toolkit", "Szoveg es listak",
        "felkover resszel", "Pontozott dia elem", "Szamozott dia elem",
        "Tabla es kiemeles", "Alapdij", "Kamat", "Ez az egy mondat", "Kep a dian",
    ], "CASE 2")
    if verdict == GREEN:
        print("    CASE 2: 4 slides, every element on the slide, no error marker.")
    return verdict


def case_3_refusal(client, work):
    print("CASE 3 - an unsupported block type is REFUSED, cleanly, with nothing written")
    path = os.path.join(work, "must-not-exist.docx")
    kind, text = call(client, "docx_build", {
        "file": path,
        "blocks": [{"type": "heading", "text": "jo blokk", "level": 1},
                   {"type": "footnote", "text": "ezt nem tudjuk"}],
    })
    print("    tool said: %s" % (text or "(no text)"))
    if kind == "harness":
        return HARNESS
    if kind == "error_trace":
        print("    CASE 3 RED: the refusal leaked a stack trace.")
        return RED
    if kind == "ok":
        print("    CASE 3 RED: the unsupported block was ACCEPTED. A builder that silently "
              "drops a block hands back a document the caller believes is complete.")
        return RED
    if "footnote" not in text:
        print("    CASE 3 RED: the refusal does not name the offending block type.")
        return RED
    if not all(word in text for word in ("heading", "paragraph", "list", "table")):
        print("    CASE 3 RED: the refusal does not name what IS supported, so the caller "
              "cannot fix the request from the message.")
        return RED
    if os.path.exists(path):
        print("    CASE 3 RED: the refused call still wrote %s - a file that exists reads "
              "as a success to whoever checks for one." % path)
        return RED
    print("    CASE 3: refused by name, supported list given, nothing written.")
    return GREEN


def case_4_overflow(client, work):
    print("CASE 4 - a slide that runs off the bottom edge is REFUSED by name")
    path = os.path.join(work, "must-not-exist.pptx")
    kind, text = call(client, "pptx_build", {
        "file": path,
        "slides": [{"title": "Tulcsordulo dia",
                    "blocks": [{"type": "list", "items": ["sor %d" % i for i in range(40)]}]}],
    })
    print("    tool said: %s" % (text or "(no text)"))
    if kind == "harness":
        return HARNESS
    if kind == "error_trace":
        print("    CASE 4 RED: the refusal leaked a stack trace.")
        return RED
    if kind == "ok":
        print("    CASE 4 RED: the overflowing slide was accepted - half the content would "
              "be off the slide with nothing in the file to show it.")
        return RED
    if "Tulcsordulo dia" not in text:
        print("    CASE 4 RED: the refusal does not name the slide, so the caller does not "
              "know which one to split.")
        return RED
    if os.path.exists(path):
        print("    CASE 4 RED: the refused call still wrote %s." % path)
        return RED
    print("    CASE 4: refused, slide named, nothing written.")
    return GREEN


def case_5_edit(client, work):
    print("CASE 5 - the edit tools change a real deck, and the change reaches the paper")
    path = os.path.join(work, "edit.pptx")
    kind, text = call(client, "pptx_build", {
        "file": path,
        "slides": [
            {"title": "Ajanlat", "blocks": [
                # Stored as TWO runs, so the value exists nowhere in the XML as one
                # string: a plain search-and-replace finds nothing here.
                {"type": "text", "runs": [
                    {"text": "A hitel osszege 12 5"},
                    {"text": "00 000", "bold": True},
                    {"text": " Ft."},
                ]},
            ]},
            {"title": "Reszletek", "blocks": [
                {"type": "table", "header": ["Tetel", "Osszeg"],
                 "rows": [["Alapdij", "1000"], ["Kamat", "REGI ERTEK"]]},
            ]},
        ],
    })
    print("    pptx_build said: %s" % (text or "(no text)"))
    if kind == "harness":
        return HARNESS
    if kind != "ok":
        print("    HARNESS: could not build the fixture deck: %s" % text)
        return HARNESS

    kind, text = call(client, "pptx_replace_text", {
        "file": path,
        "replacements": [{"find": "12 500 000", "replace": "9 900 000"}],
    })
    print("    pptx_replace_text said: %s" % (text or "(no text)"))
    if kind == "harness":
        return HARNESS
    if kind == "error_trace":
        print("    CASE 5 RED: the edit leaked a stack trace.")
        return RED
    if kind != "ok":
        print("    CASE 5 RED: the replacement failed: %s" % text)
        return RED
    if "1 occurrence" not in text:
        print("    CASE 5 RED: the tool did not report exactly one occurrence. A "
              "run-split value that reports 0 is the exact failure this tool exists "
              "to prevent.")
        return RED

    kind, text = call(client, "pptx_set_table_cell", {
        "file": path, "slide": 2, "table": 0, "row": 2, "column": 1, "text": "UJ ERTEK",
    })
    print("    pptx_set_table_cell said: %s" % (text or "(no text)"))
    if kind == "harness":
        return HARNESS
    if kind != "ok":
        print("    CASE 5 RED: the cell write failed: %s" % text)
        return RED

    out = os.path.join(work, "edit.pdf")
    verdict, _, pdf_text, pages = render(client, path, out)
    if verdict is not None:
        return verdict
    print("    rendered: %s page(s)" % pages)
    verdict = check_words(pdf_text, ["9 900 000", "UJ ERTEK", "Alapdij"], "CASE 5")
    if verdict != GREEN:
        return verdict
    # The negative half: an edit that adds the new text while leaving the old one
    # on the slide is not an edit, and a positive-only check would pass it.
    leftovers = [old for old in ("12 500 000", "REGI ERTEK") if old in pdf_text]
    if leftovers:
        print("    CASE 5 RED: the OLD text is still on the rendered page: %s"
              % ", ".join(leftovers))
        return RED
    print("    CASE 5: run-split value replaced, cell overwritten, old text gone from the PDF.")
    return GREEN


def case_6_edit_refusals(client, work):
    print("CASE 6 - the edit tools refuse what they cannot address, and touch nothing")
    path = os.path.join(work, "refuse.pptx")
    kind, _ = call(client, "pptx_build", {
        "file": path,
        "slides": [{"title": "Egyetlen dia", "blocks": [
            {"type": "text", "text": "nincs itt tabla"}]}],
    })
    if kind != "ok":
        print("    HARNESS: could not build the fixture deck.")
        return HARNESS
    with open(path, "rb") as fh:
        before = fh.read()

    checks = [
        ("a slide past the end", {"file": path, "slide": 9,
                                  "replacements": [{"find": "nincs", "replace": "x"}]},
         "pptx_replace_text", "1 slide"),
        ("a table that is not there", {"file": path, "slide": 1, "table": 0, "row": 0,
                                       "column": 0, "text": "x"},
         "pptx_set_table_cell", "0 table"),
    ]
    for label, arguments, tool, expected in checks:
        kind, text = call(client, tool, arguments)
        print("    %s -> %s" % (label, text or "(no text)"))
        if kind == "harness":
            return HARNESS
        if kind == "error_trace":
            print("    CASE 6 RED: the refusal leaked a stack trace.")
            return RED
        if kind == "ok":
            print("    CASE 6 RED: %s was ACCEPTED - an address that does not exist must "
                  "not come back as success." % label)
            return RED
        if expected not in text:
            print("    CASE 6 RED: the refusal does not state the actual size (expected to "
                  "see %r in the message), so the caller cannot fix the request from it."
                  % expected)
            return RED

    with open(path, "rb") as fh:
        if fh.read() != before:
            print("    CASE 6 RED: a refused edit still rewrote the file.")
            return RED
    print("    CASE 6: both refused with the real size named, file byte-identical.")
    return GREEN


def main():
    server_js = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "office-mcp.js")
    server_js = os.path.abspath(server_js)
    print("Office builder acceptance gate (docx_build / pptx_build)")
    print("server: %s\n" % server_js)

    work = tempfile.mkdtemp(prefix="office-build-gate-")
    client = McpClient(server_js)
    try:
        try:
            client.start()
            init = client.request("initialize", {
                "protocolVersion": PROTOCOL_VERSION,
                "capabilities": {},
                "clientInfo": {"name": "office-build-gate", "version": "1.0.0"},
            }, timeout=20)
        except McpError as exc:
            print("HARNESS: MCP did not start / handshake failed:\n%s" % exc)
            return HARNESS
        if "error" in init:
            print("HARNESS: initialize returned an error: %s" % init["error"])
            return HARNESS
        client.notify("notifications/initialized")

        try:
            listed = client.request("tools/list", {}, timeout=15)
        except McpError as exc:
            print("HARNESS: tools/list failed:\n%s" % exc)
            return HARNESS
        names = [t.get("name") for t in ((listed.get("result") or {}).get("tools") or [])]
        print("tools/list: %s\n" % (", ".join(names) if names else "(none)"))
        for needed in ("docx_build", "pptx_build", "pptx_replace_text", "pptx_set_table_cell"):
            if needed not in names:
                hr()
                print("PENDING: %s is not listed - it is not wired into the MCP yet." % needed)
                return PENDING
        if "office_to_pdf" not in names:
            print("PENDING: office_to_pdf is not listed - the build cannot be proven on paper.")
            return PENDING

        codes = {}
        codes["1"] = case_1_docx(client, work)
        print()
        codes["2"] = case_2_pptx(client, work)
        print()
        codes["3"] = case_3_refusal(client, work)
        print()
        codes["4"] = case_4_overflow(client, work)
        print()
        codes["5"] = case_5_edit(client, work)
        print()
        codes["6"] = case_6_edit_refusals(client, work)

        overall = aggregate(list(codes.values()))
        print()
        hr()
        print("SUMMARY  " + "   ".join("%s=%s" % (k, _NAME[codes[k]])
                                       for k in ("1", "2", "3", "4", "5", "6")))
        print("  1: docx renders   2: pptx renders   3: unsupported refused")
        print("  4: overflow refused   5: edit reaches the paper   6: bad address refused")
        if overall == GREEN:
            print("GREEN: the builders produce files that open and render, the editors change "
                  "them where it shows, and both refuse what they cannot do.")
        elif overall == RED:
            print("RED: see the case above - a real defect, not an environment problem.")
        elif overall == PENDING:
            print("PENDING: the render engine is not available here; re-run where it is.")
        else:
            print("HARNESS: the gate itself could not complete - not a verdict on the tools.")
        return overall
    finally:
        client.close()
        import shutil
        shutil.rmtree(work, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
