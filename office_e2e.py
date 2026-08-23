#!/usr/bin/env python3
"""Office MCP end-to-end acceptance harness.

Drives the FINISHED office MCP through its TOOL INTERFACE (JSON-RPC over stdio,
tools/call) against a real xlsx and proves the recalc chain actually computes -
not that "a call succeeded", which for a spreadsheet tool proves nothing.

This is the acceptance gate: Office may be reported usable only when a real
end-to-end run passes through the MCP's tools/call. The harness
tests the tool from OUTSIDE, across the contract (tool name + documented
behaviour), so it is independent of whoever is writing the tool's insides.

Three cases, each proving something DIFFERENT, reported separately:

  CASE A - a stale cached value must be OVERWRITTEN.
    A pre-built fixture whose formula cell A3=A1+A2 carries a deliberately WRONG
    cached value (A1=2, A2=3, stored <v>=999). If recalc runs, A3 -> 5; if it
    only trusts the cache, A3 stays 999.

  CASE B - a freshly WRITTEN formula must EVALUATE to a number (the real usage
    chain). Start from an empty workbook, tools/call xlsx_set_cells to write
    A1=2, A2=3, A3=formula "A1+A2", then tools/call xlsx_recalc. There is NO fake
    cache here, so "stale value" does not even apply: the question is whether the
    chain ends with a NUMBER, or with the formula text / an empty cell. Both
    failure modes are recognised and named separately, not just "not 5". This is
    what a caller actually does, and the two tools had never been driven together
    across the contract before.

  CASE C - a missing file must return ONE clean error (isError, no stack trace).

Runs on the CURRENT state too: if a tool is not listed, or answers a clean
"not available / error", the case reports PENDING and does not pretend to pass.
A LEAKED stack trace in a tool error is a defect (RED), not PENDING - that is a
contract violation regardless of the underlying cause.

Exit code is the AGGREGATE over the cases (an acceptance gate: 0 means "Office is
acceptable", nothing weaker). Precedence HARNESS > RED > PENDING > GREEN:
  0  GREEN    every case passed.
  1  RED      a case ran and returned a wrong/stale/empty value, or a tool leaked
              a stack trace / did not error where it must - a real defect.
  2  HARNESS  the test itself could not run (node missing, handshake/timeout,
              a fixture did not build). Not a verdict on the tool.
  3  PENDING  a tool is not there yet, or answered a clean unavailable/error
              (waiting on the engine implementation). Not a proven defect.

Self-contained: python3 stdlib only, no openpyxl, no network. The fixture
helpers are lifted verbatim from office_recalc_gate.py so this file stands alone
and does not couple to another script that may move or change.

  python3 office_e2e.py                       # office-mcp.js next to this file
  python3 office_e2e.py /path/to/office-mcp.js

The recalc ENGINE (host soffice vs container) is chosen inside the tool from its
own env (OFFICE_ENGINE / OFFICE_SOFFICE_BIN); this harness is engine-neutral and
never invokes docker. To isolate the tool logic from the container-probe path,
run it host-pinned:  OFFICE_SOFFICE_BIN=/usr/bin/soffice python3 office_e2e.py
"""

import json
import os
import queue
import re
import subprocess
import sys
import tempfile
import threading
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
PROTOCOL_VERSION = "2025-06-18"  # SDK 1.29.0 supports it; the server negotiates anyway.

# --- fixtures (lifted from office_recalc_gate.py; the deliberately-wrong cache) ----------

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

WORKBOOK = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"""

WORKBOOK_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"""

# A3 holds the formula AND a cached value of 999. A correct engine overwrites it with 5;
# an engine that just trusts the cache hands 999 straight back.
SHEET_STALE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1"><v>2</v></c></row>
<row r="2"><c r="A2"><v>3</v></c></row>
<row r="3"><c r="A3"><f>A1+A2</f><v>999</v></c></row>
</sheetData>
</worksheet>"""

# An empty Sheet1: the CASE B starting point. xlsx_set_cells fills it, then recalc.
SHEET_EMPTY = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData/>
</worksheet>"""


def _write_xlsx(path, sheet_xml):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("xl/workbook.xml", WORKBOOK)
        z.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS)
        z.writestr("xl/worksheets/sheet1.xml", sheet_xml)


def build_fixture(path):
    """Workbook with A3 carrying the deliberately-wrong 999 cache (CASE A)."""
    _write_xlsx(path, SHEET_STALE)


def build_empty_workbook(path):
    """Valid workbook with an empty Sheet1 (CASE B starting point)."""
    _write_xlsx(path, SHEET_EMPTY)


def read_a3(path):
    """Return (formula, cached_value) for A3, reading the raw sheet XML."""
    with zipfile.ZipFile(path) as z:
        name = next((n for n in z.namelist()
                     if re.match(r"xl/worksheets/sheet1\.xml$", n)), None)
        if name is None:
            return None, "no xl/worksheets/sheet1.xml in the output"
        xml = z.read(name).decode("utf-8", "replace")
    cell = re.search(r'<c[^>]*r="A3"[^>]*>(.*?)</c>', xml, re.S)
    if not cell:
        return None, "A3 missing from the output"
    body = cell.group(1)
    f = re.search(r"<f[^>]*>(.*?)</f>", body, re.S)
    v = re.search(r"<v[^>]*>(.*?)</v>", body, re.S)
    return (f.group(1) if f else None), (v.group(1) if v else None)


def read_a3_detail(path):
    """Structured read of A3: (status, formula, value).

    status: 'ok'         - an A3 cell exists (formula/value may still be None)
            'empty_cell' - a self-closing <c r="A3"/> with nothing in it
            'no_a3'      - the sheet has no A3 cell at all
            'no_sheet'   - sheet1.xml is missing from the workbook
    Distinguishing these matters: 'formula text but no value' and 'cell empty /
    absent' are two different failure modes, and the report names each.
    """
    with zipfile.ZipFile(path) as z:
        name = next((n for n in z.namelist()
                     if re.match(r"xl/worksheets/sheet1\.xml$", n)), None)
        if name is None:
            return "no_sheet", None, None
        xml = z.read(name).decode("utf-8", "replace")
    cell = re.search(r'<c[^>]*r="A3"[^>]*>(.*?)</c>', xml, re.S)
    if not cell:
        if re.search(r'<c[^>]*r="A3"[^>]*/>', xml):
            return "empty_cell", None, None
        return "no_a3", None, None
    body = cell.group(1)
    f = re.search(r"<f[^>]*>(.*?)</f>", body, re.S)
    v = re.search(r"<v[^>]*>(.*?)</v>", body, re.S)
    return "ok", (f.group(1) if f else None), (v.group(1) if v else None)


# --- minimal JSON-RPC / MCP stdio client ------------------------------------------------

class McpError(Exception):
    """The harness could not talk to the server as expected (HARNESS-class)."""


class McpClient:
    """Speaks newline-delimited JSON-RPC 2.0 to an MCP stdio server.

    The MCP stdio transport frames each message as one line of JSON on stdout;
    the server's own logging goes to stderr, which is captured separately so a
    log line never gets parsed as a protocol message.
    """

    def __init__(self, server_js):
        self.server_js = server_js
        self.proc = None
        self._q = queue.Queue()
        self._stderr = []
        self._next_id = 0

    def start(self):
        if not os.path.exists(self.server_js):
            raise McpError("office-mcp.js not found at %s" % self.server_js)
        try:
            self.proc = subprocess.Popen(
                ["node", self.server_js],
                cwd=os.path.dirname(self.server_js) or ".",
                stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                text=True, bufsize=1,
            )
        except FileNotFoundError as e:
            raise McpError("could not start node: %s" % e)
        threading.Thread(target=self._pump_stdout, daemon=True).start()
        threading.Thread(target=self._pump_stderr, daemon=True).start()

    def _pump_stdout(self):
        for line in self.proc.stdout:
            line = line.strip()
            if not line:
                continue
            try:
                self._q.put(json.loads(line))
            except json.JSONDecodeError:
                # Not a protocol message (should not happen: logs go to stderr).
                self._stderr.append("[stdout non-json] " + line)

    def _pump_stderr(self):
        for line in self.proc.stderr:
            self._stderr.append(line.rstrip())

    def stderr_text(self):
        return "\n".join(self._stderr).strip()

    def notify(self, method, params=None):
        self._send({"jsonrpc": "2.0", "method": method, "params": params or {}})

    def request(self, method, params=None, timeout=15):
        self._next_id += 1
        rid = self._next_id
        self._send({"jsonrpc": "2.0", "id": rid, "method": method, "params": params or {}})
        return self._wait_for(rid, timeout)

    def _send(self, msg):
        if not self.proc or self.proc.poll() is not None:
            raise McpError("server process is not running (exited %s). stderr:\n%s"
                           % (self.proc.returncode if self.proc else "n/a", self.stderr_text()))
        self.proc.stdin.write(json.dumps(msg) + "\n")
        self.proc.stdin.flush()

    def _wait_for(self, rid, timeout):
        """Return the JSON-RPC message whose id matches rid, ignoring others."""
        import time
        held = []
        deadline = time.monotonic() + timeout
        while True:
            get_timeout = max(0.05, deadline - time.monotonic())
            if deadline - time.monotonic() <= 0:
                break
            try:
                msg = self._q.get(timeout=get_timeout)
            except queue.Empty:
                break
            if msg.get("id") == rid:
                for h in held:
                    self._q.put(h)
                return msg
            # A message we did not ask for (notification, or a later id). Hold it.
            held.append(msg)
        if self.proc.poll() is not None:
            raise McpError("server exited (%s) while waiting for response id=%s. stderr:\n%s"
                           % (self.proc.returncode, rid, self.stderr_text()))
        raise McpError("timed out after %ss waiting for response id=%s. stderr:\n%s"
                       % (timeout, rid, self.stderr_text()))

    def close(self):
        if not self.proc:
            return
        try:
            if self.proc.stdin and not self.proc.stdin.closed:
                self.proc.stdin.close()
        except Exception:
            pass
        try:
            self.proc.terminate()
            self.proc.wait(timeout=10)
        except Exception:
            try:
                self.proc.kill()
            except Exception:
                pass


# --- report + verdict helpers -----------------------------------------------------------

GREEN, RED, HARNESS, PENDING = 0, 1, 2, 3
_NAME = {GREEN: "GREEN", RED: "RED", HARNESS: "HARNESS", PENDING: "PENDING"}


def hr():
    print("-" * 72)


def tool_result_text(result):
    """Flatten the text blocks of an MCP tools/call result."""
    parts = [b.get("text", "") for b in (result.get("content") or []) if b.get("type") == "text"]
    return "\n".join(parts).strip()


def aggregate(codes):
    """Worst-wins over the cases: HARNESS > RED > PENDING > GREEN."""
    for c in (HARNESS, RED, PENDING):
        if c in codes:
            return c
    return GREEN


def call_recalc(client, path, timeout=180):
    """tools/call xlsx_recalc; classify the RESPONSE (not yet the file).

    Returns (kind, text):
      'harness'      - the call did not answer (McpError)
      'proto_error'  - a JSON-RPC error object came back
      'error_clean'  - isError with a single clean message (no stack trace)
      'error_trace'  - isError with a LEAKED stack trace (a defect regardless of cause)
      'ok'           - success; the caller must now inspect the FILE
    """
    try:
        called = client.request("tools/call", {"name": "xlsx_recalc", "arguments": {"file": path}},
                                timeout=timeout)
    except McpError as e:
        return "harness", str(e)
    if "error" in called:
        return "proto_error", str(called["error"])
    result = called.get("result") or {}
    text = tool_result_text(result)
    if result.get("isError"):
        if "Traceback" in text or "NameError" in text or "most recent call last" in text:
            return "error_trace", text
        return "error_clean", text
    return "ok", text


def judge_a3_is_five(path, tag):
    """Read A3 from the FILE and decide GREEN/RED, naming each failure mode."""
    try:
        st, f, v = read_a3_detail(path)
    except Exception as e:
        print("    HARNESS: could not read the workbook back: %s" % e)
        return HARNESS
    print("    A3 in file: status=%s formula=%r value=%r" % (st, f, v))
    if st == "ok" and v == "5":
        print("    %s: A3 computed to 5 in the file - the value was actually calculated." % tag)
        return GREEN
    if st in ("no_sheet", "no_a3", "empty_cell") or (st == "ok" and f is None and v is None):
        print("    %s RED: A3 is empty/absent after the chain - the cell carries no value at all "
              "(not a number, not even a formula)." % tag)
        return RED
    if st == "ok" and v is None and f is not None:
        print("    %s RED: A3 still holds only the FORMULA TEXT %r with no computed value - a "
              "reader that does not recalculate sees %r, not a number." % (tag, f, "=" + f))
        return RED
    if st == "ok" and v == "999":
        print("    %s RED: A3 is still the stale 999 - recalc did not overwrite the cache." % tag)
        return RED
    print("    %s RED: A3=%r, expected 5 (wrong number)." % (tag, v))
    return RED


def verdict_from_recalc(kind, text, path, tag):
    """Turn a call_recalc outcome + the file into a case verdict, with a printed reason."""
    print("    tool said: %s" % (text or "(no text)"))
    if kind == "harness":
        print("    HARNESS: xlsx_recalc did not answer.")
        return HARNESS
    if kind == "proto_error":
        print("    PENDING: xlsx_recalc returned a protocol error (not a value defect).")
        return PENDING
    if kind == "error_clean":
        print("    PENDING: xlsx_recalc returned a clean error - recalc not available here. The "
              "tool failed loudly with a message, which is correct behaviour, not a defect.")
        return PENDING
    if kind == "error_trace":
        print("    RED: xlsx_recalc leaked a STACK TRACE into its response. Whatever the root "
              "cause, a tool must return one clean message - a caller must never receive a")
        print("    Python traceback. (This is the current NameError regression surfacing; it is "
              "a tool error-handling defect, not a fault of this test's chain.)")
        return RED
    return judge_a3_is_five(path, tag)


# --- the cases --------------------------------------------------------------------------

def case_a_stale_cache(client, work):
    print("CASE A - a stale cached value must be OVERWRITTEN")
    print("  fixture: A1=2, A2=3, A3=formula A1+A2 with a deliberately wrong cached <v>=999")
    path = os.path.join(work, "caseA.xlsx")
    try:
        build_fixture(path)
        f, v = read_a3(path)
    except Exception as e:
        print("  HARNESS: could not build the fixture: %s" % e)
        return HARNESS
    if (f, v) != ("A1+A2", "999"):
        print("  HARNESS: fixture did not build as intended (want formula A1+A2, cached 999)")
        return HARNESS
    print("  built ok (A3 cached=%r, deliberately wrong)" % v)
    print("  tools/call xlsx_recalc")
    kind, text = call_recalc(client, path)
    return verdict_from_recalc(kind, text, path, "CASE A")


def case_b_fresh_formula(client, work, names):
    print("CASE B - a freshly WRITTEN formula must EVALUATE to a number (real usage chain)")
    if "xlsx_set_cells" not in names:
        print("  PENDING: xlsx_set_cells is not listed - cannot exercise the write->recalc chain.")
        return PENDING
    path = os.path.join(work, "caseB.xlsx")
    try:
        build_empty_workbook(path)
    except Exception as e:
        print("  HARNESS: could not build the empty workbook: %s" % e)
        return HARNESS
    print("  empty workbook built (Sheet1, no data - there is NO fake cache here)")

    # 1) write the cells through the tool.
    print("  tools/call xlsx_set_cells  (A1=2, A2=3, A3=formula 'A1+A2')")
    try:
        set_call = client.request("tools/call", {"name": "xlsx_set_cells", "arguments": {
            "file": path, "sheet": "Sheet1",
            "cells": [
                {"ref": "A1", "value": 2},
                {"ref": "A2", "value": 3},
                {"ref": "A3", "formula": "A1+A2"},
            ],
        }}, timeout=60)
    except McpError as e:
        print("  HARNESS: xlsx_set_cells did not answer: %s" % e)
        return HARNESS
    if "error" in set_call:
        print("  RED: xlsx_set_cells returned a protocol error: %s" % set_call["error"])
        return RED
    sres = set_call.get("result") or {}
    stext = tool_result_text(sres)
    print("    tool said: %s" % (stext or "(no text)"))
    if sres.get("isError"):
        print("  RED: the write stage failed (xlsx_set_cells): %s" % stext)
        return RED
    try:
        st, f, v = read_a3_detail(path)
        print("    A3 after set_cells: status=%s formula=%r value=%r (formula, no cached number "
              "by design)" % (st, f, v))
    except Exception as e:
        print("  HARNESS: could not read back after set_cells: %s" % e)
        return HARNESS

    # 2) recalc, then the only thing that matters: is there a NUMBER in the file?
    print("  tools/call xlsx_recalc")
    kind, text = call_recalc(client, path)
    return verdict_from_recalc(kind, text, path, "CASE B")


def case_c_error_path(client, work):
    print("CASE C - a missing file must return ONE clean error (no stack trace)")
    missing = os.path.join(work, "does-not-exist.xlsx")
    print("  tools/call xlsx_recalc  (file=%s)" % missing)
    try:
        err_call = client.request("tools/call", {"name": "xlsx_recalc", "arguments": {"file": missing}},
                                  timeout=60)
    except McpError as e:
        print("  HARNESS: the error-path call did not answer: %s" % e)
        return HARNESS
    er = err_call.get("result") or {}
    etext = tool_result_text(er)
    is_err = bool(er.get("isError")) or ("error" in err_call)
    print("    tool said: %s" % (etext or err_call.get("error") or "(no text)"))
    has_trace = ("Traceback" in etext) or ("most recent call last" in etext)
    clean = is_err and etext and not has_trace and etext.count("\n") <= 3
    if clean:
        print("    CASE C: one clear message, isError set, no stack trace.")
        return GREEN
    if has_trace:
        print("    CASE C RED: the error path leaked a stack trace - a missing file must come back "
              "as one clean message, never a Python traceback.")
        return RED
    print("    CASE C RED: the error path is not clean (expected isError + one short message).")
    return RED


# --- the run ----------------------------------------------------------------------------

def main():
    server_js = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "office-mcp.js")
    server_js = os.path.abspath(server_js)
    print("Office MCP end-to-end acceptance harness")
    print("server: %s" % server_js)
    engine_env = os.environ.get("OFFICE_ENGINE") or (
        "host-pinned (OFFICE_SOFFICE_BIN)" if os.environ.get("OFFICE_SOFFICE_BIN") else "AUTO (tool default)")
    print("engine selection (chosen inside the tool): %s\n" % engine_env)

    work = tempfile.mkdtemp(prefix="office-e2e-")
    client = McpClient(server_js)
    try:
        # Handshake.
        try:
            client.start()
            init = client.request("initialize", {
                "protocolVersion": PROTOCOL_VERSION,
                "capabilities": {},
                "clientInfo": {"name": "office-e2e", "version": "1.0.0"},
            }, timeout=20)
        except McpError as e:
            print("HARNESS: MCP did not start / handshake failed:\n%s" % e)
            return HARNESS
        if "error" in init:
            print("HARNESS: initialize returned an error: %s" % init["error"])
            return HARNESS
        srv = (init.get("result") or {}).get("serverInfo") or {}
        print("handshake ok: server %s v%s, protocol %s"
              % (srv.get("name", "?"), srv.get("version", "?"),
                 (init.get("result") or {}).get("protocolVersion", "?")))
        client.notify("notifications/initialized")

        # tools/list.
        try:
            listed = client.request("tools/list", {}, timeout=15)
        except McpError as e:
            print("HARNESS: tools/list failed:\n%s" % e)
            return HARNESS
        names = [t.get("name") for t in ((listed.get("result") or {}).get("tools") or [])]
        print("tools/list: %s\n" % (", ".join(names) if names else "(none)"))
        if "xlsx_recalc" not in names:
            hr()
            print("PENDING: xlsx_recalc is not listed yet - the recalc tool is not wired.")
            print("Re-run once it appears; the write->recalc chain will be proven in one shot.")
            return PENDING

        # The three cases, each reported on its own.
        codes = {}
        codes["A"] = case_a_stale_cache(client, work)
        print()
        codes["B"] = case_b_fresh_formula(client, work, names)
        print()
        codes["C"] = case_c_error_path(client, work)

        overall = aggregate(list(codes.values()))
        print()
        hr()
        print("SUMMARY  " + "   ".join("%s=%s" % (k, _NAME[codes[k]]) for k in ("A", "B", "C")))
        print("  A: stale 999 overwritten   B: fresh formula -> number   C: clean error on missing file")
        if overall == GREEN:
            print("GREEN: every case passed end-to-end through tools/call. Office acceptance PASSES.")
        elif overall == RED:
            print("RED: at least one case returned a wrong value or a leaked stack trace. If a RED is")
            print("the 'leaked stack trace' note above, that is the current NameError regression in")
            print("the tool, not a fault of the chain being tested - re-run host-pinned to isolate.")
        elif overall == PENDING:
            print("PENDING: recalc is not available in this environment yet. Nothing is proven broken;")
            print("re-run once the engine is reachable.")
        else:
            print("HARNESS: the test itself could not complete - see above. Not a verdict on the tool.")
        return overall
    finally:
        client.close()
        import shutil
        shutil.rmtree(work, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
