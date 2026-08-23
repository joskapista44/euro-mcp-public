#!/usr/bin/env python3
"""Runs one DocBuilder job on the euro-office box and reports the outcome as one JSON line.

It is sent over ssh stdin per call and installs nothing: the temp directory it serves from is
removed on the way out, so the box keeps no trace of a run. That is deliberate - the access to
this machine is scoped to reaching euro-office and nothing else.

Everything happens here because the Document Server FETCHES the builder script and the input
document itself, over HTTP, from an address IT can reach. Running the call from elsewhere and
serving the files from elsewhere are two different requirements, and the second one is what
forces the work onto this host.

argv[1]  base64 of the builder script
argv[2]  optional: base64 of an input .docx. Without it a minimal document is generated here.
argv[3]  optional: the address the Document Server reaches this host at (default 172.22.0.1,
         measured on 2026-08-13 as the eo_experiment_net gateway)

The JWT is read from the deployment's own env file and never printed, never passed on a command
line, and never included in the output.
"""

import base64
import hashlib
import hmac
import html
import io
import json
import os
import re
import shutil
import socket
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.request
import zipfile
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

DS_URL = os.environ.get("EO_DS_URL", "http://127.0.0.1:8081")
# No default here: a hardcoded path would only exist on the original deployment's box and would
# fail confusingly (as a missing-file error, not a configuration error) for anyone else. Required
# via EO_ENV_FILE; the check at its first use below fails closed with a named reason if unset.
ENV_FILE = os.environ.get("EO_ENV_FILE", "")

# office-diag-e6-trace-id: set ONCE per process in main(), from argv[5]. Safe as a module global
# HERE (unlike the Node side) because this script is a brand-new OS process per job -- there is
# no second concurrent call sharing this interpreter to clobber it.
TRACE_ID = None


def fail(reason, **extra):
    print(json.dumps({"ok": False, "outcome": "nem-mert", "detail": reason, "traceId": TRACE_ID, **extra}))
    sys.exit(2)


def read_jwt():
    if not ENV_FILE:
        fail("EO_ENV_FILE is not set -- no default is provided, point it at the env file on this box")
    if not os.path.exists(ENV_FILE):
        fail(f"the env file does not exist: {ENV_FILE}")
    for line in open(ENV_FILE, encoding="utf-8", errors="replace"):
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        if key.strip() in ("JWT_SECRET", "JWT_TOKEN", "DS_JWT"):
            return value.strip().strip('"').strip("'")
    fail("no JWT key found in the env file (looked for JWT_SECRET / JWT_TOKEN / DS_JWT)")


def b64url(raw):
    return base64.urlsafe_b64encode(raw).rstrip(b"=")


def sign(secret, payload):
    header = b64url(json.dumps({"alg": "HS256", "typ": "JWT"}).encode())
    body = b64url(json.dumps(payload).encode())
    sig = b64url(hmac.new(secret.encode(), header + b"." + body, hashlib.sha256).digest())
    return (header + b"." + body + b"." + sig).decode()


# A minimal but genuinely valid docx, so the Document Server has something real to open. Written
# with the stdlib zipfile rather than a template file, because a template would be one more thing
# living on this box.
MINIMAL_DOCX_PARTS = {
    "[Content_Types].xml": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        "</Types>"
    ),
    "_rels/.rels": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        "</Relationships>"
    ),
    "word/_rels/document.xml.rels": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
    ),
    "word/document.xml": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        "<w:body><w:p><w:r><w:t>EURO-MCP bemeneti dokumentum</w:t></w:r></w:p></w:body></w:document>"
    ),
}


def write_minimal_docx(path):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, body in MINIMAL_DOCX_PARTS.items():
            z.writestr(name, body)


def detect_kind(raw):
    """What did the builder actually produce? Returns (kind, main part name).

    The type is read out of the PACKAGE, not out of the request: a caller can ask for one
    output type and a script can save another, and only the bytes know which happened.
    """
    ismertek = (
        ("docx", "word/document.xml"),
        ("xlsx", "xl/workbook.xml"),
        ("pptx", "ppt/presentation.xml"),
    )
    try:
        with zipfile.ZipFile(io.BytesIO(raw)) as z:
            nevek = set(z.namelist())
    except Exception:
        return "ismeretlen", None
    for kind, fo in ismertek:
        if fo in nevek:
            return kind, fo
    return "ismeretlen", None


# The `export`/
# `documentStats` OPERATIONS entries (lib.cjs) append their result as a new paragraph carrying one
# of these fixed, distinctive prefixes -- never occurring in ordinary content -- so each marker's
# own payload is bounded by the START OF THE NEXT KNOWN MARKER (or end of string) rather than by a
# paragraph boundary: every paragraph in the saved package lands back-to-back with NO separator
# once XML tags are stripped (an existing, pre-dated property of the `text` variable below, left
# unchanged here), so there is no other reliable boundary to bound on.
_EXPORT_MARKER_PREFIXES = (
    "__EXPORT_MARKDOWN__:",
    "__EXPORT_HTML__:",
    "__STATS_JSON__:",
    "__CUSTOM_PROPERTIES_JSON__:",
)


def _fully_unescape(s):
    """Repeats `html.unescape()` until a pass no longer changes anything (bounded).

    MEASURED (2026-08-17): the marker payloads carry DIFFERENT escaping DEPTHS depending on their
    source. `JSON.stringify()` output (the `documentStats`/`GetCustomProperties` markers) has no
    escaping of its own, so `AddText()`'s XML-escaping is the ONLY layer -- one pass fully cleans
    it. `ToMarkdown()`/`ToHtml()` output is ALREADY html-entity-escaped by the time it reaches
    `AddText()` (a plain `&` the caller typed came back as `&amp;amp;` in the raw package, not
    `&amp;` -- confirmed by comparing the raw capture against the once-unescaped result directly),
    so it needs a SECOND pass. A fixed-count pass would be wrong for one of these two cases; a
    bounded fixed-point loop self-adapts to whichever depth a given marker's source produces and
    stops as soon as a pass changes nothing, which is also why it is safe on ordinary content that
    was never escaped at all (a no-op on the first pass).

    *** A NAMED LIMIT (2026-08-17), NOT fixed on
    purpose: this cannot tell "double-escaped by AddText()" apart from "genuinely contains an
    entity-looking substring as its own content" -- a document whose TEXT is about escaping (e.g.
    exported markdown containing the literal string "&amp;quot;terv&amp;quot;" as something the
    caller wrote, not an artifact) would come back over-unescaped ("terv" instead of the intended
    "&quot;terv&quot;"). A single fixed pass would get that case right but corrupt the FAR more
    common double-escaped-by-the-engine case instead -- this trade favours the common case on
    purpose. If a caller ever needs the literal-entity case preserved, that needs a different
    signal than the text itself (this function has none), not a smaller iteration bound. ***
    """
    for _ in range(5):
        next_s = html.unescape(s)
        if next_s == s:
            return s
        s = next_s
    return s


def extract_export_fields(text):
    """Pulls the export/documentStats OPERATIONS' marker paragraphs out of the RAW (still
    XML-entity-escaped) stripped document text into named, UNTRUNCATED JSON fields.

    `text` is used only to find marker POSITIONS -- the existing `markersFound`/`documentText`/
    `outcome` fields built from it elsewhere are left byte-for-byte unchanged. Each captured
    PAYLOAD is unescaped (see `_fully_unescape`) before use: DocBuilder's `AddText()` XML-escapes
    the text it writes (measured: a JSON string like `["myMarker"]` lands in the package as
    `[&quot;myMarker&quot;]`), so `json.loads()` on the raw capture would fail on every call that
    actually has content -- this is not a hypothetical, it reproduces on every `documentStats`/
    `export` call that contains a quote, which is all of them.

    If TWO `export` operations with different formats land in the same batch (markdown AND html),
    the later prefix in `_EXPORT_MARKER_PREFIXES` wins `exportResult` -- an unmeasured, unusual
    caller pattern (the operation only ever emits one marker per call), not solved here.
    """
    boundary = "|".join(re.escape(p) for p in _EXPORT_MARKER_PREFIXES)
    fields = {"exportResult": None, "stats": None, "customProperties": None}
    for prefix in _EXPORT_MARKER_PREFIXES:
        m = re.search(re.escape(prefix) + r"(.*?)(?:" + boundary + r"|$)", text, re.S)
        if not m:
            continue
        payload = _fully_unescape(m.group(1))
        if prefix in ("__EXPORT_MARKDOWN__:", "__EXPORT_HTML__:"):
            fields["exportResult"] = {
                "format": "markdown" if prefix == "__EXPORT_MARKDOWN__:" else "html",
                "content": payload,
            }
        else:
            key = "stats" if prefix == "__STATS_JSON__:" else "customProperties"
            try:
                fields[key] = json.loads(payload)
            except json.JSONDecodeError:
                fields[key] = {"parseError": True, "raw": payload}
    return fields


def free_port():
    s = socket.socket()
    s.bind(("0.0.0.0", 0))
    port = s.getsockname()[1]
    s.close()
    return port


def main():
    global TRACE_ID
    if len(sys.argv) < 2:
        fail("no builder script given")
    script = base64.b64decode(sys.argv[1]).decode("utf-8")
    doc_b64 = sys.argv[2] if len(sys.argv) > 2 and sys.argv[2] not in ("", "-") else None
    box_ip = sys.argv[3] if len(sys.argv) > 3 and sys.argv[3] else "172.22.0.1"
    # argv[4] is docFlag ("return-doc" or "-"), checked further down where it was already read.
    TRACE_ID = sys.argv[5] if len(sys.argv) > 5 and sys.argv[5] else None

    secret = read_jwt()
    work = tempfile.mkdtemp(prefix="euro-mcp-")
    try:
        if doc_b64:
            open(os.path.join(work, "bemenet.docx"), "wb").write(base64.b64decode(doc_b64))
        else:
            write_minimal_docx(os.path.join(work, "bemenet.docx"))

        port = free_port()
        # Bound to 0.0.0.0 on purpose: a loopback-bound listener answers fine from this host and
        # is invisible to the container, which is the failure that reads as a network problem.
        handler = partial(SimpleHTTPRequestHandler, directory=work)
        httpd = ThreadingHTTPServer(("0.0.0.0", port), handler)
        threading.Thread(target=httpd.serve_forever, daemon=True).start()

        doc_url = f"http://{box_ip}:{port}/bemenet.docx"
        script = script.replace("__DOC_URL__", doc_url)
        open(os.path.join(work, "script.docbuilder"), "w", encoding="utf-8").write(script)

        # Prove our own precondition before blaming the service: if our own server does not hand
        # back the script we just wrote, the failure is ours and no DocBuilder error would say so.
        try:
            with urllib.request.urlopen(f"http://127.0.0.1:{port}/script.docbuilder", timeout=5) as r:
                own = r.read().decode("utf-8", "replace")
            if "builder.OpenFile(" not in own:
                fail("our own file server did not return the script we just wrote")
        except Exception as exc:
            fail(f"our own file server is not answering: {type(exc).__name__}")

        body = {"async": False, "url": f"http://{box_ip}:{port}/script.docbuilder"}
        token = sign(secret, body)
        body["token"] = token
        req = urllib.request.Request(
            DS_URL.rstrip("/") + "/docbuilder",
            data=json.dumps(body).encode(),
            headers={"Content-Type": "application/json", "Authorization": "Bearer " + token},
        )
        try:
            with urllib.request.urlopen(req, timeout=90) as r:
                answer = json.loads(r.read().decode())
        except urllib.error.HTTPError as exc:
            fail(f"the Document Server answered HTTP {exc.code}")
        except Exception as exc:
            fail(f"the Document Server could not be reached: {type(exc).__name__}")

        if not isinstance(answer, dict) or "urls" not in answer:
            code = answer.get("error") if isinstance(answer, dict) else None
            named = {-3: "blocked", -8: "auth", 6: "auth", -4: "fetch"}.get(code, "unknown")
            print(json.dumps({"ok": False, "outcome": named, "dsError": code, "traceId": TRACE_ID}))
            return

        out_url = list(answer["urls"].values())[0]
        with urllib.request.urlopen(out_url, timeout=60) as r:
            saved = r.read()
        # The extension follows what the builder actually produced. It used to be hard-coded to
        # .docx, which was harmless by itself -- but the verification below read word/document.xml
        # out of it, and on a spreadsheet or a presentation that raised, returned early, and the
        # caller got outcome="nem-docx" with written=null. The document existed on disk at that
        # moment and was simply dropped. That silent null is why two files were built on a
        # different route entirely before anyone noticed.
        kind, main_part = detect_kind(saved)
        saved_path = os.path.join(work, "eredmeny." + (kind if kind != "ismeretlen" else "bin"))
        open(saved_path, "wb").write(saved)

        result = {
            "savedBytes": len(saved),
            "dsKey": answer.get("key"),
            "kind": kind,
            "traceId": TRACE_ID,
        }
        if kind == "ismeretlen":
            # Still not a reason to withhold the bytes: the caller asked for a document and one
            # came back. We say we cannot classify it, and hand it over anyway.
            result["ok"] = False
            result["outcome"] = "ismeretlen-csomag"
            result["detail"] = "the produced file is not a recognised OOXML package"
        elif kind == "docx":
            # Unchanged from before for the text core: the MARKER is what says the edit happened,
            # and the marker only exists because edit_document puts one there.
            with zipfile.ZipFile(saved_path) as z:
                xml = z.read(main_part).decode("utf-8", "replace")
            text = re.sub(r"<[^>]*>", "", xml)
            markers = re.findall(r"EURO-MCP-[0-9]+-[0-9a-f]+", text)
            result["ok"] = bool(markers)
            result["outcome"] = "ok" if markers else "nem-valtozott"
            result["markersFound"] = markers
            result["documentText"] = text[:400]
            result.update(extract_export_fields(text))
        else:
            # xlsx / pptx: there is no marker convention here, so we do NOT claim the content was
            # verified. What we CAN say is that a well-formed package of the expected kind came
            # back, and how much of it there is -- which is more than the old null said, and less
            # than a bare "ok" would imply. *A response that cannot verify must say so, not go quiet.*
            with zipfile.ZipFile(saved_path) as z:
                parts = z.namelist()
                body = z.read(main_part).decode("utf-8", "replace")
            result["ok"] = True
            result["outcome"] = "kesz-tartalom-nem-ellenorizve"
            result["parts"] = len(parts)
            result["documentText"] = re.sub(r"<[^>]*>", " ", body)[:400]
        # The edited document only exists inside the temp directory this run is about to delete,
        # so a caller that wants it back has to be handed the bytes now. It is opt-in because
        # most calls only need the verdict, and a megabyte of base64 on every status check would
        # be paid for nothing.
        if len(sys.argv) > 4 and sys.argv[4] == "return-doc":
            result["savedBase64"] = base64.b64encode(saved).decode()
        print(json.dumps(result))
    finally:
        shutil.rmtree(work, ignore_errors=True)


if __name__ == "__main__":
    main()
