#!/usr/bin/env python3
"""Faithful-ish stand-in for `soffice`, modelling the recalc trap for office_recalc tests.

Invoked as an executable via OFFICE_SOFFICE_BIN, so it reacts to the REAL flags office_recalc
drives soffice with:

  --terminate_after_init : create a LibreOffice-shaped user profile (registrymodifications.xcu)
  --convert-to xlsx      : convert INPUT into --outdir. If the job profile has recalc-on-load
                           forced (OOXMLRecalcMode=0), recompute the formula cells; otherwise
                           hand back the STALE cached value, exactly as real soffice does by
                           default. That default-999 behaviour is what makes the 999->5
                           regression meaningful: it is a real trap, not a scripted pass.

FAKE_SOFFICE_BEHAVIOR (env) overrides ONLY the convert step, so profile init still succeeds and
the interesting failure lands where a real converter would hit it:

  recalc (default) | no-output | error | hang

python3 stdlib only. This is a test double; it understands just enough xlsx to move A-column
sums, which is all the gate fixture needs.
"""

import os
import re
import sys
import time
import zipfile

PROFILE_INIT_XCU = (
    '<?xml version="1.0" encoding="UTF-8"?>\n'
    '<oor:items xmlns:oor="http://openoffice.org/2001/registry">\n'
    '</oor:items>\n'
)


def _opt_value(argv, name):
    for i, a in enumerate(argv):
        if a == name and i + 1 < len(argv):
            return argv[i + 1]
    return None


def _profile_dir(argv):
    """Extract the profile path from -env:UserInstallation=file://<dir>."""
    for a in argv:
        if a.startswith("-env:UserInstallation=file://"):
            return a[len("-env:UserInstallation=file://"):]
    return None


def do_profile_init(argv):
    profile = _profile_dir(argv)
    if not profile:
        sys.stderr.write("fake_soffice_recalc: no UserInstallation given for init\n")
        return 1
    user_dir = os.path.join(profile, "user")
    os.makedirs(user_dir, exist_ok=True)
    with open(os.path.join(user_dir, "registrymodifications.xcu"), "w", encoding="utf-8") as fh:
        fh.write(PROFILE_INIT_XCU)
    return 0


def recalc_forced(argv):
    """True if the job profile has recalc-on-load forced (what office_recalc injects)."""
    profile = _profile_dir(argv)
    if not profile:
        return False
    xcu = os.path.join(profile, "user", "registrymodifications.xcu")
    if not os.path.exists(xcu):
        return False
    with open(xcu, encoding="utf-8") as fh:
        body = fh.read()
    return "OOXMLRecalcMode" in body and "<value>0</value>" in body


def _cell_values(sheet_xml):
    """ref -> literal <v> text, for cells that carry a value."""
    out = {}
    for cell in re.finditer(r'<c[^>]*\br="([A-Z]+[0-9]+)"[^>]*>(.*?)</c>', sheet_xml, re.S):
        ref, body = cell.group(1), cell.group(2)
        m = re.search(r"<v[^>]*>(.*?)</v>", body, re.S)
        if m and not re.search(r"<f[^>]*>", body):
            out[ref] = m.group(1)
    return out


def _evaluate(formula, values):
    """Evaluate the tiny subset the fixture uses: 'REF+REF'. Return a number or None."""
    m = re.fullmatch(r"\s*([A-Z]+[0-9]+)\s*\+\s*([A-Z]+[0-9]+)\s*", formula)
    if not m:
        return None
    try:
        return float(values[m.group(1)]) + float(values[m.group(2)])
    except (KeyError, ValueError):
        return None


def _fmt(num):
    return str(int(num)) if num == int(num) else repr(num)


def recompute_sheet(sheet_xml):
    """Overwrite each formula cell's cached <v> with the recomputed value."""
    values = _cell_values(sheet_xml)

    def repl(cell):
        body = cell.group(0)
        fm = re.search(r"<f[^>]*>(.*?)</f>", body, re.S)
        if not fm:
            return body
        result = _evaluate(fm.group(1), values)
        if result is None:
            return body
        if re.search(r"<v[^>]*>.*?</v>", body, re.S):
            return re.sub(r"<v[^>]*>.*?</v>", "<v>%s</v>" % _fmt(result), body, flags=re.S)
        return body.replace("</c>", "<v>%s</v></c>" % _fmt(result))

    return re.sub(r'<c[^>]*\br="[A-Z]+[0-9]+"[^>]*>.*?</c>', repl, sheet_xml, flags=re.S)


def do_convert(argv):
    outdir = _opt_value(argv, "--outdir")
    src = argv[-1] if argv else None
    if not outdir or not src or not os.path.exists(src):
        sys.stderr.write("fake_soffice_recalc: convert needs --outdir and an existing input\n")
        return 1

    forced = recalc_forced(argv)
    os.makedirs(outdir, exist_ok=True)
    dst = os.path.join(outdir, os.path.basename(src))

    with zipfile.ZipFile(src) as zin:
        names = zin.namelist()
        sheet_name = next((n for n in names if re.match(r"xl/worksheets/sheet1\.xml$", n)), None)
        data = {n: zin.read(n) for n in names}

    if forced and sheet_name:
        xml = data[sheet_name].decode("utf-8", "replace")
        data[sheet_name] = recompute_sheet(xml).encode("utf-8")
    # When not forced, the cached (stale) value is copied through untouched - the 999 trap.

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, data[n])
    return 0


def main(argv):
    behavior = os.environ.get("FAKE_SOFFICE_BEHAVIOR", "recalc")

    if "--terminate_after_init" in argv:
        # Profile init always succeeds, so the injected failure lands on convert.
        return do_profile_init(argv)

    if "--convert-to" in argv:
        if behavior == "no-output":
            return 0  # exit clean, produce nothing
        if behavior == "error":
            sys.stderr.write("Error: source file could not be loaded\n")
            return 1
        if behavior == "hang":
            while True:
                time.sleep(3600)
        return do_convert(argv)

    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
