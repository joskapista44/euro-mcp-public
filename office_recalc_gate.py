#!/usr/bin/env python3
"""soffice recalc gate.

Builds an xlsx whose formula cell carries a deliberately WRONG cached value, runs it
through headless LibreOffice, and reads the result back.

  A1 = 2
  A2 = 3
  A3 = formula "A1+A2", cached <v> = 999     <- the lie

If the engine recalculates, A3 comes back 5. If it only copies the cache, it comes back
999. There is no third outcome, and "the conversion succeeded" proves neither: this is
exactly how the OnlyOffice ConvertService passed its test while returning stale numbers.

Self-contained: python3 stdlib only, no openpyxl, no network.

  python3 office_recalc_gate.py                       # soffice from PATH
  python3 office_recalc_gate.py /path/to/soffice      # a specific binary
  python3 office_recalc_gate.py --docker <image>      # soffice inside a container

The container mode exists so the same gate can be pointed at the packaged engine once it
is built (office-lo.Dockerfile). The result is not assumed to carry over from the host: a
different LibreOffice version can behave differently, and this setting is precisely the
kind that changes between them.
"""

import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

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
SHEET = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1"><v>2</v></c></row>
<row r="2"><c r="A2"><v>3</v></c></row>
<row r="3"><c r="A3"><f>A1+A2</f><v>999</v></c></row>
</sheetData>
</worksheet>"""


def build_fixture(path):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("xl/workbook.xml", WORKBOOK)
        z.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS)
        z.writestr("xl/worksheets/sheet1.xml", SHEET)


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


# Recalculation on load is OFF by default for xlsx, and --convert-to does not override it:
# the converter happily writes the stale cached value back out. The setting lives in the
# user profile, so the only way to get a recalculation is to hand soffice a profile that
# asks for one. 0 = Always, 1 = Never, 2 = Prompt.
RECALC_ALWAYS_XCU = (
    '<item oor:path="/org.openoffice.Office.Calc/Formula/Load">'
    '<prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>'
    '<item oor:path="/org.openoffice.Office.Calc/Formula/Load">'
    '<prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>'
)


class HostRunner:
    """soffice installed on this machine."""

    def __init__(self, soffice):
        self.soffice = soffice
        self.label = soffice
        self.uid, self.gid = os.getuid(), os.getgid()

    # What soffice sees and what Python opens are the same path here.
    def seen(self, jobdir, rel):
        return os.path.join(jobdir, rel)

    def command(self, jobdir, args):
        return [self.soffice] + args


class DockerRunner:
    """soffice inside a short-lived container (see office-lo.Dockerfile).

    The job directory is the only thing mounted, so every path soffice is given has to be
    rewritten to its location inside the container. --network=none because a converter
    never needs the network and documents can carry external references; -u so the output
    belongs to the invoking user rather than root.
    """

    def __init__(self, image):
        self.image = image
        # Which uid the container process runs as. Defaults to the caller, but has to be
        # overridable: the case worth testing is a NON-root uid with no /etc/passwd entry
        # in the image, and whoever drives docker often has to do so as root - which would
        # otherwise silently test uid 0 and skip exactly the path HOME=/work exists for.
        env_uid = os.environ.get("OFFICE_GATE_UID")
        if env_uid and ":" in env_uid:
            self.user = env_uid
        elif env_uid:
            self.user = "%s:%s" % (env_uid, env_uid)
        else:
            self.user = "%d:%d" % (os.getuid(), os.getgid())
        self.uid, self.gid = (int(x) for x in self.user.split(":"))
        self.label = "docker:%s (uid %s)" % (image, self.user)

    def seen(self, jobdir, rel):
        return "/work/" + rel

    def command(self, jobdir, args):
        # HOME must point inside the mount: the uid from -u has no /etc/passwd entry in
        # the image, so LibreOffice has nowhere to put its home and the profile init can
        # fail in a way that looks unrelated to permissions. (Raised during an earlier ops review.)
        return ["docker", "run", "--rm", "--network=none",
                "-v", jobdir + ":/work",
                "-u", self.user,
                "-e", "HOME=/work",
                self.image, "soffice"] + args


def prepare_profile(runner, jobdir, recalc):
    """Create a fresh LibreOffice profile, optionally forcing recalculation on load.

    Built per job rather than copied from a prepared directory on the host: a host-side
    template would be an undeclared dependency whose absence shows up as wrong numbers
    rather than an error, which is the exact failure this whole exercise is about. The
    same reasoning keeps it out of the container image.
    """
    subprocess.run(runner.command(jobdir, [
        "--headless", "--norestore",
        "-env:UserInstallation=file://" + runner.seen(jobdir, "profile"),
        "--terminate_after_init"]), capture_output=True, text=True, timeout=300)
    if not recalc:
        return True
    xcu = os.path.join(jobdir, "profile", "user", "registrymodifications.xcu")
    if not os.path.exists(xcu):
        print("FAIL: soffice did not create a profile at %s" % xcu)
        return False
    with open(xcu, encoding="utf-8") as fh:
        body = fh.read()
    if "</oor:items>" not in body:
        print("FAIL: unexpected profile format, no </oor:items> in %s" % xcu)
        return False
    with open(xcu, "w", encoding="utf-8") as fh:
        fh.write(body.replace("</oor:items>", RECALC_ALWAYS_XCU + "</oor:items>"))
    return True


def convert(runner, jobdir, name):
    """Run the conversion and return the value soffice wrote into A3 (or None)."""
    proc = subprocess.run(runner.command(jobdir, [
        "--headless", "--norestore",
        "-env:UserInstallation=file://" + runner.seen(jobdir, "profile"),
        "--convert-to", "xlsx",
        "--outdir", runner.seen(jobdir, "out"),
        runner.seen(jobdir, name)]), capture_output=True, text=True, timeout=300)
    if proc.stderr.strip():
        print("  stderr:", proc.stderr.strip()[:300])
    # soffice is known to exit 0 without producing anything, so trust the file, not the
    # exit code.
    out = os.path.join(jobdir, "out", name)
    if not os.path.exists(out):
        print("  no output file produced (exit=%d)" % proc.returncode)
        return None
    _, value = read_a3(out)
    return value


def grant_jobdir(runner, jobdir):
    """Make the job directory writable by the identity that will run soffice.

    Only matters when the two differ, which happens when docker has to be driven as root
    while the container itself runs as an ordinary uid. Then the directory belongs to root
    and the container cannot create its profile in it - soffice reports "User installation
    [failed]", which reads like a broken image rather than a directory permission.
    (Found running the container gate with a non-root uid.)
    """
    if (runner.uid, runner.gid) == (os.getuid(), os.getgid()):
        return
    for root, dirs, files in os.walk(jobdir):
        for name in [root] + [os.path.join(root, n) for n in dirs + files]:
            try:
                os.chown(name, runner.uid, runner.gid)
            except PermissionError:
                # Not root, so ownership cannot be handed over; widen the mode instead.
                # Acceptable for a throwaway directory holding a synthetic fixture.
                os.chmod(name, 0o777 if os.path.isdir(name) else 0o666)


def run_case(runner, work, label, recalc):
    jobdir = os.path.join(work, label)
    os.makedirs(os.path.join(jobdir, "out"))
    build_fixture(os.path.join(jobdir, "gate.xlsx"))
    grant_jobdir(runner, jobdir)
    if not prepare_profile(runner, jobdir, recalc):
        return None
    value = convert(runner, jobdir, "gate.xlsx")
    print("  A3 after conversion: %r" % value)
    return value


def build_runner(argv):
    """Return the runner to test, or None with a message already printed."""
    if len(argv) > 1 and argv[1] == "--docker":
        if len(argv) < 3:
            print("FAIL: --docker needs an image name")
            return None
        if not shutil.which("docker"):
            print("FAIL: docker is not available on this machine")
            return None
        return DockerRunner(argv[2])
    soffice = argv[1] if len(argv) > 1 else shutil.which("soffice")
    if not soffice:
        print("FAIL: soffice not found on PATH (pass the path as an argument)")
        return None
    return HostRunner(soffice)


def main():
    runner = build_runner(sys.argv)
    if runner is None:
        return 2
    print("engine: %s\n" % runner.label)

    work = tempfile.mkdtemp(prefix="recalc-gate-")
    os.chmod(work, 0o755)  # the container runs as this uid but docker may traverse as root
    try:
        probe = os.path.join(work, "fixture.xlsx")
        build_fixture(probe)
        f, v = read_a3(probe)
        print("fixture A3: formula=%r cached=%r   (the cache is deliberately wrong)\n"
              % (f, v))
        if v != "999":
            print("FAIL: fixture did not build as intended")
            return 2

        # Both cases are run every time, because the interesting claim is not "it can
        # recalculate" but "it only recalculates when told to". Reporting the green case
        # alone would hide how narrow the working configuration is.
        print("case 1: plain --convert-to (default profile)")
        plain = run_case(runner, work, "plain", recalc=False)
        print("\ncase 2: --convert-to with recalc-on-load forced in the profile")
        forced = run_case(runner, work, "forced", recalc=True)

        print("\n" + "-" * 72)
        if forced == "5" and plain == "999":
            print("GATE GREEN, conditionally: the engine recalculates ONLY with "
                  "recalc-on-load forced in the profile.")
            print("Plain --convert-to returned the stale 999 - numbers that look right and "
                  "are not. Any caller that forgets the profile setting gets that silently.")
            return 0
        if forced == "5" and plain == "5":
            print("GATE GREEN: the engine recalculated in both cases. Check whether this "
                  "host carries a global profile setting - the result would then not "
                  "reproduce elsewhere.")
            return 0
        if forced == "999":
            print("GATE RED: still 999 even with recalc-on-load forced. This control path "
                  "does not recalculate; a different one (UNO/macro calculateAll) would "
                  "have to be proven before any code is written.")
            return 1
        print("GATE INCONCLUSIVE: plain=%r forced=%r (expected 999 and 5)." % (plain, forced))
        return 2
    finally:
        shutil.rmtree(work, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
