#!/usr/bin/env python3
"""Recalculate an xlsx through headless LibreOffice, driven by the office MCP.

Reads a JSON request on stdin and writes a JSON result on stdout:

    {"file": "/abs/path.xlsx", "timeout": 180}      # timeout optional, seconds

    -> {"ok": true, "file": "/abs/path.xlsx"}
    -> {"ok": false, "error": "one clean message"}

Why this exists and why it is shaped this way (verified on two LibreOffice versions,
7.3.7.2 and 26.2.4.2): LibreOffice does NOT recalculate xlsx formulas on load by default,
and `--convert-to` does not override that - a plain conversion hands back the STALE cached
value. A caller that forgets the setting gets numbers that look right and are wrong. The
only reliable control is a per-job user profile with recalc-on-load forced, so that is what
this helper builds, per job, from scratch. See office_recalc_gate.py for the regression that
proves plain=999 / forced=5.

The engine command is assembled in exactly ONE place - the Runner (HostRunner today). The
motor might later move into a container; when it does, only the command prefix and the
path translation change (DockerRunner), while the per-job profile, the xcu fix, the
conversion flags and the output-file check stay byte-for-byte identical. Every soffice call
goes through runner.command(); every path soffice is handed goes through runner.seen().

The actual process execution (timeout, kill-on-hang, trust-the-file-not-the-exit-code,
single clean error, temp cleanup) lives in office_runner, so it is tested once and reused.
python3 stdlib only; no openpyxl, no network.
"""

import json
import os
import grp
import pwd
import shlex
import shutil
import subprocess
import sys
import traceback
import zipfile
from office_paths import PathNotAllowed, check_path

from office_runner import ExternalToolError, job_workspace, run_producing_file

# Recalculation on load is OFF by default for xlsx and --convert-to does not change it. The
# setting lives in the user profile: 0 = Always, 1 = Never, 2 = Prompt. Forcing Always is the
# whole point of building a throwaway profile per job.
RECALC_ALWAYS_XCU = (
    '<item oor:path="/org.openoffice.Office.Calc/Formula/Load">'
    '<prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>'
    '<item oor:path="/org.openoffice.Office.Calc/Formula/Load">'
    '<prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>'
)


class HostRunner:
    """soffice installed on this machine.

    Command assembly and path translation live here so a container engine is a drop-in
    replacement (see DockerRunner): the flags, profile, conversion and output-check stay the
    same, only the command prefix and how paths are seen differ. `seen` is the identity here
    because what soffice opens and what Python opens are the same path.
    """

    def __init__(self, soffice):
        self.soffice = soffice
        self.label = soffice

    def prepare(self, jobdir):
        pass

    def seen(self, jobdir, rel):
        return os.path.join(jobdir, rel)

    def command(self, jobdir, args):
        return [self.soffice] + args

    def describe(self):
        return "host soffice %s" % self.soffice


class DockerRunner:
    """soffice inside a short-lived container - the likely home if the motor moves into a
    container per the fleet's containerise-larger-apps policy.

    The job directory is the only thing mounted, so every path soffice is given is rewritten
    to its location inside the container. --network=none because a converter never needs the
    network and documents can carry external references; -u so output belongs to the invoking
    user; HOME inside the mount because the -u uid has no /etc/passwd entry and LibreOffice
    otherwise has nowhere to put its profile. Structurally identical to the gate's runner.
    """

    def __init__(self, image):
        self.image = image
        self.label = "docker:" + image

    def prepare(self, jobdir):
        # The container runs as this uid but docker may traverse the mount as root.
        os.chmod(jobdir, 0o755)

    def seen(self, jobdir, rel):
        return "/work/" + rel

    def command(self, jobdir, args):
        return _docker_argv(["docker", "run", "--rm", "--network=none",
                             "-v", jobdir + ":/work",
                             "-u", "%d:%d" % (os.getuid(), os.getgid()),
                             "-e", "HOME=/work",
                             self.image, "soffice"] + args)

    def describe(self):
        return "container %s" % self.image


# The container is the production engine (the owner's decision): the host apt-LibreOffice is to be
# removed once acceptance is met, so the DEFAULT must be the container when it is available,
# with a VISIBLE fall back to host soffice. A silent fallback is exactly how a "worked
# yesterday, broken today" incident hides its cause when the host binary disappears. The image
# name is a parameter (OFFICE_IMAGE), never hardcoded at a call site.
DEFAULT_IMAGE = "euro-office-lo:1"


def resolve_soffice():
    """The soffice binary: an explicit override wins, else PATH. Never hardcoded."""
    return os.environ.get("OFFICE_SOFFICE_BIN") or shutil.which("soffice")


def _docker_group_wrap_needed():
    """True when this process lacks the docker group although the user is a member of it.

    Group membership is fixed when a process starts, so a user added to `docker` AFTER a
    long-running parent started (here: the fleet's tmux server) gets "permission denied" on
    the socket while `getent group docker` happily lists them. `sg` runs a single command
    with a group the user already holds, so the capability is reachable without restarting
    anything - and restarting that parent would kill every agent session, which is the most
    expensive move available.

    Returns False once the group IS present (the state after any such restart), so the
    wrapper disappears on its own and never has to be removed.
    """
    try:
        entry = grp.getgrnam("docker")
    except KeyError:
        return False
    if entry.gr_gid in os.getgroups():
        return False
    try:
        user = pwd.getpwuid(os.getuid()).pw_name
    except KeyError:
        return False
    return user in entry.gr_mem and shutil.which("sg") is not None


def _docker_argv(argv):
    """How a docker command must actually be spawned from this process.

    Every argument is quoted because `sg -c` takes ONE shell string: an unquoted path with a
    space would silently become two arguments, which for a `-v host:/work` mount means the
    job directory is not where soffice looks.
    """
    if not _docker_group_wrap_needed():
        return argv
    return ["sg", "docker", "-c", " ".join(shlex.quote(a) for a in argv)]


def _docker_on_path():
    """docker present, decided by a PATH lookup - does NOT invoke docker."""
    return shutil.which("docker") is not None


def _docker_image_present(image):
    """Whether the container image is available locally, via a read-only `docker image
    inspect`. Isolated in one function so tests stub it and the recalc code never invokes
    docker itself - the MCP process runs the assembled command."""
    try:
        proc = subprocess.run(
            _docker_argv(["docker", "image", "inspect", image]),
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=15,
        )
        return proc.returncode == 0
    except (OSError, subprocess.SubprocessError):
        return False


def build_runner(report=None):
    """Select the engine. Precedence, most explicit first:

      1. OFFICE_ENGINE=docker:<image>  -> that container (docker must be on PATH).
      2. OFFICE_ENGINE=host            -> host soffice.
      3. OFFICE_SOFFICE_BIN set        -> host soffice at that binary (dev/ops pin; no docker
                                          probe, so a dev box never triggers a docker call).
      4. nothing set -> auto: the CONTAINER when docker AND the image are both available,
                        otherwise a VISIBLE fall back to host soffice.

    Image name comes from OFFICE_IMAGE (default DEFAULT_IMAGE), never hardcoded at the caller.
    `report`, if given, is called with a one-line note when auto mode falls back to host.
    Swapping engines never touches the recalc logic below - only which command prefix is built.
    """
    engine = os.environ.get("OFFICE_ENGINE", "").strip()
    image = os.environ.get("OFFICE_IMAGE") or DEFAULT_IMAGE

    if engine.startswith("docker:"):
        img = engine[len("docker:"):]
        if not img:
            raise ExternalToolError("OFFICE_ENGINE=docker: needs an image name")
        if not _docker_on_path():
            raise ExternalToolError("OFFICE_ENGINE requested docker but docker is not on PATH")
        return DockerRunner(img)

    if engine == "host":
        soffice = resolve_soffice()
        if not soffice:
            raise ExternalToolError(
                "soffice not found: set OFFICE_SOFFICE_BIN or install LibreOffice"
            )
        return HostRunner(soffice)

    if engine:
        raise ExternalToolError(
            "unknown OFFICE_ENGINE=%r (use 'docker:<image>' or 'host')" % engine
        )

    override = os.environ.get("OFFICE_SOFFICE_BIN")
    if override:
        return HostRunner(override)

    # Auto: prefer the container, fall back to host - but say so.
    if _docker_on_path() and _docker_image_present(image):
        return DockerRunner(image)

    soffice = shutil.which("soffice")
    if not soffice:
        raise ExternalToolError(
            "no recalc engine available: container image %r not usable (docker on PATH: %s) "
            "and soffice not found on PATH (set OFFICE_SOFFICE_BIN, OFFICE_ENGINE=host, or make "
            "the container image available)" % (image, _docker_on_path())
        )
    if report is not None:
        report("engine: container %r unavailable, falling back to host soffice %s"
               % (image, soffice))
    return HostRunner(soffice)


def prepare_profile(runner, jobdir, timeout, force_recalc=True):
    """Create a fresh LibreOffice profile in jobdir, optionally forcing recalc-on-load.

    Built per job rather than copied from a host-side template on purpose: a missing template
    would surface as wrong numbers, not an error, which is the exact failure this whole tool
    guards against. Per-job also means parallel recalcs never share a profile.

    force_recalc=False is for callers that must render a workbook EXACTLY as stored - PDF
    export of a single extracted sheet, above all. That sheet routinely references sheets the
    extraction removed (203 of 2219 formula cells in a real workbook), and a recalculation
    would turn every one of them into #REF! on the page. Stored values are also what the PDF
    tool already promises: recalculating is a separate step, asked for on purpose.
    """
    profile_seen = runner.seen(jobdir, "profile")
    xcu_local = os.path.join(jobdir, "profile", "user", "registrymodifications.xcu")

    # The profile init IS an "external tool run": its output file is the xcu. Routing it
    # through run_producing_file gives it the same timeout / clean-error treatment as the
    # conversion, and asserts the profile actually materialised.
    run_producing_file(
        runner.command(jobdir, [
            "--headless", "--norestore",
            "-env:UserInstallation=file://" + profile_seen,
            "--terminate_after_init"]),
        xcu_local, timeout,
    )

    if not force_recalc:
        return

    with open(xcu_local, encoding="utf-8") as fh:
        body = fh.read()
    if "</oor:items>" not in body:
        raise ExternalToolError(
            "unexpected LibreOffice profile format (no </oor:items> in registrymodifications.xcu)"
        )
    with open(xcu_local, "w", encoding="utf-8") as fh:
        fh.write(body.replace("</oor:items>", RECALC_ALWAYS_XCU + "</oor:items>"))


def recalc_file(path, timeout, report=None):
    """Recalculate `path` in place. Raises ExternalToolError with one clean message on failure.

    Returns a human description of the engine that ran (e.g. "container euro-office-lo:1" or
    "host soffice /usr/bin/soffice") so the caller can report WHICH engine was used - without
    that, nobody can tell which motor a deployment is on without inspecting logs.

    The original file is only overwritten once a recalculated output exists: on any failure
    (timeout, non-zero exit, or exit-0-without-output) the original is left exactly as it was.
    """
    runner = build_runner(report=report)

    if not os.path.exists(path):
        raise ExternalToolError("file not found: " + path)
    try:
        with zipfile.ZipFile(path):
            pass
    except (zipfile.BadZipFile, OSError) as exc:
        raise ExternalToolError("not a valid xlsx (%s): %s" % (exc, path))

    name = os.path.basename(path)
    with job_workspace(prefix="office-recalc-") as jobdir:
        runner.prepare(jobdir)
        src_local = os.path.join(jobdir, name)
        shutil.copy(path, src_local)
        out_local = os.path.join(jobdir, "out", name)
        os.makedirs(os.path.join(jobdir, "out"))

        prepare_profile(runner, jobdir, timeout)

        run_producing_file(
            runner.command(jobdir, [
                "--headless", "--norestore",
                "-env:UserInstallation=file://" + runner.seen(jobdir, "profile"),
                "--convert-to", "xlsx",
                "--outdir", runner.seen(jobdir, "out"),
                runner.seen(jobdir, name)]),
            out_local, timeout,
        )
        # Move the result out of the workspace before it is torn down, overwriting the original
        # only now that we know a real output exists.
        shutil.move(out_local, path)
    return runner.describe()


def main():
    try:
        req = json.load(sys.stdin)
    except Exception as exc:  # noqa: BLE001 - report any parse failure as a clean result
        json.dump({"ok": False, "error": "invalid request JSON: %s" % exc}, sys.stdout)
        return

    path = req.get("file")
    if not path:
        json.dump({"ok": False, "error": "file is required"}, sys.stdout)
        return
    timeout = req.get("timeout", 180)

    try:
        path = check_path(path)
    except PathNotAllowed as exc:
        json.dump({"ok": False, "error": str(exc)}, sys.stdout)
        return

    try:
        engine = recalc_file(path, timeout, report=lambda m: sys.stderr.write(m + "\n"))
    except ExternalToolError as exc:
        json.dump({"ok": False, "error": str(exc)}, sys.stdout)
        return
    except Exception as exc:  # noqa: BLE001
        # A bug in the helper (a NameError, a typo, anything that is NOT a domain error) must
        # still leave the caller with clean JSON, not a raw traceback surfaced through the MCP.
        # One line in the response; the full trace goes to stderr for debugging only.
        traceback.print_exc(file=sys.stderr)
        json.dump({"ok": False, "error": "internal error: %s: %s"
                   % (type(exc).__name__, exc)}, sys.stdout)
        return
    json.dump({"ok": True, "file": path, "engine": engine}, sys.stdout)


if __name__ == "__main__":
    main()
