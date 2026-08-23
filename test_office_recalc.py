#!/usr/bin/env python3
"""Tests for office_recalc, exercised against fake_soffice_recalc (and real soffice if present).

The single most important test in this tool is test_stale_cache_recomputed: a formula cell
carrying a deliberately WRONG cached value (999) must come back recalculated (5). That is the
exact failure - numbers that look right and are stale - the whole recalc path exists to
prevent. The fake models the real trap: without recalc-on-load forced in the profile it hands
999 straight back, so a green test means the forcing actually happened, not that the fake was
told to pass.

The error paths proven at the runner layer (exit-0-without-output, non-zero exit, hang) are
re-checked here end-to-end through the recalc helper, so they stay covered once a real engine
sits behind the same interface.

Run:  python3 -m unittest test_office_recalc
"""

import contextlib
import glob
import io
import json
import unittest.mock
import os
import subprocess
import sys
import tempfile
import time
import unittest

import office_recalc
from office_recalc import DockerRunner, ExternalToolError, HostRunner, recalc_file
import grp
from office_recalc_gate import build_fixture, read_a3
import fake_soffice_recalc

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


def new_fixture():
    fd, path = tempfile.mkstemp(suffix=".xlsx", prefix="recalc-fixture-")
    os.close(fd)
    build_fixture(path)
    return path


def recalc_temp_dirs():
    return set(glob.glob(os.path.join(tempfile.gettempdir(), "office-recalc-*")))


class RecalcSemanticsTests(unittest.TestCase):
    def setUp(self):
        self.fixture = new_fixture()
        self.addCleanup(lambda: os.path.exists(self.fixture) and os.remove(self.fixture))

    def test_stale_cache_recomputed(self):
        """THE test: cached 999 comes back as the recalculated 5."""
        self.assertEqual(read_a3(self.fixture), ("A1+A2", "999"))  # the lie is in place
        with set_env(OFFICE_SOFFICE_BIN=FAKE, FAKE_SOFFICE_BEHAVIOR=None):
            recalc_file(self.fixture, timeout=30)
        self.assertEqual(read_a3(self.fixture), ("A1+A2", "5"))

    def test_default_profile_keeps_stale_value(self):
        """Without recalc-on-load forced, the fake returns 999 - proving the trap is real."""
        before = recalc_temp_dirs()
        with tempfile.TemporaryDirectory(prefix="norecalc-") as jobdir:
            profile = os.path.join(jobdir, "profile", "user")
            os.makedirs(profile)
            # A profile WITHOUT the forced OOXMLRecalcMode item.
            with open(os.path.join(profile, "registrymodifications.xcu"), "w") as fh:
                fh.write('<oor:items xmlns:oor="http://openoffice.org/2001/registry">'
                         '</oor:items>')
            outdir = os.path.join(jobdir, "out")
            os.makedirs(outdir)
            src = os.path.join(jobdir, "gate.xlsx")
            build_fixture(src)
            proc = subprocess.run(
                [sys.executable, FAKE, "--headless", "--norestore",
                 "-env:UserInstallation=file://" + os.path.join(jobdir, "profile"),
                 "--convert-to", "xlsx", "--outdir", outdir, src],
                capture_output=True, text=True, timeout=30,
            )
            self.assertEqual(proc.returncode, 0, proc.stderr)
            self.assertEqual(read_a3(os.path.join(outdir, "gate.xlsx")), ("A1+A2", "999"))
        self.assertEqual(recalc_temp_dirs() - before, set())


class RecalcErrorPathTests(unittest.TestCase):
    def setUp(self):
        self.fixture = new_fixture()
        self.addCleanup(lambda: os.path.exists(self.fixture) and os.remove(self.fixture))

    def _assert_untouched(self):
        # Original workbook must survive any failure with its stale value intact.
        self.assertEqual(read_a3(self.fixture), ("A1+A2", "999"))

    def test_exit_zero_without_output_is_error(self):
        before = recalc_temp_dirs()
        with set_env(OFFICE_SOFFICE_BIN=FAKE, FAKE_SOFFICE_BEHAVIOR="no-output"):
            with self.assertRaises(ExternalToolError) as ctx:
                recalc_file(self.fixture, timeout=30)
        self.assertIn("no output", str(ctx.exception))
        self._assert_untouched()
        self.assertEqual(recalc_temp_dirs() - before, set(), "leaked temp dir")

    def test_nonzero_exit_is_clean_single_line_error(self):
        before = recalc_temp_dirs()
        with set_env(OFFICE_SOFFICE_BIN=FAKE, FAKE_SOFFICE_BEHAVIOR="error"):
            with self.assertRaises(ExternalToolError) as ctx:
                recalc_file(self.fixture, timeout=30)
        msg = str(ctx.exception)
        self.assertIn("exit 1", msg)
        self.assertIn("could not be loaded", msg)
        self.assertEqual(msg.count("\n"), 0)  # one clean line, not a stderr dump
        self._assert_untouched()
        self.assertEqual(recalc_temp_dirs() - before, set(), "leaked temp dir")

    def test_hang_times_out_and_leaves_original_untouched(self):
        before = recalc_temp_dirs()
        start = time.monotonic()
        with set_env(OFFICE_SOFFICE_BIN=FAKE, FAKE_SOFFICE_BEHAVIOR="hang"):
            with self.assertRaises(ExternalToolError) as ctx:
                recalc_file(self.fixture, timeout=2)
        self.assertIn("timed out", str(ctx.exception))
        self.assertLess(time.monotonic() - start, 30)
        self._assert_untouched()
        self.assertEqual(recalc_temp_dirs() - before, set(), "leaked temp dir")

    def test_missing_binary_path_reported_cleanly(self):
        """OFFICE_SOFFICE_BIN points at a path that is not a runnable binary."""
        with set_env(OFFICE_SOFFICE_BIN=os.path.join(HERE, "no-such-soffice-xyz"),
                     FAKE_SOFFICE_BEHAVIOR=None):
            with self.assertRaises(ExternalToolError) as ctx:
                recalc_file(self.fixture, timeout=10)
        self.assertIn("not found", str(ctx.exception))
        self._assert_untouched()

    def test_no_soffice_anywhere_is_one_clean_message(self):
        """soffice absent from PATH and unset override: one clean line, no temp dir, no trace.

        This is the acceptance path for the fleet host possibly losing LibreOffice - the code
        must not assume the binary is present, and its absence must surface as a plain error.
        """
        before = recalc_temp_dirs()
        with set_env(OFFICE_SOFFICE_BIN=None, FAKE_SOFFICE_BEHAVIOR=None):
            with unittest.mock.patch("office_recalc.shutil.which", return_value=None):
                with self.assertRaises(ExternalToolError) as ctx:
                    recalc_file(self.fixture, timeout=10)
        msg = str(ctx.exception)
        self.assertIn("soffice not found", msg)
        self.assertIn("OFFICE_SOFFICE_BIN", msg)  # tells the caller how to fix it
        self.assertEqual(msg.count("\n"), 0)  # one clean line, not a stack trace
        self._assert_untouched()
        self.assertEqual(recalc_temp_dirs() - before, set(), "leaked temp dir")

    def test_no_soffice_json_contract_is_ok_false(self):
        """Through the stdin/stdout contract office-mcp.js uses: ok:false + clean error."""
        env = {k: v for k, v in os.environ.items() if k not in ("OFFICE_SOFFICE_BIN",)}
        env["PATH"] = ""  # nothing resolvable on PATH -> shutil.which returns None
        env.pop("FAKE_SOFFICE_BEHAVIOR", None)
        proc = subprocess.run(
            [sys.executable, os.path.join(HERE, "office_recalc.py")],
            input=json.dumps({"file": self.fixture}),
            capture_output=True, text=True, timeout=30, env=env,
        )
        self.assertEqual(proc.returncode, 0, proc.stderr)
        self.assertEqual(proc.stderr, "", "helper leaked a stack trace to stderr")
        result = json.loads(proc.stdout)
        self.assertFalse(result["ok"])
        self.assertIn("soffice not found", result["error"])
        self._assert_untouched()

    def test_nonexistent_file_reported_cleanly(self):
        missing = os.path.join(tempfile.gettempdir(), "office-recalc-absent-xyz.xlsx")
        self.assertFalse(os.path.exists(missing))
        before = recalc_temp_dirs()
        with set_env(OFFICE_SOFFICE_BIN=FAKE, FAKE_SOFFICE_BEHAVIOR=None):
            with self.assertRaises(ExternalToolError) as ctx:
                recalc_file(missing, timeout=10)
        msg = str(ctx.exception)
        self.assertIn("file not found", msg)
        self.assertEqual(msg.count("\n"), 0)  # one clean line
        self.assertEqual(recalc_temp_dirs() - before, set(), "leaked temp dir")

    def test_non_xlsx_input_reported_cleanly(self):
        fd, junk = tempfile.mkstemp(suffix=".xlsx", prefix="notzip-")
        os.write(fd, b"this is plainly not a zip archive")
        os.close(fd)
        self.addCleanup(lambda: os.path.exists(junk) and os.remove(junk))
        before = recalc_temp_dirs()
        with set_env(OFFICE_SOFFICE_BIN=FAKE, FAKE_SOFFICE_BEHAVIOR=None):
            with self.assertRaises(ExternalToolError) as ctx:
                recalc_file(junk, timeout=10)
        self.assertIn("not a valid xlsx", str(ctx.exception))
        self.assertEqual(recalc_temp_dirs() - before, set(), "leaked temp dir")


class JsonContractTests(unittest.TestCase):
    """The JSON stdin/stdout contract office-mcp.js drives the helper through."""

    def _run(self, request, behavior=None):
        env = dict(os.environ, OFFICE_SOFFICE_BIN=FAKE)
        if behavior is not None:
            env["FAKE_SOFFICE_BEHAVIOR"] = behavior
        else:
            env.pop("FAKE_SOFFICE_BEHAVIOR", None)
        proc = subprocess.run(
            [sys.executable, os.path.join(HERE, "office_recalc.py")],
            input=json.dumps(request), capture_output=True, text=True, timeout=60, env=env,
        )
        self.assertEqual(proc.returncode, 0, proc.stderr)
        return json.loads(proc.stdout)

    def test_success_result(self):
        fixture = new_fixture()
        self.addCleanup(lambda: os.path.exists(fixture) and os.remove(fixture))
        result = self._run({"file": fixture})
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["file"], fixture)
        self.assertIn("host soffice", result["engine"])  # the result names the engine that ran
        self.assertEqual(read_a3(fixture), ("A1+A2", "5"))

    def test_error_result_is_ok_false(self):
        fixture = new_fixture()
        self.addCleanup(lambda: os.path.exists(fixture) and os.remove(fixture))
        result = self._run({"file": fixture}, behavior="error")
        self.assertFalse(result["ok"])
        self.assertIn("exit 1", result["error"])

    def test_missing_file_field(self):
        result = self._run({})
        self.assertFalse(result["ok"])
        self.assertIn("file is required", result["error"])

    def test_unexpected_helper_bug_still_returns_clean_json(self):
        # A non-ExternalToolError (NameError, typo, any bug) must NOT leak a traceback through
        # the contract - the exact class of failure the container-default regression exposed.
        stdin = io.StringIO(json.dumps({"file": os.path.join(tempfile.gettempdir(), "whatever.xlsx")}))
        out, err = io.StringIO(), io.StringIO()
        with unittest.mock.patch("office_recalc.recalc_file",
                                 side_effect=RuntimeError("boom in helper")), \
                unittest.mock.patch("sys.stdin", stdin), \
                contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            office_recalc.main()
        text = out.getvalue()
        self.assertNotIn("Traceback", text)          # no stack trace on the wire
        self.assertEqual(text.count("\n"), 0)         # one clean JSON line
        result = json.loads(text)                     # still valid JSON
        self.assertFalse(result["ok"])
        self.assertIn("boom in helper", result["error"])
        self.assertIn("Traceback", err.getvalue())    # full detail kept on stderr for debugging


class FakeModelUnitTests(unittest.TestCase):
    """Guard the fake's own teeth: it must recompute A1+A2, not hardcode 5."""

    def test_recompute_sheet_sums_cells(self):
        sheet = ('<worksheet><sheetData>'
                 '<row r="1"><c r="A1"><v>7</v></c></row>'
                 '<row r="2"><c r="A2"><v>8</v></c></row>'
                 '<row r="3"><c r="A3"><f>A1+A2</f><v>999</v></c></row>'
                 '</sheetData></worksheet>')
        out = fake_soffice_recalc.recompute_sheet(sheet)
        self.assertIn("<v>15</v>", out)
        self.assertNotIn("<v>999</v>", out)


class RunnerSeamTests(unittest.TestCase):
    """The command is assembled in one place; swapping engines is that single seam.

    These pin the contract both runners honour: identical flags/paths flow through, only the
    prefix and the path translation differ. If the recalc logic ever inlines a `soffice ...`
    again instead of going through runner.command(), the container swap stops being one line.
    """

    def test_host_runner_is_identity_paths_and_bare_prefix(self):
        r = HostRunner("/usr/bin/soffice")
        self.assertEqual(r.seen("/job", "out"), os.path.join("/job", "out"))
        self.assertEqual(
            r.command("/job", ["--headless", "--convert-to", "xlsx"]),
            ["/usr/bin/soffice", "--headless", "--convert-to", "xlsx"],
        )

    def test_docker_runner_rewrites_paths_and_wraps_prefix(self):
        r = DockerRunner("office-lo:latest")
        # Paths soffice sees are inside the mount, not on the host.
        self.assertEqual(r.seen("/job", "out"), "/work/out")
        self.assertEqual(r.seen("/job", "gate.xlsx"), "/work/gate.xlsx")
        # The plain (unwrapped) invocation is what this test locks. Whether THIS host needs
        # the sg group-wrapper is a property of the host, not of the spec, so it is pinned -
        # otherwise the same code would pass or fail depending on who runs the suite.
        with unittest.mock.patch.object(office_recalc, "_docker_group_wrap_needed", return_value=False):
            cmd = r.command("/job", ["--headless", "--convert-to", "xlsx"])
        # The engine-specific args ride at the tail, byte-for-byte the same as the host case.
        self.assertEqual(cmd[-3:], ["--headless", "--convert-to", "xlsx"])
        self.assertIn("-v", cmd)
        self.assertIn("/job:/work", cmd)
        self.assertEqual(cmd[:3], ["docker", "run", "--rm"])
        self.assertIn("soffice", cmd)

    def test_docker_convert_command_matches_agreed_spec(self):
        """Lock the exact container invocation the recalc step emits, so wiring the image in
        is verified before docker is even available here. Mirrors recalc_file's convert args."""
        r = DockerRunner("office-lo:latest")
        jobdir, name = "/tmp/job", "gate.xlsx"
        convert_args = [
            "--headless", "--norestore",
            "-env:UserInstallation=file://" + r.seen(jobdir, "profile"),
            "--convert-to", "xlsx",
            "--outdir", r.seen(jobdir, "out"),
            r.seen(jobdir, name),
        ]
        with unittest.mock.patch.object(office_recalc, "_docker_group_wrap_needed", return_value=False):
            emitted = r.command(jobdir, convert_args)
        self.assertEqual(
            emitted,
            ["docker", "run", "--rm", "--network=none",
             "-v", "/tmp/job:/work",
             "-u", "%d:%d" % (os.getuid(), os.getgid()),
             "-e", "HOME=/work",
             "office-lo:latest", "soffice",
             "--headless", "--norestore",
             "-env:UserInstallation=file:///work/profile",
             "--convert-to", "xlsx",
             "--outdir", "/work/out",
             "/work/gate.xlsx"],
        )

    def test_engine_selector_defaults_to_host(self):
        with set_env(OFFICE_ENGINE=None, OFFICE_SOFFICE_BIN=FAKE):
            self.assertIsInstance(office_recalc.build_runner(), HostRunner)

    def test_engine_selector_rejects_empty_docker_image(self):
        with set_env(OFFICE_ENGINE="docker:", OFFICE_SOFFICE_BIN=FAKE):
            with self.assertRaises(ExternalToolError) as ctx:
                office_recalc.build_runner()
        self.assertIn("image name", str(ctx.exception))


class EngineSelectionTests(unittest.TestCase):
    """Auto-selection: the container is the default engine; host is the VISIBLE fallback.

    The docker probes are stubbed in every case, so these never invoke real docker even on a
    box that has it - the code only assembles the command, the MCP runs it.
    """

    def test_auto_prefers_container_when_docker_and_image_present(self):
        with set_env(OFFICE_ENGINE=None, OFFICE_SOFFICE_BIN=None, OFFICE_IMAGE=None):
            with unittest.mock.patch("office_recalc._docker_on_path", return_value=True), \
                 unittest.mock.patch("office_recalc._docker_image_present", return_value=True):
                r = office_recalc.build_runner()
        self.assertIsInstance(r, DockerRunner)
        self.assertEqual(r.image, office_recalc.DEFAULT_IMAGE)
        self.assertEqual(r.image, "euro-office-lo:1")  # sensible default, not call-site literal
        self.assertIn("container", r.describe())

    def test_office_image_env_overrides_default(self):
        with set_env(OFFICE_ENGINE=None, OFFICE_SOFFICE_BIN=None, OFFICE_IMAGE="acme/lo:9"):
            with unittest.mock.patch("office_recalc._docker_on_path", return_value=True), \
                 unittest.mock.patch("office_recalc._docker_image_present", return_value=True):
                r = office_recalc.build_runner()
        self.assertIsInstance(r, DockerRunner)
        self.assertEqual(r.image, "acme/lo:9")

    def test_auto_falls_back_to_host_visibly_when_no_docker(self):
        notes = []
        with set_env(OFFICE_ENGINE=None, OFFICE_SOFFICE_BIN=None, OFFICE_IMAGE=None):
            with unittest.mock.patch("office_recalc._docker_on_path", return_value=False), \
                 unittest.mock.patch("office_recalc.shutil.which", return_value="/opt/soffice"):
                r = office_recalc.build_runner(report=notes.append)
        self.assertIsInstance(r, HostRunner)
        self.assertIn("host soffice", r.describe())
        self.assertTrue(notes and "falling back" in notes[0], "fallback must be visible, not silent")

    def test_auto_falls_back_to_host_when_image_absent(self):
        with set_env(OFFICE_ENGINE=None, OFFICE_SOFFICE_BIN=None, OFFICE_IMAGE=None):
            with unittest.mock.patch("office_recalc._docker_on_path", return_value=True), \
                 unittest.mock.patch("office_recalc._docker_image_present", return_value=False), \
                 unittest.mock.patch("office_recalc.shutil.which", return_value="/opt/soffice"):
                r = office_recalc.build_runner()
        self.assertIsInstance(r, HostRunner)

    def test_soffice_bin_pins_host_without_probing_docker(self):
        # An explicit binary means host; docker must not even be probed (no accidental call).
        with set_env(OFFICE_ENGINE=None, OFFICE_SOFFICE_BIN="/opt/soffice", OFFICE_IMAGE=None):
            with unittest.mock.patch(
                "office_recalc._docker_image_present",
                side_effect=AssertionError("docker must not be probed when OFFICE_SOFFICE_BIN is set"),
            ):
                r = office_recalc.build_runner()
        self.assertIsInstance(r, HostRunner)
        self.assertIn("/opt/soffice", r.describe())

    def test_explicit_host_wins_over_available_container(self):
        with set_env(OFFICE_ENGINE="host", OFFICE_SOFFICE_BIN="/opt/soffice", OFFICE_IMAGE=None):
            with unittest.mock.patch("office_recalc._docker_on_path", return_value=True), \
                 unittest.mock.patch("office_recalc._docker_image_present", return_value=True):
                r = office_recalc.build_runner()
        self.assertIsInstance(r, HostRunner)

    def test_explicit_docker_requires_docker_on_path(self):
        with set_env(OFFICE_ENGINE="docker:x/y:1", OFFICE_SOFFICE_BIN=None, OFFICE_IMAGE=None):
            with unittest.mock.patch("office_recalc._docker_on_path", return_value=False):
                with self.assertRaises(ExternalToolError) as ctx:
                    office_recalc.build_runner()
        self.assertIn("docker is not on PATH", str(ctx.exception))

    def test_docker_image_present_survives_errors_without_nameerror(self):
        # Regression guard: _docker_image_present references `subprocess` (run + SubprocessError).
        # A missing `import subprocess` here NameErrors on the first real auto-mode call - which
        # is exactly what shipped and broke every call. Runs the real function, docker mocked.
        with unittest.mock.patch("office_recalc.subprocess.run", side_effect=OSError("no docker")):
            self.assertFalse(office_recalc._docker_image_present("x/y:1"))
        with unittest.mock.patch("office_recalc.subprocess.run",
                                 return_value=unittest.mock.Mock(returncode=0)):
            self.assertTrue(office_recalc._docker_image_present("x/y:1"))

    def test_auto_selects_container_through_the_real_probe(self):
        # Exercises the REAL _docker_image_present (subprocess mocked, no real docker call), so
        # the auto path itself runs here, not a stub of it.
        with set_env(OFFICE_ENGINE=None, OFFICE_SOFFICE_BIN=None, OFFICE_IMAGE=None):
            with unittest.mock.patch("office_recalc._docker_on_path", return_value=True), \
                 unittest.mock.patch("office_recalc.subprocess.run",
                                     return_value=unittest.mock.Mock(returncode=0)) as run:
                r = office_recalc.build_runner()
        self.assertIsInstance(r, DockerRunner)
        run.assert_called_once()

    def test_no_engine_at_all_is_one_clean_message(self):
        with set_env(OFFICE_ENGINE=None, OFFICE_SOFFICE_BIN=None, OFFICE_IMAGE=None):
            with unittest.mock.patch("office_recalc._docker_on_path", return_value=False), \
                 unittest.mock.patch("office_recalc.shutil.which", return_value=None):
                with self.assertRaises(ExternalToolError) as ctx:
                    office_recalc.build_runner()
        msg = str(ctx.exception)
        self.assertIn("soffice not found", msg)
        self.assertEqual(msg.count("\n"), 0)


def _real_soffice():
    """A soffice on PATH that is NOT our fake, or None."""
    import shutil
    found = shutil.which("soffice")
    return found if found and os.path.realpath(found) != os.path.realpath(FAKE) else None


@unittest.skipUnless(_real_soffice(), "no real soffice on this host")
class RealSofficeRegressionTests(unittest.TestCase):
    """Keep the real 999->5 regression alive where LibreOffice is actually installed."""

    def test_gate_is_green_on_real_soffice(self):
        proc = subprocess.run(
            [sys.executable, os.path.join(HERE, "office_recalc_gate.py")],
            capture_output=True, text=True, timeout=600,
        )
        self.assertEqual(proc.returncode, 0,
                         "gate not green:\n%s\n%s" % (proc.stdout, proc.stderr))

    def test_recalc_file_recomputes_on_real_soffice(self):
        fixture = new_fixture()
        self.addCleanup(lambda: os.path.exists(fixture) and os.remove(fixture))
        # Pin the host engine so this never probes or invokes docker on a box that has it.
        with set_env(OFFICE_ENGINE="host", OFFICE_SOFFICE_BIN=None, FAKE_SOFFICE_BEHAVIOR=None):
            engine = recalc_file(fixture, timeout=300)
        self.assertEqual(read_a3(fixture), ("A1+A2", "5"))
        self.assertIn("host soffice", engine)


if __name__ == "__main__":
    unittest.main(verbosity=2)


class DockerGroupWrapTests(unittest.TestCase):
    """The docker group may be missing from THIS process even though the user has it.

    A group is fixed at process start, so a user added to `docker` after the parent started
    sees permission denied while `getent group docker` lists them. The runner wraps the call
    in `sg` for exactly that window - and must NOT wrap in any other case, because wrapping
    when it is unnecessary would drag a shell into the path for no reason.
    """

    GROUP = grp.struct_group(("docker", "x", 4242, ["user"]))

    def _patch(self, groups, members=("user",), sg="/usr/bin/sg"):
        entry = grp.struct_group(("docker", "x", 4242, list(members)))
        return contextlib.ExitStack(), entry, groups, sg

    def _decide(self, groups, members=("user",), sg="/usr/bin/sg"):
        entry = grp.struct_group(("docker", "x", 4242, list(members)))
        with unittest.mock.patch.object(office_recalc.grp, "getgrnam", return_value=entry), \
             unittest.mock.patch.object(office_recalc.os, "getgroups", return_value=groups), \
             unittest.mock.patch.object(office_recalc.pwd, "getpwuid",
                                        return_value=type("P", (), {"pw_name": "user"})), \
             unittest.mock.patch.object(office_recalc.shutil, "which",
                                        side_effect=lambda name: sg if name == "sg" else "/usr/bin/" + name):
            return office_recalc._docker_group_wrap_needed()

    def test_no_wrap_when_the_group_is_already_present(self):
        # The state after a restart: nothing to work around.
        self.assertFalse(self._decide(groups=[4242, 1000]))

    def test_wrap_when_member_but_group_missing_from_the_process(self):
        self.assertTrue(self._decide(groups=[1000]))

    def test_no_wrap_when_the_user_is_not_a_member(self):
        # Then sg would fail anyway; the honest outcome is the plain permission error.
        self.assertFalse(self._decide(groups=[1000], members=("someone-else",)))

    def test_no_wrap_without_sg_on_path(self):
        self.assertFalse(self._decide(groups=[1000], sg=None))

    def test_no_wrap_when_there_is_no_docker_group_at_all(self):
        with unittest.mock.patch.object(office_recalc.grp, "getgrnam", side_effect=KeyError):
            self.assertFalse(office_recalc._docker_group_wrap_needed())

    def test_argv_is_untouched_when_no_wrap_is_needed(self):
        with unittest.mock.patch.object(office_recalc, "_docker_group_wrap_needed", return_value=False):
            self.assertEqual(office_recalc._docker_argv(["docker", "ps"]), ["docker", "ps"])

    def test_wrapped_argv_keeps_a_path_with_spaces_in_ONE_argument(self):
        # sg -c takes a single shell string: an unquoted mount path with a space would split
        # into two arguments and the job directory would not be where soffice looks.
        with unittest.mock.patch.object(office_recalc, "_docker_group_wrap_needed", return_value=True):
            argv = office_recalc._docker_argv(["docker", "run", "-v", "/tmp/a b:/work", "img"])
        self.assertEqual(argv[:3], ["sg", "docker", "-c"])
        self.assertEqual(len(argv), 4)
        self.assertIn("'/tmp/a b:/work'", argv[3])
        # and the quoting survives a real shell round-trip
        out = subprocess.run(["bash", "-c", "printf '%s\\n' " + argv[3].split(" ", 2)[2]],
                             capture_output=True, text=True).stdout.splitlines()
        self.assertIn("/tmp/a b:/work", out)

    def test_the_container_command_goes_through_the_wrapper(self):
        with unittest.mock.patch.object(office_recalc, "_docker_group_wrap_needed", return_value=True):
            cmd = DockerRunner("img:1").command("/tmp/job", ["--headless"])
        self.assertEqual(cmd[:3], ["sg", "docker", "-c"])
        self.assertIn("--headless", cmd[3])

