#!/usr/bin/env python3
"""Tests for office_runner, exercised against fake_soffice.

The point of the fake is that these paths are testable with no LibreOffice installed: every
failure mode the real soffice can hit is reproduced deterministically, so the runner's error
handling, cleanup and timeout-kill are verified here and stay verified once a real engine is
wired in behind the same interface.

Run:  python3 -m unittest test_office_runner    (or: python3 test_office_runner.py)
"""

import os
import subprocess
import sys
import time
import unittest

import office_runner
from office_runner import ExternalToolError, job_workspace, run_producing_file

HERE = os.path.dirname(os.path.abspath(__file__))
FAKE = os.path.join(HERE, "fake_soffice.py")


def fake_cmd(mode, outdir=None, src=None, pid_file=None):
    """Build a command that invokes fake_soffice the way the runner would soffice.

    INPUT stays last and --outdir carries the output directory, matching the contract the
    fake shares with a real `soffice --convert-to xlsx --outdir OUT INPUT`.
    """
    cmd = [sys.executable, FAKE, "--fake-mode=" + mode]
    if pid_file is not None:
        cmd += ["--pid-file", pid_file]
    cmd += ["--convert-to", "xlsx"]
    if outdir is not None:
        cmd += ["--outdir", outdir]
    if src is not None:
        cmd += [src]
    return cmd


def pid_alive(pid):
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    return True


class RunProducingFileTests(unittest.TestCase):
    # --- mode 4: correct operation -------------------------------------------------
    def test_ok_returns_output_path(self):
        with job_workspace() as work:
            src = os.path.join(work, "gate.xlsx")
            with open(src, "w", encoding="utf-8") as fh:
                fh.write("input")
            outdir = os.path.join(work, "out")
            expected = os.path.join(outdir, "gate.xlsx")

            result = run_producing_file(
                fake_cmd("ok", outdir=outdir, src=src), expected, timeout=10
            )

            self.assertEqual(result, expected)
            self.assertTrue(os.path.exists(expected))
            with open(expected, encoding="utf-8") as fh:
                self.assertEqual(fh.read(), "input")

    # --- mode 1: exit 0, no output file --------------------------------------------
    def test_exit_zero_without_output_is_error(self):
        with job_workspace() as work:
            expected = os.path.join(work, "out", "gate.xlsx")
            with self.assertRaises(ExternalToolError) as ctx:
                run_producing_file(fake_cmd("no-output"), expected, timeout=10)
            msg = str(ctx.exception)
            self.assertIn("no output", msg)
            self.assertIn(expected, msg)

    # --- mode 2: non-zero exit + stderr --------------------------------------------
    def test_nonzero_exit_reports_stderr_once(self):
        with job_workspace() as work:
            expected = os.path.join(work, "out", "gate.xlsx")
            with self.assertRaises(ExternalToolError) as ctx:
                run_producing_file(fake_cmd("error"), expected, timeout=10)
            msg = str(ctx.exception)
            self.assertIn("exit 1", msg)
            self.assertIn("could not be loaded", msg)
            # One clean line, not a dumped traceback or a multi-line stderr blob.
            self.assertEqual(msg.count("\n"), 0)

    # --- mode 3: never returns -> timeout kills the process ------------------------
    def test_hang_times_out_and_actually_kills(self):
        with job_workspace() as work:
            expected = os.path.join(work, "out", "gate.xlsx")
            pid_file = os.path.join(work, "child.pid")

            start = time.monotonic()
            with self.assertRaises(ExternalToolError) as ctx:
                run_producing_file(
                    fake_cmd("hang", pid_file=pid_file), expected, timeout=1
                )
            elapsed = time.monotonic() - start

            self.assertIn("timed out", str(ctx.exception))
            # It returned roughly at the timeout, not after the fake's 3600s sleep.
            self.assertLess(elapsed, 30)

            # The fake recorded its pid before sleeping; that process must now be gone.
            self.assertTrue(os.path.exists(pid_file), "fake never wrote its pid file")
            with open(pid_file, encoding="utf-8") as fh:
                child_pid = int(fh.read().strip())
            deadline = time.monotonic() + 5
            while pid_alive(child_pid) and time.monotonic() < deadline:
                time.sleep(0.05)
            self.assertFalse(
                pid_alive(child_pid),
                "hung process %d survived the timeout" % child_pid,
            )

    # --- missing binary: still one clean error, not a raw OSError ------------------
    def test_missing_binary_is_clean_error(self):
        expected = os.path.join(HERE, "nope", "gate.xlsx")
        with self.assertRaises(ExternalToolError) as ctx:
            run_producing_file(
                [os.path.join(HERE, "definitely-not-a-binary-xyz")], expected, timeout=5
            )
        self.assertIn("not found", str(ctx.exception))


class JobWorkspaceTests(unittest.TestCase):
    def test_workspace_removed_on_success(self):
        with job_workspace() as work:
            self.assertTrue(os.path.isdir(work))
            captured = work
        self.assertFalse(os.path.exists(captured))

    def test_workspace_removed_on_exception(self):
        captured = {}
        with self.assertRaises(RuntimeError):
            with job_workspace() as work:
                captured["path"] = work
                self.assertTrue(os.path.isdir(work))
                raise RuntimeError("boom")
        self.assertFalse(os.path.exists(captured["path"]))

    def test_every_failure_mode_leaves_no_temp_dir(self):
        """A full job in each mode: whatever the outcome, the workspace is gone after."""
        modes = ["ok", "no-output", "error"]
        leaked = []
        for mode in modes:
            work_seen = {}
            try:
                with job_workspace() as work:
                    work_seen["path"] = work
                    src = os.path.join(work, "gate.xlsx")
                    with open(src, "w", encoding="utf-8") as fh:
                        fh.write("x")
                    outdir = os.path.join(work, "out")
                    expected = os.path.join(outdir, "gate.xlsx")
                    run_producing_file(
                        fake_cmd(mode, outdir=outdir, src=src), expected, timeout=10
                    )
            except ExternalToolError:
                pass
            if os.path.exists(work_seen.get("path", "")):
                leaked.append((mode, work_seen["path"]))
        self.assertEqual(leaked, [], "temp dirs leaked: %r" % leaked)


class KillHelperTests(unittest.TestCase):
    def test_kill_process_tree_reaps_child(self):
        proc = subprocess.Popen(
            [sys.executable, "-c", "import time; time.sleep(3600)"],
            start_new_session=True,
        )
        self.assertTrue(pid_alive(proc.pid))
        office_runner._kill_process_tree(proc)
        deadline = time.monotonic() + 5
        while pid_alive(proc.pid) and time.monotonic() < deadline:
            time.sleep(0.05)
        self.assertFalse(pid_alive(proc.pid))


if __name__ == "__main__":
    unittest.main(verbosity=2)
