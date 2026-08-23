#!/usr/bin/env python3
"""External-binary runner for the office layer.

Whatever engine ends up doing the recalc/conversion (host soffice, a container, something
else), it is driven by handing a command line to a subprocess and reading back a file. That
step has four failure modes worth naming, and every one of them has to end in a single clean
error rather than a corrupt workbook or a hung session:

  1. the tool exits 0 but writes no output file   -> trust the FILE, not the exit code
  2. the tool exits non-zero with a stderr message -> surface it, trimmed, once
  3. the tool never returns                        -> time out and actually kill it
  4. the tool works                                -> return the output path

This module is engine-agnostic on purpose: it knows nothing about xlsx, recalc, or
LibreOffice flags. The caller builds the command and says which output file proves success.
python3 stdlib only.
"""

import contextlib
import os
import shutil
import signal
import subprocess
import tempfile


class ExternalToolError(Exception):
    """A run of an external tool failed in one clean, reportable way."""


@contextlib.contextmanager
def job_workspace(prefix="office-job-"):
    """A temporary directory that is ALWAYS removed, success or exception.

    Every conversion job runs in a throwaway workspace (fixture in, output out, engine
    profile alongside). A leaked temp dir is a slow leak on a long-running bridge, and a
    reused one is a way for a stale file to be mistaken for a fresh result - so cleanup runs
    on the error paths too, which is exactly where a naive implementation forgets it.
    """
    path = tempfile.mkdtemp(prefix=prefix)
    try:
        yield path
    finally:
        shutil.rmtree(path, ignore_errors=True)


def _argv0(cmd):
    return cmd[0] if cmd else "<empty command>"


def _kill_process_tree(proc):
    """SIGKILL the child's whole process group, then reap it.

    The child is started in its own session (start_new_session), so killing the group
    reaches any grandchildren a wrapper might have spawned - a plain proc.kill() would leave
    those behind and the timeout would not really cut the process off.
    """
    with contextlib.suppress(ProcessLookupError, PermissionError):
        os.killpg(os.getpgid(proc.pid), signal.SIGKILL)
    with contextlib.suppress(Exception):
        proc.wait(timeout=10)


def run_producing_file(cmd, output_path, timeout):
    """Run `cmd`, requiring it to leave a file at `output_path`.

    Returns `output_path` on success. Raises ExternalToolError - and nothing else - for each
    of the four failure modes, with a single trimmed message. `timeout` is in seconds.
    """
    try:
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            start_new_session=True,  # own process group, so the timeout can kill the tree
        )
    except FileNotFoundError as exc:
        raise ExternalToolError("external tool not found: %s" % _argv0(cmd)) from exc

    try:
        _out, err = proc.communicate(timeout=timeout)
    except subprocess.TimeoutExpired as exc:
        _kill_process_tree(proc)
        raise ExternalToolError(
            "external tool timed out after %gs and was killed: %s"
            % (timeout, _argv0(cmd))
        ) from exc

    if proc.returncode != 0:
        detail = (err or "").strip().splitlines()
        detail = detail[-1] if detail else "(no stderr)"
        raise ExternalToolError(
            "external tool failed (exit %d): %s" % (proc.returncode, detail[:300])
        )

    if not os.path.exists(output_path):
        # The exit code said success; the missing file says otherwise. The file wins.
        raise ExternalToolError(
            "external tool exited 0 but produced no output at %s" % output_path
        )

    return output_path
