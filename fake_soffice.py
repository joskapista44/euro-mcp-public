#!/usr/bin/env python3
"""A stand-in for `soffice`, used only by the office-runner tests.

It impersonates the command line the runner drives LibreOffice with
(`... --convert-to xlsx --outdir OUT INPUT`) and, selected by --fake-mode, reproduces
each behaviour the real soffice is known to inflict on a caller. No LibreOffice, no
network, python3 stdlib only.

Modes (pass `--fake-mode=<mode>` as the FIRST argument):

  ok         copy INPUT into --outdir and exit 0                    (the happy path)
  no-output  exit 0 without writing anything                       (soffice really does this,
                                                                     which is why the runner
                                                                     must check the file, not
                                                                     the exit code)
  error      write a message to stderr and exit non-zero
  hang       write our pid to --pid-file (if given), then sleep effectively forever
             so the runner's timeout has something real to kill

Contract shared with the runner: INPUT is always the LAST argument, and the output
directory is the value following --outdir. The output file keeps the input's basename,
matching how `soffice --convert-to xlsx` names its result.
"""

import os
import shutil
import sys
import time


def _opt_value(argv, name):
    """Return the argument that follows `name`, or None."""
    for i, a in enumerate(argv):
        if a == name and i + 1 < len(argv):
            return argv[i + 1]
    return None


def main(argv):
    mode = None
    if argv and argv[0].startswith("--fake-mode="):
        mode = argv[0].split("=", 1)[1]
        argv = argv[1:]

    if mode is None:
        sys.stderr.write("fake_soffice: missing --fake-mode=<ok|no-output|error|hang>\n")
        return 2

    if mode == "no-output":
        # Exit cleanly, produce nothing. The exit code lies about success.
        return 0

    if mode == "error":
        sys.stderr.write("Error: source file could not be loaded\n")
        return 1

    if mode == "hang":
        pid_file = _opt_value(argv, "--pid-file")
        if pid_file:
            with open(pid_file, "w", encoding="utf-8") as fh:
                fh.write(str(os.getpid()))
                fh.flush()
                os.fsync(fh.fileno())
        # Sleep far longer than any test timeout; the runner is expected to kill us.
        while True:
            time.sleep(3600)

    if mode == "ok":
        outdir = _opt_value(argv, "--outdir")
        if not outdir or not argv:
            sys.stderr.write("fake_soffice: ok mode needs --outdir and an input file\n")
            return 2
        src = argv[-1]
        os.makedirs(outdir, exist_ok=True)
        shutil.copy(src, os.path.join(outdir, os.path.basename(src)))
        return 0

    sys.stderr.write("fake_soffice: unknown mode %r\n" % mode)
    return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
