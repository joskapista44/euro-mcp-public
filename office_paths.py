"""Where the office tools are allowed to touch files.

The tools take an ABSOLUTE PATH from the caller, and until now they took any of
them: pointing `xlsx_set_cells` at a customer workbook, another agent's config or
anything else on the box worked exactly as well as pointing it at a scratch copy.
That is a gap the session's permission rules cannot close, because the file IO
happens in this helper process, not through the agent's Read/Write tools.

So the guard lives here, next to the IO. It is an ALLOW-list, not a deny-list: a
deny-list has to predict every sensitive location and silently loses whenever
something new appears, while an allow-list fails closed on anything unforeseen.

The default is `/tmp`, and that is not arbitrary - it is where a typical deployment's
own file-download workflow already puts these files (and where today's verified
recalc ran). Everything under /home stays out by default: repos, credential
stores, and per-agent configuration. Widen it deliberately with

    OFFICE_ALLOWED_ROOTS=/tmp:/srv/shared

not by editing this file.

Paths are compared after realpath() on BOTH sides, so a symlink into an allowed
directory cannot be used to reach a file outside it - the check sees where the
path actually lands, not what it is spelled as.
"""

import os

DEFAULT_ROOTS = "/tmp"


class PathNotAllowed(Exception):
    """The caller asked for a file outside every allowed root."""


def allowed_roots(env=None):
    """The configured roots, resolved. Empty entries are dropped, not treated as '/'."""
    raw = (env or os.environ).get("OFFICE_ALLOWED_ROOTS", DEFAULT_ROOTS)
    roots = []
    for item in raw.replace(",", ":").split(":"):
        item = item.strip()
        if not item:
            continue
        roots.append(os.path.realpath(item))
    return roots


def _within(path, root):
    # Compare on path boundaries: "/tmp/x" is inside "/tmp", "/tmpfoo" is not.
    return path == root or path.startswith(root.rstrip("/") + "/")


def check_path(path, env=None):
    """Return the resolved path, or raise PathNotAllowed.

    Resolution happens before the check, which is the point: `/tmp/link -> /home/user`
    resolves to `/home/user` and is refused, even though the spelling looked fine.
    A non-existent file still resolves (its directory does), so a write to a new
    file in a forbidden directory is refused too, not just an overwrite.
    """
    resolved = os.path.realpath(path)
    roots = allowed_roots(env)
    if not roots:
        raise PathNotAllowed(
            "no allowed roots configured: OFFICE_ALLOWED_ROOTS is empty, so every path is refused"
        )
    for root in roots:
        if _within(resolved, root):
            return resolved
    raise PathNotAllowed(
        "path is outside the allowed roots: %s (allowed: %s). "
        "Copy the file into an allowed directory, or widen OFFICE_ALLOWED_ROOTS deliberately."
        % (resolved, ", ".join(roots))
    )
