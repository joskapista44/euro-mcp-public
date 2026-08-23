# Contributing

## Before you start

This project was extracted from a larger internal deployment. If you hit a reference to
something not in this repository (an internal ticket ID in a comment, a decision attributed to
someone by name), that is expected -- it is the original reasoning left in place because it
still explains the code, not a pointer to something you need to go find.

## Workflow

1. Open an issue describing the bug or the change before sending a large patch -- for anything
   beyond a small fix, agreeing on the approach first avoids wasted work on both sides.
2. Keep changes scoped: one fix or one feature per pull request. Unrelated formatting/reordering
   in the same diff makes the actual change harder to review and harder to revert if needed.
3. Add or update a test with every behavior change. This codebase treats a test as the record of
   *why* a piece of logic exists (see the `test_*`/`test-*` file headers) -- a change without a
   test loses that record for the next person.
4. Run the relevant test files before opening the PR (see README.md's "Running the tests"
   section). There is no CI configured in this export; you are the first gate.

## Code style

- Node/CJS: no build step, no bundler -- plain `require()`. Keep new modules dependency-light;
  a new npm dependency should be justified in the PR description.
- Python: no third-party packages, no `pip install` step. If a task looks like it needs one,
  that is worth raising as a design question before adding it, not working around silently.
- Match the file you are editing: comment density, naming, and error-message language (Hungarian
  and English are both present in this codebase; a file is consistently one or the other).

## Security-sensitive changes

Anything touching path validation (`office_paths.py`), the co-editing identity/allowlist logic
(`coedit.cjs`'s `detectCallerId`), or credential handling (`runner.cjs`'s key resolution) should
call that out explicitly in the PR description, even for a change that looks small. See
`SECURITY.md` for how to report a vulnerability instead of opening a public PR.
