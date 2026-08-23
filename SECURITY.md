# Security

## Reporting a vulnerability

Please do not open a public issue for a security vulnerability. Instead, report it privately to
the project owner (see the repository's contact information on the hosting platform) with:

- the affected file(s) and, if you have one, a minimal reproduction;
- what you expected to happen and what actually happened;
- your assessment of impact (what an attacker could do with it).

You should get an acknowledgement within a few days. Please give a reasonable amount of time to
address the issue before any public disclosure.

## What this project treats as security-sensitive

- **Path validation** (`office_paths.py`): an allow-list guard on which absolute paths the
  office tools may read or write. A bypass here means a caller could point a tool at an
  arbitrary file on the host. Both a straightforward traversal and a symlink into an allowed
  directory pointing back out are in scope.
- **Co-editing identity/authorization** (`coedit.cjs`'s `detectCallerId`, the
  `EURO_COEDIT_AGENTS` allow-list): decides who may write into a live, shared document session.
  A mistake here does not surface as an error -- it surfaces as one caller silently writing into
  another's document, so it is treated with the same weight as an authorization bug in any
  multi-tenant system.
- **Credential handling** (`runner.cjs`'s `withKey`): the SSH private key used to reach the
  DocBuilder box is written to a 0600 temp file for the duration of one call and then
  overwritten before deletion. A change that makes the key outlive the call, widens its file
  permissions, or logs it is a regression here even if nothing else about the change looks
  security-related.
- **OOXML package consistency** (`package-consistency.cjs`): a package that is malformed in a
  way that a naive check would miss (a relationship pointing at a part that does not exist) can
  fail differently in different consumers -- this is a correctness/robustness boundary more than
  a classic vulnerability class, but a regression that lets an inconsistent package through is
  still worth reporting the same way.

## Out of scope

Bugs in an underlying dependency (LibreOffice/`soffice`, the `@modelcontextprotocol/sdk`, `zod`,
`playwright`) belong to that project, not this one -- please report those upstream. If you are
not sure which side a bug is on, report it here anyway and let the maintainer redirect it.
