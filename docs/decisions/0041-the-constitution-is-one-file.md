# ADR 0041: The constitution is one file, `AGENTS.md`

**Status:** Accepted 2026-10-13

## Context

The constitution was once one file behind a symlink: `CLAUDE.md` was real and `AGENTS.md`
pointed at it, so every runtime converged on the same rules by construction. That
construction did not survive Windows. Creating a symlink there needs admin rights or
Developer Mode, so a Windows checkout materialized `AGENTS.md` as a ~9-byte text file
containing the literal string `CLAUDE.md`, and an AGENTS.md-reading agent silently got no
constitution at all. A failure that quiet is the worst kind, so `AGENTS.md` became a real
file and `scripts/check-constitution.ts` replaced the symlink's guarantee with a machine
check: the two files had to stay byte-identical, line-ending-insensitively, or the build
failed.

That bought correctness at the price of two files and a gate to hold one document.
Claude Code now reads `AGENTS.md` natively, as the other agent runtimes already did, so no
runtime needs a `CLAUDE.md` to find the rules, and the mirror has nothing left to mirror.

## Decision

**`AGENTS.md` is the constitution, and there is no copy.** `CLAUDE.md` and
`scripts/check-constitution.ts` are deleted, along with the `constitution:check` package
script and the verify step that ran the check. Their whole job was keeping two files
identical; with one file, that guarantee holds by construction again.

Every citation that read `CLAUDE.md §n` now reads `AGENTS.md §n`. The section numbering is
untouched, so each citation still means what it meant when it was written. One was wrong
all along (ADR 0035 said §6 for what §1.6 says) and is corrected rather than carried over.

## Consequences

- One file to read, one file to edit, nothing to keep in sync. The failure mode the check
  existed for, a one-sided edit to the constitution, is gone with the second file.
- A tool that still looks for `CLAUDE.md` by habit finds nothing. That is the §1 clean
  break applied to our own tooling: the runtimes that matter read `AGENTS.md`, and a
  second file kept as a courtesy to the rest would be exactly the compatibility weight
  the fork exists to shed.
- The symlink story lives here now instead of in the deleted check's comment, because
  "why was this ever two files" is a question the tree can no longer answer on its own.
