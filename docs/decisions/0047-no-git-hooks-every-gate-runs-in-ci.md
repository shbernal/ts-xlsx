# ADR 0047: No git hooks; every gate runs in CI

**Status:** Accepted (2026-11-23) · supersedes [ADR 0045](./0045-hooks-install-with-plain-lefthook.md) · leaves the single gate definition of [ADR 0022](./0022-verification-is-one-cached-parallel-entrypoint.md) unchanged

## Context

Lefthook ran three kinds of hook here. `pre-commit` formatted, linted and char-checked the staged
files. `pre-push` ran `node scripts/verify.ts --full`. `commit-msg` pulled two rules from
`shbernal/lefthook-rules`: `no-ai-attribution` and `no-shell-quoting-leak`.

The first two duplicated CI. `pre-push` was the same command as `corpus.yml` by design (ADR 0022),
and the staged-scoped `pre-commit` jobs were a narrower copy of gates `verify --full` already runs
over the whole tree. A duplicate is only worth its upkeep if it is reliable, and these were not:

- **A machine-wide `core.hooksPath`** makes `lefthook install` exit 1, and since `prepare` ran it,
  every `pnpm run` failed with it. ADR 0045 accepted that and told the developer to unset the
  setting or install with `--ignore-scripts`.
- **An EDR that quarantines files in place** left the hooks unable to start lefthook, and the
  generated wrapper then exited 0. Ten commits landed with no hook run (ADR 0045's context).
- **The rules remote** was fetched on install and cached; an offline install left the
  commit-message rules silently unenforced until the next fetch.

So the local gate was the one that could vanish without a word, and CI was the one that could not.
Keeping both meant two configurations to read and the less trustworthy one running first.

## Decision

No git hooks: no lefthook, no pre-commit, no pre-push. `lefthook.yml`, the `prepare` script, the
`lefthook` devDependency and its `allowBuilds` entry are removed.

Every gate runs in CI. The code gates were already there. The one gate that existed only locally,
the commit-message check, moves to a `commit-check` job in `.github/workflows/ci.yml`, running
`commit-check/commit-check-action` over every commit of the pull request (`fetch-depth: 0`,
`message: true`, `branch: false`). Its rules are in `.github/cchk.toml`, and only one is on:
`ai_attribution = "forbid"`. Conventional Commits, the imperative subject and the subject length
checks are switched off explicitly, because the tool's defaults have them on.

The AI disclosure moves to the pull request: `.github/PULL_REQUEST_TEMPLATE.md` asks for an
`AI disclosure:` line naming the harness and model, or "written by hand". A squash merge builds its
message from the commits rather than the description, so the disclosure stays off `main`.

`no-shell-quoting-leak` is dropped, not moved. commit-check has no equivalent, and the leak it
caught (a PowerShell here-string delimiter landing as a bare `@` line) is prevented upstream by
AGENTS.md's rule that a multi-line message goes through a file and `git commit -F`.

## Consequences

- `pnpm install` no longer touches `.git/hooks`, so `core.hooksPath` and an EDR stop mattering to
  this repo. Nothing has to be unset and nothing needs `--ignore-scripts`.
- Feedback arrives later: a format slip or an em dash shows up on the runner rather than at
  commit time. `node scripts/verify.ts --full` (or `--quick`) is still the way to see it first, and
  the agent harness's Stop hook still runs `verify --full --cached` at the end of a turn. That
  hook is Claude Code's, not git's, and stays.
- A leaked here-string delimiter is no longer caught by anything. It is caught by not writing one.
- The `commit-check` job is not among the required checks of the `main` ruleset, so for now it
  reports without blocking a merge. Making it required is a repository-settings change, not a
  change to this tree.
