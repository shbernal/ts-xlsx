# ADR 0045: Hooks install with plain `lefthook install`

**Status:** Accepted (2026-10-14)

## Context

`prepare` ran `scripts/install-hooks.ts` instead of `lefthook install`. The script was added in
`968d388b` and hardened in `9bc57b03`, for three situations, all from one Windows machine:

- **A machine-wide `core.hooksPath`.** Lefthook refuses to install while it is set, and exits 1.
  Since pnpm 11 verifies dependencies before every script, a failing `prepare` breaks every
  `pnpm run`. The script skipped the install and said so instead.
- **An EDR that quarantines files in place.** The binary kept its name and size, but opening it
  failed. Lefthook's generated hooks end their search for the binary with a bare `echo`, so a hook
  that could not start lefthook printed one line and exited 0. Ten commits landed with no hooks
  run. The script rewrote that branch to exit 1.
- **A damaged `node_modules/lefthook`**, also the EDR's doing, which the script reported as a
  failure rather than as "lefthook not installed".

None of these apply on the machines this repo is developed on now: no `core.hooksPath`, no EDR.
And the script had a bug of its own. In a `git worktree`, it looked for the hooks in the
worktree's private git directory (`.git/worktrees/<name>/hooks`), which does not exist, and
crashed after lefthook had already installed into the shared `.git/hooks` correctly.

## Decision

`prepare` is `lefthook install`. `scripts/install-hooks.ts` is deleted.

Lefthook's own postinstall stays denied in `pnpm-workspace.yaml`, because it runs
`lefthook install -f`, and `-f` writes into a shared `core.hooksPath` directory.

## Consequences

- A developer with `core.hooksPath` set gets a failing `prepare`. Unset it for this repo, or
  install with `--ignore-scripts` and run the hooks some other way.
- A hook that cannot start lefthook passes silently again. On a machine where that can happen,
  restore the script (`git show 7699212c:scripts/install-hooks.ts`), with the worktree bug fixed
  by using `git rev-parse --git-common-dir`.
- Worktrees install correctly.
