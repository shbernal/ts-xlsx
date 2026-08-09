<!-- One change, one purpose, fully green (CLAUDE.md §3). A half-migration is not reviewable,
     and a PR that needs a follow-up to be correct is two PRs. -->

## What changes, and why

<!-- The diff already says what. Say why it exists: the problem it solves, and — if a shape
     changed — why the new shape is better rather than merely different. Breaking changes are
     welcome here, unannounced ones are not. -->

## How it is proved correct

<!-- The check from docs/agent-correctness-playbook.md that covers this change, and what it
     printed. A behaviour change with no test that failed before it is asserted, not proved.
     A bug fix names the reproduction that failed first. -->

## Definition of Done (CLAUDE.md §2)

- [ ] `pnpm verify` is green — lint, typecheck, unit tests, corpus, generated docs, invariants
- [ ] New behaviour ships with unit tests in this change; a fixed bug ships with the failing
      reproduction that motivated it
- [ ] Any real-world file or edge case this came from is now a permanent case under
      `test/corpus/cases`, so it cannot regress — or N/A, this touches no reader/writer path
- [ ] Writer changes validated against the OOXML oracle (`pnpm test:ooxml`) — or N/A
- [ ] Public API is precisely typed and documented at the surface; the types are the docs
- [ ] No new runtime dependency — or it is named and justified below
- [ ] `CHANGELOG.md` `[Unreleased]` records this, including any break and what callers do
      about it — or N/A, nothing observable to a consumer changed

## Anything left undone

<!-- Assumptions you made, debt you knowingly took, a check you could not run and why.
     Leave a trail, not a mess — this is where the next agent picks it up. Delete if empty. -->
