# Backlog: the drain model

The upstream backlog is a **work queue we drain once**, not a ledger we maintain.
There is deliberately **no permanent per-item disposition table here** — tracking
every issue/PR number as `captured`/`implemented` would preserve numbers that go
meaningless the moment we finish leaving that project. Instead:

- **The queue** is `backlog/issues/<n>.json` — one harvested thread per file, pulled
  in once. A file present is work still to do.
- **Draining** an item means distilling its knowledge into durable product (a corpus
  case under [`../../test/corpus/`](../../test/corpus/README.md) and/or a spec note
  under [`specs/`](specs/)), then **deleting the raw record**. Removal *is* the
  completion signal.
- **The commit message is the source of truth.** It records, in the project's own
  durable terms, what behavior was preserved — or why an item was not carried. It
  never cites an upstream number.
- **The proof nothing was silently dropped** (`CLAUDE.md` §"no silent caps") is
  structural, not a hand-maintained list:
  - [`backlog/manifest.json`](backlog/README.md) is the frozen **universe** — every
    open issue + PR captured at fork time. The fixed denominator.
  - `git log` is the per-item account of what happened to each.
  - An **empty queue** means Phase 1 is complete.

## Follow it

```
npm run harvest:status -- --clusters
```

Shows filled-vs-remaining against the manifest and the remaining queue by label, so
you can pick high-signal work. See the `harvest-triage` skill for the per-item drain
workflow and the `write-corpus-case` skill for authoring cases.

## Clusters

Group durable output by theme: `address-decoding` · tables · styles · streaming ·
pivot · images · conditional-formatting · dates · formulas · csv · types ·
security/deps.

## One-time, not continuous

We do **not** re-harvest to pick up new upstream activity — the universe is frozen at
capture. `npm run harvest:list` snapshots it once; `npm run harvest:all` (resumable)
fills the queue; then agents drain it.
