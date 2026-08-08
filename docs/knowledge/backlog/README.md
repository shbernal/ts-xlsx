# Backlog manifest — frozen provenance

`manifest.json` is a snapshot of the upstream `exceljs/exceljs` backlog as it stood at
fork time: every open issue and pull request, 794 items, captured 2026-07-09. **Nothing
reads it.** It is kept because it is *evidence*, and the claim it evidences is the one
CLAUDE.md §3 makes hardest to fake — that nothing was silently dropped.

The fork's premise (see [`../../architecture.md`](../../architecture.md)) was that the
value trapped upstream was knowledge, not code, and that the knowledge could be drained
into durable product: an implementation-blind case under
[`../../../test/corpus/`](../../../test/corpus/README.md), a behavior note under
[`../specs/`](../specs/), or a reasoned decision not to carry the item. That drain ran as
a work queue — one JSON record per thread under `issues/`, deleted as it was distilled,
so removal *was* the completion signal and no per-item disposition table had to be
maintained by hand.

**The queue reached empty, and the tooling that filled and drained it is gone.** What
survives is the pair that makes the account auditable after the fact:

- **this manifest** — the fixed denominator, so the universe cannot be quietly
  re-scoped to match whatever got done;
- **`git log`** — the per-item account, in the project's own durable terms, of what each
  item became or why it was not carried.

Do not rebuild the fetch tooling against this file. The universe is frozen at capture on
purpose: re-harvesting to pick up new upstream activity would re-couple us to a project
we have finished leaving, and the numbers in here go meaningless the moment anyone reads
them as a live tracker. They are provenance for decisions already made — ADR-0014 cites
item #141 exactly that way.

## What a record holds

Top level: `schema`, `repo`, `generatedAt`, the `total` / `issues` / `pullRequests`
counts, and `items[]`. Each item carries `number`, `type` (`issue` or `pull_request`),
`title`, `labels`, `reactions`, `comments`, and `url` — identity plus the triage signal
that set *priority* during the drain, never inclusion.

The file also carries `harvestComplete: true`, which told the fetch tooling that an
absent record meant "drained" rather than "never fetched". That flag now has no reader.
It is left in place rather than tidied away, because editing a frozen artifact to look
neater is how provenance stops being provenance.
