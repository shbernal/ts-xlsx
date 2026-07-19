# Backlog dataset

A local, queryable snapshot of the upstream `exceljs/exceljs` backlog — the
**knowledge** we forked to preserve (see [`../../architecture.md`](../../architecture.md)).
Each open issue and PR becomes one JSON record here so the backlog survives
independently of GitHub's API, upstream's availability, and the eventual removal of
the `upstream` remote.

This directory is a **work queue that drains once** — a record present is work still
to do; processing it distills the knowledge into durable product and **deletes the
record** (see [`../BACKLOG.md`](../BACKLOG.md) for the drain model).

## Layout

```
docs/knowledge/backlog/
  manifest.json                      the frozen universe — every open issue+PR at fork time
  issues/<zero-padded-number>.json   one harvested thread (issue OR pull request)
  attachments/<number>/<n>.<ext>     spreadsheet-shaped fixtures pulled from a thread
```

Numbers are zero-padded to 4 digits for stable sorting (`0140.json`); larger numbers
keep their natural width (`1695.json`).

## The toolchain (one-time harvest)

The harvest is a **one-time** operation — we never re-run it to pick up new upstream
activity. Auth comes from the `gh` CLI (must be logged in); scripts shell out to
`gh api` so no token ever enters a process.

```
pnpm run harvest:list                 # snapshot the universe → manifest.json (once)
pnpm run harvest:all                  # fill the queue: fetch every item (resumable)
pnpm run harvest:status -- --clusters # follow filled-vs-remaining; remaining by label
pnpm run harvest -- <number>          # (re)fetch a single thread — the atom
```

`harvest:all` is resumable: rerun after any interruption and it skips records already
on disk. When the queue is whole it flips `manifest.harvestComplete`, which is how
`harvest:status` knows an absent record means "drained", not "never fetched".

Attachment downloads are capped at 25 MB and restricted to spreadsheet-shaped
extensions (`xlsx`, `xlsm`, `xlsb`, `xls`, `csv`, `zip`, `ods`). Every discovered
link is still *recorded* in `attachments[]` (with `isFixture` telling you which were
downloaded), so nothing is silently dropped. Re-fetching a single thread refreshes
the raw record in place.

## Record schema — `ts-xlsx/backlog-item@1`

| Field | Meaning |
|---|---|
| `schema` | Version tag; bump when the shape changes. |
| `repo`, `number`, `type` | Source coordinates; `type` is `issue` or `pull_request`. |
| `url`, `title`, `state`, `author` | Thread identity. |
| `createdAt`, `closedAt` | ISO timestamps (`closedAt` null while open). |
| `labels`, `reactions`, `commentCount` | Triage signal — reactions/frequency set *priority*, not inclusion. |
| `body`, `comments[]` | Full text; `comments[]` = `{author, authorAssociation, createdAt, reactions, body}`. |
| `attachments[]` | `{url, ext, isFixture}` for every link found in body + comments. |
| `pr` | Present only for pull requests: `{merged, mergeable, baseRef, additions, deletions, changedFiles[]}`. We capture a PR's *intent, repro, and changed-file map* — not the diff, which targets code we are deleting. |

Disposition is deliberately **not** in the record and is **not** tracked per-item
anywhere: a record present is undrained; a record removed (with its commit) is
handled. See the [drain model](../BACKLOG.md).

## How this feeds the corpus

A harvested record is raw knowledge. Draining it distills the credible scenario into
an implementation-blind regression case under
[`../../../test/corpus/`](../../../test/corpus/README.md) (and/or a spec note under
[`../specs/`](../specs/)), and the record is then deleted. The durable case captures
the *real-world behavior* in its own terms — not the upstream number, which dies with
the fork — so we never fix the same bug twice. The commit message is the account of
what was preserved.
