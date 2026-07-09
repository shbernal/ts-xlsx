# Backlog dataset

A local, queryable snapshot of the upstream `exceljs/exceljs` backlog — the
**knowledge** we forked to preserve (see [`../../../STRATEGY.md`](../../../STRATEGY.md)).
Each open issue and PR becomes one durable JSON record here so the backlog survives
independently of GitHub's API, upstream's availability, and the eventual removal of
the `upstream` remote.

## Layout

```
docs/knowledge/backlog/
  issues/<zero-padded-number>.json   one harvested thread (issue OR pull request)
  attachments/<number>/<n>.<ext>     spreadsheet-shaped fixtures pulled from a thread
```

Numbers are zero-padded to 4 digits for stable sorting (`0140.json`); larger numbers
keep their natural width (`1695.json`).

## Producing / refreshing a record

```
node scripts/harvest/fetch-issue.mjs <number> [--repo owner/name]
```

Auth comes from the `gh` CLI (must be logged in); the script shells out to `gh api`
so no token ever enters the process. Re-running **refreshes the raw thread in place**
and never touches triage state — disposition lives only in
[`../BACKLOG.md`](../BACKLOG.md).

Attachment downloads are capped at 25 MB and restricted to spreadsheet-shaped
extensions (`xlsx`, `xlsm`, `xlsb`, `xls`, `csv`, `zip`, `ods`). Every discovered
link is still *recorded* in `attachments[]` (with `isFixture` telling you which were
downloaded), so nothing is silently dropped.

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

Disposition (`captured | superseded | out-of-scope`) is deliberately **not** in the
record; it is triage output and lives in [`../BACKLOG.md`](../BACKLOG.md).

## How this feeds the corpus

A harvested record is raw knowledge. Phase 1 distills the credible ones into
implementation-blind regression cases under [`../../../test/corpus/`](../../../test/corpus/README.md),
each linking back here via its `provenance.ref`. That link is how we honor the effort
we inherited — and how we guarantee we never fix the same bug twice.
