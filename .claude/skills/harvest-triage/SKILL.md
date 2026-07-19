---
name: harvest-triage
description: Drain items from the upstream backlog queue — turn a harvested issue/PR record into durable product (a corpus case and/or spec note), then delete the record and commit. Use when processing docs/knowledge/backlog/issues/*.json, "triaging the backlog", "harvesting the backlog", distilling upstream issues/PRs into corpus cases or specs, or draining the harvest queue during Phase 1.
---

# Harvesting the backlog (the drain)

The upstream backlog was pulled into a local **work queue** once:
`docs/knowledge/backlog/issues/<n>.json`, one thread per file, measured against
`docs/knowledge/backlog/manifest.json` (the frozen universe). Your job is to
**drain** that queue: convert each thread's *knowledge* into durable product, then
delete the raw record. The harvest is complete; see `docs/architecture.md` for the origin
story this skill served.

## The model in one breath

- **A record present in `issues/` is work still to do. Removing it is how "done" is
  recorded — there is no per-item ledger.** The empty queue means Phase 1 is
  complete; `manifest.json` proves nothing was silently dropped.
- **The commit message is the source of truth.** It states, in durable terms, what
  knowledge was preserved (or deliberately not carried) and why.
- **Never write upstream issue/PR numbers into durable artifacts** — corpus case
  text, spec notes, or commit messages. We are leaving that project; the numbers
  become meaningless. Describe the *real-world scenario* instead; that is the part
  that lasts. (A number may sit in a case's optional `provenance` block as a
  disposable trace — never as its identity, and the durable text must stand without it.)
- **Capture broadly, implement selectively.** Phase 1 hoards knowledge cheaply. Do
  **not** fix legacy `lib/` code here — that is Phase 2/3. Your output is a corpus
  case, a spec note, or a reasoned "not carried".

## Pick what to work on

```
pnpm run harvest:status -- --clusters
```

Shows how much of the queue remains and its label breakdown. Prioritize by signal
(reactions, `bug`/`help wanted` labels, real reproductions with attached fixtures)
— but priority sets *order*, not *inclusion*. Everything gets dispositioned.

Read the record you picked: `docs/knowledge/backlog/issues/<n>.json`. It has the
full body, comments, labels, reactions, discovered attachment links, and — for PRs
— the changed-file map. Any downloaded fixtures are under
`docs/knowledge/backlog/attachments/<n>/`.

## Decide the disposition, then produce the durable artifact

Judge the thread and take exactly one path:

| The thread is… | Do this |
|---|---|
| A credible bug with a reproduction | Write a **corpus case** (invoke the `write-corpus-case` skill). If it ships a sample file, promote the fixture (below). |
| A proposal / enhancement / design discussion | Write a **spec note** under `docs/knowledge/specs/<slug>.md`: desired behavior, prior art, open questions. Feeds Phase 3 design. |
| A PR | Extract the **intent, root cause, and reproduction** into a corpus case and/or spec note. **Discard the diff** — it targets code we are deleting. Capture *what correct behavior looks like*, not the patch. |
| A trivial dep bump, an already-fixed report, a duplicate, or noise | Produce **no** artifact. The removal + a one-line commit ("not carried: …, because …") is the record. |

Promote a fixture that a corpus case needs out of the disposable attachments area
into the durable corpus, then reference it from the case:

```
test/corpus/fixtures/<case-slug>/<file>
```

## Close the loop: remove the record and commit

Once the knowledge is safely in a corpus case and/or spec note (or you've judged it
not worth carrying):

```
git rm docs/knowledge/backlog/issues/<n>.json
# also remove the attachments dir if nothing durable was promoted out of it:
# git rm -r docs/knowledge/backlog/attachments/<n>
git commit
```

Write the commit as the **durable account**, in the project's own terms:

- ✅ `harvest: capture full-row/column defined-name decoding (leaks NaN into serialized addresses)`
- ✅ `harvest: spec note for streaming pivot-table writes; prior art + open questions`
- ✅ `harvest: not carried — dep bump superseded by the fflate/zip rewrite`
- ❌ `capture #140` / `implement issue 140` / `port PR 636` ← never; numbers are dead weight

One item ≈ one small, coherent commit. Related items may batch, but keep the message
specific about *what behavior* each preserves.

## Guardrails

- **Nothing silently dropped.** Even "not carried" gets a commit that says why. The
  manifest is the fixed denominator; `harvest:status` is the live count.
- **Provenance is durable knowledge, not a link.** "Real files declare defined names
  like `$A:$A` for whole columns" survives the fork; "see issue 140" does not.
- **Do not re-harvest for new upstream activity.** The universe is frozen. If a
  record is missing but in the manifest, run `pnpm run harvest:all` (resumable) — do
  not re-run `harvest:list`.
- **Don't touch legacy `lib/` behavior.** Turning a case green is Phase 2/3. Here you
  only *record the target*, with `baseline` capturing what legacy does today.
