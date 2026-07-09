# Progress

> **Live execution tracker.** This file records *where we are* and *what remains*.
> It is subordinate to the two authorities it references — keep them the source of truth:
> - [`CLAUDE.md`](CLAUDE.md) — the constitution (principles).
> - [`STRATEGY.md`](STRATEGY.md) — the authoritative phased plan (Phases 0–4).
>
> When a phase's status changes, update this file **and** `STRATEGY.md` in the same breath.
> Legend: ✅ done · 🔜 next · ⏳ pending · 🧊 deferred-on-purpose · ❓ open decision.

_Last updated: 2026-07-09._

---

## The one guardrail that governs ordering

**Freeze the legacy tree's *shape* until the backlog is drained.** The knowledge in the
~654 issues + ~139 PRs is the prize; a regression corpus preserves that knowledge across
the rewrite, but it does **not** preserve the *merge-ability* of open PRs. Every high-drift
move — whole-tree reformat, `.js`→`.ts` rename, module-layout change, dependency swap —
breaks the open PRs at once and forces expensive re-derivation. So those moves come **last**.
Harvest and bank value first; modernize only after there's nothing left in the backlog that
needs the old shape. (This is the mistake from the prior dead-project takeover: modernized
too early, drifted past the backlog, lost it.)

---

## Status by phase (canonical numbering from `STRATEGY.md`)

### ✅ Takeover / hosting independence — *`STRATEGY.md` Phase 4, pulled forward*
The infrastructure clean break is **done** (2026-07-09):
- ✅ Fresh **non-fork** public repo `shbernal/ts-xlsx` created; full history + all release
  tags mirror-pushed in; inherited upstream branches pruned to `master`.
- ✅ Fork identity repointed in-tree: `package.json` (name `@shbernal/ts-xlsx`, repo/homepage/
  bugs, original creator → `contributors`), `LICENSE` (original MIT copyright retained +
  fork copyright added), `README.md` (independent-fork notice).
- ✅ `upstream` (exceljs/exceljs) kept as a **read-only** remote (push disabled) purely to
  feed the harvest; to be dropped entirely at harvest end.
- ✅ Local env ready at `C:\Users\000023500\dev\ts-xlsx`: deps installed, `npm run build`
  green, `npm run test:unit` green (883 passing, 1 pending).

### 🔜 Phase 0 — Foundation & harvest tooling  *(additive only — does NOT touch legacy shape)*
- ✅ **Harvest toolchain** built & proven (shared core `scripts/harvest/lib.mjs`):
  - `harvest:list` → freezes the universe into `docs/knowledge/backlog/manifest.json` (654
    issues + 140 PRs). `harvest:all` → resumable queue fill. `harvest:status [--clusters]` →
    drain progress. Single-thread atom (`npm run harvest <n>`) pulls body, comments, labels,
    reactions, attachment links + spreadsheet fixtures, PR changed-file map into
    `backlog/issues/<n>.json` (schema `ts-xlsx/backlog-item@1`). Auth via `gh api`; 25 MB /
    spreadsheet-ext download cap. Documented in `docs/knowledge/backlog/README.md`.
  - ⏳ *Fill + drain* across all 794 items is Phase 1 work.
- ✅ **Regression corpus format** defined & runnable under `test/corpus/`: implementation-blind
  cases assert against an **adapter contract** vocabulary (`current` adapter binds it to
  `lib/`); each behavior carries a `baseline` (`pass`/`fail` vs legacy) so the runner tells a
  known-open bug from a real regression (exit 1 only on regression). `npm run corpus`.
  Documented in `test/corpus/README.md`. **Drain model** (no per-item ledger) documented in
  `docs/knowledge/BACKLOG.md`.
- ⏳ CI skeleton for the *additive* checks only (corpus + existing suite). **No toolchain
  rip-out yet** (see 🧊 below). — *next slice.*
- **Exit:** ✅ **MET** — issue #140 harvested end-to-end → corpus case `0140-address-decoding`
  runs red/green against current code through the adapter (2 green regression-locks: the
  `$1` crash is fixed upstream; 1 known-open red: `decodeEx('$1:$1')` still leaks
  `"$undefined$1"` / `"NaN:NaN"` into serialized addresses). Existing suite still 883/1 pending.
- 🧊 **Deferred out of Phase 0:** replacing Babel/Grunt/Mocha with TS/Vitest/Biome/tsup.
  Highest-drift action in the plan → runs last (see Phase 3/4).

### 🔜 Phase 1 — Harvest the backlog  *(one-time drain of the queue)*
Model: `harvest:list` freezes the universe (`manifest.json`, 794 items), `harvest:all` fills
the queue (`backlog/issues/*.json`), agents **drain** it — distill each thread into durable
product, delete the record, commit. No per-item ledger; the commit message is the account of
record; durable artifacts never cite upstream numbers (they die with the fork).
- ✅ **Harvest toolchain** built (`harvest:list` / `:all` / `:status` + single-thread atom;
  shared core in `scripts/harvest/lib.mjs`). Manifest snapshot taken: 654 issues + 140 PRs.
- ✅ **Agent skills** authored: `harvest-triage` (per-item drain) and `write-corpus-case`.
- ⏳ **Fill the queue:** `npm run harvest:all` (resumable) — *next slice.*
- ⏳ **Drain:** cluster by theme (tables, styles, streaming, pivot, images, conditional
  formatting, dates, formulas, csv, types, security/deps). Credible bug/repro → corpus case
  (+ fixture); PR → intent + repro + root cause (discard the diff); proposal → spec note;
  dep bump/noise → not carried (commit says why). High-value PR *code* may still be pushed as
  a `harvest/pr-<n>` branch so a patch stays cherry-pickable (branch names are transient).
- **Exit:** the queue is empty; every carried item left a corpus case and/or spec note; corpus
  runs against current code (mostly red where bugs are real). Follow via `harvest:status`.

### ⏳ Phase 2 — Stabilize-to-validate  *(time-boxed, on the frozen legacy tree)*
- Bank the cheaply-capturable value: fix high-value low-risk bugs test-first, and — pending
  the ❓ decision below — **merge the mergeable open PRs** onto `master` while the tree still
  matches their base. Land cheap security/dep fixes only if they don't cost more than the
  rewrite would anyway. No refactoring/typing/restyling of legacy code.
- **Exit:** corpus expresses agreed "correct behavior" per captured case; current-code
  pass/fail recorded as the baseline the rewrite must beat.

### ⏳ Phase 3 — The rebuild  *(discard the debt)*
- Greenfield TypeScript, corpus-driven: **core model → XML layer → xlsx r/w → streaming → csv**,
  each landed fully green before the next depends on it. Here is where the deferred toolchain
  modernization (TS/Vitest/Biome/tsup, dep swaps like `fflate`) actually happens.
- **Exit:** feature-parity-or-better on the corpus for all in-scope areas; legacy `lib/`
  deleted; dependency tree small and audit-clean.

### ⏳ Phase 4 — Independence & identity  *(only the remainder — hosting already done)*
- ✅ Hosting/repo independence (done above).
- ⏳ Drop the `upstream` remote once harvest is complete.
- ❓ **Final rebrand name** (human decision). `ts-xlsx` / `@shbernal/ts-xlsx` is a
  **provisional** working name only.
- ⏳ First-class docs from types; publish a `0.x` release under the chosen name.

---

## Open decisions ❓
1. **Merge-first vs corpus-only for open PRs** (drives Phase 2 scope). Options: (a) freeze the
   legacy tree and actually merge every clean/light-conflict PR before modernizing — max value,
   longer on legacy; (b) corpus-only per the original `STRATEGY.md` wording — cleaner, but
   re-implement everything and lose the authors' patches/review; (c) hybrid by value.
   *Recommended: (a) or (c). Not yet chosen.*
2. **Final brand name** — reserved for the human (Phase 4).

## Manual follow-ups (outside this tool's reach)
- Delete the old GitHub fork `shbernal/exceljs` (needs `delete_repo` token scope). Safe anytime —
  the harvest reads upstream `exceljs/exceljs`, not the fork.

## 🔜 Immediate next action
The end-to-end format proof is **done**, and the harvest toolchain + agent skills are built.
Next slice, in order:
1. **Fill the queue**: `npm run harvest:all` (resumable one-time bulk fetch of all 794 items).
2. **Drain it** per the `harvest-triage` skill: cluster, distill credible bugs into corpus cases
   and proposals into spec notes, delete each record, commit. Track with `npm run harvest:status`.
3. **CI skeleton** (final Phase 0 bullet): a workflow running `npm run test:unit` + `npm run corpus`
   on push/PR — additive checks only, no toolchain rip-out.
