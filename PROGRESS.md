# Progress

> **Live execution tracker.** This file records *where we are* and *what remains*.
> It is subordinate to the two authorities it references — keep them the source of truth:
> - [`CLAUDE.md`](CLAUDE.md) — the constitution (principles).
> - [`STRATEGY.md`](STRATEGY.md) — the authoritative phased plan (Phases 0–4).
>
> When a phase's status changes, update this file **and** `STRATEGY.md` in the same breath.
> Legend: ✅ done · 🔜 next · ⏳ pending · 🧊 deferred-on-purpose · ❓ open decision.

_Last updated: 2026-07-09 (labeled clusters + two unlabeled slices; 123/794)._

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

### 🔄 Phase 1 — Harvest the backlog  *(one-time drain of the queue — IN PROGRESS)*
Model: `harvest:list` freezes the universe (`manifest.json`, 794 items), `harvest:all` fills
the queue (`backlog/issues/*.json`), agents **drain** it — distill each thread into durable
product, delete the record, commit. No per-item ledger; the commit message is the account of
record; durable artifacts never cite upstream numbers (they die with the fork).
- ✅ **Harvest toolchain** built (`harvest:list` / `:all` / `:status` + single-thread atom;
  shared core in `scripts/harvest/lib.mjs`). Manifest snapshot taken: 654 issues + 140 PRs.
- ✅ **Agent skills** authored: `harvest-triage` (per-item drain) and `write-corpus-case`.
- ✅ **Queue filled:** `npm run harvest:all` fetched all 793 remaining records (attachments
  gitignored as regenerable scratch; fixtures get promoted into `test/corpus/fixtures/`).
  `manifest.harvestComplete = true`; stage is now DRAIN.
- 🔄 **Draining (as of last update: 78/794, ~10%).** Method: a parallel **triage workflow**
  reads each record and returns a structured disposition (corpus_case / spec_note /
  not_carried) — *no writes/git*; the main loop **materializes** artifacts serially so
  baselines (set by running against `lib/`) and the shared adapter contract stay controlled.
  - **Bug cluster (74): ✅ fully drained.** Zero `bug`-labeled records remain in the queue.
    Landed: **35 corpus cases**, **27 spec notes**, and the not-carried removals. Corpus now
    **74 green / 20 known-open / 0 regressions**. The adapter contract in
    `test/corpus/adapters/workbook-io.mjs` grew to a broad vocabulary:
    `roundtripWorkbook` / `inspectPackage` (now also worksheet-rels + comment/VML/table
    package facts) / `tryWriteWorkbook` / `mutateWorksheet` / `readFixtureValidations` /
    `roundtripFixtureValidationXml` / `readFixtureReport` / `roundtripFixture` /
    `inspectImageAnchors` / `readFixtureImageAnchors` / `csvRead` / `csvWrite` /
    `streamWriteSheet` / `roundtripFormulas` / `roundtripTableAppend`; the `spec` gained
    `images`, `sharedFormula`, and `note` inputs. Batches this pass: **A** xlsx-io reads
    (foreign-generator namespace-prefixed/BOM tolerance — known-open; styled-template
    round-trip fidelity — locked), **B** images (fractional-anchor EMU-vs-real-width —
    known-open; string-range anchor read/write — locked), **C** csv (delimiter/date-format/
    conservative-coercion — locked; header-mode `data.map` crash — known-open), **D**
    streaming (`addRows` missing + richText shared-string collapse — known-open; per-sheet
    memory-release + shared-string ordering — spec notes), **E** formulas/tables
    (shared-formula clone translation + orphan-master legible error + comment/table
    coexistence — locked; table-append-after-reload — known-open).
  - **Labeled clusters (help-wanted / feature / proposal / enhancement / discussion /
    question): ✅ fully drained** (93/794, ~12%). 15 records dispositioned via the triage
    workflow: **3 corpus cases** — themed-workbook mutate/write validity (regression lock),
    image pixel-extent → EMU independent of source DPI (regression lock, verified against a
    pHYs-tagged image), and full-row/full-column-span defined names dropped on read+write
    (known-open, over-strict address validation; bounded-ref control passes) — **6 spec
    notes** (chart-part passthrough, column-key persistence, hyperlink default style,
    duplicateRow formula translation, sheetView booleans vs showFormulas, column-range
    accessor), and **6 not-carried** (3 support-thread misconceptions; 3 large-file
    streaming threads already banked in existing streaming/memory spec notes). Adapter
    gained `readFixtureDefinedNames` + a `definedNames` map in the roundtrip model and a
    spec-level `definedNames` input. Corpus **79 green / 23 known-open / 0 regressions**;
    38 corpus cases + 33 spec notes.
  - **Unlabeled bulk — first slice (15) drained** (108/794, ~14%). Triaged the 15
    highest-comment attachment-bearing records; all 15 became corpus cases (12) or spec
    notes (1) with 2 folded into existing notes. Landed **12 corpus cases** across three
    committed sub-batches: value/type/pagesetup + image/column write-validation (newline
    round-trip, date-as-serial, rich-text hyperlink, fitToPage — locks; image-missing-
    extension and >16384-column emission — known-open); cell-color fidelity + foreign-sheet
    tolerance + per-sheet defined-name scope (solid-fill fg/font separation, theme+tint
    fills, missing-sheetFormatPr read — locks; same-named per-sheet defined-name collision —
    known-open); table-roundtrip corruption + loaded-table rows + streaming-read date typing
    (all known-open). Adapter grew: per-image `extension` + per-sheet `pageSetup` spec inputs,
    a pageSetup fit fact + `definedNames` map in the roundtrip model, `contentTypeDefaults`
    and per-sheet `columnGroups`/`maxColumnIndex` on inspectPackage, and new capabilities
    `readFixtureDefinedNames`/`readFixtureCellStyles`/`roundtripFixtureTableXml`/
    `readFixtureTable`/`streamReadFixture`. Corpus **98 green / 33 known-open / 0
    regressions**; 52 corpus cases + 34 spec notes.
  - **Unlabeled bulk — second slice (15) drained** (123/794, ~15%). 10 more corpus cases
    + 4 spec notes + 1 not-carried across five committed sub-batches. Spec notes: ESM
    entrypoint ergonomics, browser seekable-source streaming read, header/footer-image
    authoring, whole-column-validation bounded memory (a hang, kept out of the suite on
    purpose), sheet-protection hash compatibility. Corpus cases: foreign-read tolerance
    (MiniExcel prefixed ns, absolute-path table rel, Strict-mode ISO date — all
    known-open); unmodeled-part passthrough (header/footer image VML, vector shapes, pivot
    tables — all known-open); style fidelity (column-width/pageSetup lock; custom
    indexed-color palette + DXF-numFmt "[object Object]" — known-open); autoFilter
    bounded-range (lock). Adapter grew: `readFixtureCells`, `roundtripFixturePackageParts`,
    `roundtripFixtureStyleFacts`; spec-level `autoFilter`; inspectPackage `autoFilterRef`/
    `dimensionRef`. Corpus **104 green / 46 known-open / 0 regressions**; 62 corpus cases
    + 39 spec notes.
  - ⏳ **Next: continue the unlabeled bulk** (671 remaining) in ~15-record slices, same
    triage-workflow → materialize loop, prioritizing attachment-bearing records.
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
Drain at **123/794 (~15%)**; **all labeled clusters + two unlabeled slices are drained**.
The full pipeline is proven: parallel triage workflow → serial materialization → green corpus
(104 green / 46 known-open / 0 regressions). CI corpus check is committed
(`.github/workflows/corpus.yml`). Next slices, in order:
1. **Continue the unlabeled bulk** (686 remaining) in ~15-record slices, same triage-workflow
   → materialize loop. Prioritize by **attachment presence** (a promoted fixture is a credible
   reproduction → corpus case). Reuse the now-broad adapter vocabulary before adding surface;
   set each baseline by running `npm run corpus` (probe empirically — triage guesses are often
   wrong); commit in coherent per-cluster batches.
2. **Open decision #1** (merge-first vs corpus-only for the ~140 PRs) comes due before
   Phase 2; it does not block the issue drain.
