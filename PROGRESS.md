# Progress

> **Live execution tracker.** This file records *where we are* and *what remains*.
> It is subordinate to the two authorities it references — keep them the source of truth:
> - [`CLAUDE.md`](CLAUDE.md) — the constitution (principles).
> - [`STRATEGY.md`](STRATEGY.md) — the authoritative phased plan (Phases 0–4).
>
> When a phase's status changes, update this file **and** `STRATEGY.md` in the same breath.
> Legend: ✅ done · 🔜 next · ⏳ pending · 🧊 deferred-on-purpose · ❓ open decision.

_Last updated: 2026-07-10 (labeled clusters + nine unlabeled slices; 226/794; fixture-less bulk drain underway)._

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
  - **Unlabeled bulk — third slice (15) drained** (138/794, ~17%). 11 corpus cases + 4
    spec notes + 3 not-carried. Reader robustness: legacy .xls silent-empty, malformed-VML
    abort, missing-Company crash, missing-`r` implied position, 1900-epoch date serials
    (all known-open); empty comments read as blank notes + merges preserved/overlap-rejected
    (locks). Passthrough: pivot+slicer parts dropped, x14 cross-sheet list validation
    dropped (known-open). Spec notes: phantom row-value slot, merge-registration O(n²),
    in-cell rich-value images, streaming many-sheet drop. Not-carried: two broken HTML
    "fixtures" (1437, 2791 — downloads saved as error pages) and environmental non-defects
    (row-height OS scaling, prod-build corruption). Adapter grew: `readFixtureCells` (+note),
    `roundtripFixturePackageParts` (+slicers), spec `merges`, mutateWorksheet `mergeCells`.
    Corpus **109 green / 59 known-open / 0 regressions**; 71 corpus cases + 43 spec notes.
    GOTCHA banked: several attachments are broken HTML downloads (check magic bytes); a
    slow/hanging read (30k merges, whole-column validation) stays a spec note, never a case.
  - **Unlabeled bulk — fourth slice (15) drained** (153/794, ~19%). 12 corpus cases + 3
    folds (no net new spec notes). Validations: reference-based list source (defined name +
    cross-sheet range) and custom COUNTIF round-trips (locks). Types/formulas: escaped-literal-`m`
    format misread as Date, string-formula result under a date format written as `NaN`, streaming
    locale-keyed built-in date id degrades to a serial — folded into the streaming-date case
    (all known-open). Streaming: real sheet names dropped for positional placeholders (known-open),
    blank-row `row.number` gaps preserved (lock). Images: fractional-anchor sub-cell offset scales
    with real column width / row height (lock), single-cell-anchored picture emits a zeroed spPr
    transform that detaches it in LibreOffice (known-open). Styles: `duplicateValues` conditional
    format dropped on round-trip (known-open), themed/indexed fill+border colors survive (lock).
    Layout: multiple print areas per sheet lost on read AND write (known-open), row-splice/insert
    strands a merged range at its old indices (known-open). Folds: template-chart preservation →
    existing passthrough spec (+printerSettings); full-column-validation slow read (~7.5s) →
    existing bounded-memory spec. Adapter grew: `readFixtureValidationRules`,
    `roundtripFixtureValidationXml` (+per-rule facts), `roundtripFixtureCellXml`,
    `streamVsEagerSheetNames`/`streamVsEagerRowNumbers`, `inspectImageAnchors` (+spPr transform),
    `roundtripFixtureConditionalFormatting`, `roundtripFixtureColorFidelity`,
    `roundtripFixturePrintAreas`/`writePrintAreaDefinedName`, mutateWorksheet (+merges report).
    Corpus **127 green / 72 known-open / 0 regressions**; 83 corpus cases + 43 spec notes.
    GOTCHA banked: a mangled write can still pass a weak count check (two print areas → one range +
    a bare `A12` token) — assert range *shape* (must contain `:`), not just comma-split count.
  - **Fifth slice — the last 13 attachment-bearing records — drained** (166/794, ~21%).
    4 corpus cases + 1 added behavior + 1 new spec note + 2 folds + 6 not-carried. Foreign-file
    reads: Lark export throws on a mixed empty-`<t/>`+richtext shared string (known-open, Excel
    -normalized copy loads = control lock); CJK built-in date numFmt ids 57/31 read as bare
    numbers with empty format code (known-open); a `#fragment` in a hyperlink's `location` attr is
    dropped on read (known-open, added to the existing fragment case). Streaming: hidden-row flag
    lost (`hidden="true"` string form, known-open); worksheet-part-before-`workbook.xml` ZIP order
    reads fine (lock). New spec note: VML-comment read **hangs** (>30s, hostile-input termination).
    Folds: huge-defined-names-table **OOM** (~2.85 MB / 35k names exhausts 900 MB) → bounded-memory
    spec. Not-carried: Electron devtools quirk, old dep-audit scan, and four already-banked dups
    (omitted-`r`, legacy-`.xls`, richtext-hyperlink, streaming-names). Adapter grew:
    `readFixtureHyperlinks`, `readFixtureCells` (+numFmt), `streamVsEagerRowHidden`,
    `streamReadReport`. Corpus **131 green / 78 known-open / 0 regressions**; 87 corpus cases + 44
    spec notes. **The attachment-bearing queue is now exhausted** — all 628 remaining records have
    no promoted fixture (design discussions, feature requests, repro-less bug reports), so the next
    slices skew toward spec notes and reasoned not-carried, with corpus cases only where a behavior
    is reproducible from a spec-built workbook.
  - **Sixth slice — first fixture-less bulk slice, top-15 by comment signal — drained** (181/794,
    ~23%). 5 corpus cases + 5 spec notes + 5 not-carried. Corpus cases (all authored from
    spec-built workbooks, no fixture): list-type data validation round-trips both value-source forms
    — inline literal `"Male,Female"` and cross-sheet range `Levels!$A$2:$A$9999` (lock); per-cell
    protection — an *unlocked* cell survives as `locked=false` with `applyProtection` in cellXfs and
    `sheet.protect()` emits `<sheetProtection>` (lock; note `locked=true` collapses to the OOXML
    default); `insertRow`/`duplicateRow` strand a merged range at its old indices (known-open,
    distinct code path from the spliceRows case); streaming commit over a caller-supplied
    PassThrough/Duplex resolves and yields a valid package (lock — the reported cloud-SDK hang does
    not repro with a standard writable); streaming writer has **no image parity** (`sheet.addImage`
    absent, known-open). Adapter grew: `authorListValidations`, `authorCellProtection`,
    `streamCommitReport`, `streamWriterImageSupport`, plus `insertRow`/`duplicateRow` mutate ops.
    Spec notes: nested-property column keys; declarative nested column headers; streaming writer
    incremental HTTP delivery (latency face of backpressure); streaming reader must resolve styles
    before cells regardless of ZIP order + not stall (date behavior already a case); foreign
    read-modify-write must stay a valid package. Not-carried: column `eachCell`/style usage
    question, unzipper dep bump (zip layer being replaced), Firefox polyfill symptom, browser
    `createWriteStream` crash (→ browser-safe-io-boundary), Node-typed `.d.ts` (→
    public-types-node-stream-portability). Corpus **140 green / 82 known-open / 0 regressions**; 92
    corpus cases + 49 spec notes.
  - **Seventh slice — fixture-less bulk, top-15 by comment signal — drained** (196/794,
    ~25%). 3 corpus behaviors carried + 6 spec notes + 6 not-carried (folds now dominate). Corpus:
    the streaming writer produces a **valid zip**, not merely valid XML — every part present and
    non-empty, every entry's CRC matches its bytes, re-reads to the same sheet/values (lock — the
    reported zero-byte/bad-CRC corruption does not repro, via a new `streamWritePackageReport` that
    treats the writer's own output as an untrusted archive); reading an `autoFilter` carrying filter
    criteria (`filters/filter` value list + `customFilters/customFilter` comparison) does not throw
    and loads every cell via buffered and streaming reader (lock — historic "Unexpected xml node:
    filter" crash gone, backed by a hand-built foreign-shaped fixture); a sheet with both a note and
    a table must emit `legacyDrawing` **before** `tableParts` per CT_Worksheet — writer emits them
    reversed (known-open), caught by augmenting the existing `comment-and-table-coexist` case with a
    new `elementOrder` fact on `inspectPackage` (the tolerant reader hid it from the round-trip
    behavior). Spec notes: JS-Date timezone projection (configurable UTC/local/zone rule); document
    default font (workbook + worksheet level); partial write-side column-definition type; atomic
    writeFile (temp+rename, no zero-byte file); browser streaming write over a WHATWG sink;
    DrawingML shape authoring (roundtrip half already a known-open case). Not-carried — six folds:
    fractional-anchor EMU offset (two existing image cases), spliceRows tail removal
    (splice-rows-removes-requested-count), NodeJS-ambient `.d.ts` (public-types spec), SSR
    not-a-constructor (esm-entrypoint spec), HTML-page build (browser specs), deprecated transitives
    (CLAUDE.md §2 principle). Corpus **146 green / 83 known-open / 0 regressions**; 94 corpus cases +
    55 spec notes.
  - **Eighth slice — fixture-less bulk, top-15 by comment signal — drained** (211/794, ~27%).
    6 corpus cases + 3 spec notes (in 2 files) + 6 not-carried folds. Corpus: header/footer
    first-/even-page variants emit their child elements but omit the gating
    `differentFirst`/`differentOddEven` flags, so apps ignore them (known-open; new `headerFooter`
    fact + spec input); cross-sheet references in a formula and a data-validation list survive with
    sheet-name casing + `!` intact (lock); modern functions (FILTER/XLOOKUP/LET) must be stored with
    the `_xlfn.` prefix or Excel drops them — writer emits the plain name (known-open), explicit
    prefix not doubled (lock); the streaming reader corrupts multi-byte UTF-8 split across a chunk
    boundary — a large **CJK** payload returns U+FFFD (known-open) while a large emoji payload
    happens to survive (this is why the earlier small probe falsely showed pass — the bug only fires
    once the shared-strings XML spans chunks); loaded cells sharing an on-disk style index alias one
    mutable style object so mutating one bleeds into siblings (known-open); copying a sheet via
    `dst.model={...src.model,name}` drops merged ranges (known-open). Adapter grew: `headerFooter`
    fact/spec-input, `streamReadSpec`, `loadMutateCellStyle`, `copyWorksheetModel`. Spec notes:
    combined worksheet clone (intra) + cross-workbook merge into one requirement; extended (x14)
    conditional-formatting expression rules must round-trip formula + extended dxf. Not-carried
    folds: missing-sheetViews (→ worksheet-always-emits-sheetview), whole-column-DV OOM (→ its
    bounded-memory spec, which is deliberately a note because the read hangs), columns `push`
    TypeError (→ columns-mutable-array-ergonomics), image `{col,row}`/ext anchor (verified works,
    no defect), IE11 RegExp (dead runtime), tmp CVE dep bump (CLAUDE.md §2). Corpus **153 green /
    91 known-open / 0 regressions**; 100 corpus cases + 57 spec notes.
  - **Ninth slice — fixture-less bulk, top-15 by comment signal — drained** (226/794, ~28%).
    7 corpus behaviors (5 new cases + 2 augmentations) + 5 spec notes + 3 not-carried folds.
    Probing overturned four triage guesses: **style dedup already works** (40 identical-fill cells
    collapse to one shared style-table index, a distinct style stays separate — regression lock, not
    the reported perf bug), and **streaming row.values is already 1-based and byte-identical to the
    eager read** (lock). Genuine known-opens: a row/column delete-splice whose range reaches the last
    populated row/column silently leaves the trailing entries in place (interior splice shifts fine —
    lock); iterating a row with include-empty drops the trailing run of empty cells so positional
    reconstruction misaligns with a wider header (interior blanks surfaced — contrast lock); writing a
    modern function must not inject a spurious `@` implicit-intersection operator (lock, augmenting the
    xlfn case); spread-reassigning one loaded cell's font member bleeds into a shared sibling
    (known-open, augmenting the style-aliasing case). Table ref + part survive a load→save round-trip
    incl. empty-body/single-row shapes (lock). Adapter grew: `styleDedupReport`, `readRowCellPresence`,
    `streamVsEagerRowValues`, `roundtripSpecTableFacts`, `loadMutateCellFont`. Spec notes: streaming
    read-modify-write over a template; audit-clean dependency tree (built-ins over transitive deps);
    multiselect DV is not a native format feature; 255-char inline-list limit (warn, don't silently
    emit an invisible dropdown); apply-style-over-a-range API. Not-carried folds: two streaming
    multibyte-corruption reports (→ existing chunk-boundary case), a browser-bundle usage question.
    Corpus **165 green / 96 known-open / 0 regressions**; 105 corpus cases + 62 spec notes.
  - ⏳ **Next: continue the unlabeled bulk** (568 remaining, all fixture-less) in ~15-record
    slices, same triage-workflow → materialize loop. Ranking by comment/reaction signal; always
    check `docs/knowledge/specs/` + existing cases first — folds/dups now dominate a slice, so
    probe-then-fold is the default move. NB: size hostile-input/streaming repros realistically —
    a bug that needs data to span a chunk boundary (or a large sqref) will falsely pass a tiny
    probe. And **probe before trusting a triage "likely bug"** — several reported bugs (style
    dedup, streaming row indexing) turn out already-correct locks today.
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
Drain at **226/794 (~28%)**; **all labeled clusters + nine unlabeled slices are drained; the
attachment-bearing queue is exhausted and the fixture-less bulk drain is underway**. The full
pipeline is proven: parallel triage workflow → serial materialization → green corpus (165 green / 96
known-open / 0 regressions). CI corpus check is committed (`.github/workflows/corpus.yml`). Next
slices, in order:
1. **Continue the unlabeled bulk** (568 remaining, all fixture-less) in ~15-record slices, same
   triage-workflow → materialize loop. Attachment prioritization no longer applies (none left);
   these records are design discussions, feature requests, and repro-less bug reports. Folds now
   dominate — a slice is increasingly probe-then-fold into an existing case/spec — so a corpus case
   lands only where a fresh behavior reproduces from a spec-built (or small hand-built) fixture.
   Reuse the now-broad adapter vocabulary before adding surface; set each baseline by running
   `npm run corpus` (probe empirically); commit in coherent per-cluster batches. Always check
   `docs/knowledge/specs/` and existing cases first.
2. **Open decision #1** (merge-first vs corpus-only for the ~140 PRs) comes due before
   Phase 2; it does not block the issue drain.
