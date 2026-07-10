# Progress

> **Live execution tracker.** This file records *where we are* and *what remains*.
> It is subordinate to the two authorities it references — keep them the source of truth:
> - [`CLAUDE.md`](CLAUDE.md) — the constitution (principles).
> - [`STRATEGY.md`](STRATEGY.md) — the authoritative phased plan (Phases 0–4).
>
> When a phase's status changes, update this file **and** `STRATEGY.md` in the same breath.
> Legend: ✅ done · 🔜 next · ⏳ pending · 🧊 deferred-on-purpose · ❓ open decision.

_Last updated: 2026-07-10 (labeled clusters + thirty-three unlabeled slices; 586/794 = 74%; fixture-less bulk drain underway)._

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
  - **Tenth slice — fixture-less bulk, top-15 by comment signal — drained** (241/794, ~30%).
    5 corpus cases + 8 spec notes + 2 not-carried. Probing again overturned three triage guesses:
    **splice already carries a shifted row's font/fill/numFmt** (the reported style-loss does not
    reproduce — regression lock), and **the reader already tolerates non-address defined names**
    (a constant/#REF!/external ref no longer aborts the load; a valid sibling name survives — lock).
    Genuine known-opens: a headerless table still emits an autoFilter (only valid over a header row)
    though it correctly zeroes the header-row count; a cell note's VML textbox omits
    `mso-fit-shape-to-text:t` so multi-line notes clip; a merge overlapping a formatted-table region
    is silently written as Excel-invalid geometry instead of being surfaced. Adapter grew:
    `inspectPackage.vml` (comment textbox styles + fit-to-text flag), and `mutateWorksheet` now
    accepts styled cell inputs + a `readStyles` list. New fixture: a workbook with a mixed
    definedNames block (valid + degenerate). Spec notes: honest write-buffer return type
    (Uint8Array, not Buffer-extends-ArrayBuffer); early row-iteration termination via (async)
    iterables; first-class row sort; WHATWG Web Streams I/O for edge runtimes; Node-only path reader
    must fail clearly; usable table handle (not a `{worksheet,table}` wrapper); .xlsm VBA
    preservation; streaming-write image embedding. Not-carried: two reports of the same transitive
    inflight CVE (subsumed by the ZIP-toolchain rewrite + audit-green gate). Corpus **177 green /
    99 known-open / 0 regressions**; 110 corpus cases + 70 spec notes.
  - **Eleventh slice — fixture-less bulk, top-15 by comment signal — drained** (256/794, ~32%).
    5 corpus cases + 1 augment + 6 spec notes (3 new + 2 augments) + 6 not-carried (3 folds + 3
    non-actionable). Probing overturned three more triage guesses into locks: **addRow by dense/
    sparse/object/mixed shapes all populate** (reported empty-array-rows doesn't reproduce);
    **a whole-column-range validation writes one range-scoped `dataValidation` without throwing**;
    **background image + cell note coexist with unique rel-ids** (reported corruption doesn't
    reproduce). Genuine known-opens: the row-outline **collapsed flag is emitted on the hidden detail
    rows instead of the summary row** (outlineLevel+hidden correct — lock); loading a comment
    worksheet whose **VML drawing target is missing throws** instead of degrading (fixture authored
    by stripping the VML part). Augmented the date-serial case: a **time-of-day under a duration
    format stays a numeric fractional serial**, not text (lock). Adapter grew: inspectPackage now
    reports per-sheet row outline attrs + `hasBackgroundPicture`; spec rows accept `outlineLevel`
    and a sheet accepts a `background`; new `roundtripRangeValidation` + `appendRowShapes`. New
    fixture: comment worksheet with a dangling VML rel. Spec notes: HTML-fragment→rich-text cell;
    streaming-writer table support (+ honest type surface); CSV multi-sheet selection (don't silently
    drop). Augments folded in: write-buffer return type gained the load side + ES2024 Buffer
    genericity; browser-safe-io-boundary gained the bundle-time no-Node-builtins requirement.
    Not-carried: streaming-addImage (dup of last slice's note), yarn-install + two more transitive
    CVEs (audit-gate subsumes), browser fs.createWriteStream (entry-split already noted). Corpus
    **191 green / 103 known-open / 0 regressions**; 115 corpus cases + 73 spec notes.
  - **Twelfth slice — fixture-less bulk, top-15 by comment signal — drained** (271/794, ~34%).
    5 corpus cases + 7 spec notes (6 new + 1 augment) + 3 not-carried. Probing overturned three more
    triage guesses into locks: **append-after-last-row round-trips clean**, **sheet-protection honors
    permissive options** (`sort:true` → `sort="0"` permitted, not the reported sort-blocked defect),
    **a complex custom accounting numFmt round-trips byte-for-byte** (the reported comma-drop was user
    error). Genuine known-opens: **interleaved repeated images** (B,A,A) resolve the third anchor onto
    the wrong image via a same-as-previous heuristic (`[B,A,B]`); **CSV coerces a 20-digit number
    through Number()** and loses precision. Routed to spec notes, not cases: a >512MB worksheet
    `RangeError: Invalid string length` (huge fixture would OOM CI → folded into bounded-memory note)
    and column-declared numFmt not reaching cells (probe showed it fails in BOTH streaming and
    buffered writers, so the streaming-only framing didn't hold → note, not a fragile case). Adapter
    grew: `authorCellProtection` now returns parsed `sheetProtectionAttrs`; new `interleavedImageAnchors`
    (resolves each anchor's media by identity) + `appendRowsAfterReload`. Spec notes: internal-hyperlink
    portability, embedded-chart read/write, image-source contract, worksheet→HTML export, published-types
    resolution, column-numFmt-reaches-cells. Not-carried: strict-CSP (dup of no-unsafe-eval note),
    async-iterable read error, Webpack+IE11 chunk-load.  Corpus **203 green / 105 known-open / 0
    regressions**; 120 corpus cases + 79 spec notes.
  - **Thirteenth slice — fixture-less bulk, top-15 by comment signal — drained** (286/794, ~36%).
    5 corpus cases + 1 augment + 3 spec notes (2 new + 1 augment) + 6 not-carried. Probing again
    overturned three triage guesses into locks: **added worksheet defaults to visible** (hidden-by-
    default doesn't reproduce), **distinct per-column numFmts stay independent** (no shared-style
    collapse via the getColumn path), and a **table dynamicFilter loads tolerantly** (the reported
    parse crash doesn't reproduce — the table filterColumn isn't strictly parsed on this load path,
    so it's a tolerant-read lock). Genuine known-opens: **spliceColumns into a shared-formula sheet
    throws** "master must exist above/left of clone" (read→rewrite round-trips fine — lock); the
    **streaming writer emits <hyperlinks> before <conditionalFormatting>**, violating the CT_Worksheet
    sequence Excel repairs. Augmented the list-validation case: inline lists serialize double-quote-
    wrapped, range refs unquoted (the form Excel is strict about). Adapter grew: worksheet visibility
    `state` in buildFrom/inspectPackage; new `sharedFormulaRoundtripAndSplice` + `streamWriteCfHyperlinkOrder`.
    New fixture: a table with an injected column-level dynamicFilter. Spec notes: unsupported-input-
    format typed error (.xls BIFF → UnsupportedFormatError, no zip-internals leak); cross-worksheet row
    copy preserving styles/formulas/merges. Augment folded in: browser-safe-io-boundary gained the
    streaming-in-browser boundary (absent-or-typed-error, not present-but-broken). Not-carried: IE10
    (dead browser), four dependency-bump/CVE reports (audit gate subsumes), streaming backpressure
    (dup of the row-commit-backpressure note). Corpus **217 green / 107 known-open / 0 regressions**;
    125 corpus cases + 81 spec notes.
  - **Fourteenth slice — fixture-less bulk, top-15 by comment signal — drained** (301/794, ~38%).
    4 corpus cases + 5 spec notes + 6 not-carried (2 genuine + 4 folds). Probing overturned FOUR
    triage guesses into locks this slice: **per-cell/column alignment doesn't leak** (no shared-style
    bleed), **a chart-graphicFrame drawing loads tolerantly** (the reported "anchors" crash doesn't
    reproduce), **a workbook missing docProps/app.xml loads fine** (company/manager left unset), and
    injected **out-of-range row numbers load without throwing** (→ kept as a spec note since a
    synthetic fixture may not represent the real SpreadsheetGear trigger). Only genuine known-open:
    **illegal table names** (spaces/apostrophe/leading digit) are written through, producing a
    repair-prompting file. Adapter grew: buildFrom columns accept `alignment`. New fixtures: a chart-
    graphicFrame drawing, a package with app.xml stripped. Spec notes: whole-file password encryption
    (CFB/MS-OFFCRYPTO), lean-zip container over native CompressionStream, pivot multi-value fields,
    foreign row-number tolerance, cell-padding-vs-indent honesty. Augment folded in: published-types
    note gained the runtime interop dimension (class-called-as-function under SSR bundlers). Not-
    carried: inflight CVE, UMD-bundle Vite/Rollup, and four folds (insert/merge → existing shift
    cases, streaming addImage → existing note, browser fs → browser-safe-io, esm-class → published-
    types). Corpus **227 green / 109 known-open / 0 regressions**; 129 corpus cases + 86 spec notes.
  - **Fifteenth slice — fixture-less bulk, top-15 by comment signal — drained** (316/794, 40%).
    4 corpus cases + 1 spec note + 10 not-carried (6 folds + 4 legacy/env). Folds now DOMINATE:
    six of the fifteen were duplicates of artifacts written in earlier slices (missing-app.xml,
    xlfn IFS, chart support, streaming backpressure, whole-column-validation memory, table-name/
    coexistence multi-report). Two genuine known-opens: the databar **gradient flag is written but
    dropped on read**, and **image rotation is dropped on a load/save round-trip**. The rest flipped
    to locks under probing (per-cell fill doesn't leak to column siblings; databar type/color/cfvo
    round-trip; a totals-row-formula table column loads without the reported filter-button crash).
    1348 (sheetPr child order) → spec note: the writer doesn't emit `<outlinePr>` from summary
    settings, so the corruption isn't reachable yet. Adapter grew: table columns accept rich
    `columnDefs` (totalsRowFunction) + `displayName`; new `authorConditionalFormatting` +
    `roundtripFixtureImageRotation`. New fixture: an image anchor with an injected rotation. Corpus
    **237 green / 111 known-open / 0 regressions**; 133 corpus cases + 87 spec notes.
  - **Sixteenth slice — fixture-less bulk, top-15 by comment signal — drained** (331/794, 42%).
    6 corpus cases + 2 augments + 1 combined spec note + 6 not-carried. Five genuine known-opens this
    slice (the richest in a while): a leading-dot image extension → `image1..png` media part, image
    lost on reload; loading a foreign file with a reserved sheet name ("History") throws; rowBreaks
    ignored on read + dropped on round-trip; a non-coercible date-validation operand serializes
    `<formula1>NaN</formula1>`; assigning one base style object to two cells then mutating one's font
    bleeds to the sibling. Locks: multi-area CF sqref survives as one rule (silent-drop doesn't
    reproduce); a real Date DV operand writes a valid serial. Augments folded two reports: numFmt
    invariant-separator fidelity (percentage/date codes verbatim) and a hyphen table name (into the
    illegal-name case). Adapter grew: `imageExtensionRoundtrip`, `roundtripFixtureRowBreaks`,
    `authorDateValidation`, `sharedBaseStyleFontMutation`, + `authorConditionalFormatting` now
    reports blockCount/sqrefs/ruleCount. New fixtures: a reserved-name sheet, an injected rowBreaks
    section. Combined spec note: public type surface must expose every runtime accessor/option
    (dataValidations manager + protection spinCount). Not-carried: four dependency/license reports
    (subsumed by supply-chain stance) + a streaming-flush backpressure dup. Corpus **245 green / 120
    known-open / 0 regressions**; 139 corpus cases + 88 spec notes.
  - **Seventeenth slice — fixture-less bulk, top-15 by comment signal — drained** (346/794, 44%).
    5 corpus cases + 1 spec note + 9 not-carried (5 folds + 4 legacy). One genuine known-open cluster:
    a row splice does NOT re-pin a table's cell range or an anchored image (both stranded at old
    coords) and duplicate table column names are accepted. The rest flipped to locks under probing: a
    clean horizontal merge opens without repair (covered cells not populated); a table column numFmt
    style applies to body cells cleanly; defaultRowHeight IS serialized (symmetric with defaultColWidth);
    explicit image editAs (twoCell/oneCell/absolute) honored. Adapter grew: inspectPackage `sheetFormat`
    fact; spec worksheet `properties`, table columnDefs `style`, image range `editAs`; new
    `spliceShiftsRefs`, `mergeCleanReport`, `tableColumnStyleReport`. Spec note: preserve drawing
    shapes (autoshapes/text boxes) on round-trip — the library has no shape model so they're dropped.
    Not-carried folds: note-in-table ordering → comment-and-table-coexist; userShapes crash → chart-
    drawing tolerance (loads fine here too); streaming memory → backpressure; .xls → unsupported-format;
    browser large-write → browser-streaming. Plus four legacy build/support/deps reports. Corpus
    **256 green / 123 known-open / 0 regressions**; 144 corpus cases + 89 spec notes.
  - **Eighteenth slice — fixture-less bulk, top-15 by comment signal — drained** (361/794, 45%).
    4 corpus cases + 1 augment + 2 spec notes + 8 not-carried (6 folds + 2 non-actionable). One
    genuine known-open: NaN/Infinity/-Infinity numeric cell values serialize as bare `<v>NaN</v>`
    tokens Excel treats as unreadable. The rest flipped to locks under probing: inserted-row cells
    stay mutable with style inheritance (the "object is not extensible" freeze doesn't reproduce); a
    slave-cell write resolves to the merge master (no stray value); a numeric-looking string stays a
    string with its trailing/leading zeros; cross-sheet reference column letters stay uppercase
    (augmented into the cross-sheet case). Spec notes: CF cellIs string-literal quoting + expression
    per-cell translation; xlsx date-detection control + opt-out. Adapter grew: `insertRowThenStyle`,
    `mergeSlaveWrite`, `nonFiniteCellReport`. Not-carried folds: splice-styles → existing case,
    x:sst namespace → miniexcel case, xlsm → macro note, pivot → pivot artifacts, streaming-http →
    incremental-http note, rowspan/colspan → html-export note. Corpus **266 green / 126 known-open /
    0 regressions**; 148 corpus cases + 91 spec notes.
  - **Nineteenth slice — fixture-less bulk, top-15 by comment signal — drained** (376/794, 47%).
    The most fold-heavy slice yet: 1 corpus case + 1 augment + 13 not-carried (7 folds + 6 deps/
    support), 0 new spec notes. Two known-opens: falsy formula results (0/false/"") are dropped on
    read by a truthiness test in the copy path; a two-cell range anchor emits a zeroed spPr extent
    (cx=0 cy=0) so strict OOXML viewers render nothing (augmented into the sppr-transform image case).
    Folds this slice all hit artifacts from earlier slices: databar CF → databar case, image editAs
    (×2) → editAs case, non-consecutive image reuse → interleaved-images case, streaming addImage →
    streaming-image note, streaming OOM → browser-streaming/memory notes, cross-sheet DV load → the
    cross-sheet validation cases (probing showed it's preserved). Six of fifteen were pure
    dependency/CVE/support reports. Adapter grew: `formulaFalsyResultReport`. Corpus **268 green /
    130 known-open / 0 regressions**; 149 corpus cases + 91 spec notes.
  - **Twentieth slice — fixture-less bulk, top-15 by comment signal — drained** (391/794, 49%).
    4 corpus cases + 3 spec notes + 8 not-carried (5 folds + 3 spec folds). Three known-opens: the
    streaming writer emits <hyperlinks> before <dataValidations> (same class as the CF/hyperlinks
    order bug); an autofilter writes no `_xlnm._FilterDatabase` defined name so LibreOffice ignores
    the filter; a CSV write with a non-matching sheetName silently emits empty output. Locks: getImages
    enumerates both anchor variants (the Excel-specific empty-result doesn't reproduce for standard
    anchors); CSV matching/default sheet selection works. Spec notes: centered image placement,
    single-worksheet emit + sheet-swap, locale-dependent builtin date-format-code reporting policy.
    Adapter grew: `streamWriteDvHyperlinkOrder`, `autoFilterDefinedNameReport`,
    `enumerateImagesAfterRoundtrip`, `csvWriteSheetSelection`. Folds hit earlier artifacts: solid-fill
    fgColor, spurious-@ formula, merged-slave value, streaming-addImage, edge-runtime, enumerate-merges,
    esm-bundle. Corpus **276 green / 134 known-open / 0 regressions**; 153 corpus cases + 94 spec notes.
  - **Twenty-first slice — fixture-less bulk, top-15 by comment signal — drained** (406/794, 51%).
    7 corpus cases + 3 spec notes (2 new + 1 augment) + 5 not-carried. A rich known-open haul: an
    unstyled cell resolves to no default font; a per-cell border mutation bleeds to style-record
    siblings (border facet of the aliasing family); a content-less hidden row loses its hidden flag;
    a streaming add-row after commit throws an internal null crash not a legible error; a column
    splice strands a merge to the right. Locks: table cell edit round-trips valid, column border
    stays scoped, hidden-row-with-height survives. Spec notes: formula-recalc contract (not a formula
    engine), rich-text writer robustness; augmented the drawing-shapes note with the addShape
    authoring API. Adapter grew: unstyledCellFontReport, loadMutateCellBorder, hiddenEmptyRowReport,
    streamAddRowAfterCommit, tableCellEditRoundtrip, columnBorderScopedReport. Corpus **285 green /
    141 known-open / 0 regressions**; 160 corpus cases + 96 spec notes.
  - **Twenty-second slice — fixture-less bulk, top-15 by comment signal — drained** (421/794, 53%).
    8 corpus cases + 1 spec note + 6 not-carried (4 folds + 2 non-reproducible). A big known-open
    haul from the write-side/validation veins: table column name emits raw CR/LF (unescaped); an
    internal '#' hyperlink gets a spurious external relationship; useSharedStrings=false has no
    effect; a DV formula keeps its leading '='; a whole-type DV with a cell-reference operand reads
    back null; an x14 extLst CF crashes the writer on re-write. Locks: duplicateRow copies faithfully
    + merge-after-dup succeeds; streaming commit rejects on a bad destination (doesn't hang). Two
    fixtures (DV reference operands, x14 extLst CF). Spec note: image accessibility (alt-text +
    decorative). Adapter grew by 7 capabilities. Folds: comment auto-size → comment-note-box case,
    print areas → print-areas case, append-to-loaded-table → its case. Corpus **296 green / 148
    known-open / 0 regressions**; 168 corpus cases + 97 spec notes.
  - **Twenty-third slice — fixture-less bulk, top-15 by signal — drained** (436/794, 55%).
    3 corpus cases + 2 spec notes (+ 2 folds into existing type/style notes) + 8 not-carried. A
    lock-heavy slice: all three corpus candidates probed as already-correct and became regression
    locks — the master cell of a merged region keeps its border/numFmt/font through the merge and
    round-trip; a streaming read→write style copy preserves font/fill/numFmt and emits a loadable
    styles part; single and eight-concurrent streaming reads all fully resolve shared strings (no
    dropped entries, no hang). New spec notes: serialization must be iterative not per-element
    recursive (a large export must not throw "Maximum call stack size exceeded"); the streaming
    reader's style-caching option governs date interpretation (numFmt lives in the style table) so
    its default must not hand back plausible-but-wrong serial numbers. Type-surface note grew by two
    runtime members (streaming writer `stream`, `Range.forEachAddress`); set-style-over-range grew
    the address-iteration primitive. Adapter grew by 3 capabilities (mergeMasterBorderReport,
    streamingStyleCopyReport, streamingSharedStringsRead). Not-carried: eval/CSP + global-polyfill +
    es5-toolchain reports fold into their bundle-hygiene specs; table-repair and image-editAs fold
    into existing cases; multiselect-dropdown folds into its spec; dead reports (frontend download,
    `new Buffer` deprecation). Corpus **304 green / 148 known-open / 0 regressions**; 171 corpus
    cases + 99 spec notes.
  - **Twenty-fourth slice — fixture-less bulk, top-15 by signal — drained** (451/794, 57%).
    3 corpus cases + 1 augment + 3 spec notes + 8 not-carried. A balanced write-side haul: three
    new known-opens each paired with a lock — a structured-OBJECT numFmt serializes to
    formatCode="[object Object]" and corrupts the package (a valid string numFmt survives fine);
    a CSV with non-ASCII text is written as correct UTF-8 bytes but with NO byte-order mark, so
    Excel mis-renders it; the streaming writer keeps a master formula but drops shared-formula
    slave cells so they reload empty. Augmented alignment-flags with a numeric-indent lock (the
    OS-specific indent-loss report doesn't reproduce). Spec notes: indexed-palette colors must
    resolve to a concrete ARGB while remembering their origin; AutoFilter should support hiding the
    dropdown button per column (additive, no current behavior — filterColumn/hiddenButton absent);
    worksheet enumeration must tolerate foreign generators (files reading back as zero worksheets,
    cured only by resaving in Excel — awaits a real fixture). Adapter grew by 3 capabilities
    (numFmtObjectCorruptionReport, csvNonAsciiEncodingReport, streamingSharedFormulaReport).
    Not-carried folds: framework-ESM-build → esm-package-entrypoint; charts → embedded-chart;
    hf-images → header-footer-image-authoring; autofilter-sort (SheetJS sample) → filter-database
    case; loaded-font-mutation bleed → shared-style-aliasing case; hf-text → headerfooter case;
    plus a callbacks-as-API eachSheet (discarded by design) and an empty "Fills error" report.
    Corpus **308 green / 151 known-open / 0 regressions**; 174 corpus cases + 102 spec notes.
  - **Twenty-fifth slice — fixture-less bulk, top-15 by signal — drained** (466/794, 59%).
    3 corpus cases + 2 spec notes + 10 not-carried — a dependency/release-heavy slice. Two locks +
    one known-open: writing several adjacent equivalent columns collapses them into shared <col>
    spans without the reported "column.equivalentTo is not a function" crash (lock); a formula whose
    cached result is a date serial under a date format reads back as a valid Date, not Invalid Date
    (lock); a worksheet referencing a drawing part that doesn't resolve crashes the whole load with
    an internal null-dereference on `.anchors` (known-open, fixture: an image workbook with its
    drawing part removed but the rel/element left intact — the missing-drawing sibling of the
    missing-VML case). Spec notes: the streaming writer emits a corrupt ZIP at very large scale (a
    ZIP64 offset/size boundary defect — kept a note not a case because it needs millions of rows to
    manifest), and the in-cell vs floating image distinction (Excel "Place in Cell" rich-value
    images vs today's floating anchored drawings). Adapter grew by 2 capabilities
    (equivalentColumnCollapseReport, formulaDateResultReport). Not-carried was dominated by
    dependency-tree advisories (minimatch, node-tmp, inflight) and release logistics (publish
    request, "when new version") — all superseded by the clean-break dependency policy — plus
    support questions (column formula fill-down, HTML-into-cell, a content-free postMessage security
    report, a zero-length-buffer zip error) and a font-aliasing report that folded into
    per-cell-font-isolation. Corpus **312 green / 153 known-open / 0 regressions**; 177 corpus cases
    + 104 spec notes.
  - **Twenty-sixth slice — fixture-less bulk, top-15 by signal — drained** (481/794, 61%).
    3 corpus cases + 1 augment + 1 spec-note augment + 10 not-carried — a fold-heavy slice. Two
    locks + one known-open + an xlfn augment: workbook-level structure protection (lockStructure) is
    dropped on a read→write round-trip (known-open; worksheet-level protection survives, workbook
    doesn't — ExcelJS has no workbook-protection API at all); an unbordered cell reads back with no
    border sides and a one-sided border doesn't sprout the other three (lock — reported phantom
    borders don't reproduce); looping many sheets each with a table + DV produces a valid package
    with unique table part ids and surviving validations (lock — the reported protected-view/erased
    -DV failure doesn't reproduce). Augmented the xlfn case with a LET/LAMBDA/BYROW behavior (modern
    functions written verbatim without the _xlfn./_xlpm. prefixes Excel requires — known-open, same
    class as FILTER). Extended the merged-cell scaling spec note with the efficient-lookup
    requirement (a PR that indexed merge membership). Adapter: normalizeCell + buildFrom now carry
    per-cell `border`; new capabilities workbookProtectionRoundtrip, multiSheetTableReport.
    Not-carried was dominated by folds (full-column DV → DV-single-sqref case; streaming cell-type
    → streaming-shared-strings case; csv-write → csv-by-name case; foreign-file compat → foreign
    -generator specs) plus dependency/env/doc noise (deprecated transitive deps, RN Promise
    polyfill, server bundling, a Chinese-doc fix). Corpus **317 green / 156 known-open / 0
    regressions**; 180 corpus cases + 104 spec notes.
  - **Twenty-seventh slice — fixture-less bulk, top-15 by signal — drained** (496/794, 62%).
    1 corpus case + 2 augments + 1 spec note + 11 not-carried — very fold-heavy. One known-open + two
    locks: a conditional-formatting rule's stopIfTrue flag is silently dropped on write (known-open —
    layered-rule precedence lost). Augments: an image with no explicit editAs defaults to
    editAs="oneCell" in the singular spelling (lock, pinning the ambiguous-doc default); a
    date-looking string under a date-formatted column stays a string — a number format is a display
    instruction not a coercion, so pivotable dates need Date values (lock). Spec note: data-bar
    colors (color/negativeFillColor) + opt-in targeted "ignored errors" write support (rejecting the
    blanket-suppression hack). Adapter grew conditionalFormattingStopIfTrue. Not-carried was mostly
    already-locked behaviors that didn't reproduce (indent, strict-date, INDIRECT dropdown,
    addBackgroundImage, numeric-value-as-text) plus folds (merged-row discovery → merged-cell note,
    DataValidationType → public-type-surface) and dep/help noise. Corpus **319 green / 158 known-open
    / 0 regressions**; 181 corpus cases + 105 spec notes.
  - **Twenty-eighth slice — fixture-less bulk, top-15 by signal — drained** (511/794, 64%).
    1 corpus case + 3 spec notes + 2 augments + 9 not-carried — the most fold-heavy slice yet.
    One genuine bug locked: a whole-column (or whole-row) print area is a valid column-only OOXML
    reference ($A:$D) that the writer emits correctly, but the READER decodes each endpoint as a cell
    address and, finding no row number, injects NaN — surfacing the print area back as the corrupt
    string "ANaN:DNaN" (known-open; write + bounded-range control are locks). Adapter grew
    printAreaRoundtrip. Spec notes: streaming writer exposes sheet views as a getter-only property so
    a frozen split assigned after addWorksheet throws (no parity with the buffered writer, which
    honors both) — reclassified from corpus to spec after probing showed a TypeError, not a silent
    drop; streaming date reads ignore the workbook 1904 base and can fail under cached/compact styles
    (two root causes the existing streaming-date known-open does not exercise); the date-system flag
    must not be a required field on the formula-cell authoring type. Augments: native/intrinsic-size
    image placement (oneCellAnchor with ext from decoded pixels) onto the in-cell-vs-floating note;
    outlineProperties (summaryBelow/summaryRight) type-surface requirement onto the sheetPr note.
    Not-carried was five folds (spliceRows tail-clearing ×2 existing cases; keyed-addRows + header
    loss → column-key-roundtrip-persistence; default font → default-font-workbook-worksheet-level;
    csv single-added-sheet → csv-write-sheet-selection-by-name) plus four support/scope records
    (PDF/print, high-level header how-to, legacy-build packaging complaint, no-repro selectbox).
    Corpus **321 green / 159 known-open / 0 regressions**; 182 corpus cases + 108 spec notes.
  - **Twenty-ninth slice — fixture-less bulk, top-15 by signal — drained** (526/794, 66%).
    1 corpus case + 2 spec notes + 2 augments + 10 not-carried — fold-dominated, six of the
    not-carried were exact folds into existing artifacts. Corpus: the CSV writer silently ignores a
    requested output encoding and always emits UTF-8, so a caller asking for another encoding gets
    mojibake (known-open); multibyte emoji/CJK fidelity under the default UTF-8 path is locked green
    (adapter grew csvWriteEncodingReport). New spec notes: a first-class "time" data-validation type
    parallel to "date" (day-fraction serial bounds, full operator set, public type union adds
    "time"); rich-text alignment composition (alignment is cell-level, runs carry only character
    formatting, so per-run alignment is a no-op/type-error). Augment: browser-safe-io-boundary grew a
    "no unguarded Node-global (process/Buffer) access on a hot path" section (an env-detection helper
    reading process.version made mere in-memory API use throw in the browser). Not-carried folds:
    internal "#" hyperlink dropped (→ internal-location-hyperlink-not-external-rel); column-numfmt
    reaching cells (→ column-declared-numfmt-reaches-cells, repro was boilerplate); partial column
    definitions (→ column-definition-type-is-partial-on-write); corrupt-zip/non-atomic write (→
    atomic-writefile-no-partial-output); workbook structure protection (→
    workbook-structure-protection-survives-roundtrip case); read/decrypt encrypted workbook (→
    workbook-password-encryption). Five support/scope records: IE11 unicode-RegExp, Jest/Windows
    hang, Electron writeBuffer error, legacy-UMD Promise.finally clobber, wrap-text how-to.
    Corpus **322 green / 160 known-open / 0 regressions**; 183 corpus cases + 110 spec notes.
  - **Thirtieth slice — fixture-less bulk, top-15 by signal — drained** (541/794, 68%).
    2 corpus cases + 1 spec note + 12 not-carried (nine exact folds + three support/noise) — the
    most fold-heavy split yet. Both new cases are write-side preservation defects, each locked
    read-works/write-drops: (1) a What-If-Analysis data-table formula (`<f t="dataTable">` with
    input-cell refs) is recognized on read (shareType dataTable + range + cached result) but the
    writer drops `t="dataTable"` on a read-modify-write, so the cell stops recalculating (known-open;
    adapter grew dataTableFormulaRoundtrip via in-adapter XML injection, no fixture file); (2)
    fullCalcOnLoad set on the streaming writer never reaches the output calcPr while the in-memory
    writer emits it (known-open, in-memory + flag-unset as passing controls; adapter grew
    streamingFullCalcOnLoadReport). Spec note: inline table-column width in the column definition
    (ergonomics, equivalent to the sheet-column width). Nine folds: table+comment corruption (→
    comment-and-table-coexist-on-same-sheet), multi-section numFmt (→
    custom-numfmt-string-roundtrips-verbatim, already a 4-section format), string-not-coerced (→
    numeric-looking-string-preserved-as-string), browser F_OK (→ browser-safe-io-boundary, literally
    its title), whole-column DV (→ whole-column-data-validation-bounded-memory), import interop (→
    esm-package-entrypoint-ergonomics), form controls (→ form-controls-roundtrip-preserved), indexed
    color (→ indexed-palette-colors-resolve-to-concrete), fractional-image-resized (→
    image-anchor-fractional-offset-respects-cell-size). Three support/noise: name-mismatch worksheet
    lookup, a thank-you post, addRow on EOL Node 6. Corpus **325 green / 162 known-open / 0
    regressions**; 185 corpus cases + 111 spec notes.
  - **Thirty-first slice — fixture-less bulk, top-15 by signal — drained** (556/794, 70%).
    2 corpus cases + 1 augment + 2 spec notes + 10 not-carried (eight folds + two support). New
    known-opens: (1) no first-class note removal — note=null throws, undefined/empty-object leave the
    comment part and VML drawing in the package so a "removed" note still shows a marker (adapter grew
    removeCellNoteReport); (2) a row from an array built in a foreign realm (Node vm context)
    populates no cells because array detection is realm-bound identity, not structural, though
    Array.isArray recognizes it (augmented add-row-array-and-object-shapes-populate; adapter grew
    crossRealmArrayRow). New lock: CSV export sizes each row by the sheet max column extent, so a wide
    row after a narrow first row keeps all its fields (reported truncation doesn't reproduce). Two
    reclassify-on-probe: CSV row width was already correct → lock not known-open; multi-letter
    currency numFmt (CHF) round-trips both quoted and unquoted through the library, so no round-trip
    bug — the corruption is Excel-side and auto-quoting is undecided → spec note not corpus. Other
    spec note: template placeholder replacement. Eight folds: sheet-scoped defined names (→
    same-named-defined-names-scoped-per-sheet), defaultRowHeight (→
    worksheet-default-row-height-applied-on-write), table rows/ref on read (→ loaded-table-exposes-
    data-rows), charts dropped (→ chart-parts-survive-template-roundtrip), chart authoring (→
    embedded-chart-read-write), large-file memory (→ bounded-memory-large-workbook-read), table row
    insert (→ splice-rows-updates-table-and-image-refs), image via writeFile (→ image survival
    cases). Corpus **329 green / 164 known-open / 0 regressions**; 187 corpus cases + 113 spec notes.
  - **Thirty-second slice — fixture-less bulk, top-15 by signal — drained** (571/794, 72%).
    4 corpus cases + 2 spec notes + 9 not-carried (four folds + five noise) — an unusually
    corpus-rich slice, all four confirmed by probing. One lock + three known-opens: default theme
    part is emitted for a no-theme workbook (lock, regression guard against a corruption class); a
    CSS-style "#"-prefixed fill ARGB is written verbatim as a malformed 9-char rgb that renders black
    (known-open); table style theme "None" emits a bogus name="None" instead of an unstyled table
    (known-open); the boolean font-flag parser ignores the val attribute so <b val="0"/> reads as
    bold=true (known-open). Adapter grew fillArgbHashPrefixReport, tableStyleThemeReport,
    fontExplicitFalseBoldReport (styles-injection); default-theme reused inspectPackage hasThemePart.
    Spec notes: rich-text shared-string reader robustness (a foreign empty-string accumulator crashes
    the reader — library reads its own fine, so a foreign fixture is needed) and block row
    duplication (copy N contiguous rows M times). Four folds: cross-sheet list validation (→ three
    existing validation cases), CSV headers-option crash (→ csv-read-with-header-rows, same upstream
    ref), table roundtrip corruption (→ existing-table-roundtrip-fidelity), inline-list length cap (→
    list-validation-inline-formula-length-limit). Five noise: Excel-not-imported, minimist dep vuln,
    React StrictMode double-invoke, jszip release request, IF how-to. Corpus **337 green / 167
    known-open / 0 regressions**; 191 corpus cases + 115 spec notes.
  - **Thirty-third slice — fixture-less bulk, top-15 by signal — drained** (586/794, 74%).
    2 corpus cases + 2 spec notes + 1 augment + 10 not-carried (four folds + six noise). One lock +
    two known-opens: worksheet outline summary-position (summaryBelow/summaryRight) serializes into
    <outlinePr> and round-trips (lock — probing resolved a stale assumption in
    sheetpr-child-order-outline-before-pagesetup that outlinePr was never emitted; corrected that
    note); inserting a row above a noted/outlined row drops the cell note entirely and leaves the
    outline level pinned to the old absolute row index (two known-opens). Adapter grew
    outlinePropertiesRoundtrip and rowInsertPreservesNoteAndOutline. Spec notes: streaming-reader
    temp-file leak on early for-await abort (2147+2148 same bug — real leak, but tmpdir-counting is
    flaky and a corpus case would pollute the CI temp dir, so a note with the .return()/finally
    mechanism); CSV write-to-stream incremental flush + backpressure. Four folds: Vietnamese CSV
    diacritics (→ csv-write-honors-requested-encoding multibyte fidelity), fill-mutate-after-load
    read-only (→ loadMutateCellStyle, doesn't reproduce), stream+sharedStrings valid (→ streaming
    valid-package cases), image merged-cell stretch (→ image-embedded-in-cell-vs-floating-anchor).
    Six noise: non-repro getRow, UI-tool discovery, screenshot-only reports, misfiled Office Scripts
    bug, placeholder XYZ, buffer-load screenshot. Corpus **340 green / 169 known-open / 0
    regressions**; 193 corpus cases + 117 spec notes.
  - ⏳ **Next: continue the unlabeled bulk** (208 remaining, all fixture-less) in ~15-record
    slices, same triage-workflow → materialize loop. Ranking by comment/reaction signal; always
    check `docs/knowledge/specs/` + existing cases first — folds/dups now dominate a slice, so
    probe-then-fold is the default move. NB: size hostile-input/streaming repros realistically —
    a bug that needs data to span a chunk boundary (or a large sqref) will falsely pass a tiny
    probe, and a genuinely-huge repro (>512MB worksheet) belongs in a spec note, never a corpus
    fixture. And **probe before trusting a triage "likely bug"** — reported bugs (style dedup,
    streaming row indexing, splice style-loss, non-address defined-name crash, empty-array addRow,
    whole-column validation, background+note rel collision, sort-blocked protection, numFmt comma-
    drop) repeatedly turn out already-correct locks today. Late-slice spec notes increasingly FOLD
    into notes written a slice or two earlier (streaming image, browser bundling, buffer types, CSP,
    large-file memory) — augment the earlier note + not-carry. When a triage's "streaming-specific"
    framing is probed and the defect ALSO reproduces in the buffered path, prefer a spec note over a
    mislabeled case.
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
Drain at **586/794 (74%)**; **all labeled clusters + thirty-three unlabeled slices are drained; the
attachment-bearing queue is exhausted and the fixture-less bulk drain is underway**. The full
pipeline is proven: parallel triage workflow → serial materialization → green corpus (340 green / 169
known-open / 0 regressions). CI corpus check is committed (`.github/workflows/corpus.yml`). Next
slices, in order:
1. **Continue the unlabeled bulk** (208 remaining, all fixture-less) in ~15-record slices, same
   triage-workflow → materialize loop. Attachment prioritization no longer applies (none left);
   these records are design discussions, feature requests, and repro-less bug reports. Folds now
   dominate — a slice is increasingly probe-then-fold into an existing case/spec — so a corpus case
   lands only where a fresh behavior reproduces from a spec-built (or small hand-built) fixture.
   Reuse the now-broad adapter vocabulary before adding surface; set each baseline by running
   `npm run corpus` (probe empirically); commit in coherent per-cluster batches. Always check
   `docs/knowledge/specs/` and existing cases first.
2. **Open decision #1** (merge-first vs corpus-only for the ~140 PRs) comes due before
   Phase 2; it does not block the issue drain.
