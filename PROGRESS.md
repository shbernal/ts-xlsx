# Progress

> **Live execution tracker.** This file records *where we are* and *what remains*.
> It is subordinate to the two authorities it references — keep them the source of truth:
> - [`CLAUDE.md`](CLAUDE.md) — the constitution (principles).
> - [`STRATEGY.md`](STRATEGY.md) — the authoritative phased plan (Phases 0–4).
>
> When a phase's status changes, update this file **and** `STRATEGY.md` in the same breath.
> Legend: ✅ done · 🔜 next · ⏳ pending · 🧊 deferred-on-purpose · ❓ open decision.

_Last updated: 2026-07-13 (**Phase 3 rebuild underway** — `src/core/address.ts` + the core in-memory model (`value`/`cell`/`worksheet`/`workbook`) + the buffered `.xlsx` **writer** and now a buffered **reader** (`src/io/xlsx/`, on `fflate`; XML read via a hand-written SAX parser, ADR 0004) green vs `--adapter rewrite` at 155 green / 0 regressions (styles round-trip pattern fills, number formats, fonts, borders, **alignment, and per-cell protection** through an interned, composed style table, with dedup exposed to the corpus — alignment and protection are xf *children*, interned into the xf signature and emitted as body elements in schema order; the SAX reader now does XML §2.11 end-of-line normalization; **now also sheet-level protection** — `sheet.protect(password, options)` emits `<sheetProtection>` with author-facing allow-flags inverted to OOXML "forbidden" booleans, guarded by an OOXML-agile SHA-512 password credential derived via `node:crypto`; **and copy-on-write style aliasing** — loaded cells that shared a style record expose independent facet objects, so a single-cell facet edit never bleeds into a sibling, hard-locked by `style-isolation.test.ts`; **and column-scope inheritance of every style facet** — a column now defaults a fill/font/border/alignment/protection to its cells, not just a number format, scoped to its own column; **and the foreign-fixture reading family (now fully closed)** — the reader proves itself against real non-Excel input: theme+tint/indexed fill & border colours survive a pure open-then-save, the workbook default font resolves onto every unstyled cell, a real styled template survives a no-op round-trip, foreign explicit-off font/alignment forms are honoured (incl. the `<u val="none">` → falsy fix), and ~two dozen real fixtures with namespace-prefixed roots / BOMs / unusual part order / missing optional parts read without crashing; **the merged-cells family is fully closed** — slave→master addressing, overlap rejection, and lossless `model` export/import; **and the structural-edit family** — `spliceRows`/`spliceColumns`/`insertRow`/`duplicateRow` shift cells with their styles, row/column metadata, and merged ranges; **and cell notes** — a per-cell `note` facet written as the legacy comments part (`comments{n}.xml` + a VML drawing + `<legacyDrawing>` + rels + content-types) and read back through the sheet's own relationships; **and anchored images** — workbook-wide media (`workbook.addImage` → id; `sheet.addImage(id, {tl, br})`) written as `xl/media/` bytes + a DrawingML `xl/drawings/drawing{n}.xml` (two-cell anchors) + rels + `<drawing>` element + content-types and read back with bytes intact, with tables and image anchors now shifting through a row/column splice (and duplicate table column names rejected at construction), which closes the last structural-edit case; **and reading `<sheetProtection>` back** — the reader now parses worksheet-level protection into the model (agile credential preserved verbatim, flags un-inverted), closing the round-trip on a passthrough save; **and structured date values** — a `Date` serialises to a 1900-system serial (phantom-leap-day quirk reproduced) under a date number format and reads back as a `Date`, an Invalid Date is written value-less rather than dropping siblings, and the reader also handles Strict-mode `t="d"` ISO cells and locale-specific built-in date format ids; **and a worksheet tab colour** — a `Worksheet.tabColor` (ARGB/theme) writes `<sheetPr><tabColor>` and reads back verbatim, with an uncoloured sheet never fabricating one; **and worksheet outline summary positions** — `Worksheet.outline` (`summaryBelow`/`summaryRight`) folds an `<outlinePr>` into the same `<sheetPr>` (after `<tabColor>` per CT_SheetPr order), emitting only the flags the author set and reading them back; **and worksheet page setup** — a `Worksheet.pageSetup` (`fitToPage`/`fitToWidth`/`fitToHeight`/`scale`/`orientation`/`pageOrder`) splits across `<pageSetUpPr>` (the fit-to-page flag, inside `<sheetPr>` after `<outlinePr>`) and `<pageSetup>` (between `<pageMargins>` and `<headerFooter>`), emitting only the attributes the author set and reading them back — verified against a real Excel fixture for both fractional column widths and print-scaling attributes, including `paperSize` (Excel's 1-based paper index, the last live `<pageSetup>` attribute); **and colour-string normalization** — a CSS-habit `#`-prefixed ARGB (`#FFBFBFBF`) is stripped to a bare 8-hex value at `colorAttrs`, the single choke point every fill/font/border/tab colour serializes through, so no colour is ever emitted as the 9-character value strict consumers render black (bare and foreign-lowercase values pass through verbatim, keeping round-trip fidelity intact); **and formatted-but-empty cells** — a cell given only a style (a fill/border, no value) is now serialised as a valueless `<c r=.. s=../>` instead of being dropped by the value-only row filter, and the reader finalises a self-closing `<c/>` from its `open` handler (it fires no `close`), so a style-only cell survives a round-trip while a cell with neither value nor style is still never fabricated; **and stricter ARGB validation** — the same `normalizeArgb` choke point now promotes a 6-hex RGB (a colour written without its alpha channel) to a fully-opaque 8-hex ARGB and rejects a value that is neither 6 nor 8 hex digits at the API surface, rather than emitting a colour Excel silently renders black; **and the modern-function `_xlfn.` prefix** — a function added after OOXML froze its formula grammar (FILTER, XLOOKUP, LET, LAMBDA, IFS, …) is stored in the sheet XML under the required `_xlfn.` name-mangling prefix so current Excel accepts it, applied by a shared `src/core/formula.ts` primitive that skips legacy and already-prefixed names and never touches a string literal or injects an `@` operator, and stripped back to the plain name on read so the model only ever holds the readable form — **now including the dotted 2010 statistical rename family** (NORM.DIST, T.DIST.2T, …), where the tokenizer treats `.` as part of a function name so the whole dotted name is prefixed once (`_xlfn.NORM.DIST`) rather than its trailing segment, **and the bare-name 2010/2013 additions** (SEC/CSC/COT, the BITAND & IMxxx families, AGGREGATE, XOR, ISOWEEKNUM, GAMMA/GAUSS/PHI, …), which need no tokenizer work — only accurate membership — so the prefix follows a function's vintage not its spelling and a pre-2007 cousin (SIN, GAMMALN) is left bare, at 203 green / 0 regressions); independent Microsoft 365 OOXML validation added as a required CI oracle; Phase 1 harvest complete at 245 cases + 150 spec notes)._

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
  - **Thirty-fourth slice — fixture-less bulk, top-15 by signal — drained** (601/794, 76%).
    4 corpus cases + 1 spec note + 10 not-carried (four folds + six noise) — a lock-heavy slice: all
    four corpus cases are regression-guard LOCKS (the reported bugs do not reproduce on current
    code). frozen top row emits <pane ySplit="1" state="frozen"/> and round-trips; a tab color given
    as 8-digit ARGB (alpha first) round-trips verbatim with no spurious color on uncolored sheets;
    images anchored to single cells with interleaved addRow resolve one-to-one with no off-by-one
    drift; a 5-column table reads back all five columns (no 3-column cap). Adapter grew
    frozenTopRowRoundtrip, tabColorRoundtrip, cellAnchoredImagePositionReport,
    wideTableColumnReadReport. Spec note: carrying images between workbooks must copy the media +
    rels or fail loudly (dangling drawing rel → broken image). Four folds: numeric-looking string
    quote-prefix (→ numeric-looking-string-preserved-as-string), richText shared-string collision (→
    streaming-write-richtext-shared-strings-distinct), CSP eval shim (→ no-unsafe-eval-csp-compatible,
    already existed), rowCount after clearing (→ worksheet-row-count-reflects-data, by-design). Six
    noise: multi-sheet how-to, no-repro protection, MIME user-side, TS decl packaging, no-repro
    getCell, HTTP transport. Corpus **349 green / 169 known-open / 0 regressions**; 197 corpus cases
    + 118 spec notes.
  - **Thirty-fifth slice — fixture-less bulk, top-15 by signal — drained** (616/794, 78%).
    6 corpus cases/augments + 3 spec notes + 2 augments + 4 not-carried (all folds). Four locks +
    two known-opens on the corpus side: csv `map` option controls value coercion (default coerces
    "007"→7, an identity map preserves strings — 2 locks); a merged child cell's display text
    mirrors the master and never throws (lock); out-of-order `<col>` tags still bind width/hidden to
    the right column index (lock); row and column outline levels round-trip on both axes (lock);
    streaming write emits `sheetProtection` before `autoFilter` (reload control locked, ordering
    known-open); loading vs constructing a reserved "History" sheet split apart — the guard
    over-applies to API construction (known-open) while genuinely-invalid names stay rejected (lock).
    Adapter grew csvReadMapReport, mergedCellDisplayTextReport, outOfOrderColumnsReport,
    rowColumnOutlineLevelRoundtrip, streamAutoFilterProtectionOrder, addReservedSheetNameReport.
    Spec notes: streaming writer's worksheet is a narrowed forward-only type (random-access ops must
    be absent, not throw); streaming and buffered I/O should be one implementation with parity as an
    invariant; a cell should report its merge role (master/child/none). Augments: validate the image
    definition at entry (no-payload/bad-extension/format-mismatch → add-image-source-contract); a
    runtime value export like the range class needs a value-level declaration
    (→ public-type-surface-matches-runtime). Four folds: inefficient merge-conflict check
    (→ merged-cell-registration-linear-not-quadratic), multiple print areas
    (→ multiple-print-areas-one-sheet-roundtrip), fill-after-border style bleed
    (→ cell-border-mutation-does-not-bleed-to-style-siblings), generic slow-writeBuffer report
    (no repro, Phase 3 perf). Corpus **359 green / 171 known-open / 0 regressions**; 202 corpus cases
    + 121 spec notes.
  - **Thirty-sixth slice — fixture-less bulk, top-15 by signal — drained** (631/794, 79%).
    0 new corpus cases + 4 spec notes + 2 augments + 9 not-carried — the most fold-dominated slice
    yet: every one of the three triage "corpus_case" candidates collapsed on probing. Empty-table-rows
    rejection was *rejected* — we already lock the opposite, better contract (an empty-body table
    writes a VALID full-header-ref file, see empty-table-writes-valid-file), so the reject-at-add PR is
    superseded, not adopted. Image editAs was an exact fold into image-range-anchor-edit-as-mode-honored
    (the in-cell editAs-undefined case is an Excel/Numbers app limitation, not a lib bug). Falsy formula
    results were an exact fold into formula-cell-preserves-falsy-result-values. New spec notes: table
    AutoFilter should be an explicit toggle independent of the header row (the corruption path is
    already a case); authoring API for workbook structure protection (the round-trip is already a
    case); explicit comment-box sizing (auto-fit is already a case); streaming reader must surface
    merged-cell ranges. Augments: streaming reader `name` and Row/Column `.values` array typing
    (→ public-type-surface-matches-runtime + sheet-values-sparse-array-return-type). Not-carried also
    folded web-native streams (→ web-streams-io-surface), the TS5 question, and three dependency/Node
    reports (→ minimal-audit-clean-dependency-tree + a modern runtime baseline — the whole reason we
    forked). Corpus unchanged at **359 green / 171 known-open / 0 regressions**; 202 corpus cases
    + 125 spec notes.
  - **Thirty-seventh slice — fixture-less bulk, top-15 by signal — drained** (646/794, 81%).
    2 new corpus cases + 2 spec notes + 11 not-carried. The two corpus cases came from *reframing*
    a triage disposition against the real code: (1) the `skipStyles` "crash" was a red herring —
    that option isn't consumed anywhere in the reader; the durable bug is that a worksheet `<col>`
    carrying a styleId with **no styles part** crashes reconcile on `undefined.getStyleModel(...)`,
    so the case is a hostile-input tolerance lock backed by a styles-stripped fixture (baseline
    fail). (2) spliceColumns *insertion* mode actually works and round-trips cleanly, so it's locked
    as a regression guard (baseline pass) for the previously-untested splice direction. Two triage
    "corpus_case" candidates collapsed on probing: duplicateRow→unmerge→re-merge "already merged"
    does **not** reproduce on current code (the phantom-merge invariant is already locked and
    passing), and the splice-columns merged-cell PR is an exact fold into the existing
    baseline-fail re-anchor lock. New spec notes: dynamic-array (spill) formulas must not be
    downgraded to legacy CSE array formulas on write (a model gap — the writer wraps them in braces);
    pivot per-field aggregation metrics (Count and the rest via `dataField/@subtotal`), answering the
    multi-value-fields note's per-measure open question. Not-carried folded pageSetUpPr order
    (→ sheetpr-child-order), table-name validation (exact case fold), spliceRows reorder
    (→ splice-rows-removes-requested-count), shallow-model sheet copy (→ worksheet-clone), drawings
    dropped (→ preserve-drawing-shapes-on-roundtrip), and two dependency reports
    (→ minimal-audit-clean-dependency-tree). Corpus **360 green / 173 known-open / 0 regressions**;
    204 corpus cases + 127 spec notes.
  - **Thirty-eighth slice — fixture-less bulk, top-15 by signal — drained** (661/794, 83%).
    4 new corpus cases + 3 new spec notes + 2 fold-augments + 6 not-carried. This slice broke the
    late-drain fold pattern: probing four triage "corpus_case" calls against the real reader/writer
    confirmed all four as genuine, reproducing bugs (no reframes needed) — (1) an alignment element
    carrying only an explicit-false boolean (`wrapText="0"`/`shrinkToFit="0"`) reads back as an
    `{ wrapText:false }` object instead of no alignment, because the raw `"0"` attribute string is
    truthy in JS; (2) `getWorksheet` matches case-sensitively while `addWorksheet` rejects duplicates
    case-insensitively, so a name reported absent by lookup throws on add; (3) an in-workbook
    `#Sheet2!A1` hyperlink is written as an *external*-mode relationship AND a `location`, doubling
    the target (`#Sheet2!A1##Sheet2!A1`) in strict consumers — an assertable, serialized upgrade of
    the `internal-hyperlink-target-portability` spec (whose defect was cross-app navigation the corpus
    couldn't exercise); (4) a comments part at a non-canonical OPC path (reachable only via the rels)
    crashes reconcile on `undefined.comments` because parts are located by filename glob, not
    relationship type. Four new adapter capabilities added. New spec notes: defined names may hold
    arbitrary formulas (INDEX/MATCH), not only ranges, wired to data-validation list sources; pivot
    page fields (report filters, `axisPage` + `<pageFields>`); the library's streams must honor the
    Node pipe contract (`pipe` returns the destination — `StreamBuf.pipe` returns `undefined`,
    breaking `.pipe(dest).on('finish')` and `stream.pipeline`). Fold-augments: extended-CF copy path
    drops x14 just like a raw round-trip; large-write XML-fragment accumulation (3 array entries per
    cell) hits a hard array-size ceiling — must flush incrementally. Not-carried: JSZip-not-a-
    constructor (bundler/ESM dep-resolution env issue, and JSZip is being replaced), streaming
    addTable (already in the streaming-table spec), a vague .NET-SDK-interop meta-report (no concrete
    construct; dups foreign-generator corpus), a "wrong defaults" fix with no diff/repro/detail in the
    record, a foreign-file streaming regression the record says already reads on master (no fixture),
    and streaming multibyte-UTF-8 chunk-boundary (already locked by an existing corpus case). Corpus
    **362 green / 180 known-open / 0 regressions**; 208 corpus cases + 130 spec notes.
  - **Thirty-ninth slice — fixture-less bulk, oldest-15 of the tail — drained** (676/794, 85%).
    Signal has gone uniformly zero across the remaining tail, so ordering shifted to oldest-first.
    2 new corpus cases + 1 new spec note + 1 umbrella-augment (3 records) + 9 not-carried — the fold
    rate is now dominant. Probing reshaped triage heavily: of six "corpus_case" calls, only two held.
    (1) The streaming writer exposes its own output stream (`writer.stream`, a public property), but
    that stream's `pipe(dest)` returns `undefined` instead of the destination, breaking
    `writer.stream.pipe(out).on('finish', …)` and `stream.pipeline` — a control confirms the bytes
    still flow, isolating the defect to the return value. This is the assertable *upgrade* of last
    slice's `stream-pipe-returns-destination` spec (a duplicate report surfaced here). (2) A table
    declaring a *calculated column* (`<calculatedColumnFormula>` on the column definition, as Excel
    emits) crashes the reader — it loses its place at the nested element, truncates the column list,
    then dereferences a missing column during autoFilter reconciliation
    (`…setting 'filterButton'`); fixture authored by injecting the element into a real table part.
    New spec note: the image-anchor *authoring* surface should accept rotation/extent/offset (rotation
    round-trips on read today but can't be *set* via `addImage`, so rebuilding a sheet through the
    public API flattens placement). Umbrella-augment into `public-type-surface-matches-runtime`
    (three type-only PRs): streaming worksheet-reader `id`/`name`/`state` (runtime populates all three
    — confirmed a hidden sheet reports `state:'hidden'`); optional `WorkbookReader` options argument;
    CSV read/write options types matching the documented shape. Not-carried (9): lint-only cleanup;
    date1904-optional and Row.values type fixes (already fully captured by the formula-value and
    sheet-values type specs); multibyte chunk-boundary (existing case); `outlineProperties`
    summaryBelow (existing outline case already flags the type-surface gap); regenerator/CSP (existing
    no-unsafe-eval spec); sheet-scoped whole-row defined name `'Sheet'!15:15` dropped (SAME over-strict
    address validation already locked by `defined-name-full-row-column-span` — probe: `15:15` dropped,
    `$A$15:$D$15` kept); Firefox `constructor` permission-denied (app/polyfill Xray, no isolable
    library defect); and dataValidation list formula >255 chars — **does not reproduce**: no length
    guard exists in current code (a 351-char inline list writes and round-trips cleanly), and the
    long-list→defined-name workaround is already a case. Two new adapter capabilities. Corpus
    **363 green / 183 known-open / 0 regressions**; 210 corpus cases + 131 spec notes.
  - **Fortieth slice — fixture-less bulk, oldest-15 of the tail — drained** (691/794, 87%).
    4 new corpus cases + 2 augmented cases + 1 new spec note + 3 spec-augments + 3 not-carried folds.
    New cases: (1) a digit-only JS string persists as a *text* cell and a number as a *numeric* cell
    — the library preserves the declared type and never coerces a zero-padded code (baseline pass,
    a type-fidelity regression lock; the "number stored as text" advisory is the honest consequence,
    not a bug). (2) Unfreezing a frozen view by replacing it with a normal view emits valid
    sheetViews with no leftover `<pane>` and reloads as state 'normal' (baseline pass) — new adapter
    `unfreezeViewRoundtrip`. (3) A foreign `core.xml` that binds the core-properties namespace as the
    *default* xmlns (so `lastModifiedBy`/`lastPrinted` are unprefixed) must load and read its
    last-modified-by — today the reader throws `Unexpected xml node in parseOpen` (baseline fail),
    fixture authored by injecting the default-namespace core part. (4) After a row splice, `lastRow`
    must resolve to the last *populated* row, not the trailing empty slot the delete leaves behind
    (baseline fail; empirically the delete shifts data up but keeps the `_rows` length, so `lastRow`
    reads an emptied slot — a no-delete control passes). Augments: whole-row/column protection carries
    `locked=false` to its band's cells while off-band cells stay default-locked (baseline pass,
    `authorCellProtection` extended); the 1900 phantom-leap boundary — serial 59 must read 1900-02-28
    (lib reads it one day early, baseline fail) and serial 61 reads 1900-03-01 across the phantom day
    (baseline pass), added onto the existing serials fixture. Spec: new
    `data-validation-message-length-limits` (over-limit prompt/error/title corrupts the file — 255
    body / 32 title — validate at authoring, distinct from the inline list-formula 255 limit). Spec
    augments: `public-type-surface-matches-runtime` items 9 (`Row.dimensions` is a `{min,max}`
    column-span, not a number — folds 2 records) and 10 (media `type` is a discriminated union, not
    an open string); `databar-colors-and-ignored-errors-write` gains the x14 negative-color child
    order (fill before border, or the package corrupts); `bounded-memory-large-workbook-read` gains
    the single large-AREA defined name (filter-database over ~140M cells → OOM; must round-trip by
    corners) and the eager `writeFile` O(document-size) buffering. Not-carried folds (3): reading
    page breaks is already pinned by `read-manual-row-page-breaks` (probe: write emits breaks, read
    returns `[]`); the fractional image-anchor-over-custom-width offset by
    `image-anchor-fractional-offset-respects-cell-size`; the non-streaming writeFile OOM by the new
    bounded-memory write-path bullet. Triage's 2621 "corpus_case" was correctly re-routed to a spec
    (OOM repro would kill CI). Corpus **373 green / 188 known-open / 0 regressions**; 214 cases + 132
    specs. Two adapter extensions (`mutateWorksheet.lastRow`, `readFixtureReport.lastModifiedBy`) plus
    the new `unfreezeViewRoundtrip` capability.
  - **Forty-first slice — fixture-less bulk, oldest-15 of the tail — drained** (706/794, 89%).
    2 new corpus cases + 1 new spec note + 3 spec-augments + 9 not-carried folds — the fold-heavy
    tail continues. New cases: (1) a list dropdown whose source range lives on *another* sheet is
    serialized by spreadsheet apps in the 2009 `x14` data-validation extension (worksheet `extLst`),
    not the plain `<dataValidation>`; the reader understands only the standard form, so the cross-sheet
    rule is silently dropped — the cell reports no validation and a read→write round-trip loses the
    dropdown (baseline fail; same-sheet standard validation is the passing control; fixture built by
    injecting the x14 extension onto `Sheet1!A1` → `Sheet2!$D$3:$D$5`). (2) Anchoring a floating image
    over a cell range advanced the row-append cursor, so a subsequent `addRows` appended *below* the
    anchored range instead of filling from the top and the layout depended on add order (baseline fail;
    new `imageAnchorRowAppendReport` capability reports the first data cell for both orders). New spec
    `xlsb-binary-format-output` (read/write binary .xlsb as a second codec over the shared model;
    OPC/ZIP + BIFF12 records, RK/Ptg encodings, read-before-write, bounded-allocation posture). Spec
    augments: `public-type-surface-matches-runtime` — `protect()` typed for every call shape it already
    accepts (no-args / password-only / password+options, all returning a Promise; the hand-written
    types wrongly required a password); `streaming-read-emits-all-worksheets` — foreign (openpyxl)
    files lost their sheet names (bind names via the relationship graph) and a stream-end race dropped
    trailing rows/sheets (completion must mean input truly exhausted; guard in a soak harness);
    `image-carry-between-workbooks-requires-media-registration` — recorded today's opaque
    `Cannot read properties of undefined (reading 'name')` crash on cross-workbook `model` transplant
    (merge loss on the same path already cased). Not-carried folds (9): webpack `process` polyfill
    (charter precludes the class), the x14 databar negative-color reorder (legacy emits no negative
    colors → latent, already in `databar-colors-…`), the databar-color `.d.ts` one-liner, an empty
    template, the JSZip `new Function` CSP patch (`no-unsafe-eval-…`), an OOM report with no repro
    (`bounded-memory-…`), a package-rename PR (reserved human decision), a "protect one of two sheets"
    usage question (`sheet-protection-…`), and library suggestions (sort → `sort-rows-by-column`;
    column move user-acknowledged trivial). Corpus **376 green / 192 known-open / 0 regressions**;
    216 cases + 133 specs. One adapter extension (`imageAnchorRowAppendReport`).
  - **Forty-second slice — fixture-less bulk, oldest-15 of the tail — drained** (721/794, 91%).
    7 new corpus cases + 1 case augment + 3 spec-augments + 3 not-carried folds. Probing flipped
    most triage guesses: only 3 of the batch reproduced as *open* bugs, the rest are already-correct
    behavior locked as regression guards. Real bugs (baseline fail): (1) a rich-text run with an
    **empty text string** serializes to an empty `<t>` element that Excel flags as corrupt — must be
    dropped, surrounding runs kept (`rich-text-empty-substring-run`, new `richTextRoundtripReport`);
    (2) a **whitespace-only CSV field** coerces to numeric `0` via `Number('   ')===0` instead of
    staying a string (`csv-whitespace-only-cell-preserved-as-string`, reuses `csvRead`); (3) explicit-
    off **font toggles** (`<i val="0"/>`, `<strike val="0"/>`, `<u val="none"/>`) read as enabled —
    the presence-based-parsing gotcha, extending the existing bold-only case with italic/strike/
    underline via new `fontExplicitOffFlagsReport`. Regression locks (correct today, guarding silent-
    corruption classes): cell `col`/`row` are numeric 1-based indices at runtime; leading rich-text
    run formatting is position-independent; a row-level fill stays scoped to its row (no whole-sheet
    bleed — extended `buildFrom` with `row.fill`); a merge into a trailing empty final row is iterated
    and resolves to its master; an image added to a *loaded* (not fresh) worksheet persists on
    re-serialize. Spec augments: `xlsx-date-detection-control` — a literal string cell stays a string
    even when it reads like a date, and date-vs-string classification is uniform (the reported
    per-column "inconsistency" is expected, not a reader bug); `cell-full-address-descriptor-numeric-
    row-col` — the decoded-`Address` descriptor type must match the runtime field-for-field (numeric
    row/col, real optionality for the sheet qualifier and `$`-markers), enforced by a type-level test;
    `image-embedded-in-cell-vs-floating-anchor` — reading floating images is a first-class enumerable
    affordance (media + anchor kind + offset) that works for foreign files. Not-carried (3):
    dependency-version pinning (policy tied to upstream's dep tree, already answered by the fork's
    dependency-hygiene stance); a COUNTIF duplicate-detection support question whose repro targets
    another library; browser `writeFile` "fs.createWriteStream is not a function" (already covered by
    `browser-safe-io-boundary` — path methods are Node-only and must fail actionably). Corpus **392
    green / 198 known-open / 0 regressions**; 223 cases + 133 specs. Five adapter capabilities added.
  - **Forty-third slice — fixture-less bulk, oldest-15 of the tail — drained** (736/794, 93%).
    5 new corpus cases + 3 new spec notes + 1 spec-augment + 5 not-carried. Four new adapter
    capabilities. Probing was decisive again — every corpus-case guess reproduced, and it flipped
    2777 from "bug is dropped mode" to the true schema violation. Real bugs (baseline fail): (1)
    **editAs on a one-cell image anchor** — a top-left+extent placement makes a `oneCellAnchor`, onto
    which legacy stamps `editAs` even though the drawing schema defines that attribute only on
    `twoCellAnchor`; a top-left+bottom-right placement round-trips `editAs` on a two-cell anchor,
    locked (`image-editas-only-valid-on-two-cell-anchor`, reuses `inspectImageAnchors`); (2) **cell
    style setters mutate the shared/deduplicated style object in place**, bleeding alignment/numFmt/
    protection into style-sharing siblings — extends the fill/font/border copy-on-write family to the
    remaining facets via parametric `loadMutateCellFacet`
    (`cell-style-setter-isolates-alignment-numfmt-protection`); (3) the **streaming reader loses the
    tail of a many-sheet workbook** — at ~180 sheets a worksheet part is parsed before the workbook
    model is built and the reader throws `this.model.sheets` undefined; a 3-sheet workbook streams
    clean (`streaming-read-emits-all-worksheets-at-scale`, new `streamReadManySheets` reads from a
    scratch file — only the file-path input reproduces it, `Readable.from(buffer)` crashes even at 5
    sheets); (4) the **streaming reader drops the hidden flag on columns** (and rows), reporting all
    visible — column companion to the existing row case (`streaming-reader-preserves-hidden-column`,
    new `streamVsEagerColumnHidden`); (5) the **quotePrefix cell-format flag** (force literal text) is
    neither written nor preserved (`quote-prefix-cell-flag-roundtrip`, new `quotePrefixReport`). Spec
    notes: `worksheet-paper-size-type-and-custom-dimensions` (new — paperSize type omits A3/other OOXML
    codes + custom paperWidth/paperHeight, folds two records); `data-validation-scope-on-row-insert`
    (new — separates the app's UI-side inheritance from the library's own sqref-shift policy);
    `cell-value-raw-and-displayed-accessor` (new — unified raw/displayed accessor + kind inspection);
    `image-embedded-in-cell-vs-floating-anchor` augmented with a query-images-by-anchor-cell bullet.
    Not-carried (5): import-a-subset/package-too-large (ESM tree-shaking + minimal dep tree already the
    stance); unfilled `[BUG] XYZ` template (trivial A1=7 round-trip); `update-dependency-version` and a
    transitive `async` 3.2.5→3.2.6 bump (fork sheds the legacy dep tree); a style-cache uninitialized
    defensive patch (code-internal guard on a deleted file; style-part-absent robustness already
    locked). Corpus **400 green / 210 known-open / 0 regressions**; 228 cases + 136 specs.
  - **Forty-fourth slice — fixture-less bulk, oldest-15 of the tail — drained** (751/794, 95%).
    3 new corpus cases + 4 new spec notes + 8 not-carried (5 of them folds into existing
    artifacts). Two new adapter capabilities. Fold pressure now dominates: 3 of the 5 triaged
    "corpus_case" candidates were already locked and became not-carried after the mandatory
    check-existing-first sweep. Real bugs (baseline fail): (1) **named-style (cellStyleXfs) fill
    dropped** — a cell whose fill lives only in the named-style layer (a `cellXfs` `xfId` into
    `cellStyleXfs`) reads back with `pattern:'none'` and the whole non-default `cellStyleXfs` layer
    collapses to count=1 on write, so formatting supplied through named cell styles is lost on both
    read and save (`cellstylexfs-named-style-fill-roundtrip`, new `namedStyleFillReport`; crafted
    fixture where A1's yellow lives only in `cellStyleXfs[1]` via `xfId=1`); (2) **empty comment-rel
    Target crashes the read** — a worksheet comments relationship with `Target=""` throws
    `TypeError …reading 'comments'` during reconcile, a *distinct* path from the already-locked
    missing-VML-drawing case (`worksheet-comment-rel-empty-target-tolerated`, reuses
    `readFixtureReport`; crafted fixture). Regression lock (baseline pass): (3) a **list validation
    with a cross-sheet source range applied per-cell down many rows does NOT drift** — every cell
    keeps `Lookup!A1:A5` verbatim and the identical rules collapse to a single `sqref`; the reported
    "shrinking dropdown" is a desktop-app relative-reference effect, not ours
    (`list-validation-source-range-stable-across-rows`, new `listValidationSourceRangeAcrossRows`).
    Spec notes: `worksheet-default-cell-protection-unlock` (express unlocked-by-default in the
    default cell format so unlocking N cells costs zero per-cell deviations);
    `native-iteration-protocol` (`Symbol.iterator`/`asyncIterator` on collections; sparse-vs-dense
    stays explicit because `[Symbol.iterator]` takes no options); `streaming-csv-row-reader`
    (bounded-memory async-iterable CSV read, the counterpart to the existing CSV write stream);
    `style-dedup-value-based-and-cell-add-style` (value-based dedup as the single default —
    identity caching measured ~5× slower than none — collision-free key as a correctness invariant,
    plus merge-style-onto-cell respecting inheritance). Not-carried folds: image rotation on export
    (already in `image-anchor-authoring-rotation-extent-offset`); streaming writer embeds images
    (already `streaming-writer-image-parity` + `streaming-write-add-image`); boolean font `val="0"`
    (already `font-boolean-flag-honors-explicit-false`). Not-carried noise: Android WeChat addImage
    (third-party runtime, reporter's own scaling math a no-op); iOS Safari download UX (browser Blob
    plumbing); unzipper dup bump + two CI-matrix records (upstream toolchain the fork replaces).
    Corpus **404 green / 215 known-open / 0 regressions**; 231 cases + 140 specs.
  - **Forty-fifth slice — fixture-less bulk, oldest-15 of the tail — drained** (766/794, 96%).
    A bug-HEAVY slice, breaking the recent fold-dominated pattern: 7 new corpus cases + 3 new spec
    notes + 1 spec-augment + 4 not-carried (3 folds). Five new adapter capabilities. Probing decided
    disposition repeatedly: 3 of the 7 corpus cases are regression LOCKS (baseline pass) where the
    reported bug is already correct in the fork's lib, and 3 triaged "corpus_case" candidates were
    already locked and became not-carried. Real bugs (baseline fail): (1) a **conditional-formatting
    rule with no formula crashes serialization** — the writer indexes into an absent formula list and
    throws `TypeError …reading '0'` (`conditional-format-rule-without-formula`, reuses
    `authorConditionalFormatting`); (2) an **explicit column width equal to the conventional default
    (9) is silently dropped** — no `<col>` emitted, reads back `undefined`, while every other width
    survives; the writer treats "equals the magic default" as "skip"
    (`explicit-column-width-equal-to-default-magic-value-survives`, new
    `columnWidthDefaultCollisionReport`); (3) a **dirty image extension** (a URL-derived
    `png?alt=media&token=…`) leaks straight into the content-type `Default` Extension/ContentType,
    producing an invalid package — distinct from the missing-extension case (non-empty but malformed)
    (`image-dirty-extension-sanitized-in-content-type`, reuses `inspectPackage`); (4) **no way to
    remove an added image** — `ws.removeImage` is absent (`worksheet-image-removal`, new
    `removeImageReport`, mirrors the streaming-image-parity known-open shape). Regression locks
    (baseline pass, guarding silent-loss / environment-failure classes): (5) **password protection
    works under Node** — `ws.protect(pw)` resolves with a valid SHA-512 hash/salt/spin protection and
    real salt randomness (two protects differ), NOT the reported browser-only-secure-random throw
    (`worksheet-password-protection-hashes-in-node`, new `worksheetPasswordProtectionReport`); (6) a
    **table displayName round-trips** to the table XML and back, distinct from the internal name — the
    reported typo bug is absent (`table-display-name-roundtrips`, new `tableDisplayNameReport`); (7)
    **hidden/veryHidden sheet state survives a write** and veryHidden does not degrade
    (`worksheet-hidden-state-preserved-on-write`, new `worksheetStateReport`). Spec notes:
    `worksheet-default-style` (new — worksheet/workbook default style mapped to OOXML `<sheetFormatPr>`/
    `<cols>` defaulting, O(deviations) not O(cells); the umbrella over the default-font and
    default-protection notes); `indexing-convention-accessor-naming` (new — 1-based accessors must not
    be named `index`; the type surface carries the convention); `range-to-image-render` (new — render a
    cell range to a styled SVG/PNG, dependency-light draw-to-SVG over a headless browser);
    `cell-value-raw-and-displayed-accessor` augmented with the numFmt-preserved-on-read invariant
    (folds the "numFmt not reflected in output" report). Not-carried folds: streaming shared-strings/
    sheet-names race (`streaming-read-resolves-shared-strings-without-race`); multibyte-UTF-8 chunk
    boundary (`stream-read-multibyte-utf8-chunk-boundary`); image anchor EMU with custom col/row size
    (`image-anchor-fractional-offset-respects-cell-size`). Not-carried noise: a repro-less "buffer load
    crashes with forEach of undefined". Corpus **416 green / 222 known-open / 0 regressions**; 238
    cases + 143 specs.
  - **Forty-sixth slice — fixture-less bulk, oldest-15 of the tail — drained** (781/794, 98%).
    Back to a fold-dominated slice: 3 new corpus cases + 3 new spec notes + 1 spec-augment + 8
    not-carried (7 folds + 1 noise). Two new adapter capabilities + one crafted fixture. Real bugs
    (baseline fail): (1) **duplicate table column names are emitted verbatim** — supplying colliding
    column names writes `name="foo"` three times, which OOXML forbids (Excel repairs on open); the
    writer must disambiguate to a unique set (`table-duplicate-column-names-disambiguated`, new
    `tableDuplicateColumnNamesReport`; a distinct-names control passes). NB the library's own reader
    reloads the corrupt file fine, so the durable assertion is written-name uniqueness, not reload.
    (2) an **invalid Date under a date numFmt serializes `<v>NaN</v>`** — the date→serial conversion
    runs because of the format and NaN corrupts the cell; string/null under the same format are
    unaffected (`date-numfmt-nonnumeric-value-serializes-valid-xml`, new `dateNumFmtValueReport`;
    NaN-leak baseline fail, string/null/Invalid-Date-literal controls pass). (3) a **table autoFilter
    filterColumn colId pointing outside the declared columns crashes the read** — `Cannot set
    properties of undefined (setting 'filterButton')` aborts the whole load; the dangling reference
    must be tolerated (`table-filter-column-out-of-range-tolerated`, crafted fixture, reuses
    `readFixtureReport`). Spec notes: `first-worksheet-accessor-after-deletion` (new — deletion-safe
    first-sheet accessor; ids are stable and not renumbered, so by-id '1' fails after deletion;
    records the id-vs-order distinction and the empty→undefined contract); `large-workbook-write-
    performance` (new — hundreds-of-thousands-of-rows writes must be linear-time/bounded-memory via
    the streaming path; a perf/hang concern, so spec not a flaky corpus case); `modular-parser-
    entrypoints-for-tree-shaking` (new — format-agnostic core + per-format tree-shakeable entry
    points; a build/bundle concern outside the corpus); extended-CF spec **augmented** with the
    mixed/custom icon-set x14 case. Probes flipped two guesses: the "streaming drops mixed icons" bug
    is dropped by BOTH writers (missing x14 feature → spec-augment, not a streaming corpus case), and
    the out-of-range filter column is a genuine distinct crash, not a fold of the dynamicFilter
    tolerance case. Not-carried folds: streaming batch addRows (`streaming-write-add-rows-batch`);
    fractional image column offset (`fractional-image-anchor-positioning` +
    `image-anchor-fractional-offset-respects-cell-size`); table style 'None'/null theme
    (`table-style-none-produces-unstyled-table`); row/cell missing `r` attribute
    (`cells-without-r-attribute-imply-position`); table dynamicFilter node
    (`table-dynamic-filter-tolerated-on-load`); huge validation range on read
    (`whole-column-data-validation-bounded-memory` spec +
    `data-validation-whole-column-range-writes-single-sqref`); sheet-protection permission flags
    (`sheet-protection-permits-requested-operations`). Not-carried noise: an empty "[F] XYZ" feature
    template. Corpus **420 green / 226 known-open / 0 regressions**; 241 cases + 146 specs.
  - **Forty-seventh (FINAL) slice — the last 13 records — drained; the queue is EMPTY** (794/794,
    100%). 4 new corpus cases + 4 new spec notes + 5 not-carried (2 folds + 3 noise/policy). Two new
    adapter capabilities. Real bugs (baseline fail): (1) a **cell formula supplied with a leading '='
    is stored in `<f>` verbatim** (`=1+2`) and reads back with it — OOXML expects no '='; Excel
    tolerates it but Google Sheets/WPS reject the file (`formula-stored-without-leading-equals-for-
    portability`, reuses `inspectPackage`+`roundtripFormulas`; probe FLIPPED the triage's likely-pass
    guess). (2) a **minimal dataBar CF rule (type+priority only) crashes the writer** on the absent
    cfvo collection — it should default a min/max cfvo and a bar color like Excel (`databar-
    conditional-formatting-minimal-defaults`, reuses `authorConditionalFormatting`). (3) the
    **streaming reader drops merged cells** (returns null) that the buffered reader exposes
    (`streaming-reader-surfaces-merged-cells`, new `streamReadMergesReport` — now backs the existing
    `streaming-read-surfaces-merged-cells` spec with a corpus lock). (4) **pivot-cache shared-item
    strings are written raw** (a bare `&`), corrupting the pivotCacheDefinition XML; a null source
    value writes fine (`pivot-cache-escapes-xml-special-characters`, new `pivotCacheSpecialCharsReport`;
    ts-xlsx's experimental pivot writer). Spec notes: `protection-allow-sort-autofilter-on-locked-sheet`
    (granting sort/AutoFilter is insufficient — sorting rewrites locked cells, so the fix needs an
    unlocked window or `<protectedRanges>`; the flag encoding is already locked, this is the missing
    semantics); `multiple-pivot-tables-from-shared-source` (N pivot tables — 5 correctness constraints
    for when authoring lands); `pivot-table-preserve-worksheet-column-widths` (OOXML
    `applyWidthHeightFormats` as a typed boolean); `external-workbook-reference-formulas` (cross-book
    refs need an externalLink part + `TargetMode=External` rel + indexed `<externalReferences>`, not a
    literal `[file.xlsx]Sheet!A1` string). Not-carried folds: append-to-loaded-table →
    `table-loaded-from-file-accepts-appended-rows` (append already known-open there); `x:`-namespace
    prefixed XML → `miniexcel-prefixed-namespace-reads-without-crashing` + missing-props tolerance
    cases. Not-carried noise/policy: a comment typo fix, an exceljs.org impersonation alert (branding,
    no code), and an npm-audit/toolchain modernization (folds into the clean-break dependency stance).
    Corpus **428 green / 233 known-open / 0 regressions**; **245 cases + 150 specs**.
  - ✅ **Phase 1 harvest COMPLETE.** The upstream ExcelJS backlog work queue
    (`docs/knowledge/backlog/issues/`) is fully drained: all 794 frozen manifest items (654 issues +
    140 PRs) are now a corpus case, a spec note, or a reasoned not-carried. `manifest.json` remains as
    proof nothing was silently dropped. The durable product: **245 implementation-blind corpus cases**
    (regression-locking current-correct behavior and known-open bugs the rewrite must fix) + **150 spec
    notes** (Phase-3 design targets). **Next: Phase 2/3** — begin the TypeScript-first rewrite behind
    the corpus contract (a `rewrite.mjs` adapter binds the same vocabulary to new code; every existing
    case runs unchanged). Reserved for the human: open decision #1 (merge-first vs corpus-only for the
    ~140 PRs) and the final rebrand name.
- **Exit:** the queue is empty; every carried item left a corpus case and/or spec note; corpus
  runs against current code (mostly red where bugs are real). Follow via `harvest:status`.

### ✅ Phase 2 — Stabilize-to-validate  *(satisfied by the harvest — no legacy stabilization undertaken)*
- **Disposition (2026-07-11):** the exit criterion — corpus expresses agreed "correct behavior"
  per case with current-code pass/fail as the baseline — was **already met** by Phase 1. All 245
  cases carry a run-verified `baseline` against the legacy oracle (428 green / 233 known-open). Per
  the phase's own "it is scaffolding — don't invest in legacy" rule and `CLAUDE.md` §"we do not keep
  legacy code", we spend **zero** further effort on `lib/` and go straight to the rebuild. We do not
  fix bugs in a tree we are deleting.
- Open decision #1 (merge-first vs corpus-only for the ~140 PRs) is therefore **moot for
  correctness/knowledge** — the corpus already captured every PR's intent. It survives only as a
  human call about whether to salvage the authors' *patches/review credit* on the legacy tree; the
  rewrite does not depend on it. Left open below, downgraded from "gating".

### 🔄 Phase 3 — The rebuild  *(discard the debt — IN PROGRESS as of 2026-07-11)*
- **Kickoff landed:** greenfield tree under `src/` (strict TS, ESM via `src/package.json`, legacy
  CommonJS root untouched). Corpus adapter `test/corpus/adapters/rewrite.mjs` binds the same
  vocabulary to the new code; not-yet-built capabilities are **skipped (`∅`)** so the whole corpus
  runs against the partial rewrite (runner grew skip semantics + adapter-aware `↑` messaging).
- **First module: address decoding** (`src/core/address.ts`) — the col/row primitive everything
  else stands on. Resolves the legacy known-open (`decodeRange('$1:$1')` no longer leaks
  `undefined`/`NaN`; legacy still fails it → the rewrite already beats legacy). 12 native
  `node --test` unit tests.
- **Second module: the core in-memory model foundation** — `src/core/value.ts` (the typed cell-value
  model: a discriminated `CellValue` union + `detectValueType`/`coerceCellValue`, no stringly sentinels,
  no silent kind-coercion — a numeric-looking string *stays* a string), `src/core/cell.ts` (immutable
  numeric 1-based `row`/`col`, value routed through the model so `type` is always consistent),
  `src/core/worksheet.ts` (sparse row-major cell grid + A1 `getCell`), `src/core/workbook.ts`
  (worksheet ownership, Excel-faithful sheet-name validation, case-insensitive lookup + by-id). 22
  more unit tests. Lights the first **pure in-memory** corpus capability (`cellColRowTypes`) →
  `cell-col-row-are-numeric-indices` green under `--adapter rewrite`. Rows/columns/merges/defined-names
  and the styles surface (fills/borders/alignment) are follow-up slices; the round-trip-shaped cases
  stay skipped until the writer lands.
- **Third module: the buffered `.xlsx` writer** (`src/io/xlsx/`) — the first vertical slice to the
  write path. `write.ts` serialises a `Workbook` into a valid OPC package (content types, rels,
  workbook, per-sheet XML with `dimension`/`sheetViews`, the default theme + stylesheet, core/app
  props) via **`fflate`** (`zipSync`); `xml.ts` is the audited escaping surface (text + attribute,
  `xml:space="preserve"`); `static-parts.ts` ships the default Office theme/styles. Serialises
  number/string/boolean/formula cells; **refuses** a zero-sheet workbook, a non-finite number, or a
  value kind it can't yet represent (rather than emitting a lossy/corrupt package). Formula `=`
  normalisation moved into the value model (canonical OOXML stored form). Proven valid by the legacy
  ExcelJS *reader* accepting the output. Zip/XML-write dependency decision recorded in
  [`docs/decisions/0003-zip-and-xml-write-path.md`](docs/decisions/0003-zip-and-xml-write-path.md)
  (`fflate`; the *parser* choice stays deferred to the reader slice). The corpus's `inspectPackage`
  fact-extractor was factored into a shared, implementation-blind
  [`test/corpus/adapters/ooxml-facts.mjs`](test/corpus/adapters/ooxml-facts.mjs) so both adapters
  derive identical facts. The rewrite adapter **feature-gates** unsupported spec features → those
  behaviors skip (`∅`), never falsely regress.
- **Rows & columns slice landed** (2026-07-12): the model grew `getColumn`/`getRow` (metadata stored
  apart from the cell grid — a hidden column or grouped row costs no phantom cells) and
  `Worksheet.properties` (default row height / column width); the writer emits `<cols>`, per-`<row>`
  height/hidden/outline/collapsed attrs, and a real `<sheetFormatPr>`. The writer drops any column
  past XFD (16384) rather than serialize a range Excel treats as corrupt. Adapter feature-gate
  widened (`columns`/`rows`/sheet `properties`, gated at key granularity).
- **Page-margins slice landed** (2026-07-12): the model grew `Worksheet.pageMargins` (a `PageMargins`
  bag stored like `properties`); the writer emits `<pageMargins>` after `<sheetData>`, all-or-nothing —
  when any margin is set it writes all six (untouched sides fall back to Excel's Normal-preset
  defaults), and omits the element entirely when none is set. Adapter feature-gate widened with the
  `pageMargins` sheet key + per-side allowlist.
- **Header/footer slice landed** (2026-07-12): the model grew `Worksheet.headerFooter` (a
  `HeaderFooter` bag of odd/even/first header & footer strings); the writer emits `<headerFooter>`
  after `<pageMargins>` in CT_HeaderFooter child order and — crucially — derives the gating flags
  `differentOddEven="1"` / `differentFirst="1"` from which variants are present, without which Excel
  silently ignores the even/first content. Omitted when no variant is set. Adapter gate widened with
  the `headerFooter` sheet key + per-child allowlist.
- **Corpus vs rewrite reprioritization:** merges was next in the nominal writer order, but a survey of
  the merge cases showed *every* one needs the reader (`roundtripWorkbook`, `mergeCleanReport`,
  `mutateWorksheet`, model reports) or also needs `tables` — so merges lights up **zero** corpus cases
  today. The writer-only cases whose `inspectPackage` fact-extractors are already pre-wired
  (`pageMargins`, `headerFooter`, `tables`) turn known-opens green *now*, so the writer order is
  page-margins → header/footer → tables; merges/defined-names land alongside the reader that lets their
  cases assert.
- **Tables slice landed** (2026-07-12): a new `src/core/table.ts` (`Table` model — validated Excel
  identifier name, ≥1 column, derived geometry: `ref`/`autoFilterRef`/`region` computed from an anchor
  + column/row counts, not a stored range string, so an empty/headerless/totals-bearing table each
  refs correctly); `Worksheet` grew `addTable`/`tables` + minimal `mergeCells`/`merges`. The writer
  now emits a full table sub-package: a global-numbered `xl/tables/tableN.xml` part (child order
  autoFilter → tableColumns → tableStyleInfo), a per-sheet `xl/worksheets/_rels/sheetN.xml.rels`,
  `<tableParts>` (after `<headerFooter>` per CT_Worksheet), and a content-type override; it derives
  `headerRowCount="0"` + **no** autoFilter for a headerless table, drops autoFilter over the totals
  row, and — as OOXML gatekeeper — **rejects a merge that overlaps a table** (Excel-invalid geometry).
  Also emits `<mergeCells>` after `<sheetData>`. `tryWriteWorkbook` now reports `{ok:true}` on a
  successful write (no case reads cell-survival from it — that stays the reader's job). Deliberate
  deferral: the cell-reference-name-collision rule (`A1`) is **not** enforced — the corpus treats
  `T1`/`T2` as valid table names, so enforcing it would regress a baseline-pass fixture; noted in
  `table.ts`. Real-consumer check: legacy ExcelJS loads our table package and finds the table by name.
- **Reader slice landed** (2026-07-12): the writer-only vein was exhausted, so the XML **read** path
  opened — the biggest single unlock (~48 reader-dependent cases assert through read-back). ADR 0004
  settles the parser choice deferred by ADR 0003: **a lean, hand-written SAX pull parser**
  (`src/io/xlsx/xml-read.ts`), no XML dependency, single O(n) pass, entities *decoded but never
  expanded* (only the five predefined + numeric refs; DTDs/`<!ENTITY>` skipped), so billion-laughs and
  XXE are structurally absent, not merely mitigated. The reader (`src/io/xlsx/read.ts`, on `fflate`
  `unzipSync` with a **declared-size inflate cap** as the naïve-zip-bomb guard) reconstructs what the
  writer emits: sheet names/order, number/string/boolean/formula cells (inlineStr + `t="s"` shared
  strings), per-column width/visibility, per-row height/visibility, merges, page margins, and core
  document properties (`docProps/core.xml`). The model grew `Worksheet.rowCount`/`actualRowCount`
  (used-range extent that spans gaps vs populated-row tally). The adapter's `roundtripWorkbook` now
  binds write→read→normalize to this reader, mirroring `current.mjs`'s JSON model exactly.
- **Styles slice — pattern fills, both directions (2026-07-12):** the first cut into the largest `∅`
  cluster (styles). The style model grew a real `Fill`/`PatternFill`/`FillPatternType` (`src/core/style.ts`),
  `Cell.fill`, and `RowProperties.fill`; each cell owns its own fill object, so a fill set on one cell
  *cannot alias* a neighbour's style — the classic style-bleed bug is absent by construction, not patched.
  The writer now generates `styles.xml` per-workbook via a new **`StyleRegistry`** (`src/io/xlsx/styles.ts`)
  that **interns** fills and cell formats: identical fills collapse to one `<fills>`/`<cellXfs>` entry
  (bounded write cost on large lightly-formatted sheets), distinct fills stay separate. Cells emit `s="N"`;
  a formatted row emits `s="N" customFormat="1"`. The static `STYLES_XML` was deleted (dead). The reader
  parses `styles.xml` into an xf→fill table and resolves each cell's `s` (a cell without its own `s`
  inherits its `customFormat` row's fill). A hostile-input note surfaced during the slice: the SAX parser
  emits no close event for self-closing tags, so the colourless `patternFill` (none/gray125) must be pushed
  on *open* to keep fill-id slots aligned — otherwise every foreign styles.xml mis-indexes.
- **Latest corpus vs rewrite: 198 green / 5 known-open / 88 legacy known-opens resolved / 0 regressions /
  370 skipped; test:src 260/260.** The **paperSize passthrough** closes the last live `<pageSetup>` attribute
  that was dropped on write: `PageSetup.paperSize` carries Excel's 1-based paper enumeration index (9 = A4) as
  an opaque integer, emitted as the leading `<pageSetup>` attribute (CT_PageSetup order) and read back with the
  same NaN-drop guard as the other numerics. The real-Excel fidelity fixture already carries `paperSize="9"`, so
  its round-trip now proves it survives (the assertion was strengthened in place — no new corpus behavior). 2 new
  `read.test.ts` tests (paperSize round-trips and leads the attributes; a non-numeric `paperSize` is dropped, not
  stored as NaN). Still not modelled from `<pageSetup>`: the printer-settings `r:id` (an opaque binary part —
  its own future slice).
- **The page-setup slice** teaches the sheet its print scaling: a new
  `PageSetup` on `Worksheet.pageSetup` (`fitToPage`/`fitToWidth`/`fitToHeight`/`scale`/`orientation`/`pageOrder`/`paperSize`,
  mutable in place) lets an author fit columns/rows onto a fixed page count or a fixed zoom and pick paper
  size, orientation, and page order. The fit-to-page flag rides `<pageSetUpPr>` inside `<sheetPr>` (after `<outlinePr>`
  per CT_SheetPr order); the rest ride `<pageSetup>`, placed between `<pageMargins>` and `<headerFooter>` in
  CT_Worksheet order. Only the attributes the author set are emitted (no element at all when the sheet keeps its
  defaults), and the reader sets only those the source carried, so a re-write stays byte-clean. The model threads
  `pageSetup` so a `dst.model = src.model` clone preserves it. Lights `pagesetup-fit-to-page-round-trips` (2
  behaviors, declarative) **and** `column-width-and-pagesetup-roundtrip-fidelity` (2 behaviors) — the latter via a
  new `roundtripFixtureStyleFacts` adapter method that reads a real Excel fixture (`scale="96" fitToHeight="0"
  pageOrder="overThenDown" orientation="landscape"`), writes it back, and re-reads it, proving both the fractional
  column widths and the print-scaling attributes survive a no-op round-trip. That un-skipped adapter method also
  surfaced 2 pre-existing known-opens (DXF number formats, custom indexed-color palette — both baseline `fail`,
  unrelated features). 8 `read.test.ts` tests (fit-to-page round-trip, `<pageSetUpPr>`/`<pageSetup>` split,
  CT_SheetPr order after `<outlinePr>`, CT_Worksheet placement between margins and header/footer, orientation +
  pageOrder emit-only-when-set, default → no element, a `<pageSetUpPr>` present only for `autoPageBreaks` leaves
  `fitToPage` unset, model clone).
- The **outline-summary-position slice** teaches the sheet its grouping
  layout: a new `OutlineProperties` on `Worksheet.outline` (`summaryBelow`/`summaryRight`, mutable in place)
  lets an author place summary rows above their detail and summary columns to their left, inverting Excel's
  defaults. The writer folds `<outlinePr>` into the existing `<sheetPr>`, after `<tabColor>` per CT_SheetPr
  order, emitting only the flags the caller set (and no element when the sheet keeps the defaults); the reader
  parses them back, setting only those the source carried so a re-write stays byte-clean. The model threads
  `outline` so a `dst.model = src.model` clone preserves it. Lights `worksheet-outline-summary-position-round-trips`
  (2 behaviors, both `✓` green regression locks — legacy round-trips this too). Adapter `outlinePropertiesRoundtrip`;
  6 `read.test.ts` tests (round-trip, `<sheetPr>` payload, shared-element CT_SheetPr order, single-flag-only,
  default → no element, model clone).
- The **worksheet-tab-colour slice** adds a `Worksheet.tabColor` facet (an
  ARGB/theme `Color`): the writer emits `<sheetPr><tabColor>` as the leading worksheet child when a tab colour
  is set and nothing when it is unset, and the reader parses it back — so a coloured tab round-trips verbatim
  and an uncoloured sheet never fabricates one. The colour-attribute serializer (`colorAttrs`) is shared
  between the stylesheet and the tab-colour writer so the two paths can't drift, and the worksheet model
  carries `tabColor` so a `dst.model = src.model` clone preserves it. Lights `worksheet-tab-color-argb-round-trips`
  (3 behaviors, all `✓` green regression locks — legacy round-trips this too). Adapter `tabColorRoundtrip`;
  5 `read.test.ts` tests (ARGB round-trip, `<sheetPr>` position, no-colour → no element, theme+tint, model clone).
- The **structured-date-values slice** teaches both directions that a
  date is a numeric serial plus a date number format. New `src/core/date.ts` owns the date domain in one
  place: `dateToSerial`/`serialToDate` reproduce the 1900 date system including the phantom 1900-02-29
  leap-year quirk (serial 1 is 1900-01-01, not 1899-12-31), `isDateFormat` detects a date/time format code,
  and `DEFAULT_DATE_NUMFMT` is the format a bare `Date` acquires so it renders and reads back as a date. The
  writer serialises a valid `Date` to its serial under a date format (an explicit cell/column format wins) and
  writes an Invalid Date value-less rather than throwing, so one bad date never drops its siblings; the reader
  surfaces a plain numeric cell under a date format as a `Date` (a string/boolean/formula result keeps its
  own kind). Also closes two foreign date encodings the reader ignored: a Strict-mode `t="d"` cell (ISO 8601
  value parsed literally) and the locale-specific built-in date/time format ids (27..36, 50..58) a CJK locale
  styles date cells with, resolved from the built-in table. Lights `invalid-date-cell-tolerated-on-write` and
  the date behavior of `custom-numfmt-string-roundtrips-verbatim` (green), and resolves the legacy known-opens
  `date-serial-1900-epoch-leap-year`, `strict-mode-iso8601-date-parses-correctly`, and
  `builtin-cjk-date-numfmt-ids-resolve-to-date-format`. Adapter gained `readFixtureCells`, structured
  date-value building, and a `tryWriteWorkbook` survival report; `src/core/date.test.ts` + `read.test.ts`
  round-trip tests lock the math and both foreign encodings.
- The **read-sheet-protection slice** closes the round-trip on a capability
  the writer already had: the reader now parses `<sheetProtection>` back into the model, so a protected sheet
  read then re-written stays locked instead of silently unlocking on a passthrough save. The OOXML flag/default
  encoding table moved to `src/core/protection.ts` as the single canonical `SHEET_PROTECTION_FLAGS` (shared by
  both the writer's allow-flag→attribute inversion and the reader's attribute→allow-flag one, so the two
  directions can't drift). `Worksheet.restoreProtection` is the deserialization counterpart to `protect`: the
  reader reinstates the agile credential (algorithm/hash/salt/spinCount) **verbatim** — there is no plaintext
  password to recover, so it is never re-hashed — while `protect` remains the only password-hashing entry point.
  A `sheet="0"` element restores nothing; only flag attributes actually present are carried, mirroring the
  writer's omit-when-default rule so a re-write is byte-identical. New case
  `sheet-protection-survives-roundtrip` (4 behaviors, all `✓` green regression locks — legacy round-trips this
  too), adapter `sheetProtectionRoundtrip` on both oracles, and 4 `read.test.ts` round-trip tests.
- The **anchored-images slice** closes the last structural-edit case,
  `splice-rows-updates-table-and-image-refs`. Images live as workbook-wide media (`Workbook.addImage`
  returns a numeric id; the bytes are stored once) that a sheet anchors with `sheet.addImage(id, {tl, br})`
  — a two-cell grid anchor (0-based). The writer plans the referenced media into `xl/media/image{n}.{ext}`
  parts (one per image, deduped, unreferenced images not written), emits one `xl/drawings/drawing{n}.xml`
  per sheet (a DrawingML `<xdr:twoCellAnchor>` per image), its `_rels` mapping each `r:embed` to its media,
  a `<drawing r:id>` element in the sheet XML (schema order: before `<legacyDrawing>`, before
  `<tableParts>`), the sheet→drawing relationship (numbered after any table rels, before the note rels), a
  `<Default>` content-type per image extension, and the drawing-part override. The reader reaches a sheet's
  drawing through its own rels (`.../drawing`), parses the anchors, resolves each embed to its media bytes,
  and re-registers them (deduped by media path so a shared image stays one workbook image) — so an image
  round-trips with its bytes intact. Tables and images now **shift with a row/column splice** alongside
  merges: `Table.shiftRows/shiftColumns` re-pin the range (a splice above moves it, one inside grows/shrinks
  the data rows, one deleting every row drops the table), and anchor points move like merge edges. Authoring
  a table with **duplicate column names is now rejected** at construction (case-insensitively) — Excel treats
  a collision as corruption. New `src/core/image.ts` (anchor model), `src/io/xlsx/images.ts` (drawing writer +
  reader), `src/core/table.test.ts` (dup rejection + shift), and `src/io/xlsx/images.test.ts` (round-trip,
  splice-shift, shared-media). The case's three behaviors are `↑` FIXED (baselines stay `fail`). **No
  structural-edit case remains `∅`.**
- The **cell-notes slice** gives a cell a `note` (plain-text comment)
  facet and teaches the writer/reader the legacy-comment package parts. A `Cell.note` is metadata anchored
  to the cell independent of its value — a cell can carry a note while empty, and it is owned per-cell (no
  bleed) and copied on a structural shift, so `copyCellContent`, the worksheet `model`, and every splice/
  insert/duplicate already carry it. The writer emits, for each sheet that has notes, an `xl/comments{n}.xml`
  part (one anonymous author, one `<comment ref>` per noted cell), an `xl/drawings/vmlDrawing{n}.vml`
  companion (the legacy VML text-box shape Excel still needs to render a note, one hidden shape per note), a
  `<legacyDrawing r:id>` element in the sheet XML (in schema order, before `<tableParts>`), the two sheet-
  local relationships (VML + comments, numbered after any table rels), the `vml` extension default, and the
  comments-part content-type override — so a noted workbook is a complete, valid package, and a note-free
  one grows no comment/VML parts. The reader reaches a sheet's comments through the sheet's own rels
  (`_rels/sheet{n}.xml.rels` → the `.../comments` relationship, target resolved relative to the sheet's
  directory), parses the text runs, and applies each note by A1 reference. This closes
  `row-insert-preserves-note-and-outline-level`: the row-insert control is a `✓` green lock, and the two
  fail-baseline behaviors (note survives the insert, outline level follows its row — outline was already
  end-to-end, only notes were missing) are `↑` FIXED (baselines stay `fail`). New `src/io/xlsx/comments.ts`
  (writer + VML + reader) with seven `comments.test.ts` round-trip/isolation/shape tests, plus two model-
  level `worksheet.test.ts` cases (note travels a splice, note survives model export/import). (The
  `splice-rows-updates-table-and-image-refs` case that once remained `∅` here has since been closed by
  the anchored-images slice above.) The
  **structural-edit slice** (spliceRows/spliceColumns/insertRow/
  duplicateRow on `Worksheet`) is the family that opens up bulk grid editing. A row splice removes `count`
  rows at `start` and inserts new ones in their place; rows below shift by `inserts.length - count`, so a
  delete pulls the tail up, an insert pushes it down, and an over-large count clears the tail rather than
  silently no-op-ing. A shifted cell is a fresh cell at the new coordinates carrying the original's value and
  **every style facet** (`Cell` fixes its position at construction; facet objects pass by reference, safe
  under copy-on-write), so a splice never blanks a styled row. Row/column *metadata* (heights, widths,
  hidden/outline) shifts the same way, and merged ranges re-anchor: a range wholly past the cut shifts by the
  net delta, one whose covered rows/columns are entirely deleted drops, a straddling range clamps to the cut
  as best-effort, and an unbounded whole-row/column merge passes through. `insertRow` is `spliceRows(pos, 0,
  values)`; `duplicateRow` copies a row `count` times (inserting-and-shifting, or overwriting) and the copies
  carry no merge of their own, so a range can be merged onto a duplicate afterward. Added a `columnCount`
  getter (the used-range width, mirroring `rowCount`). Lights nine cases — the pass-baseline halves are green
  regression locks (`✓`: small-count splice, interior delete, style travel on shift, column shift-left/insert,
  trailing-data shift, faithful duplicate + merge, inheritance-inserted cells stay mutable, splice-below/
  insert-below controls), the fail-baseline halves are `↑` FIXED (rewrite-only, baselines stay `fail`:
  over-large count clears the tail, delete-at-end removes trailing rows/columns, splice/insert/duplicate
  shift merges, `lastRow` tracks the tail past a delete). Covered by twelve new `src/core/worksheet.test.ts`
  cases. The rewrite adapter now binds all four `mutateWorksheet` ops plus `duplicateRowReport` and
  `insertRowThenStyle`, and reads style facets + `columnCount` back. The `row-insert-preserves-note-and-
  outline-level` case (cell notes) has since been closed by the cell-notes slice above, and the last one,
  `splice-rows-updates-table-and-image-refs`, by the anchored-images slice — no structural-edit case is `∅`.
  The **worksheet model export/import slice** (the merged-cells family's
  third and final real source feature — the family is now fully closed) gives `Worksheet` a `model` getter
  that snapshots the sheet's transferable content (state, page setup, columns, rows, cells with their
  per-cell styles, merges, tables, sheet protection) and a setter that reproduces it. The getter emits and
  the setter consumes exactly the same fields, so `dst.model = src.model` clones a sheet losslessly — closing
  a long-standing asymmetry where the exported model exposed merges under one property the importer never
  read, silently dropping every merged range on a model copy. On import, cells load at their exact positions
  with merges re-applied afterward (a slave cell's value cannot be misrouted mid-load), and the assignment
  replaces the target's content wholesale, leaving no residue; identity (`name`, `id`) is not part of the
  model and is untouched. `Table` gained an `options` getter so a table round-trips through its constructor
  shape. Lights `worksheet-model-preserves-merged-cells`: its precondition behaviour (baseline `pass`) is a
  green lock, its import behaviour (baseline `fail`) is `↑` FIXED (rewrite-only; baseline stays `fail`).
  Covered by six new `src/core/worksheet.test.ts` cases (merge round-trip, cell values + facets, column/row/
  page metadata, tables + protection, wholesale replace leaves no residue, snapshot does not alias the
  source). This was the springboard for the structural-edit slice above, which has since landed and unlocked
  those cases. The **merge-overlap rejection slice** (the merged-cells family's second
  real source feature) makes `Worksheet.mergeCells` reject a bounded range that overlaps an already-merged
  region — overlapping merges are Excel-invalid geometry that opens as corrupt. The check runs against the
  decoded merge rectangles already kept for slave-cell addressing (two inclusive rectangles overlap when
  neither is fully left/right/above/below the other); a rejected range never enters the merge list, and
  edge-abutting merges that share no cell are still accepted. Whole-row/column merges are unbounded, carry
  no rectangle, and are not overlap-checked. The reader stays tolerant of a corrupt file that declares
  overlapping merges: the offending range is dropped, never allowed to abort the parse (one bad merge must
  not cost the rest of the sheet's geometry). The rewrite corpus adapter now binds the `mutateWorksheet`
  structural-edit contract for the `mergeCells` op only; the still-unbuilt splice/insert/duplicate ops (and
  style read-back) tag their error as `notImplemented`, so those ~nine cases skip per-behavior rather than
  fail — the partial-bind trap avoided — while the no-op control that shares the last-row case is satisfied
  by deriving `lastRow` from the row iterator. Lights the overlap behaviour of
  `many-merged-cells-preserved-and-overlap-rejected` plus that last-row control (two new green locks, both
  baselines `pass`). Covered by four new `src/core/worksheet.test.ts` cases (overlap, containment,
  edge-abutting-allowed, unbounded-not-checked). At the time this was the last-but-one merged-cell case; the
  worksheet model export/import copy (see the model slice above) has since closed the family.
  The **merge-aware cell addressing slice** (the first real source feature of
  the merged-cells family) makes `getCell` resolve a covered address to its region's master (top-left) cell:
  the worksheet keeps a decoded rectangle beside each merge, and addressing any cell inside a merged block
  returns the one master cell. So a value written by addressing a non-master (slave) cell — e.g. the
  bottom-right of `A1:B2` — lands on the master, only the master ever materialises, the serialized sheet
  carries a value on exactly one cell (no stray value on a covered cell, which spreadsheets treat as
  malformed), and reading either address returns the region's value
  (`merged-range-slave-cell-write-resolves-to-master`, three green regression locks). Resolution is consulted
  at access time, so it applies to merges declared after cells already exist; an unbounded whole-row/column
  merge is still declared but redirects nothing (no bounded rect). Covered by `src/core/worksheet.test.ts`
  (5 model-level tests) plus the corpus round-trip. Merge-vs-merge overlap rejection followed as the family's
  second source feature (see the merge-overlap slice above), and the worksheet model export/import copy (see
  the model slice above) closed the family as its third.

- **The merge write-cleanliness slice** (adapter-only, no source change)
  binds two merged-cell reports proving the writer emits clean merges and the master cell keeps its style:
  a merged horizontal span emits the range exactly once with a value only on the anchor, so no covered cell
  carries conflicting content — the clean shape that opens without Excel's "recover?" repair prompt — and the
  anchor's value and alignment survive the round-trip (`merged-range-opens-without-repair-prompt`); the
  master (top-left) cell of a merged region keeps its border, number format, and font through the merge and
  round-trip, so the region renders its intended outline (`merged-region-master-cell-border-survives`). The
  writer already emits only populated cells (from `sheet.rows()`) and the reader round-trips per-cell styles,
  so binding `mergeCleanReport` + `mergeMasterBorderReport` lights both cases green with no source change —
  six green regression locks (all baselines `pass`). Slave-cell writes resolving to the master (merge-aware
  cell addressing) followed as the family's first real source feature, then merge-vs-merge overlap rejection
  (see the merge-overlap slice above), and the worksheet model export/import copy (see the model slice above)
  closed the family as its third.

- The **`roundtripFixture` slice** closes the last deferred piece of the
  foreign-fixture reading family (adapter-only, no source change): a real Excel-authored styled template
  (banded solid fills, bold headers, custom column widths, merged-cell layout across three sheets) read from
  disk and written straight back keeps its sheet names, custom column widths, and per-cell
  fill/font/numFmt/alignment/border — the mainstream "open a styled template, fill it in, save it" path is
  format-preserving. The reader/writer already round-trip it faithfully; binding the capability proves it.
  In the rewrite's model a column stores a width only for a custom width, so "has a width" is exactly "is a
  custom width". Three green regression locks (baseline `pass`) — `template-styles-survive-read-write-roundtrip`.
  With this the foreign-fixture family is **fully closed**.

- The earlier **default-font resolution slice** closed the first of the two
  deferred foreign-fixture pieces: an unstyled cell renders in the workbook default face (Calibri 11), yet
  the reader surfaced no font for it — font id 0 was skipped as if it were an absence, the way border id 0
  genuinely is, but font 0 is a *real* font. Reader: an xf naming font 0 now resolves to that default face,
  and a bare cell (no style of its own nor an inherited row/column one) falls back to xf 0 rather than to
  nothing. Writer: a cell carrying exactly the default font interns back to id 0, so a read→write round-trip
  stays byte-stable instead of accreting a redundant `<font>` entry. Isolation is untouched — in-memory
  cells are never stamped, and the default face carries no bold/italic/colour to bleed (the one src test
  that used `font===undefined` as an isolation proxy now asserts true isolation *plus* the resolved default;
  two new src tests lock the resolution and the round-trip dedup). Lights the two
  `default-font-applies-to-unstyled-cells` behaviours (rewrite-only `↑`; baselines stay `fail`).

- The **foreign-fixture reading family** (115 → 143, now fully closed) is the fourth
  *capability family*, delivered in three bulk slices plus two follow-ups — the reader now proves itself
  against real, non-Excel input, mostly with no source change:
  1. **Real-fixture colour reading** (115 → 122, adapter-only): bound `readFixtureCellStyles` and
     `roundtripFixtureColorFidelity` so three real Excel-authored fixtures measure the reader/writer.
     A solid-pattern fill exposes its visible colour on `fgColor` with the automatic indexed `bgColor`,
     and the font colour is a wholly separate facet; a pure open-then-save preserves every cell's fill
     and border-edge colours, **including theme+tint and indexed-palette references**, verbatim. The
     reader already parsed theme/tint/indexed/argb — this locks it end-to-end. Green regression locks
     (baseline `pass`): `solid-fill-foreground-vs-font-color`, `theme-and-rgb-fill-colors-read-faithfully`,
     `fill-border-color-survives-roundtrip`.
  2. **Foreign explicit-off font/alignment forms** (122 → 124, one small reader fix): a package-patching
     adapter helper (`reloadPatched`) feeds the reader the explicit-off forms real producers emit but the
     writer never generates — `<b val="0"/>`, an alignment element carrying only `wrapText="0"`/`shrinkToFit="0"`,
     `<u val="none"/>`. The reader already honoured `val="0"` on b/i/strike and treated an all-false alignment
     as *no* alignment; the one **source fix**: `<u val="none"/>` (the explicit ABSENCE of an underline) now
     reads back falsy instead of the truthy string `"none"` that a consumer's `if (font.underline)` would
     mistake for underlined — bare `<u/>` still true, a named variant still carries through. Locked by a new
     `read.test.ts` foreign-underline test. Fixes `font-boolean-flag-honors-explicit-false` and
     `alignment-false-boolean-attrs-yield-no-alignment` (baselines stay `fail` — legacy still fails; these
     are rewrite-only `↑`).
  3. **`readFixtureReport`** (124 → 140, adapter-only): the broad reader-robustness probe over ~two dozen
     real fixtures, capturing any crash as data (`{ok, error, sheetNames, lastModifiedBy, creator}`). The
     lean SAX reader is robust **by construction** to every shape that crashes tag-literal / fixed-part-order
     parsers: namespace-prefixed roots (`<x:workbook>`), a leading BOM, non-ASCII sheet names, `workbook.xml`
     ordered after a worksheet (parts read from a map), unprefixed core.xml (`lastModifiedBy` still read),
     cells lacking `r` (positions inferred), missing optional parts/elements (app.xml `Company`,
     `sheetFormatPr`, styles, drawing, VML/comment targets), and foreign boolean spellings / mixed `<si>`
     shapes. 28 legacy known-opens resolved in one bind.

  **This family is now fully closed.** Both deferred pieces are done: `unstyledCellFontReport` (the broad
  default-font resolution) and `roundtripFixture` (`template-styles-survive-read-write-roundtrip`) — see the
  `roundtripFixture` and default-font slices above.

  The **column-scope style-inheritance slice** (111 → 115) is the third
  *capability family* past the per-cell facet surface. A column carried only a number format as a default
  for its cells; every other facet a column can declare — fill, font, border, alignment, protection — was
  dropped on write and never inherited on read. Real files routinely scope a border or an alignment to a
  whole column, so `ColumnProperties` now carries the full facet bundle (documented as per-cell defaults
  with cell-over-row-over-column precedence, symmetric with a row's fill), the writer composes each cell's
  style inheriting every column facet (not just numFmt) and interns the full bundle into the column's own
  `<col>` style so Excel applies it to the column's empty cells too, and the reader mirrors a `<col style>`'s
  full facet bundle onto the column model (a bare cell already inherited the full column xf). This flips the
  two column-scope behaviours to green **regression locks** (baseline `pass`): `alignment-does-not-leak-across-cells`
  (a column alignment stays scoped to its own cells) and `column-border-style-scoped-to-declaring-column`
  (a column border does not bleed into later width-only columns). Locked in src by three new
  `read.test.ts` round-trip tests (column alignment inheritance/isolation, column border scoping, and a
  one-facet override that preserves the column's other default). The **copy-on-write style-aliasing slice**
  (106 → 111) is the second *capability family* past the per-cell facet surface, and it required **no source change** — the rewrite
  already gets isolation by construction, and this slice proved it. On disk, identically-formatted cells
  deduplicate to one shared style record, so a loaded workbook hands several cells the same style; legacy
  then bled a single-cell facet edit into every sibling that shared the record. The rewrite cannot: each
  `Cell` owns independent facet fields, every facet setter REPLACES the field (never edits in place), and
  the facet types are `readonly` so a shared record is inert. Wiring five adapter methods
  (`sharedBaseStyleFontMutation`, `loadMutateCellBorder`, `loadMutateCellStyle`, `loadMutateCellFont`,
  `loadMutateCellFacet` — the rewrite has no `.style` aggregate, so a shared "base style" is decomposed
  onto per-facet setters) resolved 11 legacy known-opens across `loaded-cells-shared-style-object-aliasing`,
  `shared-base-style-font-mutation-isolated`, `cell-border-mutation-does-not-bleed-to-style-siblings`, and
  `cell-style-setter-isolates-alignment-numfmt-protection` (all three facets: alignment, numFmt, protection).
  Because those baselines are `fail` (the legacy bleed), a rewrite regression would silently drop them to
  `○` known-open rather than `✗`, so the guarantee is **hard-locked in src** by a new
  `src/io/xlsx/style-isolation.test.ts` (5 tests: fill/font/border/facet isolation through the real
  write→read path, including the same-base-object aliasing trap and the facets-compose-not-replace case).
  The **sheet-level protection slice** (94 → 106) is the first *capability family* past the
  per-cell facet surface, and it lit three cases at once: `sheet.protect(password?, options?)` (on
  `Worksheet`) stores a `SheetProtection` overlay, and the writer emits a self-closing `<sheetProtection>`
  after `<sheetData>` (CT_Worksheet order). The option surface is stated in the **author's** terms — each
  flag answers "may a user still do this while protected?" — and the writer INVERTS it to OOXML's
  "forbidden" booleans (attr "1" locks, "0"/omission permits), writing only values that differ from each
  attribute's per-attribute default (most editing ops default forbidden under protection; selecting cells
  defaults permitted). A password derives an **OOXML-agile SHA-512 credential** (`src/core/protection.ts`,
  the one module touching `node:crypto`): salt is `randomBytes(16)`, hashed `spinCount` (default 100000)
  times with a little-endian uint32 counter — the plaintext is never retained, and two protects with one
  password differ because the salt is real randomness. Whole-column/row unlock **bands** are realized by the
  adapter stamping the flag onto the listed band cells (the same end-state a per-cell override yields;
  native column-scope inheritance stays a separate family). This lit `cell-protection-locked-flag-and-sheet-protection`
  (all 5), `sheet-protection-permits-requested-operations` (all 3), and the security case
  `worksheet-password-protection-hashes-in-node` (all 4). The earlier **alignment slice** (86 → 94) round-trips a cell's alignment — horizontal/vertical
  placement, `textRotation`, `indent`, and the `wrapText`/`shrinkToFit` flags — through the shared style
  table: `Cell.alignment` (an `Alignment`, each cell owning its own so alignment never bleeds onto
  siblings) → because `<alignment>` is a *child of the xf* (not a shared sub-table like fills/fonts/borders),
  the `StyleRegistry` interns its serialised attribute string directly into the xf signature and emits it as
  the xf's body (`applyAlignment="1"`), with an all-default alignment contributing nothing and collapsing to
  xf 0 → the reader holds each `<cellXfs>` xf open from start to close so an `<alignment>` child can attach
  before the xf commits, parsing flags by their true boolean (`wrapText="0"` is off, never a spurious
  `{wrapText:false}`). Lit `alignment-flags-round-trip` (wrapText/shrinkToFit/indent) and the cell-level
  behavior of `alignment-does-not-leak-across-cells`; the column-scoped and copy-on-write alignment cases
  (`cell-style-setter-isolates-alignment-numfmt-protection`) gate on the column-scope/aliasing families for a
  later slice. The slice also uncovered — and fixed — a latent newline gap: enabling the `alignment` cell
  key made `cell-value-newline-line-break-roundtrip` runnable, which needs `\r\n`→`\n` normalization. The
  fix is spec-correct XML end-of-line handling (§2.11) in the SAX reader's text path (a literal CRLF/CR
  normalizes to LF, before entity decoding so a deliberate `&#13;` survives; CDATA stays verbatim), taking
  that case fully green too. Before it, the **borders slice** (85 → 86) round-trips a cell border — the four sides plus a diagonal,
  each an independent edge with an optional colour — through the shared style table: `Cell.border` (a
  `Border`, each cell owning its own so a border never bleeds onto siblings) → `StyleRegistry` interns each
  distinct `<border>` in schema edge order (left, right, top, bottom, diagonal; id from 1) and composes it
  into the cell's one xf alongside fill+numFmt+font, with an all-styleless border collapsing to the empty
  default border id 0 → the reader parses `<borders>` (edges, their colour children, and the
  diagonalUp/diagonalDown direction) and resolves an xf's `borderId` (id 0 = empty default = no border), so
  an unbordered cell reads back borderless and a one-sided border never fabricates the other three sides.
  `unbordered-cell-has-no-phantom-border` is green; the aliasing/column-scope/fixture border cases
  (`cell-border-mutation-does-not-bleed-to-style-siblings`, `column-border-style-scoped-to-declaring-column`,
  `merged-region-master-cell-border-survives`, `fill-border-color-survives-roundtrip`) gate on their own
  capabilities for a later slice. Before it, the **fonts slice** (79 → 85) round-trips a cell font — bold/italic/underline, size,
  colour, typeface — through the shared style table: `Cell.font` (a `Partial<Font>`, each cell owning its
  own so a font never bleeds onto siblings) → `StyleRegistry` interns each distinct `<font>` (id from 1,
  after the default) and composes it into the cell's one xf alongside fill+numFmt → the reader parses
  `<fonts>` and resolves an xf's `fontId` (id 0 = default = no explicit font). Boolean flags are read
  honouring their `val` (`<b/>`/`<b val="1"/>` on, `<b val="0"/>` off — presence alone is not truth), so
  the explicit-false reader path is already correct for a later foreign-XML slice. `per-cell-font-isolation`
  is green; the aliasing/foreign-XML font cases (`fontExplicitFalse*`, `unstyledCellFontReport`,
  `sharedBaseStyleFontMutation`) stay for a later slice. Before it, the **style-dedup slice** (76 → 79)
  bound the corpus's `styleDedupReport` to the rewrite by reading the written package (count `<cellXfs>`
  entries + each cell's resolved index), lighting `shared-styles-deduplicated-in-written-package`: 40
  identically-styled cells collapse to one index, the table never inflates to one entry per cell, and a
  distinct style keeps its own index. The **number-format slice** lit up the numFmt cluster (63 → 76): a custom accounting
  format round-trips byte-for-byte (no comma-drop, quoted literals/escapes intact), distinct per-column
  numFmts stay independent, and — the shared-style-bleed fix — assigning a fill to one cell in a
  formatted column keeps the column's number format on that cell and every sibling. The writer composes
  each cell's *full* style (cell overrides atop row/column defaults) into one interned xf, so overriding
  one facet never drops another; the reader parses `<numFmts>` (+ the ECMA-376 built-in id table for
  foreign files) and resolves cell/column/row style precedence. Only the date-value numFmt behavior stays
  `∅` (gated on structured date values). The earlier styles slice lit up the two pure-fill round-trips
  (58 → 63): a solid fill stays local to its cell (no bleed onto row/column/sheet siblings), and a
  row-level fill is inherited only by that row's cells. The
  reader earlier lit up 23 round-trip cases (35 → 58): scalar-type fidelity (a digit-string
  `"10"`/zero-padded `"007"` stays a string, `15` stays a number), formulas with markup-significant
  operators (`<`, `>`, `&`) round-trip verbatim, merged-range count/geometry survives, explicit row
  heights persist, table geometry round-trips intact, workbook core properties (creator/lastModifiedBy/
  created/modified) read back unchanged, and `rowCount` spans a row gap while `actualRowCount` excludes
  it. The `_xlfn.` modern-function prefix bug is still `○` (legacy fails too); outline summary-row
  `collapsed` inference remains open. Legacy oracle unchanged: **424 green / 233 known-open / 0
  regressions** (an occasional flake in the *legacy* streaming-control `streamReadManySheets(3)` is a
  `lib/` temp-file/resource timing artefact — the rewrite adapter skips streaming, so it is unrelated to
  the rewrite). Gates all green: `typecheck` clean, `test:src` **158/158** (+10 over the sheet-protection
  slice: 6 writer tests covering the `<sheetProtection>` element — omitted when unprotected, self-closing
  `sheet="1"` after `<sheetData>`, agile credential attrs from a password, allow-flag→forbidden inversion
  writing only non-defaults, default flags omitted, `unprotect()` clearing it — plus 4 `deriveCredential`
  tests including an independent re-implementation of the OOXML-agile hash proving the digest is genuinely
  derived from the reported salt, and real-randomness/custom-spin-count checks; +8 earlier over the alignment
  slice: protection interning as the second xf child + all-default→xf-0 + alignment+protection as two ordered
  xf children + protection/facet independence + cell-protection round-trip of the meaningful flags +
  default-locked→no-protection + foreign explicit-`locked="1"`→no-protection). The per-cell **protection**
  slice completed the per-cell style-facet surface (fill, numFmt, font, border, alignment, protection);
  `locked` defaults to TRUE in OOXML, so only an explicitly *unlocked* cell (`locked="0"`) carries
  information, and the reader never fabricates `{locked: true}` from a default or explicit-`locked="1"` cell.
  The **sheet-level protection** slice then delivered `<sheetProtection>` + password hashing (see the metrics
  narrative above), lighting the three protection cases the per-cell facet alone could not.
- **Gates wired:** `npm run typecheck` (strict `tsc --noEmit`, TypeScript 5.x), `npm run test:src`
  (native `node --test` on `.ts`), `npm run corpus:rewrite`. Toolchain rationale in
  [`docs/decisions/0001-rewrite-runtime-and-toolchain.md`](docs/decisions/0001-rewrite-runtime-and-toolchain.md):
  Node 24 runs `.ts` directly, so **no bundler/build step** on the runtime path yet — `tsc` is the
  type gate only. Vitest/Biome/tsup + the `fflate`/XML dep swaps are a deferred toolchain-standup
  slice.
- **Independent OOXML oracle:** `npm run test:ooxml` validates representative buffered and
  streaming `.xlsx` output with Microsoft's `OpenXmlValidator` 3.5.1 targeting Microsoft 365. The
  repo-owned .NET 10 wrapper emits structured JSON and runs in a separate required CI job. One
  exact legacy default-font ordering error is baselined; any additional error or a stale baseline
  fails. Decision and limits: `docs/decisions/0002-ooxml-validation-oracle.md`.
- Build order from here: **core model → XML layer → xlsx r/w → streaming → csv**, each landed fully
  green before the next depends on it. Drive each cluster's known-opens to `✓`/`↑` under
  `--adapter rewrite`.
- **Exit:** feature-parity-or-better on the corpus for all in-scope areas (no `○`/`∅` left under
  the rewrite adapter); legacy `lib/` deleted; dependency tree small and audit-clean.

### ⏳ Phase 4 — Independence & identity  *(only the remainder — hosting already done)*
- ✅ Hosting/repo independence (done above).
- ⏳ Drop the `upstream` remote once harvest is complete.
- ❓ **Final rebrand name** (human decision). `ts-xlsx` / `@shbernal/ts-xlsx` is a
  **provisional** working name only.
- ⏳ First-class docs from types; publish a `0.x` release under the chosen name.

---

## Open decisions ❓
1. **Merge-first vs corpus-only for open PRs** — **downgraded from gating to optional** (2026-07-11).
   The corpus already captured every PR's intent, so the rewrite loses no *knowledge* either way;
   this now only decides whether to salvage the authors' *patches/review credit* on a legacy tree
   we are deleting. Default if unanswered: **corpus-only** (we do not invest in legacy). Still open
   only as a courtesy call for the human.
2. **Final brand name** — reserved for the human (Phase 4).

## Manual follow-ups (outside this tool's reach)
- Delete the old GitHub fork `shbernal/exceljs` (needs `delete_repo` token scope). Safe anytime —
  the harvest reads upstream `exceljs/exceljs`, not the fork.

## 🔜 Immediate next action
**Phase 3 rebuild is underway.** The engine is proven end-to-end: `src/` (strict TS) → `rewrite.mjs`
adapter → the corpus runs against it, the first module (`address.ts`) is green **and already beats a
legacy known-open**, with 0 regressions on the legacy oracle. Continue module by module, in the
`STRATEGY.md` build order (**core model → XML layer → xlsx r/w → streaming → csv**):

1. **Deepen the styles surface — both directions, one facet per corpus win.** Writer and reader now
   round-trip **pattern fills** (per-cell + row-inherited), **number formats** (`numFmt` on cells and
   columns, cell-inherits-column, `<numFmts>` + ECMA-376 built-in id table), **fonts** (bold/italic/
   underline/size/colour/typeface → `<fonts>`, boolean flags read honouring `val`), **borders** (four
   sides + diagonal, per-edge colour + diagonal direction → `<borders>` in schema edge order, empty border
   → id 0), **alignment** (horizontal/vertical/`textRotation`/`indent`/`wrapText`/`shrinkToFit` as an
   `<alignment>` child of the xf), and **protection** (`locked`/`hidden` as a `<protection>` child of the
   xf, after `<alignment>`) through a shared `StyleRegistry` that interns numFmts/fills/fonts/borders/
   cellXfs and composes each cell's full style; dedup is exposed to the corpus via `styleDedupReport`.
   **The per-cell style-facet surface is now complete** (fill, numFmt, font, border, alignment, protection).
   What remains for styles is not new *facets* but the cross-cutting *capability families* the deferred
   cases gate on — and these now recur across facets, so building each capability lights several cases at
   once:
   - ✅ **Sheet-level protection** (`sheet.protect(password, options)` → `<sheetProtection>`, plus column/row
     protection bands) — **DONE.** `authorCellProtection` + `worksheetPasswordProtectionReport` lit
     `cell-protection-locked-flag-and-sheet-protection` (all 5), `sheet-protection-permits-requested-operations`
     (all 3), and `worksheet-password-protection-hashes-in-node` (all 4). Author-facing allow-flags invert to
     OOXML forbidden booleans; password → OOXML-agile SHA-512 credential via `src/core/protection.ts`
     (`node:crypto`). Band unlock is currently adapter-expanded onto listed cells; native column-scope
     inheritance of protection is folded into the column-scope family below. **Reading `<sheetProtection>`
     back into the model is not yet done** (no case needs it — a read→re-write of a password-protected sheet
     would today lose the credential we cannot re-derive; capture that when foreign-fixture reading lands).
   - ✅ **Copy-on-write style aliasing** (`loadMutateCellFacet`/`loadMutateCellStyle`/`loadMutateCellFont`/
     `loadMutateCellBorder`/`sharedBaseStyleFontMutation`) — **DONE, no source change needed.** The rewrite
     already isolates by construction (each cell owns independent facet fields, every setter REPLACES the
     field, facet types are `readonly` so a shared record is inert), so wiring the five adapter methods
     resolved 11 legacy known-opens across `loaded-cells-shared-style-object-aliasing`,
     `shared-base-style-font-mutation-isolated`, `cell-border-mutation-does-not-bleed-to-style-siblings`, and
     `cell-style-setter-isolates-alignment-numfmt-protection` (alignment+numFmt+protection). Because those
     baselines are `fail`, a future regression would drop to `○` not `✗`, so the guarantee is hard-locked in
     src by `src/io/xlsx/style-isolation.test.ts` (5 tests, real write→read path).
   - ✅ **Column-scope style inheritance** for facets beyond numFmt — **DONE.** `ColumnProperties` now carries
     the full facet bundle (fill/font/border/alignment/protection, not just numFmt); the writer composes each
     cell's style inheriting every column facet (cell ?? row ?? column) and interns the full bundle into the
     `<col>` style; the reader mirrors a `<col style>`'s full bundle onto the column model (a bare cell already
     inherited the full column xf). Flipped `alignment-does-not-leak-across-cells` (column behaviors) and
     `column-border-style-scoped-to-declaring-column` to green regression locks, plus three src round-trip
     tests in `read.test.ts`. (Native column-scope inheritance of protection bands — currently adapter-expanded
     onto listed cells — now has the machinery; wire it when a case demands it.)
   - ✅ **Foreign-fixture reading** (fourth capability family, 115 → 143) — **FULLY CLOSED.** Three bulk slices,
     mostly adapter-only: (1) real-fixture colour reading (`readFixtureCellStyles`/`roundtripFixtureColorFidelity`
     → `solid-fill-foreground-vs-font-color`, `theme-and-rgb-fill-colors-read-faithfully`,
     `fill-border-color-survives-roundtrip` — theme+tint/indexed colours survive a pure open-then-save);
     (2) foreign explicit-off font/alignment forms (`fontExplicitFalse*`/`alignmentFalseBooleanReport` via a
     `reloadPatched` package-patching helper → `font-boolean-flag-honors-explicit-false`,
     `alignment-false-boolean-attrs-yield-no-alignment`; the ONE source fix was `<u val="none">` → falsy);
     (3) `readFixtureReport` — the broad reader-robustness probe, resolving 28 legacy known-opens across ~two
     dozen fixtures (namespace-prefixed roots, BOMs, non-ASCII names, unusual part order, missing optional
     parts, foreign boolean spellings) with **no source change** — the lean SAX reader is robust by construction.
     Plus the two deferred follow-ups, both now DONE: `unstyledCellFontReport` (workbook default font for every
     unstyled cell — its own default-font slice) and `roundtripFixture` (`template-styles-survive-read-write-roundtrip`
     — a real styled template survives a no-op round-trip; adapter-only, no source change).

   The remaining **border** case needing its own capability is merge-master survival
   (`merged-region-master-cell-border-survives`). (Column-scope, shared-style aliasing, and foreign-fixture
   colour-fidelity border cases — incl. `fill-border-color-survives-roundtrip` — are now green; see the
   column-scope, copy-on-write, and foreign-fixture families.) (All harvested **font** cases are now green or `↑`: `unstyledCellFontReport` — resolve `fontId` 0
   to the concrete default font for every unstyled cell — landed in the default-font slice, alongside
   `fontExplicitFalse*` and `sharedBaseStyleFontMutation`.) Also worth a small slice: ARGB validation (reject/normalize
   a `#`-prefixed fill colour — `solid-fill-argb-rejects-hash-prefix`), and writing a filled-but-empty cell
   (today the row loop drops a cell with a fill/font but null value — needed before `probeCellFonts` could
   round-trip a valueless styled cell).
2. ✅ **Foreign-fixture reading** — **FULLY CLOSED** (see the foreign-fixture family under §1: real-fixture
   colour reading, foreign explicit-off forms, `readFixtureReport`, plus `unstyledCellFontReport` default-font
   resolution and `roundtripFixture`; 115 → 143).
3. **Remaining writer/core-model widening:** **`addRow`/`addRows` shapes** (dense/sparse-array +
   keyed-object — `add-row-array-and-object-shapes-populate`), **defined-name storage** →
   `<definedNames>`. Small follow-ups: modern-function `_xlfn.` prefix, outline summary-row `collapsed`
   inference (both `○`), and the **streaming reader** (where the *running-counter* inflate bound lands —
   the buffered reader only has the declared-size cap; see ADR 0004).
4. **Toolchain-standup slice** (do when `src/` is large enough to justify it): Vitest + Biome +
   an ESM/`.d.ts` bundler, and schedule the legacy Grunt/Babel/Mocha rip-out. Until then the gates
   are `npm run typecheck` + `npm run test:src` + `npm run corpus:rewrite`.

**Reserved for the human (not blocking the rewrite):** open decision #1 (now optional — see above)
and the final brand name (Phase 4). **Housekeeping:** per `STRATEGY.md` we no longer track
`exceljs/exceljs` (frozen universe, no re-harvest); the `upstream` remote can be dropped anytime.
