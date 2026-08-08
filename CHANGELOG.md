# Changelog

All notable changes to this project are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project adheres to
[Semantic Versioning](https://semver.org/) — see
[ADR-0015](docs/decisions/0015-publishing-name-semver-and-first-version.md) for the
versioning policy, including why the first published version is `1.0.0` rather than a
`0.x` series.

This file tracks changes from its introduction forward. Earlier history — the full
ExcelJS-to-`ts-xlsx` rewrite — is recorded in `git log` and the [ADR series](docs/decisions/).

## [Unreleased]

## [1.1.0] — 2026-08-08

The first release since 1.0.0 to add API surface. Six additions, each of them something a
caller was otherwise assembling by hand: a worksheet lookup that throws instead of returning
`undefined`, the sheet's extent as a single range, one plain-text rendering of a cell value,
the grid's measured geometry ceilings, a wrapped-line count for authoring a row height, and
the `CellValue` type guards — which the reader and writer had always used but the barrel never
exported. Nothing is removed and no existing signature changes. The CSV writer now shares the
value renderer rather than keeping a near-copy of it, which shifts two of its behaviours; both
are recorded below.

### Added

- **`Workbook.requireWorksheet` — the total counterpart of `getWorksheet`.** Same lookup, by name
  (case-insensitive) or numeric id, but a miss throws an `AuthoringError` naming every sheet the
  workbook does have instead of returning `undefined` that flows on into a `?.` chain and fails
  several steps later with nothing left to say. A lookup miss is a typo, a stale template or a
  renamed tab, and all three are answered by seeing the real names. The CSV writer now uses it, so
  its "no worksheet named …" message gains that listing and loses the "to write as CSV" suffix.
- **`estimateWrappedLines` — how many lines a wrapped cell takes.** A row that states no height
  records no geometry, so its height is whoever opens the file's answer. Excel's answer is an
  auto-fit computed on open — measured, and prompt: it saturates at `MAX_ROW_HEIGHT`, so past ~28
  lines it stops answering, and a consumer that does not implement it (this library's reader
  included) has no height at all. Writing one settles the geometry, and needs a line count. Counts
  characters against a character-unit width — the same unit a column's `width` is in, so no font
  metric is assumed and the deferred metric-table question stays deferred. Exact for a monospaced
  face that wraps mid-word, and against Excel a shade low (5 lines where Excel laid out 6), which
  the doc states. A hard break opens a line of its own; a zero or non-finite width is a `RangeError`
  rather than an `Infinity` that would land in a row height. The auto-fit measurements are in
  [`docs/knowledge/specs/rows-with-no-stated-height-are-autofitted-on-open.md`](docs/knowledge/specs/rows-with-no-stated-height-are-autofitted-on-open.md).
- **`MAX_ROW_HEIGHT` and `MAX_COLUMN_WIDTH` — the grid's geometry limits, measured.** The
  companions to `MAX_ROW`/`MAX_COLUMN`: how large a line may be set, where those two bound where
  a cell may be. `ht` and `width` are a bare `xsd:double` in the schema, so the ceiling is Excel's
  own, and it is not the one Microsoft's specifications table publishes — Excel Desktop accepts a
  row height of 409.5 and refuses 409.6, against a documented "409 points". Column width is exactly
  255. Both were measured over COM rather than quoted; the probe and its numbers are in
  [`docs/knowledge/specs/grid-geometry-limits-are-excels-not-the-schemas.md`](docs/knowledge/specs/grid-geometry-limits-are-excels-not-the-schemas.md).
  They bound *assignment*, not a file, and the two behave differently once one exceeds them: Excel
  opens an over-limit package clean, silently clamping a row to 409.6 (a tick above what it lets
  you set) while honouring a `width` of 1000 and re-saving it verbatim. Nothing enforces them at
  either end, therefore — `Row.height` and `Column.width` are the reader's path into a foreign
  file as well as an author's, so a bound that threw would refuse files Excel opens. There is
  deliberately no `DEFAULT_COLUMN_WIDTH` beside them: the default width follows the workbook's
  default font (8.43 for Calibri 11, 8.09 for Aptos Narrow 11), so a constant would be a wrong
  answer wearing a right one's name.
- **`Worksheet.usedRange` — the sheet's extent as a handle.** `rowCount` and `columnCount` said
  once: `A1` through the last row and column carrying anything, or `undefined` when the sheet spans
  no rectangle. It replaces the `` `A1:${numberToColumn(sheet.columnCount)}${sheet.rowCount}` ``
  every caller was assembling by hand, and it is what an auto-filter over a whole sheet wants —
  `sheet.autoFilter = sheet.usedRange.address`, where a header-only ref yields dropdowns that
  filter nothing. Anchored at `A1` and inheriting both counts' definition of *used*, so it is not
  the tight `<dimension>` box a written package records; the doc comment states the difference.
- **`cellValueToText` and `Cell.text` — one plain-text rendering of a value, for everyone.**
  `cellValueToText` is total over `CellValue`: the empty cell and an invalid `Date` give `""`, a
  boolean gives Excel's `TRUE`/`FALSE`, an error its literal, rich text its runs concatenated, a
  hyperlink its label, and any of the three formula kinds the text of its cached result. It is the
  value's text, not the cell's *displayed* text — no number format is applied, because the format
  lives on the style. `cell.text` is the same answer for the cell you are holding. The CSV writer
  now renders its fields through it rather than through a private near-copy, so a CSV field and
  `cell.text` cannot disagree about the same cell; `dateFormat`/`dateUTC` remain a CSV-only
  deviation. One consequence of the merge: a data-table formula's cached result now reaches a CSV
  field, where it used to render as empty.
- **The `CellValue` type guards are public.** `isErrorValue`, `isFormulaValue`,
  `isSharedFormulaValue`, `isDataTableFormulaValue`, `isRichTextValue` and `isHyperlinkValue` have
  always existed — the reader and the writer discriminate the union with them — but the `/core`
  barrel published only `detectValueType`. A consumer holding a `CellValue` therefore had no
  narrowing primitive at all: `detectValueType` classifies but does not narrow, so reading
  `.richText` off a value meant hand-rolling `'richText' in value` (which admits far more than the
  guard does) or an `as` cast. Now exported, each with the narrowing target pinned by a type-level
  contract.

### Changed

- **A release now publishes unattended.** The `npm-publish` deployment environment required a
  human to approve the job before it could reach the registry. On a single-maintainer project
  that was never a second pair of eyes — the only login that could approve was the one that had
  just cut the release — so it delayed every publish to re-ask a decision already made. Removed
  in `environment.yml`; deployments are still restricted to `v*` tags, which is now the whole
  gate. Nothing about the published package changes, and the OIDC trusted-publishing identity is
  untouched: authentication was never what the approval provided. See
  [ADR-0026](docs/decisions/0026-releasing-is-a-github-release-and-npm-follows.md) for what that
  trades away.

## [1.0.3] — 2026-08-05

Nothing in the library's behaviour changed: `src/` differs from 1.0.2 only in doc comments.
What reaches a consumer is the JSDoc the package ships — which now states what each documented
error is — and `dist/` built by a different compiler.

### Changed

- **The toolchain is on TypeScript 7.** `typescript@^7.0.2` — the native Go compiler — replaces
  the 6.0 line that [ADR-0008](docs/decisions/0008-typescript-6-upgrade.md) settled on, now that
  the docs generator no longer needs the printer API that 7 does not ship. The two scripts using
  the compiler programmatically moved to `typescript/unstable/*`; the type gate got ~4.5× faster.
  `src/` is untouched and the published API is unchanged. The emitted JS in `dist/` now quotes
  import specifiers with single quotes rather than double — the emitter's choice, no behavioural
  difference, and the full corpus passes against the new output. See
  [ADR-0028](docs/decisions/0028-typescript-7-adoption.md).

- **Contributor tooling: authored text may no longer contain characters that render as
  something other than what they mean.** `src/core/range.ts` carried a literal U+0000 byte
  as a sentinel where the six-character escape was meant. The two are identical to the
  compiler, and opposite to every text tool downstream: grep answered `Binary file … matches`
  instead of the matching lines, so the file silently dropped out of searches while still
  appearing to have been searched — which is how its two `@throws` tags survived the audit
  below. The sentinel's spelling changed, not its value. `scripts/check-source-text.ts` joins
  the invariants gate, refusing C0 controls other than tab/LF/CR, DEL, and bidirectional
  overrides (CVE-2021-42574) across `src`, `scripts`, `test`, `tools` and `docs`. Nothing in
  the published package changes.

### Fixed

- **The API reference now says what each documented error is, and no longer truncates the
  description.** Every `**Throws**` line in `docs/api/` was missing its error type: the generator
  matched the `{ErrorType}` slot only to strip it, so the reader was told a throw happens but
  never told what is thrown. Separately, 25 tags across 10 modules were written
  `@throws {@link SomeError}` — TypeScript parses the braces after `@throws` as a type
  expression, a `{@link …}` is not one, and the parse ran past the close brace and ate the rest
  of the comment. Those descriptions were missing from editor hovers too, not only from the
  generated pages; six of them rendered as the bare text `{`. The tags now use the brace-slot
  spelling the other 45 already used, the type is rendered rather than discarded, and `gen-docs`
  fails the build on a slot that is not a type name so the shape cannot return. Links
  mid-sentence were never affected and are unchanged.

  Two further faults in the same generator are fixed, both `docs/api/` only. A class member
  rendered its description and none of its tags, so ~40 `@throws` — including every one on
  `Workbook` and `Worksheet`, the two pages a caller is most likely to open — reached the
  reference never; members now get their own block, as top-level functions always had. And
  `{@link Target}` was parsed and then flattened to a bare code span, leaving the reference
  with zero links outside its index; 433 now resolve to the page and heading that documents
  the target, and `gen-docs` fails before writing anything if one would dangle.

- **A publish rehearsal now fails on a rejected identity.** `npm publish --dry-run` demotes a
  failed OIDC token exchange to a warning and exits `0`, so the rehearsal reported success for
  the one failure it exists to catch. That is not hypothetical: the 2026-07-30 rehearsal went
  green while its exchange returned `404`, which is how a misconfigured trusted publisher
  survived the check and cost two version numbers before anyone read the log. `publish.yml`
  now reads the verbose log and fails the job when the exchange was rejected — or when it was
  never attempted at all, the shape that made 1.0.1 fail `ENEEDAUTH`. Nothing in the library
  changed; `src/` is untouched.

  (1.0.2 itself reached npm on 2026-08-05, with provenance, once the publisher's `environment`
  field was corrected on npmjs.com. The workflow needed no change for it.)

## [1.0.2] — 2026-07-29

### Changed

- **Still nothing in the library — `src/` remains byte-identical to 1.0.0.** 1.0.1 was
  tagged and released on GitHub but **never reached npm**: the publish workflow failed
  `ENEEDAUTH`, having dropped `setup-node`'s `registry-url` on the incorrect theory that it
  would make npm send a placeholder token instead of exchanging an OIDC one. It is in fact
  what tells npm which registry to authenticate against, and without it the exchange never
  starts. Restored.

  1.0.2's own first attempt then failed too, for a second and unrelated reason — this one
  outside the repository. The trusted publisher configured on npmjs.com named
  `environment.yml`, the workflow that *provisions* the deployment environment, where it had
  to name the environment itself, `npm-publish`. npm answers a rejected identity with a 404,
  which reads as "no such package" and hides which claim failed to match. With the publisher
  corrected on npm, this tag publishes on a re-run; nothing here changed for it.

  The tag and release for 1.0.1 are left standing rather than rewritten: a published claim
  that turned out to be wrong is corrected in the open, not deleted. Install 1.0.2.

## [1.0.1] — 2026-07-29 · never published to npm

### Changed

- **Nothing in the library. This release exists to prove the release path.** `src/` is
  byte-identical to 1.0.0 — `npm diff` between the two will show only the version field.
  What changed is how a release reaches you: publishing a GitHub release now publishes the
  package, authenticated by a short-lived token minted from the workflow's OIDC identity
  rather than by any credential stored in the repository, and the result carries a
  provenance attestation tying it to the commit and the run that built it. 1.0.0 was
  published by hand and has no such attestation.

  A patch bump with no behavior change is the honest way to exercise that path end to end:
  the alternative is discovering a broken publish on a release that actually matters. See
  [ADR-0026](docs/decisions/0026-releasing-is-a-github-release-and-npm-follows.md).

## [1.0.0] — 2026-07-29

The first published release. `ts-xlsx` is an independent hard fork of ExcelJS, rebuilt from
the ground up in strict TypeScript; it carries no backwards-compatibility guarantee with its
ancestor — see [migrating from ExcelJS](docs/migrating-from-exceljs.md). The **BREAKING**
markers below describe changes made against the unreleased development line while this
section accumulated, not against any previously published version; there is no earlier
release of this package to break. From here forward the project follows
[SemVer](https://semver.org/) strictly — see
[ADR-0015](docs/decisions/0015-publishing-name-semver-and-first-version.md).

Charts, vector shapes, slicers and legacy form controls are **round-trip-only** in 1.0: a
workbook carrying them survives a load/edit/save byte-faithfully, but there is no API to
author a new one ([ADR-0014](docs/decisions/0014-charts-shapes-slicers-are-round-trip-only-for-1-0.md)).

### Added

- **`writeXlsxAsync` — the same package, deflated off the calling thread.** `writeXlsx` spends the
  whole cost of DEFLATE on the caller's thread, which on a large workbook means the event loop does
  not tick for seconds. The async writer shares part-building with the sync one and differs only in
  handing the part map to `fflate`'s worker-backed zip: every part compresses to identical bytes.
  Measured on a ~42 MB part map, a single large sheet takes the same wall-clock but the longest stall
  falls from the entire write to ~17 ms, and a twenty-sheet workbook finishes ~2.4× sooner because
  its parts deflate in parallel.

  There is deliberately no `readXlsxAsync` to match: reading is dominated by XML parsing and model
  building, which no worker can take, and the reader's zip-bomb ceiling depends on counting output
  between synchronous input slices. The asymmetry is the decision, not an omission —
  [ADR-0024](docs/decisions/0024-async-is-one-writer-not-a-mirrored-pair.md).

- **BREAKING: `Row` and `Column` are real classes, and `getRow`/`getColumn` return them.** They were
  formatting bags: `sheet.getRow(2)` handed back a `RowProperties` record with no way to reach the
  row's cells, so every row-oriented consumer hand-wrote an address-encoding loop, and `getRow` and
  `addRow` disagreed about what a row even was. A handle now carries both — `row.getCell('B')`,
  `row.cells`, `row.values` alongside `row.height`/`hidden`/`outlineLevel`/`collapsed`/`fill`, and the
  column equivalent including `key`, `width` and the six style facets it defaults for its cells.

  Formatting reads and writes exactly as before (`sheet.getRow(2).height = 20`,
  `Object.assign(sheet.getColumn(1), {key, width})`), and destructuring the iterators is unchanged
  (`for (const {number, cells} of sheet.rows())`). What breaks: the declared return type,
  `sheet.rowProperties(n)` / `sheet.columnProperties(n)` (use `getRow(n).properties` /
  `getColumn(n).properties`, which no longer fabricate either), and spreading a handle —
  `{...sheet.getRow(2)}` is no longer that row's properties.

- **`getRow` and `getColumn` no longer extend the used range.** They created a format record on
  access, so merely *asking* about row 500 made `rowCount` 500 and put an empty record in the
  worksheet model. The record is now created on first write. Reading a `<row r="5"/>` that states no
  attributes likewise leaves nothing behind, which is the honest reading of an element that says
  nothing.

- **`Cell.setRichText` — rich-text runs that inherit the cell's font.** A run's `<rPr>` is a
  *complete* character format: a facet it omits falls back to the workbook default font, **not** to
  the cell's. Verified against Excel — a cell set to Courier New 16 whose first run carries only
  `<b/>` renders that run in the workbook default face at the default size. So authoring
  `{bold: true}` on a run beside a styled cell silently loses the typeface, and the only fix was to
  restate the whole font on every run.

  `cell.setRichText([{text: 'Note:', font: {bold: true}}, {text: ' the rest'}])` composes each run
  over the cell's own font, per facet. Assigning `cell.value` directly is unchanged and stays the
  bare path, for a caller who wants a run that deliberately falls back to the workbook default —
  which is the format's rule, so the writer still emits exactly the facets a run carries.

- **`Range` — style a rectangular block in one call.** `sheet.getRange('A2:A33').border = {...}`, or
  `getRange(2, 1, 33, 1)` by inclusive corners. The third handle beside `Row` and `Column`, with the
  same contract: constructing one creates nothing, `addresses()` walks the block as a generator, and
  `cells` reports only what already exists. The six style facets plus a composing `style` accessor and
  `clearStyle()` mirror `Cell`'s semantics exactly — assigning a facet replaces that facet, assigning
  `style` composes facet by facet — so there is no second convention to learn.

  Writing materialises every position in the block, because a styled-but-valueless cell is the only
  way an empty cell renders with a fill; a uniformly styled block still collapses to one shared style
  entry. The cost is bounded by construction: `A:A` and `1:1` are refused, pointing at
  `getColumn`/`getRow`, which state a whole-axis default in one attribute instead of a million cells.
  A block overlapping a merged region restyles the region's master rather than stranding a style on a
  covered cell.

- **The workbook's default font is a first-class part of the model.** `Workbook.setDefaultFont`
  authors the face, size and colour every cell with no font of its own renders in — **empty cells
  included** — merging like `setTheme`, so `setDefaultFont({size: 14})` keeps the resolved face.
  `Workbook.defaultFont` reports the resolved, complete result, and `Workbook.declaredDefaultFont`
  reports what a source package stated (`undefined` when it stated nothing).


- **Seven subpath entry points, and `"sideEffects": false`.** `@shbernal/ts-xlsx/core`, `/xlsx`,
  `/xlsb`, `/csv`, `/vba`, `/customui` and `/errors` are published alongside the bare package name,
  which still exports everything it did. Additive — nothing moves or breaks. With a bundler the
  bare name remains the right default (`sideEffects: false` now lets it prune per symbol, which
  beats any subpath); reach for a subpath when there is no bundler, or when the module graph should
  state the dependency. `/errors` carries every error class the library throws and costs 12 KB, so
  classifying a failure never loads a parser; `/core` is 332 KB against the package's 902 KB.
  `/xlsx` is only marginally cheaper than everything, because `readXlsx` sniffs the bytes and
  dispatches a binary package to the BIFF12 reader — see the table in the README, and
  [ADR-0023](docs/decisions/0023-subpath-entry-points-and-disjoint-barrels.md) for why the split
  is shaped this way.

- **`tools/vba-compiler`** — an offline build tool that produces genuinely compiled, source-matched VBA
  p-code by driving a real headless Excel (VBIDE). Emits a `vbaProject.bin` (attach via
  `Workbook.vbaProjectBytes`) or a whole edited `.xlsm`. Windows + licensed Excel only; never in CI.


- **BREAKING: every deliberate failure now descends from one `XlsxError`, and carries a `code`.**
  `catch (e) { if (e instanceof XlsxError) … }` is the whole answer to "was that this library?" —
  previously five typed classes shared no ancestor, and the model's own validation threw bare `Error`
  distinguishable only by string-matching the message. `error.code` is the coarse branch
  (`'unsupported-format'` | `'malformed-input'` | `'authoring'` | `'internal'`); `error.name` and
  `instanceof` remain the exact one. See "How a failure is reported" in
  [docs/architecture.md](docs/architecture.md).

  New classes: `AuthoringError` (a document that cannot exist — 50 sites that used to throw bare
  `Error`), `PackageReadError` (a refused inflate, previously recognised by a message prefix),
  `XmlParseError` (malformed markup, previously a native `SyntaxError` that escaped `readXlsx`
  indistinguishable from a caller's own), `XlsxParseError`, and `InternalError` (an invariant of ours
  that did not hold). The five existing classes — `UnsupportedFormatError`, `XlsbParseError`,
  `VbaParseError`, `VbaAuthorError`, `CustomUiParseError` — keep their names, messages and fields and
  gain the ancestry.

  What did **not** change: scalar argument validation stays native `RangeError` / `SyntaxError` /
  `TypeError`, and every error message is byte-identical. The break is the *type* of a caught error,
  which matters if you catch `SyntaxError` around XML parsing or switch on `constructor`.

- **`Workbook.addTableStyle({name, elements})` — custom table styles are authorable.** A workbook can
  now define its own named table styles beside Excel's built-in gallery, and a table reaches one by
  putting that name in `TableStyleInfo.name`. Each element names a region (`wholeTable`, `headerRow`,
  `firstRowStripe`, … — all 28 of `ST_TableStyleType`) and carries a `DifferentialStyle`, interned into
  the same shared `<dxfs>` table conditional formatting uses, so two elements painted alike cost one
  entry. A stripe may set its band width with `size`. Authoring a name a source file already defined
  overrides that definition rather than adding an ambiguous second one.

  Verified against Excel Desktop, not just the schema: a table style is a cross-part correspondence
  every part of which can be valid while the table still renders unstyled. Excel registers the
  authored style in the workbook's gallery and paints from it — see
  `docs/knowledge/specs/custom-table-styles.md`.

  An unnamed style, or a `size` outside the four stripe types, throws — both otherwise produce a file
  Excel opens cleanly and then ignores. `TableStyleInfo.name` itself stays unvalidated, deliberately;
  the type documents why.

- **`Workbook.setTheme({colors, fonts})` — the workbook's palette is authorable.** A colour picked
  from a spreadsheet's theme row is written as `theme="4"`, a reference resolved at render time, so
  setting `accent1` restyles every cell, chart and table style that follows the theme at once — the
  only way to recolour a workbook without touching a cell. Any subset of the twelve colour-scheme
  slots and either of the two typefaces can be set, and calls merge.

  It generates *over* the existing theme rather than replacing it. The format scheme — the gradient,
  line and effect styles a designer authored — rides through untouched, a slot left unnamed keeps its
  source encoding (`dk1`/`lt1` stay `<a:sysClr>`, so they still follow the viewer's window colours),
  and a theme that references a picture keeps that relationship. `Workbook.themeColors` and
  `themeFonts` report the effective theme. A malformed colour throws at the setter, because Excel does
  not report one — it renders the slot as flat black.

- **`Workbook.resolveColor(color)` — a themed or indexed colour now resolves to a concrete ARGB.**
  A `Color` read from a file often carries no colour at all, only a reference: `{theme: 4}` into the
  workbook theme's scheme, or `{indexed: 2}` into the legacy 64-entry palette, either optionally with
  a `tint`. Resolution follows the workbook's *own* theme and its own custom `<indexedColors>` when it
  declares one, and applies the tint last. `Workbook.themeColors` exposes the scheme it resolves
  against. Two things it deliberately does not do: `indexed="64"`/`65` (the system foreground and
  background) resolve to `undefined` rather than to invented black and white, and nothing is written
  back into the model — the `Color` keeps the encoding its file used, so a round-trip still emits
  `theme="4" tint="0.4"` and the cell keeps its link to the theme.

  Note the index order: `theme="0"` is `lt1` and `theme="1"` is `dk1`, which is *not* the order the
  slots appear in the theme part. Verified against Excel Desktop; see
  `docs/knowledge/specs/theme-color-index-order.md`.

### Changed

- **BREAKING: the styles part's font 0 is the workbook's own default font, not an assumed Calibri.**
  The writer used to splice a `Calibri 11` constant into `<fonts>` with no workbook input, and two
  things followed from that.

  A themed workbook rendered every unstyled cell in the wrong face: `setTheme({fonts: {minor}})`
  wrote the theme part correctly and could not reach a cell, because font 0 went on claiming
  `scheme="minor"` — *I am the theme's body face* — while naming Calibri outright, and Excel resolves
  the explicit name. Working around it meant setting `font` on every column and naming the face in
  every rich-text run. And reading a package whose font 0 was Aptos Narrow and writing it back
  *unmodified* replaced the declared default with Calibri, re-adding the real face as a redundant
  custom entry — so populated cells still rendered right while empty cells, and the metric every
  character-unit `<col width>` is expressed in, quietly changed.

  Font 0 now resolves from the workbook: an authored default font, else an authored theme body face,
  else the file's own font 0, else the theme's. `scheme` and `family` are derived rather than
  copied — carried while the resolved face really is the theme's body face and dropped when it is
  not, which is what Excel itself writes — and the emitted entry always states a size and a colour,
  the absence foreign readers report as a "missing default font".

  What breaks: a workbook with an authored theme font, or read from a package whose font 0 was not
  Calibri, now emits different bytes. A plain `new Workbook()` is byte-identical.
  [ADR-0025](docs/decisions/0025-the-default-font-is-declared-not-assumed.md).

- **BREAKING: a corrupt or truncated package is a `PackageReadError`, not an unsupported format.** A
  `PK`-headed archive the zip layer rejects used to surface as `UnsupportedFormatError` with format
  `'unknown'`, whose message reads *"not a valid .xlsx package: no OOXML workbook part was found"* —
  a check that never ran, since nothing inflated. It now carries `code: 'malformed-input'` alongside
  the zip-bomb refusal, which is what the taxonomy already said it was: the container is the right
  kind of thing and we cannot unpack it. Code branching on `format === 'unknown'` for a truncated
  file must catch `PackageReadError` instead.

  Two messages get honest with it. A non-ZIP blob now says the input is not a ZIP rather than blaming
  a missing part, and *"no OOXML workbook part was found"* is left to the one case where it is true —
  a package that inflated and carries neither `xl/workbook.xml` nor `xl/workbook.bin`. Unchanged: the
  zip library's own text is discarded, never folded into the message and never attached as `cause`,
  because it can name internals or an absolute filesystem path.


- `removeVbaModule` and `addVbaReference` (and their `Workbook`/`editXlsxVba*` wrappers) no longer reset
  the `_VBA_PROJECT` stream — that reset crashed the VBA load on a project with real p-code. They now
  leave it byte-for-byte untouched; the `dir` stream carries the structural change. These edits are
  retained precisely because they never touch a module's compiled p-code.

- **Contributor tooling: the gate set has one name.** `node scripts/verify.ts` runs every gate
  concurrently and is what `pnpm test`, lefthook's `pre-push` and the turn-boundary hook all
  invoke; `--quick` is the inner loop and `--cached` exits immediately when the working tree is
  byte-for-byte the one that last passed. The corpus runner gained `--case`/`--json` and now
  prints only what needs attention (`--verbose` for the old listing), and the OOXML validator
  builds on demand instead of paying `dotnet run`'s project re-evaluation per call. Nothing in
  the published package changes. See
  [ADR 0022](docs/decisions/0022-verification-is-one-cached-parallel-entrypoint.md).

### Removed

- **BREAKING: the pure-TS VBA source-authoring API is gone.** `writeVbaProject`,
  `Workbook.setVbaProject`, `editVbaModuleSources`, `Workbook.setVbaModuleSource`,
  `editXlsxVbaModuleSource`, `editXlsxVbaModuleSources`, `addVbaModule`, `Workbook.addVbaModule`, and
  `editXlsxVbaAddModule` are removed, along with the `VbaModuleSource`/`VbaProjectSpec` types. They
  emitted modules with no compiled p-code on the theory that Excel recompiles from source on open —
  which it does not: such files either fail to load ("Invalid data format") or silently run stale code.
  See [ADR 0019](docs/decisions/0019-vba-authoring-needs-real-pcode-recompile-cookie-retracted.md).

### Fixed

- **A workbook's theme part is no longer destroyed on round-trip.** Reading a file and writing it
  straight back replaced its theme with the default Office one, so every `theme="n"` colour and
  `scheme="major|minor"` font in the file silently re-rendered in the wrong brand — a branded
  workbook came back in Office blue. The source theme is now preserved verbatim, reached through the
  workbook's `.../theme` relationship rather than the conventional `xl/theme/theme1.xml` path, and
  re-emitted with the parts it references (a picture used as a themed fill) so its `r:embed` does not
  dangle. Exposed on the model as `Workbook.themePart` / `restoreThemePart`, opaque preserved XML in
  the same spirit as `restoreDifferentialStyles`.

- **Custom table styles and recent-colour swatches are no longer dropped on round-trip.** A workbook's
  `<tableStyles>` definitions and the default table/pivot styles it nominates were discarded when the
  stylesheet was regenerated, so a table asking for a custom style by name was left referencing
  nothing and rendered completely unstyled — a file that opens clean and looks wrong.
  `<colors><mruColors>` (the "Recent Colors" swatches) went the same way. Both are now preserved
  verbatim and exposed as `Workbook.tableStyles` / `restoreTableStyles` and `Workbook.mruColors` /
  `restoreMruColors`. Each `tableStyleElement`'s `dxfId` keeps resolving because the differential-style
  table is re-emitted at its original indices, and the namespace prefixes Excel stamps on a table style
  (`xr9:uid`) are re-declared on the stylesheet root rather than left dangling.

[Unreleased]: https://github.com/shbernal/ts-xlsx/compare/v1.1.0...HEAD
[1.1.0]: https://github.com/shbernal/ts-xlsx/compare/v1.0.3...v1.1.0
[1.0.3]: https://github.com/shbernal/ts-xlsx/compare/v1.0.2...v1.0.3
[1.0.2]: https://github.com/shbernal/ts-xlsx/compare/v1.0.1...v1.0.2
[1.0.1]: https://github.com/shbernal/ts-xlsx/compare/v1.0.0...v1.0.1
[1.0.0]: https://github.com/shbernal/ts-xlsx/releases/tag/v1.0.0
