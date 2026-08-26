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

### Added

- **`Workbook.exportImages(sheet)` / `Workbook.importImages(sheet, images)` — carry a sheet's
  pictures to another workbook.** An anchored image holds a media *id* into one workbook's registry,
  and that id names a different picture, or none, in the next. Copying a sheet has therefore always
  left its images behind: `dst.model = src.model` is a semantic copy by design
  ([ADR-0005](docs/decisions/0005-worksheet-model-is-semantic-only.md)), and there was no affordance
  for the bytes. Now there is one, and carrying a sheet whole is the two together:

  ```ts
  destination.model = source.model;
  destinationWorkbook.importImages(destination, sourceWorkbook.exportImages(source));
  ```

  `exportImages` resolves a sheet's anchored images and its background out of the media registry into
  a workbook-independent `WorksheetImages` value; `importImages` registers those bytes in the
  destination and re-anchors against the ids they land on. Registration is content-addressed, so a
  picture already held is re-used rather than stored twice — twenty sheets sharing one logo cost one
  media part. An import replaces the destination sheet's pictures rather than appending to them, the
  same direction a `model` assignment goes.

  `exportImages` throws `AuthoringError` for a sheet whose image ids this workbook never registered —
  which is what a sheet belonging to *another* workbook looks like from here. That pairs with the
  existing write-time check, so an anchor holding a foreign media id is refused by name at both ends
  instead of being emitted as a drawing relationship pointing at media that was never written.

  This covers floating anchors and the sheet background. Header/footer images are byte-preserved
  parts and ride the preserved-reference machinery; charts, vector drawings, pivots and slicers stay
  round-trip-only ([ADR-0014](docs/decisions/0014-charts-shapes-slicers-are-round-trip-only-for-1-0.md)) —
  their bytes carry the source sheet's identity (a pivot table's workbook-unique name, a slicer
  cache's `tabId`, a chart series' `Sheet1!$A$1`), so they cannot be relocated by copying bytes.

### Fixed

- **A malformed frozen-pane split no longer surfaces from the writer.** A `<pane>` whose `xSplit`
  or `ySplit` spelled a fraction, a negative, or a word was stored verbatim, though `freeze()`
  refuses all three. The read succeeded and the *serializer* then threw a `RangeError` naming a
  column the file never mentioned. A split that is not a non-negative integer is now dropped where
  it is read, leaving the rest of the sheet view intact.

- **A `<col>` span wider than the sheet no longer hangs the reader.** `<col min="1"
  max="99999999"/>` is one line of XML that named more columns than the format has, and the
  buffered reader walked it verbatim: roughly 16.7 million column records, twenty-five seconds,
  and then `RangeError: Map maximum size exceeded`. A denial of service on a one-line input.
  The span is now clamped to the last real column and an element wholly outside the grid is
  dropped, which is what the streaming reader already did with the same element. A `<row r>`
  past the last row is dropped for the same reason.

- **A control character in a string no longer produces a package Excel reports as damaged.**
  The writer escaped `& < > " '` and nothing else, so every other character reached the file
  byte for byte — including the C0 controls XML 1.0 forbids outright and the U+FFFE/U+FFFF
  noncharacters. A cell value holding U+0001 emitted a sheet part Microsoft's
  `OpenXmlValidator` rejects. Strings arriving from a database column, a CSV field, or a user
  form are exactly where these turn up.

  A **cell value** is now escaped with SpreadsheetML's `_xHHHH_` convention, which is what
  Excel writes and what Excel reads back, so the character survives the round-trip instead of
  breaking the file. That covers inline strings, pooled strings, rich-text runs, legacy note
  text, and the cached result of a string formula.

  A lone surrogate rode along in the same fix. It never produced an invalid package — the
  UTF-8 encoder substituted U+FFFD for it, so the file validated and the character was
  already gone. It is now escaped and preserved like the rest.

  Because `_xHHHH_` now carries meaning, a cell value that legitimately contains the literal
  text `_x0041_` has its underscore escaped as `_x005F_x0041_`, or it would read back as the
  letter `A`. Text that only resembles an escape (`_`, `_x`, `_xZZZZ_`) is untouched.

  The reader undoes the same convention, which is what makes the round-trip claim above true
  and also changes how *foreign* files read. A workbook Excel authored with a control character
  in a cell previously read back as the literal seven-character text `_x0001_`; it now reads
  back as the character. That applies to inline strings, the shared-strings pool, rich-text
  runs, legacy note text, and the cached result of a string formula, in both the buffered and
  the streaming reader.

  The decode is one left-to-right pass, so `_x005F_x0041_` reads as the literal `_x0041_`
  rather than collapsing to `A`, and the grammar is exact: `_x041_`, `_x00041_` and `_xZZZZ_`
  are ordinary text and stay that way. Every one of those outcomes was measured against Excel
  Desktop before it was implemented, and the probe workbook is committed as a corpus fixture
  (`docs/knowledge/specs/spreadsheetml-xhhhh-escape-is-decoded-on-read.md`).

  Two more carriers of human-typed prose have since been measured and joined the group: a
  **threaded comment's message** and a **print header/footer definition**. Both had been guessed
  the other way and refused such a character on write. Excel decodes and re-emits the escape in
  each, verbatim to the cell-text grammar, so both now escape on write and decode on read. Two
  consequences for callers: a header or a threaded message may carry any character a cell may,
  and a foreign file whose header reads `&C_x0041_` now loads as `&CA` rather than as the
  seven-character literal.

- **A frozen pane no longer disappears when a sheet is copied through `model`.**
  `WorksheetModel` was missing `view`, so `dst.model = src.model` reproduced the cells,
  merges, tables, autofilter and page setup — and silently unfroze the header row. The
  frozen/split pane now rides the model like every other sheet-level field, and clearing it
  works in the same direction: assigning a normal-view model over a frozen sheet unfreezes it
  rather than leaving a stale pane.

  This is additive to `WorksheetModel`. Code that *constructs* a model literal by hand rather
  than reading one from a sheet must now supply `view` (`{}` is a normal view); code that does
  the usual `dst.model = src.model` is unaffected.

  The boundary is now a rule rather than a list
  ([ADR-0005 amendment](docs/decisions/0005-worksheet-model-is-semantic-only.md)): **a field
  belongs in the model when its value means the same thing on any sheet of any workbook.** By
  that test threaded comments stay out — a comment's author is an id into the *workbook's*
  `persons` registry — and they are now named on the out-of-scope side instead of being absent
  from both lists.

- **A foreign file spelling an OOXML boolean `"false"` is read correctly.** `xsd:boolean`
  permits `true`/`false` alongside `1`/`0`. Excel writes the digit, so five readers that
  tested the attribute by hand against `'0'` had never been caught reading the long spelling
  as *true*: a `<col customWidth="false">` gained a width it does not have, a
  `<row customHeight="false">` a height, an x14 data bar's `gradient="false"` stayed a
  gradient, and a `cfRule aboveAverage="false"` read as above-average. All five now go
  through the reader's existing `boolPresent`, which has handled both spellings all along.
  Files written by Excel are unaffected; files from other producers are the point.

### Changed

- **`Workbook.authoredThemeXml()` is now `Workbook.themeOverrides`.** The old method handed back
  theme part *text*, which made the model the place that knew how a theme is spelled. The getter
  returns the colour slots and typefaces {@link setTheme} authored, or `undefined` when none were,
  and the serializer composes them onto the part it is about to write. Same result in the file;
  a caller who needs the text can compose it with `applyThemeOverrides`.

- **`parseThemeColorScheme` moved from `/core` to `/xlsx`,** which is where the part it parses is
  read, and is joined there by `parseThemeFontScheme` and `DEFAULT_THEME_XML`. Importing from the
  package root is unaffected; a subpath import needs the new one.

- **An authored theme typeface is escaped properly.** The theme writer carried its own attribute
  escape, weaker than the library's: it handled `& < > "` and left the apostrophe, the newline, the
  carriage return and the tab raw, and it did not refuse a character XML 1.0 cannot represent. A
  typeface carrying any of those produced a malformed part. It now uses the same escape as every
  other attribute in the package. A custom number format code, whose escape *deliberately* differs
  (a bare apostrophe round-trips), gained the representability guard it was also missing.

- **A row or column addressed by number is now bounded by the grid, as one addressed by letters
  already was.** `getCell('XFE1')` threw and `getColumn(16385)` did not, so whether the library
  refused a position outside the spreadsheet grid depended on how the caller spelled it, and the
  numeric spelling let the model hold, and the writer emit, a package Excel must repair. `Row`,
  `Column` and `Cell` now reject an index outside `1..1048576` / `1..16384` with a `RangeError`
  worded the same way the letter path words it.

  This is a new throw on calls that used to return. Code placing content at a synthetic index past
  the grid was producing a file Excel rejects, so the failure has moved to the call that causes it.
  Reading is unaffected: a file declaring an out-of-grid position is bounded on read rather than
  refused, so no file that opened before stops opening.

- **Error messages no longer contain em dashes.** Where a message used ` — ` to weld two
  clauses together it now uses a colon, a semicolon, or a pair of commas, whichever the
  sentence wanted: `column 16385 is out of bounds: Excel supports 1..16384` rather than
  `column 16385 is out of bounds — Excel supports 1..16384`. No message changed meaning, none grew,
  and the error *types* and `code` values are untouched, so anything branching on the
  taxonomy is unaffected. Code that matched on message text is not: match on the class or
  on `code` instead, which is what they are for.

- **Writing a character XML cannot represent into anything but a cell value now throws
  `AuthoringError`.** Only cell values have the `_xHHHH_` escape; a sheet name, a defined
  name, a table column name, a formula, a document property, a print header and a threaded
  comment do not. There is no faithful way to write U+0001 into a sheet tab — `Sheet_x0001_A`
  is a different name, not the one that was asked for — so the writer refuses, naming the
  code point and its offset.

  This is a new throw on a call that used to return. It used to return a file Excel reports
  as damaged, so the failure has moved earlier and got louder rather than appearing from
  nowhere. Tab, LF, CR and U+007F are valid XML 1.0 and are unaffected; so are the C1
  controls, which only XML 1.1 forbids.

- **The OOXML conformance oracle is now the shared `ooxml-validate` package**, and the
  repository-owned .NET tool behind it is gone
  ([ADR-0033](docs/decisions/0033-the-ooxml-oracle-is-a-shared-package.md)). Nothing in
  `src/` moved and nothing a consumer installs changed — this is development and CI
  tooling, as it always was.

  The point is that `ts-pptx` used to carry a *different* validator on a different Open XML
  SDK version, so two sibling projects were enforcing two rule sets while both calling it
  "Microsoft's validator". They now share one oracle, one pin, and one report contract.

  For contributors: **there is no .NET requirement any more.** The package fetches a
  prebuilt, checksum- and provenance-verified binary on first use and caches it, so
  `pnpm run validate:ooxml file.xlsx` and `pnpm run test:ooxml` work from a plain dev
  install. `global.json` and CI's `setup-dotnet` step are gone. Both spellings work —
  `pnpm run validate:ooxml file.xlsx` and `pnpm run validate:ooxml -- file.xlsx` — which is
  why the dependency floor is `ooxml-validate` 0.0.3.

  The baseline did not move. This repo was already validating at `Microsoft365` against
  `DocumentFormat.OpenXml` 3.5.1, so `test/ooxml-validation/allowed-errors.json` is
  unchanged and still empty.

- **A truncated VBA project fails closed instead of reading as zeros.** The `vbaProject.bin`
  reader's little-endian integer reads were open-coded in four modules, and the shared shape
  indexed past the end of the buffer without noticing: `undefined | (undefined << 8)` is `0`,
  so a truncated compound file could parse as a structure full of valid-looking zeros. Every
  caller happened to bounds-check first, so no known file was misread — but that is the wrong
  place for the check in the one subsystem that parses a blob straight out of an untrusted
  `.xlsm`. The bound now belongs to the read itself, which raises `VbaParseError` naming the
  offset and the buffer length.

## [1.3.1] — 2026-08-11

`1.3.0` was set in `package.json` and cut no release, so this is the first published version of
the line and carries what was staged for it as well as the fix below.

### Added

- **`SheetView.showGridLines` turns the on-screen grid off**, the one sheet-view attribute an
  authoring consumer reaches for first and the only facet of `<sheetView>` this library could not
  express. A workbook built as a *deliverable* — its own fills, its own borders, a title band —
  reads as a spreadsheet rather than a document while Excel's grey grid shows through it, and
  there was no way to say so:
  [`PrintOptions.gridLines`](docs/api/page-setup.md) is the neighbouring question about *printing*
  and Excel exposes the two as separate checkboxes because the answers differ.

  ```js
  sheet.view.showGridLines = false;
  ```

  Only an explicit `false` writes an attribute, and reading one back records only
  `showGridLines="0"`. Excel's default is on, so "unset" and "on" are the same state: recording
  `true` would make every re-written sheet fabricate an attribute its source never carried, which
  is the round-trip noise `<pane>` is already careful to avoid. Both arms of `sheetViewsXml` carry
  it, so a frozen sheet can hide its grid too.

  Found by the library's first real authoring consumer, which hides the grid on all seven sheets
  of the two workbooks it ships. The remaining `<sheetView>` booleans (`showRowColHeaders`,
  `showRuler`, `showFormulas`, `showZeros`, `rightToLeft`) are still unmodelled — see
  `docs/knowledge/specs/sheetview-boolean-flags-and-showformulas.md`, which called this set out and
  is now one item shorter.

- **`WorkbookProperties.title` and `.company`**, the two remaining fields of Excel's File ▸ Info
  panel a deliverable actually sets. Both were silently unwritable: a caller could assign neither,
  and a file that carried them lost them on a round-trip.

  ```js
  wb.properties.title = 'Planning Ateliers';
  wb.properties.company = 'Acme & Co';
  ```

  `title` is `dc:title` in `docProps/core.xml`, emitted **before** `dc:creator` because
  `cp:coreProperties` is a schema *sequence* and Excel repairs a file whose children are out of
  order. `company` is `<Company>` in `docProps/app.xml` — the one document property OOXML keeps
  outside the core part, which is why `appPropsXml` now takes the properties at all. Both are
  omitted entirely when unset, and both now read back.

### Fixed

- **Writing a workbook is now reproducible: the same model always produces the same bytes.** A zip
  entry stores a modification time, `fflate` defaults it to `Date.now()`, and so every package this
  library wrote was stamped with the moment it ran. Two writes of an unchanged workbook differed in
  a handful of bytes per entry and in nothing else — invisible to Excel, and corrosive to everything
  around the file. A committed `.xlsx` deliverable churned in `git diff` on every regeneration, so
  the diff stopped being read and a real change to it had somewhere to hide; nothing downstream
  could cache on output bytes; and a byte-comparison gate could not tell "the writer changed" from
  "the clock moved", which is the same as not having one.

  Every entry is now stamped 2001-01-01 12:00 — see
  [ADR 0032](docs/decisions/0032-package-output-is-reproducible.md) for why that date and why there
  is no option to go back to the clock. All four writing paths are covered: `writeXlsx`,
  `writeXlsxAsync`, `WorkbookStreamWriter`, and the package-level VBA edits, which re-zip after
  splicing and had nothing to preserve because unzipping drops the original times. The streamed
  container takes the stamp as a field on the entry rather than as an option, which is exactly the
  kind of difference that leaves one call site behind — hence a test that decodes the DOS date out
  of every entry rather than only comparing two writes, which would pass within the same
  two-second bucket regardless.

  Consequences worth stating plainly: package bytes changed once, at this release, so a stored hash
  of a previously written file no longer matches — nothing about the content did. `writeXlsx` and
  `writeXlsxAsync` are now byte-identical, which pins them to the same compression settings in a way
  comparing inflated parts never could. And a file manager listing the archive's contents shows
  2001-01-01 against each entry, the same trade every reproducible-build toolchain makes; the file's
  own filesystem timestamp is untouched.

  Found by the library's first authoring consumer, which commits its generated workbooks.
  [ADR 0024](docs/decisions/0024-async-is-one-writer-not-a-mirrored-pair.md) had met this while
  comparing the two writers, declined to pin `mtime` for a test's convenience, and left the real
  question open; a consumer answered it.

## [1.2.0] — 2026-08-09

No exported symbol changes in this release. What changes is what the tarball contains and what
one error says: the package now ships a skill for reporting a defect back here, because the
agent that hits one is working in someone else's repository and has no reason to believe filing
is its job. Alongside it, the published tarball is 42% smaller — the same code with 438 KB of
implementation comments no longer addressed to anyone.

### Added

- **The package ships a skill for reporting a bug upstream, and `InternalError` prints where to
  send one.** Most code using this library is written by an agent, and an agent that hits a
  library defect writes a workaround: that is the rational move from where it stands, since the
  workaround unblocks its user today and the issue tracker belongs to a repository it is not in.
  So the report never happens, the corpus never gets the case, and the next consumer rediscovers
  the same defect. Three layers answer three different reasons the report dies.
  `skills/ts-xlsx-upstream/` is now in `files`, so `npx skills add
  ./node_modules/@shbernal/ts-xlsx` works offline and always matches the installed version. The
  skill covers triage (is this ours or your file's?) and spends most of its length on reducing a
  failure to a script that *builds its own input* — a spreadsheet in a real project holds
  salaries and customer lists and the tracker is public, so the rule is synthesize or describe in
  prose, never redact. It files without interrupting you once the reproduction stands on its own,
  and passes `--repo shbernal/ts-xlsx` on every `gh` call, because `gh` in a consumer's checkout
  defaults to the consumer's tracker and would file our bug there.
- **`InternalError` messages now carry the tracker URL**, appended by the constructor rather than
  by its six throw sites, so a site added later cannot forget it. `'internal'` is the one code in
  the taxonomy that already declares whose bug it is, which makes an unconditional "report this"
  a statement of fact rather than a nag — and a stack trace is the only artefact of that failure
  that reaches whoever is debugging it, so a pointer anywhere else is one they must already be
  looking for. The other three codes deliberately get no banner: malformed input is the routine
  outcome for a library reading untrusted files, and a report pointer on every corrupt file
  teaches the reader to skip the line, including the one time it always means something. They get
  the test in prose instead, on the `XlsxErrorCode` TSDoc and therefore in `dist/index.d.ts` —
  report it when the file opens cleanly in Excel but not here, because that combination means the
  gap is ours.

### Changed

- **The published tarball is 42% smaller, and nothing was removed to make it so.** `build` was
  one `tsc` pass emitting both halves of `dist/`, and it kept comments — necessarily, because
  that is what puts JSDoc into the `.d.ts` hovers. The cost was that the same flag also shipped
  every implementation comment in the tree into the runtime JS: measured, 47% of it, 438 KB
  addressed to nobody. `removeComments` is whole-emit, so a single pass could only ever give up
  one audience or the other. The build is now two passes that split by audience —
  `tsconfig.build.json` emits JS with the prose stripped, `tsconfig.build.dts.json` emits
  declarations with it kept. Runtime JS falls from 918 KB to 480 KB and the tarball from 392 KB
  to 227 KB gzipped, while all 112 `.d.ts` files are byte-identical to the single-pass output —
  every one of the 237 KB of JSDoc across 90 files still reaches an editor hover. No module left
  any entry's closure, no signature changed, and the corpus's 838 behaviours run green against
  the emitted `dist/`.
- **The per-entry size budgets are halved, which makes them stricter rather than looser.** They
  were measured against JS that was ~47% comment, so a codec crossing a module boundary — the
  failure the tripwire exists to catch — could have arrived inside a release's ordinary comment
  churn without moving them. Rebaselined onto comment-free emit, they measure code.
- **The `npm-publish` environment is now checked rather than provisioned, and this repository
  holds no credential at all.** `environment.yml` wrote the environment through an
  administration endpoint, which `GITHUB_TOKEN` may not reach, so it needed a PAT —
  `ENV_ADMIN_TOKEN`. That secret was never created: the workflow failed on its own first step
  every day of its life, and the reviewer removal it was supposed to apply was done by hand
  instead. Restoring the secret would have put a long-lived credential able to rewrite the gate
  that decides who may publish next to a publish path that deliberately carries none. Reading
  an environment needs only read access, so the write was dropped: the workflow now fails when
  the live environment differs from what the file declares — weekly, on any edit to the
  declaration, and on demand. Drift is detected rather than corrected, which is the half that
  was load-bearing. Nothing about the published package changes. See
  [ADR-0026](docs/decisions/0026-releasing-is-a-github-release-and-npm-follows.md).

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

[Unreleased]: https://github.com/shbernal/ts-xlsx/compare/v1.3.1...HEAD
[1.3.1]: https://github.com/shbernal/ts-xlsx/compare/v1.2.0...v1.3.1
[1.2.0]: https://github.com/shbernal/ts-xlsx/compare/v1.1.0...v1.2.0
[1.1.0]: https://github.com/shbernal/ts-xlsx/compare/v1.0.3...v1.1.0
[1.0.3]: https://github.com/shbernal/ts-xlsx/compare/v1.0.2...v1.0.3
[1.0.2]: https://github.com/shbernal/ts-xlsx/compare/v1.0.1...v1.0.2
[1.0.1]: https://github.com/shbernal/ts-xlsx/compare/v1.0.0...v1.0.1
[1.0.0]: https://github.com/shbernal/ts-xlsx/releases/tag/v1.0.0
