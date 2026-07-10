# Regression corpus

The corpus is **the product's spine** (see [`../../STRATEGY.md`](../../STRATEGY.md)).
It encodes "correct behavior" as a set of implementation-blind cases that run against
*any* implementation through a thin adapter — so it survives the Phase 3 rewrite and
proves the new code is at least as correct as the old, plus everything the old one got
wrong. A bug without a corpus case is a bug that will return.

## Layout

```
test/corpus/
  cases/*.case.mjs      one harvested behavior cluster, implementation-blind
  adapters/<name>.mjs   binds the contract vocabulary to a concrete implementation
  run.mjs               discovers cases, runs them against an adapter, reports red/green
```

Run it:

```
node test/corpus/run.mjs [--adapter current]
```

## A case

A case module default-exports:

```js
{
  id: 'whole-column-defined-names',              // durable descriptive slug — no number prefix
  cluster: 'address-decoding',
  description: '…',
  provenance: { source: 'upstream-issue' },      // OPTIONAL, disposable trace — never the identity
  behavior: [
    { name, baseline: 'pass' | 'fail', expect(api, assert) { … } },
  ],
}
```

- **`id` / `description`** carry the durable identity: a descriptive slug and the
  *real-world scenario* in prose. Do **not** encode upstream issue/PR numbers here —
  they go meaningless when we finish leaving that project (`harvest-triage` skill).
- **`provenance`** is optional and disposable — a trace of where a case came from, never
  its identity. The durable text must stand entirely without it.
- **`behavior[]`** — each is one assertion about observable behavior. `expect` receives
  the **adapter** (`api`) and Node's strict `assert`; it throws to fail, returns to pass.
- **`baseline`** records what **today's legacy code** does for this behavior:
  `pass` = green now (a *regression lock*), `fail` = a known-open bug the rewrite must
  fix. This is what lets the corpus be "mostly red where bugs are real" *without* a red
  build.

## The adapter contract

Cases never import `lib/`. They call a small, growing vocabulary of capabilities that
the adapter provides — the adapter is the **only** place that knows how an
implementation is shaped. Current vocabulary:

| Capability | Meaning |
|---|---|
| `decodeAddress(ref)` | Decode a single cell/row/column reference → `{col, row, …}` (absent axis = `undefined`). |
| `decodeRange(ref)` | Decode a range reference → corners + serialized dimensions. |
| `probeCellFonts({apply, read})` | On a fresh worksheet, assign a font to each `apply` cell, then return `{ <address>: font }` for the `read` cells — for asserting per-cell style stays local. |
| `roundtripWorkbook(spec)` | Build a workbook from a declarative `spec`, write it to a buffer, read it back, and return a normalized JSON model (`{properties, definedNames, sheets}`, each sheet including per-cell `fill`/`alignment`, per-column `width`/`hidden`/`numFmt`/`alignment` (column alignment applied via `getColumn`, so a case can assert alignment isolation), per-row `height`/`hidden`, a `pageSetup` `{fitToPage, fitToWidth, fitToHeight, scale}` fact, `autoFilter`, `merges`, and `rowCount`/`actualRowCount`) — for asserting content survives write→read. |
| `inspectPackage(spec)` | Build + write a `spec`, unzip the package, and return raw OOXML-part facts (worksheet-declaration consistency, `pageMargins`, `sheetViews`, table XML, per-cell formula text, per-sheet `columnGroups`/`maxColumnIndex` from the `<col>` table, per-sheet `autoFilterRef`/`dimensionRef`, well-formedness, a `styles` fact recording whether a theme part backs any theme-color font reference, `worksheetRels`, per-`sheetEntries` `state` (workbook-declaration visibility: `visible`/`hidden`/`veryHidden`, `null` when absent = visible), a `contentTypeDefaults` list of `<Default>` extension/media-type pairs, a `vml` fact recording each comment `<v:textbox>` style string + an `allTextboxesFitToText` flag (a note box carrying `mso-fit-shape-to-text:t` grows to its text instead of clipping), a per-sheet `rows` map of row-element outline attributes (`outlineLevel`/`hidden`/`collapsed` — a collapsed group must carry the collapsed flag on its summary row, not on the hidden detail rows), a per-sheet `hasBackgroundPicture` flag, and a `packageParts` fact recording comment/VML/table parts + worksheet-rel-id uniqueness) — for asserting on what is actually serialized. |
| `tryWriteWorkbook(spec)` | Build + attempt to write a `spec`; return `{ok, error, survivingCells, …}` — for asserting pathological input neither throws nor drops sibling cells. |
| `roundtripRangeValidation({range, type, formula})` | Apply one data validation over a multi-cell range (e.g. a whole column) and write; return `{writeOk, writeError, sqrefs, count, reloadOk}` — for asserting a whole-column dropdown emits a single range-scoped `dataValidation`, does not throw on write, and reloads. |
| `appendRowShapes()` | Append rows as a dense array, a sparse 1-based array, an object, a typed array, and a mixed batch; return `{rows: {<n>: {A, B, C, E}}}` after reload — for asserting every row shape lands its data (not only object-keyed rows) and types survive. |
| `appendRowsAfterReload(initial, append)` | Author initial rows, write+reload, append rows past the last populated row, write+reload; return `{loadedRowCount, finalRowCount, rows}` — for asserting appended rows land at contiguous indices with no gap/overwrite and the originals survive. |
| `interleavedImageAnchors(placement)` | Place two distinct images in an interleaved order (default `'BAA'`) and resolve each anchor's referenced media (embed rId → drawing rel → media); return `{placed, resolvedLetter, distinctMediaCount, distinctRelTargets}` — for asserting every anchor renders the image it was placed with and a reused image maps to one stable relationship. |
| `authorConditionalFormatting(cf)` | Author a conditional-formatting rule (e.g. a `dataBar` with `cfvo` anchors, color, gradient) and return `{writeOk, xml:{hasDataBar, cfvoCount, hasColor, wellFormed}, reload:{type, color, gradient, cfvo}}` — for asserting a rule emits well-formed XML and round-trips. |
| `roundtripFixtureImageRotation(rel)` | Read a fixture, load-rewrite it, and report the image drawing-anchor rotation (`rot` on `<a:xfrm>`, 1/60000-deg) before/after → `{sourceRot, rewrittenRot}` — for asserting an image rotation survives a round-trip. |
| `sharedFormulaRoundtripAndSplice()` | Author a shared-formula master/slave group, read+rewrite it, and splice a column in; return `{roundtripOk, preservedFormulas, spliceOk, spliceError}` — for asserting a lossless shared-formula round-trip and surfacing the splice "master must exist above/left of clone" throw. |
| `streamWriteCfHyperlinkOrder()` | Stream-write a sheet carrying both a conditional-formatting rule and a hyperlink; return `{posConditionalFormatting, posHyperlinks, conditionalFormattingBeforeHyperlinks, reloadOk}` — for asserting the CT_Worksheet child order (conditionalFormatting must precede hyperlinks). |
| `mutateWorksheet({cells, ops, read, readStyles})` | Build a fresh sheet (cells may carry `font`/`fill`/`numFmt`), apply structural mutations (`spliceRows`/`spliceColumns`/`mergeCells`/`insertRow`/`duplicateRow`), and return `{rowCount, columnCount, cells, styles, merges, error}` — for asserting in-memory model edits behave predictably (a throwing op, e.g. an overlapping merge, is reported as `error`, not propagated; `merges` lets a case assert a splice SHIFTS a merged range and keeps it merged; `readStyles` reports each requested cell's `{value, font, fill, numFmt}` after the edit, so a case can assert a shifted row carries its style rather than being blanked). |
| `readFixtureValidations(rel)` | Read a fixture `.xlsx` (path relative to `fixtures/`) and return `{cells, count}` — per-cell data validations the reader exposes, keyed `<sheet>!<addr>`. |
| `readFixtureHyperlinks(rel)` | Read a fixture and return the first sheet's hyperlink cells → `{<addr>: {hyperlink, text}}` — for asserting a real file's hyperlink is reconstructed in full, including a `#fragment` carried in the hyperlink element's `location` attribute (separate from the relationship Target), which must not be dropped. |
| `readFixtureValidationRules(rel)` | Read a fixture and return the DISTINCT validation rules each sheet declares → `{sheets: {<name>: {rules:[{rule, coverageCount}], ruleCount}}}`, read from the worksheet model (so a validation over an empty range is still seen) and de-duplicated by content — for asserting a reference-based list source (a defined name, a cross-sheet range) is surfaced as its formula text rather than `"[object Object]"`. |
| `roundtripFixtureValidationXml(rel)` | Read a fixture, write it back, unzip, and return data-validation facts of the re-serialized package — standard `<dataValidation>` and extended `<x14:dataValidation>` (extLst) counts + `<xm:sqref>` targets, plus per-standard-rule `standardRules:[{type, sqref, errorTitle, error, formula1}]` — for asserting a validation's type, source reference, target range, and error strings survive a round-trip. |
| `readFixtureReport(rel)` | Read a fixture and return `{ok, error, sheetNames}` — the read either succeeds (with sheet names) or its error is captured as data — for asserting the reader tolerates foreign-generator files (namespace-prefixed roots, BOMs, non-ASCII sheet names, unusual zip ordering) without crashing. |
| `roundtripFixture(rel)` | Read a fixture, write it back unchanged, read it again, and return `{sheetNames, columns, styleSurvival}` before/after — for asserting sheet names, custom column widths, and per-cell styles survive the format-preserving "open a template and re-save" path. Style comparison is key-order-insensitive. |
| `inspectImageAnchors(spec)` | Build a workbook whose sheets place images (`sheets[].images:[{range}]`, `range` a string like `"B2:D6"` or `{tl, br?, ext?}`), write it, and return the serialized `{anchors:[{anchorType, editAs, from, to, ext, spPr}]}` drawing geometry — for asserting fractional/whole/string anchors map to correct OOXML col/colOff offsets against real column/row size, and that an anchored picture's own `spPr` shape transform (`{hasXfrm, off, ext, zeroedTransform}`) does not carry a zeroed placeholder that a strict consumer would honor over the anchor. |
| `readFixtureImageAnchors(rel)` | Read a fixture and return `{images:[{sheet, editAs, tl, br}], count}` with integer cell coordinates — for asserting a file whose images use (string) range anchors reads without crashing and normalizes to an object range. |
| `csvRead({csv, options})` | Parse a CSV string with reader `options` → `{ok, error, rows}`, a 2-D array of typed cell values (a Date becomes `{date: iso}`, an error `{error}`, empties `null`) — for asserting delimiter handling, value coercion, and header-mode behavior. A broken option path is captured as `{ok:false, error}`. |
| `csvWrite({spec, options})` | Write a `{rows:[[cell,…]]}` spec (a cell is a primitive, `{date: iso}`, `{formula, result}`, or `{error}`) to CSV with writer `options` → `{ok, error, text}` — for asserting field delimiter and date formatting on genuinely-typed cells. |
| `streamWriteSheet({useSharedStrings, ops, read})` | Drive the streaming workbook writer through row ops (`{op:'addRow'\|'addRows', value}`), commit, read the package back → `{ok, error, cells, rowCount}` — for asserting streaming-only behavior (batch add, richText shared-string handling). A throwing op is captured as `{ok:false, error}`. |
| `roundtripFormulas(spec)` | Build + round-trip formula cells (a cell may carry `formula` or `sharedFormula`) → per-cell `{formula, sharedFormula, result}` from the model getters — for asserting a shared-formula clone reads back a concrete, address-translated formula. |
| `roundtripTableAppend(spec, {tableName, appendRows})` | Build a table, round-trip it, fetch it by name, and try appending rows to the *reloaded* table → `{hasTable, loadedRowCount, addError, committed, finalRowCount}` — for asserting a table rehydrated from a file is mutable, not throwing on append. |
| `readFixtureDefinedNames(rel)` | Read a fixture and report the workbook-level defined names the reader exposes → `{names: {<name>:[ranges]}, count, modelCount}` — for asserting a full-row/full-column-span named range, or two same-named names scoped to different sheets, are read back rather than dropped by over-strict validation or scope collision. (`roundtripWorkbook`'s model also carries a `definedNames` map for the write→read path.) |
| `readFixtureCellStyles(rel, cells)` | Read a fixture and report the resolved `{fill, fontColor}` of each requested `"Sheet!Addr"` cell — for asserting real-file cell colors (a solid fill's visible `fgColor`, an automatic indexed `bgColor`, a theme+tint color, a separate font color) are read faithfully and not conflated. |
| `roundtripFixtureTableXml(rel)` | Read a table-bearing fixture, write it back unchanged, and report each table's raw-XML facts before/after → `{tables:[{name, source, rewritten}]}` (autoFilter presence/ref, header-row count, totalsRowShown, empty-filterColumn, column count) — for asserting a no-op round-trip does not corrupt the table part (inject an autoFilter, flip the header row, spuriously enable totalsRowShown) so Excel does not strip the table. |
| `readFixtureTable(rel, tableName)` | Read a fixture and report a named table's rehydration → `{found, columns, rowCount}` — for asserting a table loaded from a real file exposes its column names *and* its data rows, not a half-loaded model with an undefined rows array. |
| `streamReadFixture(rel, cells)` | Read a fixture's first sheet through the *streaming* reader and report requested cells' `{type, value}` (type as a stable label) — for asserting streaming read applies number formats like the full read (a date-formatted cell streams as a date, not a raw serial). |
| `streamVsEagerSheetNames(rel)` | Read a fixture eagerly and via the streaming reader and report the sheet names each path surfaces → `{eager, streaming}` — for asserting streaming exposes the real declared names (joining each worksheet part to the workbook sheet declaration), not generic positional placeholders. |
| `streamVsEagerRowNumbers(rel)` | Read a fixture eagerly and via the streaming reader and report the first sheet's row numbers each path yields → `{eager, streaming}` — for asserting streaming preserves each data row's true index across interior blank rows (1 then 6, not a resequenced 1,2). |
| `streamVsEagerRowHidden(rel)` | Read a fixture eagerly and via the streaming reader and report each first-sheet row's `{number, hidden}` from both paths → `{eager, streaming}` — for asserting streaming surfaces a row's hidden flag (interpreting the string-form `"true"/"false"`), agreeing with the eager read rather than reporting every row visible. |
| `streamReadReport(rel)` | Read a fixture through the streaming reader end-to-end → `{ok, error, sheetNames, totalRows}` — for asserting it tolerates a package whose ZIP places a worksheet part before `workbook.xml` (yielding every worksheet and all rows) rather than crashing on an unbuilt workbook model. |
| `readFixtureCells(rel, cells)` | Read a fixture's first sheet with the *full* reader and report requested cells' `{type, value, numFmt, note}` — for asserting real-file cell values/types (e.g. a Strict-mode ISO-8601 `t="d"` date parses to the stated date; a 1900-epoch serial maps to Excel's calendar date; an empty comment reads as an empty-string note; an escaped-literal-`m` scaling format is not mis-detected as a date; a locale-specific built-in date format id resolves to a non-empty date `numFmt`, not empty). |
| `roundtripFixtureCellXml(rel, cells)` | Read a fixture, write it back, and report requested cells' raw serialized `<c>` facts → `{cells: {<addr>: {t, formula, value}}, hasNaNToken}` — for asserting a round-trip does not corrupt cell content (a string-typed formula cell whose style carries a date format must not lose its `t="str"` type and emit the invalid literal `NaN` as its value). |
| `roundtripFixturePackageParts(rel)` | Read a fixture, write it back unchanged, and report package-part facts `{source, rewritten}` — family counts (drawings, VML, media, pivot tables/cache, slicers, comments) plus worksheet/drawing ref flags (`hasLegacyDrawingHF`, `hasDrawingRef`, `hasHeaderFooterImageToken`, `drawingHasShape/Picture`) — for asserting a no-op round-trip preserves parts the reader does not model (header/footer images, vector shapes, pivot tables) rather than dropping them. |
| `roundtripFixtureStyleFacts(rel)` | Read a fixture, write it back, and report style-fidelity facts `{source, rewritten}` — column widths, pageSetup (scale/fit/order/orientation), custom indexed-color palette, and conditional-format differential (`dxf`) number codes — for asserting a no-op round-trip preserves them and never serializes a numFmt as the literal `"[object Object]"`. Tolerant of a rewritten package the reader chokes on (raw styles.xml facts come from the buffer). |
| `roundtripFixtureConditionalFormatting(rel)` | Read a fixture, write it back, and report first-sheet conditional-formatting facts `{source, rewritten}` — `{blockCount, rules:[{type, dxfId, priority}]}` — for asserting a no-op round-trip preserves a cfRule (even an unmodeled type like `duplicateValues`) rather than dropping it or emitting an empty `conditionalFormatting` shell (which corrupts the file). |
| `roundtripFixtureColorFidelity(rel)` | Read a fixture, write it back, reload, and report how many styled cells' VISIBLE fill/border colors changed → `{checked, fillMismatches, borderMismatches, fillSample, borderSample}` — ignoring the benign `patternFill pattern="none"` the writer adds — for asserting themed/indexed fill and border colors survive a pure open-then-save. |
| `roundtripFixturePrintAreas(rel)` | Read a fixture whose sheet declares multiple print areas (one comma-separated `_xlnm.Print_Area` name) and report `{sourceRangeCount, readPrintArea, rewrittenRangeCount}` — for asserting both ranges are recovered on read and re-emitted on write, not truncated to the first. |
| `writePrintAreaDefinedName(printArea)` | Build a workbook with a (possibly comma-separated) `printArea`, write it, and report the emitted `_xlnm.Print_Area` ranges → `{rangeCount, ranges}` — for asserting authoring two print areas emits two proper rectangular ranges in one sheet-scoped name, not a truncated/mangled one. |
| `authorListValidations(validations)` | Author list-type data validations (each `{ref, formula, error?, allowBlank?}`), round-trip, and report `{readBack: {<ref>: {type, formulae}}, xml: {count, wellFormed, formula1}}` — for asserting both value-source forms (an inline quoted literal `"Male,Female"` and a cross-sheet range `Levels!$A$2:$A$9999`) survive verbatim and emit one well-formed `<dataValidation>` per range. |
| `authorCellProtection(cells, protect?)` | Author per-cell protection (`cells: [{ref, value?, protection?}]`) plus an optional protected sheet (`{password?, options?}`), round-trip, and report `{readBack: {<ref>: {locked}}, hasApplyProtection, sheetProtection, sheetProtectionAttrs}` — for asserting an unlocked cell survives (default is locked), the flag is carried in cellXfs, and worksheet protection emits `<sheetProtection>` (`sheetProtectionAttrs` gives the parsed permission booleans, where `"0"` PERMITS an operation and `"1"` LOCKS it, so a permissive `sort:true` option shows `sort="0"`). |
| `streamCommitReport({duplex?, timeoutMs?})` | Drive the streaming writer over a caller-supplied `PassThrough` (or `Duplex`) sink and report `{settled, timedOut, bytes, valid}` — for asserting streaming-to-a-remote-sink commit resolves within bounded time and delivers a complete, re-openable package rather than hanging on a finish signal. |
| `streamWriterImageSupport(range?)` | Report the streaming writer's image-parity surface and (if supported) the streamed package's parts → `{writerAddImage, sheetAddImage, error, mediaParts, drawingParts}` — for locking image parity with the in-memory writer (anchor a registered image on a streamed sheet; media + drawing parts appear). |
| `streamWritePackageReport({rows?})` | Assemble a whole package via the streaming writer, then treat the bytes as an untrusted archive → `{partCount, emptyParts, crcValid, reloadOk, sheetNames, firstCol}` — for asserting the streamed output is a valid zip (no zero-byte parts, per-entry CRC matches, re-reads cleanly), not merely valid XML. |
| `streamReadSpec(spec, cells?)` | Write a `spec`, read it back through the STREAMING reader over real chunk boundaries, and pair with an eager read → `{streamed, eager}` — for asserting multi-byte UTF-8 (CJK/emoji) survives the streaming path byte-exact rather than splitting into U+FFFD at a chunk boundary. |
| `loadMutateCellStyle({sharedFill?, mutateTo?})` | Author cells sharing one on-disk style index, load, mutate one cell's fill, read a sibling → `{sibling, original, bled, diskSibling, diskBled}` — for asserting loaded cells get independent style objects rather than aliasing the shared record. |
| `copyWorksheetModel({merges?, cells?})` | Copy a worksheet via the `model` export/import contract (`dst.model = {...src.model, name}`) and report merge survival → `{srcMerges, dstMerges, error}` — for asserting a model-cloned sheet keeps its merged ranges. |
| `styleDedupReport(spec, cells?)` | Write a `spec` and report the style-table size + per-cell resolved style index → `{cellXfCount, indices}` — for asserting identical cell styles dedup to one shared entry while a distinct style stays separate. |
| `readRowCellPresence(spec, rows?)` | Load a written `spec` and report, per row, the column indices a full (`includeEmpty`) iteration yields → `{rows: {<n>: {cols, cellCount, valuesLength}}, columnCount}` — for asserting trailing empty cells are surfaced up to the declared width. |
| `streamVsEagerRowValues(spec, rowNumbers?)` | Read a `spec`'s rows via both the eager and streaming readers → `{eager, streamed}` (sparse holes → null) — for asserting the streaming reader exposes the same 1-based row-values indexing as the full load. |
| `roundtripSpecTableFacts(spec)` | Write a `spec`'s table, round-trip it, report table facts before/after → `{write, roundtrip, loadOk, loadError}` (each `{ref, name, wellFormed}`) — for asserting a defined table's ref range and part survive a load→save cycle, including empty-body/single-row shapes. |
| `loadMutateCellFont({original?, mutateTo?})` | Author cells sharing one font, load, spread-reassign one cell's font (`{...cell.font, color}`), read the sibling → `{edited, sibling, original, bled}` — the font companion to `loadMutateCellStyle`. |

`inspectPackage`'s per-sheet fact also carries `elementOrder` (raw positions of `drawing` /
`legacyDrawing` / `tableParts` plus the `legacyBeforeTableParts` etc. adjacency invariants) so a
case can assert the CT_Worksheet child-element order, and a `headerFooter` fact (the odd/even/first
header/footer child text plus the `differentOddEven`/`differentFirst` gating flags). The `spec`
worksheet input accepts a `headerFooter` block mirroring those children.

The `spec` shape consumed by the three workbook capabilities is documented at the top of
`adapters/workbook-io.mjs` (worksheets with cells, columns, rows, page margins, tables).

Add capabilities only as cases demand them, and add them to **every** adapter. When the
rewrite lands, a `rewrite.mjs` adapter binds the same vocabulary to the new code and
every existing case runs unchanged — the corpus does not move, the implementation does.

## What the runner does with baselines

| baseline | actual | status | fails build? |
|---|---|---|---|
| pass | pass | `✓` green | no |
| fail | fail | `○` known-open | no — this is the corpus's job on the frozen tree |
| pass | fail | `✗` regression | **yes** (exit 1) |
| fail | pass | `↑` newly-fixed | no — but flip the baseline to `pass` |

The rewrite's finish line for an area: **every baseline in it flips to `pass`.**
