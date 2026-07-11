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
| `inspectPackage(spec)` | Build + write a `spec`, unzip the package, and return raw OOXML-part facts (worksheet-declaration consistency, `pageMargins`, `sheetViews`, table XML, per-cell formula text, per-sheet `columnGroups`/`maxColumnIndex` from the `<col>` table, per-sheet `autoFilterRef`/`dimensionRef`, well-formedness, a `styles` fact recording whether a theme part backs any theme-color font reference, `worksheetRels`, per-`sheetEntries` `state` (workbook-declaration visibility: `visible`/`hidden`/`veryHidden`, `null` when absent = visible), a per-sheet `sheetFormat` fact (`defaultRowHeight`/`defaultColWidth`/`customHeight` from `<sheetFormatPr>`; a sheet spec accepts `properties` for these), a `contentTypeDefaults` list of `<Default>` extension/media-type pairs, a `vml` fact recording each comment `<v:textbox>` style string + an `allTextboxesFitToText` flag (a note box carrying `mso-fit-shape-to-text:t` grows to its text instead of clipping), a per-sheet `rows` map of row-element outline attributes (`outlineLevel`/`hidden`/`collapsed` — a collapsed group must carry the collapsed flag on its summary row, not on the hidden detail rows), a per-sheet `hasBackgroundPicture` flag, and a `packageParts` fact recording comment/VML/table parts + worksheet-rel-id uniqueness) — for asserting on what is actually serialized. |
| `tryWriteWorkbook(spec)` | Build + attempt to write a `spec`; return `{ok, error, survivingCells, …}` — for asserting pathological input neither throws nor drops sibling cells. |
| `roundtripRangeValidation({range, type, formula})` | Apply one data validation over a multi-cell range (e.g. a whole column) and write; return `{writeOk, writeError, sqrefs, count, reloadOk}` — for asserting a whole-column dropdown emits a single range-scoped `dataValidation`, does not throw on write, and reloads. |
| `appendRowShapes()` | Append rows as a dense array, a sparse 1-based array, an object, a typed array, and a mixed batch; return `{rows: {<n>: {A, B, C, E}}}` after reload — for asserting every row shape lands its data (not only object-keyed rows) and types survive. |
| `appendRowsAfterReload(initial, append)` | Author initial rows, write+reload, append rows past the last populated row, write+reload; return `{loadedRowCount, finalRowCount, rows}` — for asserting appended rows land at contiguous indices with no gap/overwrite and the originals survive. |
| `interleavedImageAnchors(placement)` | Place two distinct images in an interleaved order (default `'BAA'`) and resolve each anchor's referenced media (embed rId → drawing rel → media); return `{placed, resolvedLetter, distinctMediaCount, distinctRelTargets}` — for asserting every anchor renders the image it was placed with and a reused image maps to one stable relationship. |
| `authorConditionalFormatting(cf)` | Author a conditional-formatting rule (e.g. a `dataBar` with `cfvo` anchors, color, gradient) and return `{writeOk, xml:{hasDataBar, cfvoCount, hasColor, wellFormed}, reload:{type, color, gradient, cfvo}}` — for asserting a rule emits well-formed XML and round-trips. |
| `roundtripFixtureImageRotation(rel)` | Read a fixture, load-rewrite it, and report the image drawing-anchor rotation (`rot` on `<a:xfrm>`, 1/60000-deg) before/after → `{sourceRot, rewrittenRot}` — for asserting an image rotation survives a round-trip. |
| `imageExtensionRoundtrip(extension)` | Add an image whose `extension` may carry a leading dot; return `{mediaParts, doubledSeparator, reloadedImageCount}` — for asserting a `.png` extension does not produce an `image1..png` media part the reader fails to discover. |
| `roundtripFixtureRowBreaks(rel)` | Report a fixture's manual row page breaks as read, then after a load-rewrite → `{sourceBreaks, loadedBreaks, rewrittenBreaks}` — for asserting rowBreaks are read and preserved, not dropped. |
| `authorDateValidation(operand)` | Author a date-type validation with a Date (or `'invalid'`) bound; return `{formula1, hasNaN}` — for asserting the formula is a real serial and never the token `NaN`. |
| `sharedBaseStyleFontMutation()` | Assign one base style object to two cells, mutate one cell's font; return `{a1Color, a2Color, bled}` — for asserting copy-on-write style isolation (no aliasing bleed). |
| `spliceShiftsRefs()` | Insert a row above a table + anchored image; return `{tableRef, imageFromRow, dupColumnNamesRejected}` — for asserting a splice re-pins table ranges and image anchors and that duplicate table column names are rejected. |
| `mergeCleanReport({anchor, range, value})` | Author a horizontal merge with an anchor value+alignment; return `{mergeCount, populatedCoveredCells, anchorValue, anchorAlignment}` — for asserting a clean merge (covered cells not populated) opens without a repair prompt. |
| `tableColumnStyleReport(numFmt)` | Author a table with a per-column numFmt style; return `{writeOk, reloadOk, styledBody, unstyledBody}` — for asserting the column style is merged into body cells without corrupting the package. |
| `insertRowThenStyle(styleMode)` | Insert a row with a style-inheritance mode then assign numFmt/font to an inserted cell; return `{error, numFmt}` — for asserting inherited-style cells stay mutable (no "object is not extensible" throw). |
| `mergeSlaveWrite({range, slave, value})` | Merge a range and write a value to a non-master cell; return `{cellsWithValue, merges, masterValue, slaveValue}` — for asserting the slave write resolves to the master with no stray slave value. |
| `nonFiniteCellReport(kind)` | Assign a non-finite number (`'NaN'`/`'Infinity'`/`'-Infinity'`); return `{writeOk, token, hasNonFiniteToken, reloadOk}` — for asserting the writer never emits a bare non-finite token into a numeric cell. |
| `formulaFalsyResultReport()` | Author formula cells with falsy results (0/false/"") + a truthy control; return `{zero, boolFalse, emptyString, truthy}` each `{isFormula, hasResult, result}` — for asserting a round-trip preserves a formula's result regardless of truthiness. |
| `streamWriteDvHyperlinkOrder()` | Stream-write a sheet with a hyperlink + a data validation; return `{posDataValidations, posHyperlinks, dataValidationsBeforeHyperlinks, reloadOk}` — for asserting the CT_Worksheet child order (dataValidations before hyperlinks). |
| `autoFilterDefinedNameReport(ref)` | Set an autofilter and report `{autoFilterRef, hasFilterDatabase, filterDatabaseHidden, filterDatabaseFormula}` — for asserting the hidden `_xlnm._FilterDatabase` defined name is emitted so LibreOffice recognizes the filter. |
| `enumerateImagesAfterRoundtrip()` | Author two-cell + one-cell anchored images, round-trip; return `{count, images:[{tl, hasMedia}], mediaCount}` — for asserting `getImages()` enumerates every image across anchor variants. |
| `csvWriteSheetSelection(sheetName)` | Write a chosen worksheet of a multi-sheet workbook to CSV; return `{ok, error, text, rowCount}` — for asserting a bad sheet selector does not silently yield empty output. |
| `unstyledCellFontReport()` | Write a plain unstyled value, round-trip; return `{hasFont, fontName, fontSize}` — for asserting an unstyled cell resolves to the workbook default font. |
| `loadMutateCellBorder()` | Author style-sharing cells, mutate one cell's border, round-trip; return `{a1, a2, a3, bled}` — for asserting a per-cell border mutation does not bleed to siblings. |
| `hiddenEmptyRowReport()` | Set hidden/height/outline on blank rows, round-trip; return `{row3Hidden, row4Hidden, row4Height, row5Hidden, row5Outline}` — for asserting a content-less hidden row keeps its flag. |
| `streamAddRowAfterCommit()` | Commit a streaming sheet then add a row; return `{rejected, legibleRejection, internalCrash, reloadOk}` — for asserting a post-commit add is rejected legibly, not an internal crash. |
| `tableCellEditRoundtrip()` | Edit a cell inside a table's range, round-trip; return `{writeOk, reloadOk, tablePresent, editedValue, relUnique, hasTablePart}` — for asserting the edit stays valid. |
| `columnBorderScopedReport()` | Author a column with a border style + plain columns, round-trip; return `{a1, b1, c1}` right-border presence — for asserting the border is scoped to its column. |
| `tableColumnNameControlChars(name)` | Author a table column name with CR/LF; return `{writeOk, rawControlChars, firstColumnTag, reloadOk}` — for asserting the name is XML-escaped, not raw control chars. |
| `internalHyperlinkReport(target)` | Author a `#`-prefixed internal hyperlink; return `{hasLocation, location, hasExternalRel, hasRid, reloadOk}` — for asserting an internal link is a location, not an external relationship. |
| `sharedStringsOption(use)` | Write with a `useSharedStrings` option; return `{hasSharedStringsPart, cellType, isSharedRef, isInline}` — for asserting the option controls string storage. |
| `dvFormulaLeadingEquals(formula)` | Author a DV formula with a leading `=`; return `{formula1, hasLeadingEquals}` — for asserting the writer strips the leading `=`. |
| `duplicateRowReport()` | Duplicate a row (default args) then merge on it; return `{dupError, mergeError, rowCount, row1, row2, merges}` — for asserting a faithful copy and a successful merge. |
| `streamCommitBadDestination()` | Stream-commit to an unwritable destination; return `{outcome, rejected, carriesIoError, error}` — for asserting a failed sink rejects rather than hanging. |
| `roundtripFixtureWriteReport(rel)` | Load a fixture and try to write it back; return `{loadOk, writeOk, writeError, sheetNames}` — for asserting a foreign construct round-trips without the writer crashing. |
| `mergeMasterBorderReport()` | Border a cell, make it a merge master, round-trip; return `{hasTopBorder, hasBottomBorder, topStyle, bottomStyle, numFmt, fontBold, merges}` — for asserting a merge keeps the master's border + style. |
| `streamingStyleCopyReport()` | Stream-read a styled workbook and copy cells + styles to the streaming writer; return `{copyError, loadOk, fontBold, fontColor, numFmt, hasFill}` — for asserting streaming style copy preserves per-cell styles. |
| `streamingSharedStringsRead(rowCount, concurrency)` | Read a shared-strings workbook once and then concurrently; return `{singleComplete, singleLength, concurrentAllComplete, concurrentLengths}` — for asserting the streaming reader never skips the shared-strings part or races. |
| `numFmtObjectCorruptionReport()` | Set a cell numFmt to a structured object (plus a valid-string control with other facets), write; return `{stylesHasObjectObject, objectNumFmtReload, controlNumFmtReload}` — for asserting a non-string numFmt never serializes to formatCode="[object Object]". |
| `csvNonAsciiEncodingReport(text)` | Write a CSV with non-ASCII text; return `{hasBom, bytesDecodeToText}` — for asserting the CSV carries a UTF-8 BOM so spreadsheet apps detect the encoding. |
| `streamingSharedFormulaReport(rows)` | Stream-write a master + shared-formula slave block, reload; return `{masterHasFormula, slaveResolved, slaveValue}` — for asserting streamed shared-formula slaves aren't dropped to empty. |
| `equivalentColumnCollapseReport()` | Define adjacent equivalent columns, write; return `{writeOk, writeError, colSpanCount, reloadOk}` — for asserting equivalent-column collapse doesn't crash and coalesces into shared `<col>` spans. |
| `formulaDateResultReport(serial)` | A formula whose cached result is a date serial under a date format, reload; return `{isValidDate, resultIso, keepsFormula}` — for asserting a numeric formula result under a date format reads as a valid Date. |
| `workbookProtectionRoundtrip()` | Load a workbook declaring `<workbookProtection lockStructure="1">`, write it back; return `{sourceHadProtection, rewrittenHasProtection, rewrittenLocksStructure}` — for asserting workbook structure protection survives a round-trip. |
| `multiSheetTableReport(sheetCount)` | Build several sheets each with a table + data validation, write, reload; return `{writeOk, writeError, tableCount, idsUnique, reloadOk, firstSheetDvSurvives}` — for asserting looped multi-sheet tables produce a valid package (unique table ids) with surviving validations. |
| `conditionalFormattingStopIfTrue()` | Apply a stopIfTrue CF rule, write, reload; return `{xmlHasStopIfTrue, reloadStopIfTrue}` — for asserting the stopIfTrue flag is serialized on the cfRule and round-trips. |
| `authorConditionalFormatting(cf)` also returns `xml.{blockCount, sqrefs, ruleCount}` so a case can assert a multi-area sqref survives as one rule. |
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
| `csvWriteEncodingReport({encoding, text})` | Two CSV write-side character facets → `{emojiRoundtrips, requestedEncoding, decodesAsRequested, decodesAsUtf8}` — for asserting multibyte (emoji/CJK) fidelity survives a UTF-8 round-trip and that a requested non-UTF-8 output encoding is actually applied rather than silently emitted as UTF-8. |
| `streamWriteSheet({useSharedStrings, ops, read})` | Drive the streaming workbook writer through row ops (`{op:'addRow'\|'addRows', value}`), commit, read the package back → `{ok, error, cells, rowCount}` — for asserting streaming-only behavior (batch add, richText shared-string handling). A throwing op is captured as `{ok:false, error}`. |
| `streamingFullCalcOnLoadReport()` | Request `fullCalcOnLoad` on the streaming writer and report whether it reaches the output vs the in-memory writer → `{streamSetThrew, streamHasFlag, streamDefaultHasFlag, memoryHasFlag}` — for asserting recalc-on-load parity between the two writers. |
| `dataTableFormulaRoundtrip()` | Round-trip a What-If-Analysis data-table formula (`<f t="dataTable">`, injected into a package) → `{readShareType, readRef, readResult, reloadOk, outHasDataTable}` — for asserting the reader recognizes the data-table kind and the writer re-emits it on a read-modify-write. |
| `removeCellNoteReport()` | Attach a note, attempt to clear it, and report what the package retains → `{commentPartPresent, vmlPartPresent, readNoteAfter, neighborNoteIntact, cleanHasCommentPart}` — for asserting a removed note leaves no comment/VML artifact while other notes survive. |
| `crossRealmArrayRow()` | Add a row from an array built in a foreign realm (a Node `vm` context) → `{isArrayCrossRealm, a, b, c}` — for asserting row-input detection is structural, not realm-bound identity. |
| `fillArgbHashPrefixReport()` | Write a solid fill ARGB clean and `#`-prefixed → `{validRgb, validReRead, hashRgb, hashReRead}` — for asserting a valid ARGB serializes as 8 hex digits and a `#`-prefixed value is not emitted as a malformed rgb that renders black. |
| `tableStyleThemeReport()` | Write a table with a real theme, `'None'`, and `null` → `{real, none, nullTheme}` each `{ok, name, hasStripes}` — for asserting `'None'` produces an unstyled table (no name), not `name="None"`. |
| `fontExplicitFalseBoldReport()` | Read a bold flag serialized as a bare / explicit-true / explicit-false tag (injected into styles) → `{bareTag, valOne, valZero}` — for asserting `<b val="0"/>` reads as false, not true. |
| `outlinePropertiesRoundtrip()` | Set worksheet outline summary-position properties, write, and report serialization + round-trip → `{outlinePrEmitted, reReadSummaryBelow, reReadSummaryRight}` — for locking that `summaryBelow`/`summaryRight` reach `<outlinePr>` and read back. |
| `rowInsertPreservesNoteAndOutline()` | Insert a row above a noted, outlined row → `{dataShifted, noteFollowsRow, outlineFollowsRow}` — for asserting a cell note and an outline level track their logical row through an insert rather than being dropped or left at the old index. |
| `frozenTopRowRoundtrip()` | Freeze the first row, write, report pane + round-trip → `{paneEmitted, reReadState, reReadYSplit, reReadXSplit}` — for locking a frozen-header sheet view survives. |
| `tabColorRoundtrip()` | Set a worksheet tab color (ARGB, alpha first) alongside an uncolored sheet → `{tabColorArgbWritten, reReadArgb, uncoloredHasTab}` — for locking tab-color ARGB round-trip. |
| `cellAnchoredImagePositionReport()` | Anchor images to single cells with interleaved `addRow` calls → `{anchorCount, froms}` — for locking a cell-range image anchor resolves to its exact cell with no off-by-one drift. |
| `wideTableColumnReadReport()` | Write a wide (5-column) table, read it back → `{colCount, colNames}` — for locking a loaded table exposes every column, not a fixed cap. |
| `csvReadMapReport()` | Read a CSV with the default and an identity map → `{default, identity}` each `{a, aType, b, bType}` — for asserting the map option governs per-value coercion (identity preserves raw strings). |
| `addReservedSheetNameReport()` | Add a `"History"` sheet and an invalid name → `{addThrew, addError, roundtripName, invalidRejected}` — for asserting an Excel-UI-reserved name is a valid document-model name while truly-invalid names are rejected. |
| `streamAutoFilterProtectionOrder()` | Stream-write a worksheet with autoFilter + sheet protection → `{protectThrew, sheetProtectionBeforeAutoFilter, reloadOk}` — for asserting CT_Worksheet order (`sheetProtection` before `autoFilter`). |
| `mergedCellDisplayTextReport()` | Read master + merged-child cell text → `{masterText, childText, childThrew}` — for asserting a merged child's text mirrors the master and does not throw. |
| `outOfOrderColumnsReport()` | Reverse the `<col>` tags in written sheet XML, reload → `{w1, w2, w3, hidden2}` — for asserting column properties stay bound to the right index regardless of `<col>` document order. |
| `rowColumnOutlineLevelRoundtrip()` | Set outline levels on a row and a column, round-trip → `{rowOutline, colOutline}` — for asserting outline (grouping) levels survive on both axes. |
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
| `printAreaRoundtrip(printArea)` | Author a worksheet with a `printArea` string, write it, reload, and report what the reader recovers → `{writtenDefinedName, reReadPrintArea, reloadOk}` — for asserting a whole-column/whole-row print area survives the round-trip instead of decoding the missing row bound as `NaN` (a corrupt `ANaN:DNaN`). |
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
| `alignmentFalseBooleanReport()` | Read a cell whose alignment carries only an explicit-false boolean (`wrapText="0"` / `shrinkToFit="0"`) → `{wrapTextZero, shrinkZero}` (reloaded alignment or null) — for asserting an explicit-false attribute yields no alignment, not an `{wrapText:false}` object (the raw `"0"` string is truthy in JS). |
| `worksheetNameLookupReport()` | Add a sheet named `Sheet`, probe a case-variant → `{foundExact, foundVariant, addVariantThrew}` — for asserting `getWorksheet` and `addWorksheet` agree on name identity (a name reported absent by lookup must be addable). |
| `internalHyperlinkSerializationReport()` | Serialize an in-workbook `#Sheet2!A1` hyperlink → `{hasWorksheetRels, hyperlinkHasRid, hyperlinkLocation, relTargetMode, reReadHyperlink}` — for asserting an internal link is written location-only with no external relationship, so consumers don't double the target. |
| `nonCanonicalCommentsPartReport()` | Read a package whose comments part lives at a non-canonical path (`xl/sheet1_comments.xml`) referenced only by the rels → `{ok, error, note}` — for asserting the reader locates parts by relationship type, not filename glob. |
| `streamWriterPipeContract()` | Pipe the streaming writer's own output stream into a `PassThrough` → `{pipeReturnsDestination, bytes, valid}` — for asserting the writer's stream honors Node's pipe contract (pipe returns the destination) while still delivering the full payload. |
| `loadFixtureTableColumns(rel, tableName)` | Load a fixture whose table declares a calculated column (`<calculatedColumnFormula>`) → `{loaded, error, columnCount, columnNames}` — for asserting the table reader consumes the nested formula element and keeps every column instead of truncating and crashing. |
| `unfreezeViewRoundtrip()` | Author a frozen view, unfreeze it (replace with a normal view), write again → `{frozenHasPane, normalHasPane, reloadedState, reloadedHasSplit}` — for asserting that removing a sheet's frozen state emits a clean normal view (no leftover `<pane>`) that opens without repair and round-trips as state `normal`. |
| `imageAnchorRowAppendReport()` | Anchor a floating image over a cell range, then append rows (both call orders) → `{imageFirst, rowsFirst}` each `{rowCount, firstDataCell}` — for asserting anchoring an image does not advance the row-append cursor, so a later `addRows` fills from the top and the layout is identical regardless of add order. |
| `cellColRowTypes(ref?)` | Report a populated cell's `col`/`row` accessors and their runtime types → `{col, row, colType, rowType}` — for locking that cell position indices are 1-based numbers at runtime (legacy types declared them as string). |
| `richTextRoundtripReport(runs)` | Round-trip a caller-supplied rich-text run array and report serialization + read-back → `{emptyTextRunInXml, runCount, runs:[{text, bold, italic, underline}]}` — for asserting empty-text runs are dropped on write (Excel rejects an empty `<t>` run) and that a leading/interior run's formatting survives. |
| `fontExplicitOffFlagsReport()` | Read explicit-off font toggles injected into styles (`<i val="0"/>`, `<strike val="0"/>`, `<u val="none"/>`) → `{italic, strike, underline}` — the companion to `fontExplicitFalseBoldReport()` for the remaining flags. |
| `trailingMergedRowIterationReport()` | Round-trip a worksheet whose merged range extends into a trailing empty final row and enumerate its cells → `{rowCount, visited, a3:{isMerged, master, visited}}` — for asserting the leading cell of the trailing merged row is visited and resolves to its master. |
| `addImageToLoadedWorksheetReport(range?)` | Load a workbook from bytes, add an image to the loaded worksheet, re-serialize → `{hasMedia, hasDrawing, reloadImageCount}` — for locking that `addImage` on a loaded (not freshly-created) worksheet persists the image. |

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
