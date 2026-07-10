// Adapter: binds the corpus's implementation-blind contract vocabulary to the
// *current* (legacy) code under lib/. This is the only file in the corpus that is
// allowed to know how today's implementation is shaped. When the Phase 3 rewrite
// lands, a sibling adapter (e.g. `rewrite.mjs`) binds the same vocabulary to the
// new code and every case runs unchanged against it.
//
// The contract is deliberately tiny and grows only as cases demand capabilities.
// Each method returns plain, JSON-serializable data so cases can assert on the
// observable shape without reaching into any implementation's internals.

import {createRequire} from 'node:module';
import {
  roundtripWorkbook,
  inspectPackage,
  tryWriteWorkbook,
  mutateWorksheet,
  readFixtureValidations,
  readFixtureValidationRules,
  readFixtureHyperlinks,
  roundtripFixtureValidationXml,
  readFixtureReport,
  roundtripFixture,
  inspectImageAnchors,
  interleavedImageAnchors,
  appendRowsAfterReload,
  readFixtureImageAnchors,
  csvRead,
  csvWrite,
  csvWriteEncodingReport,
  streamWriteSheet,
  streamingFullCalcOnLoadReport,
  dataTableFormulaRoundtrip,
  removeCellNoteReport,
  crossRealmArrayRow,
  fillArgbHashPrefixReport,
  tableStyleThemeReport,
  fontExplicitFalseBoldReport,
  outlinePropertiesRoundtrip,
  rowInsertPreservesNoteAndOutline,
  sharedFormulaRoundtripAndSplice,
  streamWriteCfHyperlinkOrder,
  roundtripFormulas,
  roundtripTableAppend,
  readFixtureDefinedNames,
  readFixtureCellStyles,
  roundtripFixtureTableXml,
  readFixtureTable,
  streamReadFixture,
  streamVsEagerSheetNames,
  streamVsEagerRowNumbers,
  streamVsEagerRowHidden,
  streamReadReport,
  readFixtureCells,
  roundtripFixtureCellXml,
  roundtripFixturePackageParts,
  roundtripFixtureStyleFacts,
  roundtripFixtureConditionalFormatting,
  authorConditionalFormatting,
  roundtripFixtureImageRotation,
  imageExtensionRoundtrip,
  roundtripFixtureRowBreaks,
  authorDateValidation,
  sharedBaseStyleFontMutation,
  spliceShiftsRefs,
  mergeCleanReport,
  tableColumnStyleReport,
  insertRowThenStyle,
  mergeSlaveWrite,
  nonFiniteCellReport,
  formulaFalsyResultReport,
  streamWriteDvHyperlinkOrder,
  autoFilterDefinedNameReport,
  enumerateImagesAfterRoundtrip,
  csvWriteSheetSelection,
  unstyledCellFontReport,
  loadMutateCellBorder,
  hiddenEmptyRowReport,
  streamAddRowAfterCommit,
  tableCellEditRoundtrip,
  columnBorderScopedReport,
  tableColumnNameControlChars,
  internalHyperlinkReport,
  sharedStringsOption,
  dvFormulaLeadingEquals,
  duplicateRowReport,
  streamCommitBadDestination,
  roundtripFixtureWriteReport,
  mergeMasterBorderReport,
  streamingStyleCopyReport,
  streamingSharedStringsRead,
  numFmtObjectCorruptionReport,
  csvNonAsciiEncodingReport,
  streamingSharedFormulaReport,
  equivalentColumnCollapseReport,
  formulaDateResultReport,
  workbookProtectionRoundtrip,
  multiSheetTableReport,
  conditionalFormattingStopIfTrue,
  roundtripFixtureColorFidelity,
  roundtripFixturePrintAreas,
  writePrintAreaDefinedName,
  printAreaRoundtrip,
  authorListValidations,
  roundtripRangeValidation,
  appendRowShapes,
  authorCellProtection,
  streamCommitReport,
  streamWriterImageSupport,
  streamWritePackageReport,
  streamReadSpec,
  loadMutateCellStyle,
  copyWorksheetModel,
  styleDedupReport,
  readRowCellPresence,
  streamVsEagerRowValues,
  roundtripSpecTableFacts,
  loadMutateCellFont,
} from './workbook-io.mjs';

const require = createRequire(import.meta.url);
const colCache = require('../../../lib/utils/col-cache.js');

export default {
  name: 'current',

  // Decode a single cell/row/column reference (e.g. 'B2', '$1', '$A') into
  // { col, row } where an absent axis is `undefined`, plus the serialized forms.
  decodeAddress(reference) {
    return colCache.decodeAddress(reference);
  },

  // Decode a range reference (e.g. 'A1:B2', '$1:$1', '$A:$A') into its corners
  // and serialized dimensions.
  decodeRange(reference) {
    return colCache.decodeEx(reference);
  },

  // Build a fresh single worksheet, assign a font to each `apply` cell, then read
  // back the resolved font of each `read` cell. Returns { <address>: font } as
  // plain JSON. Lets a case verify that per-cell styling stays local to the cell
  // it was set on and does not bleed across untouched cells of the sheet.
  probeCellFonts({apply = [], read = []}) {
    const ExcelJS = require('../../../lib/exceljs.nodejs.js');
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('sheet');
    for (const {cell, font} of apply) {
      sheet.getCell(cell).font = font;
    }
    const fonts = {};
    for (const address of read) {
      fonts[address] = sheet.getCell(address).font ?? null;
    }
    return JSON.parse(JSON.stringify(fonts));
  },

  // Build a workbook from a declarative spec, write it to a buffer, read it back
  // with the same implementation, and return a normalized JSON model — for
  // asserting that content survives a write→read round-trip. See workbook-io.mjs
  // for the spec shape.
  roundtripWorkbook,

  // Build + write a workbook from a spec, unzip the produced package, and report
  // raw OOXML-part facts (worksheet-declaration consistency, pageMargins, sheet
  // views, table XML, per-cell formula text) — for asserting on what is actually
  // serialized, independent of what the reader defaults back in.
  inspectPackage,

  // Build + attempt to write a workbook from a spec; return { ok, error, … }
  // including which cells survived — for asserting that pathological input (an
  // invalid date, an empty workbook) neither throws nor silently drops siblings.
  tryWriteWorkbook,

  // Build a fresh worksheet, apply a sequence of structural mutations (row/column
  // splices), and report { rowCount, columnCount, cells, error } — for asserting
  // that in-memory model edits behave predictably. See workbook-io.mjs.
  mutateWorksheet,

  // Read a corpus fixture `.xlsx` and report the per-cell data validations the reader
  // exposes — for asserting a real file's validations are read back on every cell.
  readFixtureValidations,

  // Read a fixture and report the DISTINCT validation rules each sheet declares, read from the
  // worksheet model (so a validation over an empty range is still seen) and de-duplicated with a
  // per-rule coverage count — for asserting a reference-based list source (defined name,
  // cross-sheet range) surfaces as its formula text, not "[object Object]". See workbook-io.mjs.
  readFixtureValidationRules,

  // Read a fixture and report the first sheet's hyperlink cells → { <addr>: {hyperlink, text} } —
  // for asserting a real file's hyperlink is reconstructed in full, including a `#fragment` carried
  // in the hyperlink element's location attribute separate from the relationship target (which must
  // not be dropped). See workbook-io.mjs.
  readFixtureHyperlinks,

  // Read a fixture, write it back, and report data-validation facts of the re-serialized
  // package (standard + extended x14 forms) — for asserting validations survive a
  // read→write round-trip. See workbook-io.mjs.
  roundtripFixtureValidationXml,

  // Read a fixture and report { ok, error, sheetNames } — for asserting the reader
  // tolerates foreign-generator files (prefixed OOXML roots, BOMs, non-ASCII sheet
  // names, unusual zip ordering) without crashing or mis-reading. See workbook-io.mjs.
  readFixtureReport,

  // Read a fixture, write it back unchanged, read it again, and report whether sheet
  // names, column widths, and per-cell styles survive the no-op round-trip — the
  // format-preserving "open a styled template and re-save" path. See workbook-io.mjs.
  roundtripFixture,

  // Build a workbook with images at given ranges, write it, and report the serialized
  // drawing-anchor geometry (from/to/ext/editAs) — for asserting fractional and string
  // anchors map to correct OOXML offsets against real column/row size. See workbook-io.mjs.
  inspectImageAnchors,

  // Place two DISTINCT images in an interleaved order and resolve, per anchor in placement order,
  // which media part it references (embed rId → drawing rel → media) → { placed, resolvedLetter,
  // distinctMediaCount, distinctRelTargets } — for asserting every anchor renders the image it was
  // placed with and a reused image maps to one stable relationship. See workbook-io.mjs.
  interleavedImageAnchors,

  // Author initial rows, write+reload, append rows past the last populated row, write+reload again
  // → { loadedRowCount, finalRowCount, rows } — for asserting appended rows land at contiguous
  // indices with no gap/overwrite and the originals survive. See workbook-io.mjs.
  appendRowsAfterReload,

  // Read a fixture and report each image's normalized anchor range — for asserting a file
  // whose anchors were authored as (string) cell ranges reads without crashing and yields
  // an object range with integer cell coordinates. See workbook-io.mjs.
  readFixtureImageAnchors,

  // Parse a CSV string with given options → { ok, error, rows } of typed cell values — for
  // asserting delimiter handling, value coercion, and header-mode behavior. See workbook-io.mjs.
  csvRead,

  // Write a declarative row spec to CSV with given options → { ok, error, text } — for
  // asserting field delimiter and date formatting on genuinely-typed cells. See workbook-io.mjs.
  csvWrite,

  // Two facets of CSV write-side character handling → { emojiRoundtrips, requestedEncoding,
  // decodesAsRequested, decodesAsUtf8 } — for asserting multibyte fidelity survives a UTF-8
  // round-trip and that a requested non-UTF-8 output encoding is actually applied, not silently
  // ignored. See workbook-io.mjs.
  csvWriteEncodingReport,

  // Drive the streaming workbook writer through row ops (addRow/addRows), commit, read back
  // → { ok, error, cells, rowCount } — for asserting streaming-only behavior like batch add
  // and richText shared-string handling. See workbook-io.mjs.
  streamWriteSheet,

  // Request fullCalcOnLoad on the streaming writer and report whether it reaches the output vs the
  // in-memory writer → { streamSetThrew, streamHasFlag, streamDefaultHasFlag, memoryHasFlag } — for
  // asserting recalc-on-load parity between the two writers. See workbook-io.mjs.
  streamingFullCalcOnLoadReport,

  // Round-trip a What-If-Analysis data-table formula (`<f t="dataTable">`) → { readShareType,
  // readRef, readResult, reloadOk, outHasDataTable } — for asserting the reader recognizes a
  // data-table formula and the writer preserves its kind on a read-modify-write. See workbook-io.mjs.
  dataTableFormulaRoundtrip,

  // Attach a note, attempt to clear it, and report what the package retains → { commentPartPresent,
  // vmlPartPresent, readNoteAfter, neighborNoteIntact, cleanHasCommentPart } — for asserting a
  // removed note leaves no comment/VML artifact. See workbook-io.mjs.
  removeCellNoteReport,

  // Add a row from an array built in a foreign realm (a Node vm context) → { isArrayCrossRealm, a,
  // b, c } — for asserting row-input detection is structural, not realm-bound identity. See
  // workbook-io.mjs.
  crossRealmArrayRow,

  // Write a solid fill ARGB both clean and '#'-prefixed → { validRgb, validReRead, hashRgb,
  // hashReRead } — for asserting a valid ARGB serializes as 8 hex digits and a '#'-prefixed value is
  // not passed through as a malformed rgb that renders black. See workbook-io.mjs.
  fillArgbHashPrefixReport,

  // Write a table with a real theme, 'None', and null → { real, none, nullTheme } each { ok, name,
  // hasStripes } — for asserting 'None' produces an unstyled table (no name), not name="None". See
  // workbook-io.mjs.
  tableStyleThemeReport,

  // Read a bold flag serialized as a bare / explicit-true / explicit-false tag → { bareTag, valOne,
  // valZero } — for asserting an explicit-false <b val="0"/> reads as false, not true. See
  // workbook-io.mjs.
  fontExplicitFalseBoldReport,

  // Set worksheet outline summary-position properties, write, and report serialization + round-trip
  // → { outlinePrEmitted, reReadSummaryBelow, reReadSummaryRight } — for locking the outlinePr write
  // path is faithful. See workbook-io.mjs.
  outlinePropertiesRoundtrip,

  // Insert a row above a noted, outlined row and report what followed → { dataShifted,
  // noteFollowsRow, outlineFollowsRow } — for asserting a cell note and an outline level track their
  // logical row through an insert. See workbook-io.mjs.
  rowInsertPreservesNoteAndOutline,

  // Author a shared-formula master/slave group, read-then-rewrite it, and splice a column in →
  // { roundtripOk, preservedFormulas, spliceOk, spliceError } — for asserting a lossless
  // shared-formula round-trip and surfacing the splice "master must exist above/left of clone"
  // throw. See workbook-io.mjs.
  sharedFormulaRoundtripAndSplice,

  // Stream-write a sheet carrying both a conditional-formatting rule and a hyperlink → report the
  // relative order of <conditionalFormatting> vs <hyperlinks> and reload success. OOXML requires
  // conditionalFormatting first; the streaming writer emits them reversed. See workbook-io.mjs.
  streamWriteCfHyperlinkOrder,

  // Build + round-trip formula cells → per-cell { formula, sharedFormula, result } — for
  // asserting a shared-formula clone reads back a translated concrete formula. See workbook-io.mjs.
  roundtripFormulas,

  // Build a table, round-trip it, and try appending rows to the reloaded table →
  // { hasTable, loadedRowCount, addError, committed, finalRowCount } — for asserting a
  // table rehydrated from a file is mutable, not throwing on append. See workbook-io.mjs.
  roundtripTableAppend,

  // Read a fixture and report the workbook-level defined names the reader exposes →
  // { names, count, modelCount } — for asserting a full-row/full-column-span named range, or
  // same-named names scoped to different sheets, are read back rather than silently dropped.
  readFixtureDefinedNames,

  // Read a fixture and report specific cells' resolved fill + font color, keyed <sheet>!<addr>
  // → { fill, fontColor } — for asserting a real file's cell colors (solid-fill foreground,
  // theme+tint, a separate font color) are read faithfully and not conflated. See workbook-io.mjs.
  readFixtureCellStyles,

  // Read a table-bearing fixture, write it back unchanged, and report each table's raw-XML facts
  // before/after → { tables:[{name, source, rewritten}] } — for asserting a no-op round-trip does
  // not corrupt the table part (inject an autoFilter, flip the header row, spuriously set
  // totalsRowShown, or emit an empty filterColumn). See workbook-io.mjs.
  roundtripFixtureTableXml,

  // Read a fixture and report a named table's rehydration → { found, columns, rowCount } — for
  // asserting a table loaded from a real file exposes its columns AND its data rows. See workbook-io.mjs.
  readFixtureTable,

  // Read a fixture's first sheet through the streaming reader and report requested cells'
  // { type, value } — for asserting streaming read applies date/number formats like the full
  // read (a date-formatted cell is a Date, not a raw serial). See workbook-io.mjs.
  streamReadFixture,

  // Read a fixture eagerly and via the streaming reader and report the sheet names each path
  // surfaces → { eager, streaming } — for asserting streaming exposes the real declared names,
  // not generic positional placeholders. See workbook-io.mjs.
  streamVsEagerSheetNames,

  // Read a fixture eagerly and via the streaming reader and report the first sheet's row numbers
  // each path yields → { eager, streaming } — for asserting streaming preserves true row indices
  // across interior blank rows. See workbook-io.mjs.
  streamVsEagerRowNumbers,

  // Read a fixture eagerly and via the streaming reader and report each first-sheet row's
  // { number, hidden } from both paths → { eager, streaming } — for asserting the streaming reader
  // surfaces the hidden flag (interpreting string-form "true"/"false"), agreeing with the eager
  // read. See workbook-io.mjs.
  streamVsEagerRowHidden,

  // Read a fixture through the streaming reader end-to-end → { ok, error, sheetNames, totalRows } —
  // for asserting it tolerates a package whose ZIP places a worksheet part before workbook.xml
  // rather than crashing on an unbuilt workbook model. See workbook-io.mjs.
  streamReadReport,

  // Read a fixture's first sheet with the full reader and report requested cells' { type, value }
  // → for asserting real-file cell values/types (e.g. a Strict-mode ISO-8601 date parses to the
  // right date, not a 1900-epoch serial). See workbook-io.mjs.
  readFixtureCells,

  // Read a fixture, write it back, and report requested cells' raw serialized <c> facts
  // ({t, formula, value}) plus a package-wide hasNaNToken flag — for asserting a round-trip does
  // not coerce a string-typed formula result into the invalid literal "NaN". See workbook-io.mjs.
  roundtripFixtureCellXml,

  // Read a fixture, write it back unchanged, and report package-part facts before/after →
  // { source, rewritten } — for asserting a no-op round-trip preserves unmodeled parts
  // (header/footer images + their VML, vector shapes, pivot tables + caches) rather than
  // dropping them. See workbook-io.mjs.
  roundtripFixturePackageParts,

  // Read a fixture, write it back, and report style-fidelity facts before/after → { source,
  // rewritten } — column widths, pageSetup, custom indexed-color palette, and conditional-format
  // differential number codes — for asserting a no-op round-trip preserves them (and never
  // serializes a numFmt as "[object Object]"). See workbook-io.mjs.
  roundtripFixtureStyleFacts,

  // Read a fixture, write it back, and report conditional-formatting facts before/after →
  // { source, rewritten } with each cfRule's { type, dxfId, priority } — for asserting a no-op
  // round-trip preserves a rule (even an unmodeled type like duplicateValues) rather than dropping
  // it or emitting an empty conditionalFormatting shell. See workbook-io.mjs.
  roundtripFixtureConditionalFormatting,

  // Author a conditional-formatting rule (e.g. a dataBar) and report the emitted CF XML facts +
  // reader read-back → { writeOk, xml:{hasDataBar,cfvoCount,hasColor,wellFormed}, reload:{type,color,
  // gradient,cfvo} } — for asserting a dataBar round-trips with valid XML. See workbook-io.mjs.
  authorConditionalFormatting,

  // Read a fixture and load-rewrite it, reporting the image drawing-anchor rotation before/after →
  // { sourceRot, rewrittenRot } (1/60000-deg units) — for asserting an image rotation survives a
  // load/save round-trip rather than being dropped. See workbook-io.mjs.
  roundtripFixtureImageRotation,

  // Add an image whose extension may carry a leading dot (".png") → { mediaParts, doubledSeparator,
  // reloadedImageCount } — for asserting a leading-dot extension does not produce an "image1..png"
  // media part that the reader then fails to discover. See workbook-io.mjs.
  imageExtensionRoundtrip,

  // Read a fixture's manual row page breaks, load-rewrite it → { sourceBreaks, loadedBreaks,
  // rewrittenBreaks } — for asserting rowBreaks are read and preserved, not silently dropped. See
  // workbook-io.mjs.
  roundtripFixtureRowBreaks,

  // Author a date-type data validation with a Date (or 'invalid') bound → { formula1, hasNaN } —
  // for asserting a date validation writes a real serial, never the token NaN. See workbook-io.mjs.
  authorDateValidation,

  // Assign one base style object to two cells, mutate one cell's font → { a1Color, a2Color, bled }
  // — for asserting copy-on-write style isolation (no aliasing bleed to siblings). See workbook-io.mjs.
  sharedBaseStyleFontMutation,

  // Insert a row above a table + anchored image and report the shifted table ref / image anchor row +
  // whether duplicate table column names are rejected → { tableRef, imageFromRow,
  // dupColumnNamesRejected } — for asserting a splice re-pins table/image refs. See workbook-io.mjs.
  spliceShiftsRefs,

  // Author a horizontal merge with an anchor value+alignment → { mergeCount, populatedCoveredCells,
  // anchorValue, anchorAlignment } — for asserting a clean merge (covered cells not populated) that
  // opens without a repair prompt. See workbook-io.mjs.
  mergeCleanReport,

  // Author a table with a per-column numFmt style → { writeOk, reloadOk, styledBody, unstyledBody }
  // — for asserting the column style is merged into the body cells without corrupting the package.
  // See workbook-io.mjs.
  tableColumnStyleReport,

  // Insert a row with a style-inheritance mode, then assign numFmt/font to an inserted cell →
  // { error, numFmt } — for asserting inherited-style cells stay mutable (no "object is not
  // extensible" throw). See workbook-io.mjs.
  insertRowThenStyle,

  // Merge a range and write a value to a NON-master (slave) cell → { cellsWithValue, merges,
  // masterValue, slaveValue } — for asserting the slave write resolves to the master with no stray
  // slave value. See workbook-io.mjs.
  mergeSlaveWrite,

  // Assign a non-finite number (NaN/Infinity) to a cell → { writeOk, token, hasNonFiniteToken,
  // reloadOk } — for asserting the writer never emits a bare NaN/Infinity token into a numeric
  // cell. See workbook-io.mjs.
  nonFiniteCellReport,

  // Author formula cells with falsy cached results (0/false/"") + a truthy control → { zero,
  // boolFalse, emptyString, truthy } each { isFormula, hasResult, result } — for asserting a
  // round-trip preserves a formula's result regardless of truthiness. See workbook-io.mjs.
  formulaFalsyResultReport,

  // Stream-write a sheet with a hyperlink + a data validation → report the order of
  // <dataValidations> vs <hyperlinks> (dataValidations must precede) + reload. See workbook-io.mjs.
  streamWriteDvHyperlinkOrder,

  // Set an autofilter and report the sheet autoFilter ref + whether the hidden _xlnm._FilterDatabase
  // defined name is emitted → { autoFilterRef, hasFilterDatabase, filterDatabaseHidden, formula } —
  // for asserting LibreOffice-portable autofilter output. See workbook-io.mjs.
  autoFilterDefinedNameReport,

  // Author two-cell + one-cell anchored images, round-trip, and report getImages() enumeration →
  // { count, images:[{tl, hasMedia}], mediaCount } — for asserting every image is enumerated across
  // anchor variants. See workbook-io.mjs.
  enumerateImagesAfterRoundtrip,

  // Write a chosen worksheet of a multi-sheet workbook to CSV → { ok, error, text, rowCount } — for
  // asserting a bad sheet selector does not silently yield empty output. See workbook-io.mjs.
  csvWriteSheetSelection,

  // Write a plain unstyled value and report the read-back font → { hasFont, fontName, fontSize } —
  // for asserting an unstyled cell resolves to the workbook default font. See workbook-io.mjs.
  unstyledCellFontReport,

  // Author cells sharing a style record, mutate one cell's border, round-trip → { a1, a2, a3, bled }
  // — for asserting a per-cell border mutation does not bleed to style-sharing siblings. See workbook-io.mjs.
  loadMutateCellBorder,

  // Set hidden/height/outline on blank rows, round-trip → { row3Hidden, row4Hidden, row4Height,
  // row5Hidden, row5Outline } — for asserting a blank hidden row keeps its flag. See workbook-io.mjs.
  hiddenEmptyRowReport,

  // Commit a streaming worksheet then add a row → { rejected, legibleRejection, internalCrash,
  // reloadOk } — for asserting a post-commit add is rejected legibly, not an internal crash. See workbook-io.mjs.
  streamAddRowAfterCommit,

  // Build a table, edit a cell inside its range, round-trip → { writeOk, reloadOk, tablePresent,
  // editedValue, relUnique, hasTablePart } — for asserting a table cell edit stays valid. See workbook-io.mjs.
  tableCellEditRoundtrip,

  // Author a column with a border style + plain columns, round-trip → { a1, b1, c1 } right-border
  // presence — for asserting the column border is scoped to its column. See workbook-io.mjs.
  columnBorderScopedReport,

  // Author a table column name with CR/LF → { writeOk, rawControlChars, firstColumnTag, reloadOk } —
  // for asserting the name is XML-escaped, not raw control chars. See workbook-io.mjs.
  tableColumnNameControlChars,

  // Author a '#'-prefixed internal hyperlink → { hasLocation, location, hasExternalRel, hasRid,
  // reloadOk } — for asserting an internal link is a location, not an external relationship. See workbook-io.mjs.
  internalHyperlinkReport,

  // Write with a useSharedStrings option → { hasSharedStringsPart, cellType, isSharedRef, isInline }
  // — for asserting the option controls string storage (false → inline). See workbook-io.mjs.
  sharedStringsOption,

  // Author a DV formula supplied with a leading '=' → { formula1, hasLeadingEquals } — for asserting
  // the writer strips the leading '=' from formula1. See workbook-io.mjs.
  dvFormulaLeadingEquals,

  // Duplicate a row (default args) then merge on it → { dupError, mergeError, rowCount, row1, row2,
  // merges } — for asserting a faithful copy and that merging on it succeeds. See workbook-io.mjs.
  duplicateRowReport,

  // Stream-commit to an unwritable destination → { outcome, rejected, carriesIoError, error } — for
  // asserting a failed sink rejects rather than hanging. See workbook-io.mjs.
  streamCommitBadDestination,

  // Load a fixture and try to write it back → { loadOk, writeOk, writeError, sheetNames } — for
  // asserting a foreign construct round-trips without the writer crashing. See workbook-io.mjs.
  roundtripFixtureWriteReport,

  // Border a cell, make it a merge master, round-trip → { hasTopBorder, hasBottomBorder, numFmt,
  // fontBold, merges } — for asserting a merge keeps the master's border + style. See workbook-io.mjs.
  mergeMasterBorderReport,

  // Stream-read a styled workbook and copy cells + styles to the streaming writer → { copyError,
  // loadOk, fontBold, fontColor, numFmt, hasFill } — for asserting streaming style copy preserves
  // per-cell styles. See workbook-io.mjs.
  streamingStyleCopyReport,

  // Read a shared-strings workbook once and then concurrently → { singleComplete, concurrentAllComplete,
  // concurrentLengths } — for asserting the streaming reader never skips the shared-strings part or
  // races under concurrency. See workbook-io.mjs.
  streamingSharedStringsRead,

  // Set a cell numFmt to a structured OBJECT (plus a valid-string control with other facets) → write →
  // { stylesHasObjectObject, objectNumFmtReload, controlNumFmtReload } — for asserting a non-string
  // numFmt never serializes to formatCode="[object Object]". See workbook-io.mjs.
  numFmtObjectCorruptionReport,

  // Write a CSV with non-ASCII (Hebrew) text → { hasBom, bytesDecodeToText } — for asserting the CSV
  // carries a UTF-8 BOM so spreadsheet apps detect the encoding. See workbook-io.mjs.
  csvNonAsciiEncodingReport,

  // Build via the streaming writer with a master + shared-formula slaves, reload → { masterHasFormula,
  // slaveResolved, slaveValue } — for asserting streamed shared-formula slaves aren't dropped to empty.
  streamingSharedFormulaReport,

  // Define adjacent equivalent columns and write → { writeOk, writeError, colSpanCount, reloadOk } —
  // for asserting equivalent-column collapse does not crash and coalesces into shared <col> spans.
  equivalentColumnCollapseReport,

  // A formula whose cached result is a date serial under a date format → { isValidDate, resultIso,
  // keepsFormula } — for asserting a numeric formula result under a date format reads as a valid Date.
  formulaDateResultReport,

  // Load a workbook declaring workbook-level structure protection, write it back → { sourceHadProtection,
  // rewrittenHasProtection, rewrittenLocksStructure } — for asserting workbook protection survives a
  // read→write round-trip.
  workbookProtectionRoundtrip,

  // Build several sheets each with a table + a data validation, write, reload → { writeOk, tableCount,
  // idsUnique, reloadOk, firstSheetDvSurvives } — for asserting looped multi-sheet tables produce a
  // valid package (unique table ids) with surviving validations.
  multiSheetTableReport,

  // Apply a stopIfTrue conditional-formatting rule, write, reload → { xmlHasStopIfTrue,
  // reloadStopIfTrue } — for asserting the stopIfTrue flag is serialized and round-trips.
  conditionalFormattingStopIfTrue,

  // Read a fixture, write it back, reload, and report how many styled cells' VISIBLE fill/border
  // colors changed → { checked, fillMismatches, borderMismatches, … } (ignoring a benign
  // pattern="none" the writer adds) — for asserting themed/indexed colors survive a pure
  // open-then-save. See workbook-io.mjs.
  roundtripFixtureColorFidelity,

  // Read a fixture whose sheet declares multiple print areas and report { sourceRangeCount,
  // readPrintArea, rewrittenRangeCount } — for asserting a single Print_Area defined name holding a
  // comma-separated range list is recovered and re-emitted with all ranges, not truncated to the
  // first. See workbook-io.mjs.
  roundtripFixturePrintAreas,

  // Build a workbook with a (possibly comma-separated) printArea, write it, and report the emitted
  // Print_Area defined name's ranges → { rangeCount, ranges } — for asserting authoring two print
  // areas emits both ranges in one sheet-scoped name. See workbook-io.mjs.
  writePrintAreaDefinedName,

  // Author a worksheet with a printArea string, write it, reload, and report what the reader
  // recovers → { writtenDefinedName, reReadPrintArea, reloadOk } — for asserting a whole-column /
  // whole-row print area survives a round-trip instead of coming back as a NaN-laced address. See
  // workbook-io.mjs.
  printAreaRoundtrip,

  // Author list-type data validations (inline quoted literal + cross-sheet range reference),
  // round-trip, and report per-cell read-back plus serialized <dataValidations> facts
  // → { readBack, xml:{count, wellFormed, formula1} } — for asserting both value-source forms
  // survive verbatim and emit Excel-acceptable XML. See workbook-io.mjs.
  authorListValidations,

  // Apply one data validation over a multi-cell RANGE (e.g. a whole column) and write
  // → { writeOk, writeError, sqrefs, count, reloadOk } — for asserting a whole-column dropdown
  // emits a single range-scoped dataValidation and neither throws on write nor fails to reload.
  roundtripRangeValidation,

  // Append rows as a dense array, a sparse 1-based array, an object, and a mixed batch; read back
  // → { rows: { <n>: { A, B, C, E } } } — for asserting every row shape lands its data (not just
  // the object-keyed ones) and types survive the round-trip.
  appendRowShapes,

  // Author per-cell protection + a protected sheet, round-trip, and report read-back locked
  // flags, whether cellXfs carries applyProtection, and the emitted <sheetProtection>
  // → { readBack, hasApplyProtection, sheetProtection } — for asserting an unlocked cell
  // survives and worksheet protection is emitted. See workbook-io.mjs.
  authorCellProtection,

  // Drive the streaming writer over a caller-supplied PassThrough/Duplex sink and report
  // whether the workbook-commit promise settles within a bounded time plus package validity
  // → { settled, timedOut, bytes, valid } — for asserting streaming-to-a-remote-sink commit
  // resolves and yields a complete package rather than hanging. See workbook-io.mjs.
  streamCommitReport,

  // Report the streaming writer's image-parity surface and, if supported, the streamed
  // package's media/drawing parts → { writerAddImage, sheetAddImage, error, mediaParts,
  // drawingParts } — for locking image parity with the in-memory writer once delivered.
  streamWriterImageSupport,

  // Assemble a whole package via the streaming writer, then report its zip-container integrity
  // → { partCount, emptyParts, crcValid, reloadOk, sheetNames, firstCol } — for asserting the
  // streamed output is a valid archive (no zero-byte parts, CRCs match, re-reads cleanly), not
  // merely valid XML. See workbook-io.mjs.
  streamWritePackageReport,

  // Write a spec, read it back through the streaming reader over real chunk boundaries, and pair
  // with an eager read → { streamed, eager } — for asserting multi-byte UTF-8 text survives the
  // streaming path byte-exact (no U+FFFD from a chunk split mid-character).
  streamReadSpec,

  // Author cells sharing one on-disk style index, load, mutate one cell's style, read a sibling
  // → { sibling, mutatedTo, bled, diskBled } — for asserting loaded cells get independent style
  // objects rather than aliasing the shared record so one edit corrupts the rest.
  loadMutateCellStyle,

  // Copy a worksheet via the model export/import contract and report merge survival
  // → { srcMerges, dstMerges, error } — for asserting a model-cloned sheet keeps its merged ranges.
  copyWorksheetModel,

  // Write a spec and report the style-table size + per-cell resolved style index
  // → { cellXfCount, indices } — for asserting identical cell styles dedup to one shared entry
  // while a genuinely distinct style stays separate (the OOXML shared-table expectation).
  styleDedupReport,

  // Load a written spec and report, per row, the column indices a full (includeEmpty) iteration
  // yields → { rows, columnCount } — for asserting trailing empty cells are surfaced so every row
  // aligns column-for-column with a wider header.
  readRowCellPresence,

  // Read a spec's rows via both the eager and streaming readers → { eager, streamed } — for
  // asserting the streaming reader exposes the same 1-based row-values indexing as the full load.
  streamVsEagerRowValues,

  // Write a spec's table, round-trip it, and report table facts before/after → { write, roundtrip,
  // loadOk } — for asserting a defined table's ref range and part survive a load→save cycle.
  roundtripSpecTableFacts,

  // Author cells sharing one font, load, spread-reassign one cell's font, read the sibling
  // → { edited, sibling, bled } — the font companion to loadMutateCellStyle for aliasing.
  loadMutateCellFont,
};
