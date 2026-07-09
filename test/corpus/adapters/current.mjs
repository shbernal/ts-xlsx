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
  readFixtureImageAnchors,
  csvRead,
  csvWrite,
  streamWriteSheet,
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
  roundtripFixtureColorFidelity,
  roundtripFixturePrintAreas,
  writePrintAreaDefinedName,
  authorListValidations,
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

  // Drive the streaming workbook writer through row ops (addRow/addRows), commit, read back
  // → { ok, error, cells, rowCount } — for asserting streaming-only behavior like batch add
  // and richText shared-string handling. See workbook-io.mjs.
  streamWriteSheet,

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

  // Author list-type data validations (inline quoted literal + cross-sheet range reference),
  // round-trip, and report per-cell read-back plus serialized <dataValidations> facts
  // → { readBack, xml:{count, wellFormed, formula1} } — for asserting both value-source forms
  // survive verbatim and emit Excel-acceptable XML. See workbook-io.mjs.
  authorListValidations,

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
