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
  roundtripFixtureValidationXml,
  readFixtureReport,
  roundtripFixture,
  inspectImageAnchors,
  readFixtureImageAnchors,
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
};
