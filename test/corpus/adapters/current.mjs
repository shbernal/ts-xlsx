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
};
